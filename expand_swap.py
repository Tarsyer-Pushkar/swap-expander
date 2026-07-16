#!/usr/bin/env python3
"""
expand_swap.py -- Standalone swap expander for Raspberry Pi devices
====================================================================
Independent, self-contained version of the "Expand Swap" feature from the
Trent Support Hub dashboard. Has no dependency on the dashboard/Flask app --
only needs `paramiko` and the zram-tools .deb package sitting next to it.

WHAT IT DOES
------------
  1. SSHes into the target device.
  2. Skips devices that are already configured with a 1024 MB zram swap.
  3. Otherwise:
       - Removes the old dphys-swapfile based swap.
       - Uploads and installs zram-tools (zramswap).
       - Configures zramswap for a 1024 MB (1G) swap via lz4 compression.
       - Enables + restarts the zramswap service.
  4. Verifies the result by parsing `free -m` on the device and confirming
     the reported Swap total is exactly 1024 MB (prints PASS/FAIL clearly).

REQUIRED FILES (must sit in the same folder as this script)
-------------------------------------------------------------
  - expand_swap.py                       (this script)
  - zram-tools_0.3.3.1-1_all.deb         (the package to install on the device)
  - devices.xlsx                         (your device list -- see EXCEL FORMAT)

REQUIREMENTS
------------
  pip install paramiko openpyxl

EXCEL FORMAT
------------
  An .xlsx file with a header row containing (case-insensitive, any order):
    - a "store code" column   (accepted headers: "store code", "store", "store_code")
    - a "device ip" column    (accepted headers: "device ip", "ip", "ip address", "device_ip")

  Example:
      | Store Code | Device IP        |
      |------------|------------------|
      | Z900       | 10.110.197.41    |
      | W120       | 10.110.198.12    |

USAGE
-----
  python expand_swap.py --excel devices.xlsx
  python expand_swap.py --excel devices.xlsx --user pi --password industrail
  python expand_swap.py --device 10.110.197.41              # single IP, no excel needed
  python expand_swap.py --devices-file devices.txt          # one IP per line, no excel needed

  If no --excel/--device/--devices-file is given, the script looks for
  "devices.xlsx" in its own folder automatically.

EXIT CODE
---------
  0 if every targeted device ends up verified at 1024 MB swap (or was
  already configured), non-zero otherwise.
"""

import argparse
import re
import socket
import sys
from functools import partial
from pathlib import Path

import paramiko
import openpyxl

print = partial(print, flush=True)

# -- Config ---------------------------------------------------------------
SSH_USER_DEFAULT = "pi"
SSH_PASS_DEFAULT = "industrail"
CONNECT_TIMEOUT  = 10
SWAP_SIZE_MB     = 1024

SCRIPT_DIR    = Path(__file__).resolve().parent
DEB_FILE      = SCRIPT_DIR / "zram-tools_0.3.3.1-1_all.deb"
REMOTE_DEB    = "/tmp/zram-tools_0.3.3.1-1_all.deb"
DEFAULT_EXCEL = SCRIPT_DIR / "devices.xlsx"

STORE_HEADERS = {"store code", "store", "store_code", "storecode"}
IP_HEADERS    = {"device ip", "ip", "ip address", "device_ip", "ip_address"}

# -- Counters ---------------------------------------------------------------
success = failed = skipped = 0


# -- SSH helpers --------------------------------------------------------------

def connect(ip: str, user: str, password: str) -> paramiko.SSHClient:
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(
        hostname=ip,
        username=user,
        password=password,
        timeout=CONNECT_TIMEOUT,
        banner_timeout=CONNECT_TIMEOUT,
        auth_timeout=CONNECT_TIMEOUT,
    )
    return ssh


def run(ssh: paramiko.SSHClient, cmd: str):
    """Run a command and return (exit_code, stdout, stderr)."""
    stdin, stdout, stderr = ssh.exec_command(cmd)
    try:
        rc  = stdout.channel.recv_exit_status()
        out = stdout.read().decode(errors="ignore")
        err = stderr.read().decode(errors="ignore")
        return rc, out, err
    finally:
        stdin.close()
        stdout.close()
        stderr.close()


def upload_deb(ssh: paramiko.SSHClient):
    sftp = ssh.open_sftp()
    try:
        sftp.put(str(DEB_FILE), REMOTE_DEB)
    finally:
        sftp.close()


def already_configured(ssh: paramiko.SSHClient) -> bool:
    rc, _, _ = run(
        ssh,
        'swapon --show | grep -q "/dev/zram0" && '
        'zramctl | grep -q "1G" && '
        'systemctl is-enabled zramswap >/dev/null 2>&1',
    )
    return rc == 0


def install(ssh: paramiko.SSHClient, password: str):
    """Remove dphys-swapfile, install zram-tools, configure a 1024MB swap."""
    script = """
set -e
echo '%s' | sudo -S swapoff -a || true
echo '%s' | sudo -S systemctl disable --now dphys-swapfile 2>/dev/null || true
echo '%s' | sudo -S apt purge -y dphys-swapfile || true
echo '%s' | sudo -S rm -f /var/swap /swapfile
echo '%s' | sudo -S dpkg -i %s || echo '%s' | sudo -S apt --fix-broken install -y
cat <<EOF | sudo tee /etc/default/zramswap >/dev/null
ALGO=lz4
SIZE=%d
PRIORITY=100
EOF
echo '%s' | sudo -S systemctl enable zramswap
echo '%s' | sudo -S systemctl restart zramswap
sleep 2
""" % (
        password, password, password, password,
        password, REMOTE_DEB, password,
        SWAP_SIZE_MB,
        password, password,
    )

    stdin, stdout, stderr = ssh.exec_command(script, get_pty=True)
    try:
        while True:
            line = stdout.readline()
            if not line:
                break
            print("    " + line.rstrip())
        rc = stdout.channel.recv_exit_status()
        err = stderr.read().decode(errors="ignore")
        return rc, err
    finally:
        stdin.close()
        stdout.close()
        stderr.close()


# -- Verification --------------------------------------------------------------

def verify_swap_1024mb(ssh: paramiko.SSHClient) -> tuple[bool, int | None, str]:
    """
    Parse `free -m` on the device and confirm the Swap total is exactly
    SWAP_SIZE_MB (1024). Returns (passed, actual_total_mb, raw_free_output).
    """
    rc, out, err = run(ssh, "free -m")
    if rc != 0:
        return False, None, err or out

    # Typical `free -m` line:  "Swap:          1024           0        1024"
    match = re.search(r"^Swap:\s+(\d+)", out, re.MULTILINE)
    if not match:
        return False, None, out

    total_mb = int(match.group(1))
    return total_mb == SWAP_SIZE_MB, total_mb, out


# -- Excel input --------------------------------------------------------------

def read_devices_from_excel(path: Path) -> list[dict]:
    """
    Read (store_code, device_ip) pairs from an .xlsx file.
    Header row is matched case-insensitively against STORE_HEADERS / IP_HEADERS.
    Rows missing an IP are skipped; rows missing a store code fall back to
    using the IP as the store label.
    """
    wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    ws = wb.active

    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if header is None:
        raise ValueError("Excel file is empty")

    store_col = ip_col = None
    for idx, cell in enumerate(header):
        name = str(cell).strip().lower() if cell is not None else ""
        if name in STORE_HEADERS and store_col is None:
            store_col = idx
        elif name in IP_HEADERS and ip_col is None:
            ip_col = idx

    if ip_col is None:
        raise ValueError(
            f"Could not find a device IP column in header {header!r}. "
            f"Expected one of: {sorted(IP_HEADERS)}"
        )

    devices = []
    for row in rows:
        if row is None or ip_col >= len(row):
            continue
        ip = row[ip_col]
        if ip is None or not str(ip).strip():
            continue
        ip = str(ip).strip()

        store = None
        if store_col is not None and store_col < len(row) and row[store_col] is not None:
            store = str(row[store_col]).strip()

        devices.append({"store": store or ip, "ip": ip})

    wb.close()
    return devices


# -- Per-device flow ------------------------------------------------------------

def process_device(store: str, ip: str, user: str, password: str):
    global success, failed, skipped

    print("\n" + "=" * 70)
    print(f"Processing [{store} | {ip}]")
    print("=" * 70)

    ssh = None
    try:
        print("Connecting...")
        ssh = connect(ip, user, password)
        print("[OK] Connected")

        was_already_configured = already_configured(ssh)
        if was_already_configured:
            print("[OK] Already configured with zram swap")
        else:
            print("Uploading zram-tools package...")
            upload_deb(ssh)
            print("Installing and configuring zramswap...")
            rc, err = install(ssh, password)
            if err.strip():
                print("    stderr: " + err.strip())
            if rc != 0:
                print(f"[FAIL] Install script exited with code {rc}")
                failed += 1
                return

        print("\nVerifying swap size...")
        passed, total_mb, raw = verify_swap_1024mb(ssh)
        print("-- free -m output --")
        for line in raw.strip().splitlines():
            print("    " + line)

        if passed and was_already_configured:
            print(f"[SKIP] Already at {total_mb} MB -- nothing to do")
            skipped += 1
        elif passed:
            print(f"[PASS] Swap is {total_mb} MB (expected {SWAP_SIZE_MB} MB)")
            success += 1
        else:
            shown = total_mb if total_mb is not None else "unknown"
            print(f"[FAIL] Swap is {shown} MB, expected {SWAP_SIZE_MB} MB")
            failed += 1

    except socket.timeout:
        print("[FAIL] Connection timed out")
        failed += 1
    except paramiko.AuthenticationException:
        print("[FAIL] SSH authentication failed")
        failed += 1
    except Exception as e:
        print(f"[FAIL] {e}")
        failed += 1
    finally:
        if ssh:
            ssh.close()


# -- Main ------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Expand swap to 1024MB via zram-tools on Pi devices.")
    parser.add_argument("--excel", help="Path to an .xlsx file with Store Code / Device IP columns")
    parser.add_argument("--device", help="Single target device IP")
    parser.add_argument("--devices-file", help="Path to a text file with one device IP per line")
    parser.add_argument("--user", default=SSH_USER_DEFAULT, help="SSH username (default: pi)")
    parser.add_argument("--password", default=SSH_PASS_DEFAULT, help="SSH password (default: industrail)")
    args = parser.parse_args()

    if not DEB_FILE.exists():
        print(f"[FAIL] Required file not found: {DEB_FILE}")
        print("       Place zram-tools_0.3.3.1-1_all.deb next to this script and retry.")
        sys.exit(1)

    devices: list[dict] = []

    if args.device:
        devices.append({"store": args.device.strip(), "ip": args.device.strip()})

    if args.devices_file:
        path = Path(args.devices_file)
        if not path.exists():
            print(f"[FAIL] Devices file not found: {path}")
            sys.exit(1)
        for line in path.read_text().splitlines():
            line = line.strip()
            if line:
                devices.append({"store": line, "ip": line})

    excel_path = Path(args.excel) if args.excel else (DEFAULT_EXCEL if DEFAULT_EXCEL.exists() else None)
    if args.excel and not excel_path.exists():
        print(f"[FAIL] Excel file not found: {excel_path}")
        sys.exit(1)
    if excel_path:
        print(f"Reading devices from {excel_path} ...")
        try:
            devices += read_devices_from_excel(excel_path)
        except Exception as e:
            print(f"[FAIL] Could not read {excel_path}: {e}")
            sys.exit(1)
        print(f"[OK] Loaded {len(devices)} device(s) from Excel")

    if not devices:
        print("[FAIL] No target devices given. Use --excel <path>, --device <ip>, or --devices-file <path>.")
        sys.exit(1)

    for dev in devices:
        process_device(dev["store"], dev["ip"], args.user, args.password)

    print("\n" + "=" * 70)
    print("Expand Swap -- Finished")
    print("=" * 70)
    print(f"Successful : {success}")
    print(f"Skipped    : {skipped}")
    print(f"Failed     : {failed}")

    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
