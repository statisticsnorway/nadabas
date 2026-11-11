"""Own one empty Excel COM session in a separate, short-lived test process."""

import json
import sys
import time
from pathlib import Path

import win32com.client
import win32process


def main():
    directory = Path(sys.argv[1]).resolve(strict=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    try:
        _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        (directory / "excel-owner.json").write_text(json.dumps({"pid": pid}))
        deadline = time.monotonic() + 120
        while not (directory / "close-owned-excel").exists():
            if time.monotonic() >= deadline:
                raise TimeoutError("Test controller did not request Excel shutdown")
            time.sleep(0.2)
    finally:
        excel.Quit()
        del excel


if __name__ == "__main__":
    main()
