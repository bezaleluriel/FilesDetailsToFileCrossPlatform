import sys


def main() -> int:
    try:
        from qt_app import main as qt_main
    except Exception as e:  # noqa: BLE001
        print("This project now uses the Qt app (PySide6).", file=sys.stderr)
        print("Install dependencies then run: python qt_app.py", file=sys.stderr)
        print(f"Error: {e}", file=sys.stderr)
        return 1
    return int(qt_main())


if __name__ == "__main__":
    raise SystemExit(main())

