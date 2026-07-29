"""Compatibility launcher for the Progress Studio desktop application."""

from progress_studio.entrypoints import desktop_main


def main() -> int:
    return desktop_main()


if __name__ == "__main__":
    raise SystemExit(main())
