"""Compatibility launcher for the Progress Studio CLI."""

from progress_studio.entrypoints import cli_main


def main() -> int:
    return cli_main()


if __name__ == "__main__":
    raise SystemExit(main())
