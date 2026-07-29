"""Module entry point for ``python -m progress_studio``."""

from .entrypoints import desktop_main


if __name__ == "__main__":
    raise SystemExit(desktop_main())
