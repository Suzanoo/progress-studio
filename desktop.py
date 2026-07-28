from progress_studio.presentation.gui.app import ProgressStudioDesktopApp


def main() -> int:
    app = ProgressStudioDesktopApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
