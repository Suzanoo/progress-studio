# macOS Packaging

Reserved for the macOS application bundle configuration.

Before adding bundle/signing definitions, verify:

- the wheel contains all JSON configuration and dashboard icon resources;
- resource lookup does not depend on the current working directory;
- Tk/Excel/file-dialog behavior is verified on the supported macOS target;
- signing/notarization requirements are documented before release.
