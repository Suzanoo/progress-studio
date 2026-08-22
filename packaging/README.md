# Packaging

Packaging definitions live here. CL-2D only establishes the packaging boundary; it does not create platform installers.

- `windows/` — future Windows executable/installer configuration.
- `macos/` — future macOS app bundle/signing/notarization configuration.

Production builds must be created from a clean repository and must pass the package-content smoke check before platform-specific packaging.
