# V1.1 product experience

Authorized from V1 stable `d0db778`. Implementation requires separate PO Desktop
acceptance before Git integration/release. No Desktop PASS is implied.

## Presentation

Keep Tkinter/ttk and existing service boundaries. Home uses A's two-by-two cards,
C's charcoal/warm-white/gold theme, and A's teal S-curve icon. Create is primary;
Mapping/Payment are optional. Home and persistent Help routes share one guide.
Finance and AI remain registered but absent from normal navigation, including
Finance's Home and Tools entries. Their underlying behavior is retained.

Rebuild: Workbook → Target → conditional options → Generate/Refresh.
Progress/Payment retain Snapshot/Live; EV retains Live and its two monetary
sources. No Rebuild All or Snapshot EV. Existing engines remain separate.

## Localization

`presentation/gui/strings.py` owns language selection, English fallback and named
placeholder rendering. `translations.py` owns bilingual resources. New views use
namespaced keys; existing English message IDs remain stable translation keys.
The launch language remains unchanged while the app is open. ENG/THA selection
persists through the existing layout preferences and applies next launch.
Missing/invalid legacy language defaults to ENG. Layout saves preserve language.

Translate presentation only. Canonical radio values, enums, paths, source field
identities, user data, formulas and sheet names are not translated. Detailed
unexpected service diagnostics may remain English under a localized message.
Dates/numbers retain accepted formatting; no automatic locale conversion.

## Protection

Custom password is cancelled. Worksheet protection still uses `okmd`, not file
encryption. The workbook README explains manual Excel changes may be reset by
processing paths using shared protection. No password propagation architecture.

## Branding

Canonical asset: `progress_studio/assets/brand/progress_studio.svg`.
Run `python scripts/build-brand-assets.py` to produce both PNG assets and the
multi-resolution ICO using Pillow. Existing EXE/installer paths remain unchanged.
The taskbar/pinned shortcut and Windows cache behavior require Windows inspection.

## PO Desktop acceptance (not automated PASS)

1. Launch; inspect S-curve icon, Home hierarchy and all normal navigation.
2. Confirm Finance absent from sidebar/Home/Tools and AI placeholder hidden.
3. Open Help from Home and sidebar; verify first-time/cycle flow and back route.
4. Start with fresh preferences: ENG. Select THA, continue current work, close/reopen:
   Thai UI. Switch back to ENG and restart. Check settings/layout persistence.
5. Check Thai marks/wrapping, dialogs, Tab/shortcuts and 100/125/150/200% DPI where practical.
6. Create P6 and MSP workbooks with Equal/Duration/Amount; inspect README protection note.
7. Verify Mapping/export and Payment Input/Breakdown still work.
8. Rebuild Progress Snapshot and Live; Payment Snapshot and Live.
9. EV Activity Amount and BOQ Mapping: appropriate validation/source controls and Live behavior.
10. Excel first-open without repair; S-curve, F9 and Save/Close/Reopen; EV Status Date independence.
11. Verify window/portable EXE/installer/taskbar icons on Windows; run package smoke.

See the implementation handoff for actual automated counts and limitations.
