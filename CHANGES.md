# Changelog

## v2.0 – 2026-03-11  (optimization & crash-fix release)

### Bug Fixes / Stability
- **Crash: double-generation** – added `threading.Lock` so clicking Generate
  twice no longer spawns concurrent threads that corrupt the output folder.
- **Crash: temp file leak** – replaced bare `os.remove` with
  `tempfile.NamedTemporaryFile` + `finally` block; temp PNG is always cleaned
  up even if PDF writing raises an exception.
- **Crash: missing font fallback** – `_get_font()` always returns a valid font
  object; no more `NoneType` errors when a `.ttf` is absent.
- **Crash: drag position desync** – placeholder dict now stores live `x`/`y`
  and is updated on every drag event; coordinates no longer drift.
- **Crash: load_project with missing files** – graceful warning instead of
  unhandled `FileNotFoundError`.
- **Crash: Excel open_in_readonly** – `read_only=True` + `wb.close()` prevent
  file-lock errors on Windows.

### Performance / Binary Size
- `ttkthemes` removed (unused, added ~8 MB to binary).
- `colorsys` import removed (logic inlined).
- `read_only=True, data_only=True` on openpyxl reduces memory for large sheets.
- PIL images now converted to RGB before PDF embedding, saving an extra
  RGBA→RGB conversion pass inside fpdf2.
- `build.spec` excludes matplotlib, numpy, PyQt5/6, PySide, wx, IPython and
  other heavy packages – expected binary reduction ≥ 40 %.
- `optimize=2` in PyInstaller strips docstrings and asserts from bytecode.
- `strip=True` + `upx=True` compress native libraries.

### Code Quality
- Monolithic class refactored into clearly labelled sections.
- All internal helpers prefixed with `_`; public API unchanged.
- Dependency guard at import time with user-friendly error dialogs.
- `resource_path()` helper centralises PyInstaller/dev path resolution.
- Legacy method aliases (`update_status`, `update_info`,
  `get_placeholder_positions`) kept for backward compatibility with saved
  `.certwiz` project files.

### UI
- Color preview swatch next to each field now **live-updates** when a colour
  is picked.
- CMYK picker window is non-resizable and modal (no more ghost windows).
- Progress bar updates via `root.after()` instead of
  `root.update_idletasks()` in the worker thread (safer cross-thread UI).
- Status bar and log panel use thread-safe `root.after(0, ...)` scheduling.
