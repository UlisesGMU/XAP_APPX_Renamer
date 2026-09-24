# XAP/APPX Renamer

A Windows desktop tool that renames `.xap`, `.appx`, and `.appxbundle` files
based on the real app name and version read from inside the package itself,
instead of whatever cryptic filename it originally shipped with.

Built with the assistance of Claude (Anthropic).

## What it does

- Reads each package's manifest directly (no installation required) to pull
  out the app's display name, publisher, and version.
- Resolves indirect name references (`ms-resource:...`) against the
  package's own bundled resource data, rather than showing the raw,
  unresolved reference string.
- For `.appxbundle` files, opens the relevant inner package to resolve a
  real name, since a bundle's own manifest doesn't carry one on its own.
- Batch renaming across a whole folder, with per-column filters, sortable
  columns, inline editing, and per-row lock/unlock before committing.
- Shows an app's tile icon on hover, and can reveal any file directly in
  Explorer.

## Requirements

- Windows 10 or 11.
- Python 3.9+ (Tkinter included with a standard Windows install).
- No administrator rights needed — this only reads package files and
  renames them on disk; it never installs, modifies, or removes anything.

## Running it

```
python xap_appx_renamer.py
```

## Building a standalone .exe (optional)

```
pip install pyinstaller
pyinstaller --onefile --noconsole --name XapAppxRenamer xap_appx_renamer.py
```

The resulting `.exe` (in `dist\`) runs without Python installed, but only
on machines matching the same architecture (32-bit vs 64-bit) as whichever
Python built it.

## Known limitations

- Full name resolution for indirect (`ms-resource:`) references works best
  when PowerShell is available (used internally to resolve them via the
  same underlying mechanism Windows itself uses) — this ships with every
  copy of Windows 10/11, so this normally isn't something you need to set
  up separately.
- A handful of very old or unusually-structured packages may not expose a
  resolvable display name at all, in which case the tool falls back to the
  package's raw internal identity name instead.

## License

GPL-3.0. See the `LICENSE` file (or add one from
[gnu.org](https://www.gnu.org/licenses/gpl-3.0.txt) if you haven't yet) —
note that GPL-3.0 is copyleft: anything you build on top of this and
distribute must also be released under GPL-3.0.
