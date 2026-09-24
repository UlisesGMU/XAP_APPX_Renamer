#!/usr/bin/env python3
"""
XAP / APPX / APPXBUNDLE Renamer
--------------------------------
Scans .xap, .appx, and .appxbundle files (they're all ZIP archives),
reads the embedded manifest to find the real application name and
version, and lets you rename the files to "Name_Version.ext".

No third-party dependencies - just the Python standard library
(zipfile, xml.etree, tkinter).

Usage:
    python xap_appx_renamer.py
"""

import os
import re
import shutil
import glob
import tempfile
import subprocess
import zipfile
import functools
import io
import time
import queue
import threading
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

SCRIPT_VERSION = "1.17.0"
SUPPORTED_EXTS = ('.xap', '.appx', '.appxbundle')
INVALID_CHARS = r'<>:"/\|?*'


def sanitize(name: str) -> str:
    """Make a string safe to use as a filename component."""
    if not name:
        return name
    name = name.strip()
    for ch in INVALID_CHARS:
        name = name.replace(ch, '')
    name = re.sub(r'\s+', ' ', name)
    return name.strip()


def select_in_explorer(path):
    """Opens the containing folder with this exact file/folder highlighted,
    via the actual Windows Shell API (SHParseDisplayName +
    SHOpenFolderAndSelectItems) - the same mechanism Explorer, browsers,
    and download managers use internally for "Show in folder". This avoids
    shelling out to explorer.exe's own "/select," command-line syntax
    entirely, which turned out to be unreliable via subprocess regardless
    of how that argument was quoted. Raises OSError on failure."""
    import ctypes
    path = os.path.normpath(os.path.abspath(path))
    shell32 = ctypes.windll.shell32
    ole32 = ctypes.windll.ole32
    ole32.CoInitialize(None)
    try:
        pidl = ctypes.c_void_p()
        hr = shell32.SHParseDisplayName(path, None, ctypes.byref(pidl), 0, None)
        if hr != 0 or not pidl:
            raise OSError(f"SHParseDisplayName failed (0x{hr & 0xFFFFFFFF:08X}) for: {path}")
        try:
            hr2 = shell32.SHOpenFolderAndSelectItems(pidl, 0, None, 0)
            if hr2 != 0:
                raise OSError(f"SHOpenFolderAndSelectItems failed (0x{hr2 & 0xFFFFFFFF:08X}) for: {path}")
        finally:
            ole32.CoTaskMemFree(pidl)
    finally:
        ole32.CoUninitialize()


def find_manifest(namelist, target_lower):
    for n in namelist:
        if n.lower() == target_lower:
            return n
    return None


def local(tag):
    """Strip the XML namespace from a tag, e.g. '{ns}Identity' -> 'Identity'."""
    return tag.split('}', 1)[-1] if '}' in tag else tag


def find_child(root, tagname):
    """Find the first descendant element with this local tag name, ignoring namespaces."""
    for el in root.iter():
        if local(el.tag) == tagname:
            return el
    return None


def _find_logo_hint(root):
    """Finds the logo/tile image path hint from a parsed AppxManifest.xml
    root element. Tries <Properties><Logo> first (present in virtually all
    manifests), then falls back to the Square150x150Logo/Square44x44Logo
    attributes on <uap:VisualElements> - some manifests, especially
    converted or unusual ones, only declare it that way."""
    props = find_child(root, 'Properties')
    if props is not None:
        logo_el = find_child(props, 'Logo')
        if logo_el is not None and logo_el.text and logo_el.text.strip():
            return logo_el.text.strip()

    for el in root.iter():
        if local(el.tag) == 'VisualElements':
            for attr in ('Square150x150Logo', 'Square44x44Logo'):
                val = el.attrib.get(attr)
                if val and val.strip():
                    return val.strip()
    return None


def _resolve_logo_bytes(z, names, logo_hint):
    """Given an open ZipFile, its namelist, and a logo hint path from the
    manifest, finds and returns the best-matching image's bytes. The hint
    is only a HINT, not necessarily an exact filename - modern apps ship
    several scaled variants (Logo.scale-100.png, Logo.scale-200.png, ...)
    rather than one file at that exact path, so this falls back to a
    best-match search alongside an exact-path check."""
    logo_hint_norm = logo_hint.replace('\\', '/')

    for n in names:
        if n.replace('\\', '/').lower() == logo_hint_norm.lower():
            return z.read(n)

    base_dir = os.path.dirname(logo_hint_norm)
    base_name = os.path.basename(logo_hint_norm)
    stem = base_name.split('.')[0]
    ext = os.path.splitext(base_name)[1] or '.png'

    candidates = [
        n for n in names
        if os.path.dirname(n.replace('\\', '/')).lower() == base_dir.lower()
        and os.path.basename(n).lower().startswith(stem.lower())
        and n.lower().endswith(ext.lower())
    ]
    if not candidates:
        # Widen to the whole archive - cheap since names is already in
        # memory (no extra disk I/O for this list itself).
        candidates = [
            n for n in names
            if os.path.basename(n).lower().startswith(stem.lower())
            and n.lower().endswith(ext.lower())
        ]

    if candidates:
        for c in candidates:
            if 'scale-100' in c.lower():
                return z.read(c)
        return z.read(candidates[0])
    return None


@functools.lru_cache(maxsize=256)
def find_manifest_icon_bytes(zip_path):
    """Extract the app's logo/tile image bytes from a .xap/.appx/.appxbundle
    zip file, for the hover-preview tooltip. Returns bytes or None. Cached
    by path since the same file is hovered repeatedly and re-reading/
    re-parsing the zip each time would be wasteful.

    A .appxbundle's OUTER zip never contains AppxManifest.xml directly - it
    only has an AppxBundleManifest.xml plus several NESTED .appx files
    (each themselves a full zip), and the real per-app manifest (with the
    actual Logo reference) lives inside one of those. So bundles need an
    extra step: read each nested .appx's bytes and open THAT as its own
    zip, rather than only ever looking at the outer archive's file list.
    """
    try:
        with zipfile.ZipFile(zip_path) as z:
            names = z.namelist()

            manifest_name = find_manifest(names, 'appxmanifest.xml')
            if manifest_name:
                root = ET.fromstring(z.read(manifest_name))
                hint = _find_logo_hint(root)
                if hint:
                    icon = _resolve_logo_bytes(z, names, hint)
                    if icon:
                        return icon

            manifest_name = find_manifest(names, 'wmappmanifest.xml')
            if manifest_name:
                root = ET.fromstring(z.read(manifest_name))
                icon_el = find_child(root, 'IconPath')
                if icon_el is not None and icon_el.text and icon_el.text.strip():
                    icon = _resolve_logo_bytes(z, names, icon_el.text.strip())
                    if icon:
                        return icon

            # .appxbundle case - see docstring above.
            for nested_name in [n for n in names if n.lower().endswith('.appx')]:
                try:
                    with zipfile.ZipFile(io.BytesIO(z.read(nested_name))) as nz:
                        nnames = nz.namelist()
                        nmanifest = find_manifest(nnames, 'appxmanifest.xml')
                        if not nmanifest:
                            continue
                        nroot = ET.fromstring(nz.read(nmanifest))
                        nhint = _find_logo_hint(nroot)
                        if nhint:
                            icon = _resolve_logo_bytes(nz, nnames, nhint)
                            if icon:
                                return icon
                except Exception:
                    continue

            return None
    except Exception:
        return None


class WrappingFrame(ttk.Frame):
    """A frame that lays out its children left to right, wrapping to a new
    row when they'd overflow the frame's current width - used for the
    per-column filter row, since several filter boxes side by side would
    otherwise overflow the window instead of wrapping."""

    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self._items = []
        self.bind('<Configure>', self._reflow)

    def add_widget(self, widget):
        self._items.append(widget)
        self._reflow()

    def _reflow(self, event=None):
        width = self.winfo_width()
        if width <= 1:
            self.after(10, self._reflow)
            return
        x = 0
        row = 0
        col = 0
        for w in self._items:
            w.update_idletasks()
            w_width = w.winfo_reqwidth()
            if x + w_width > width and col > 0:
                row += 1
                col = 0
                x = 0
            w.grid(in_=self, row=row, column=col, sticky='w', padx=(0, 4), pady=2)
            x += w_width + 4
            col += 1


class IconTooltip:
    """Shows a small floating icon preview near the cursor when hovering a
    Treeview row. icon_provider(row_id) should return image bytes (PNG) or
    None; the byte-fetching itself should do its own caching by path since
    row ids are regenerated on every re-render (e.g. search filtering)."""

    def __init__(self, tree, icon_provider):
        self.tree = tree
        self.icon_provider = icon_provider
        self.tip_win = None
        self.tip_image = None  # keep a reference so Tk doesn't garbage-collect it
        self.current_row = None
        tree.bind('<Motion>', self._on_motion, add='+')
        tree.bind('<Leave>', self._on_leave, add='+')

    def _on_motion(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id != self.current_row:
            self._hide()
            self.current_row = row_id
            if row_id:
                self._maybe_show(row_id, event)
        elif row_id and self.tip_win:
            self._reposition(event)

    def _on_leave(self, event):
        self._hide()
        self.current_row = None

    def _maybe_show(self, row_id, event):
        try:
            raw_bytes = self.icon_provider(row_id)
        except Exception:
            raw_bytes = None
        if not raw_bytes:
            return
        try:
            img = tk.PhotoImage(data=raw_bytes)
        except Exception:
            return  # not a format Tk's built-in PhotoImage can decode (e.g. JPEG)
        w = img.width()
        target = 64
        if w > target:
            factor = max(1, w // target)
            img = img.subsample(factor, factor)
        self.tip_image = img
        self.tip_win = tk.Toplevel(self.tree)
        self.tip_win.wm_overrideredirect(True)
        try:
            self.tip_win.attributes('-topmost', True)
        except tk.TclError:
            pass
        tk.Label(self.tip_win, image=self.tip_image, borderwidth=1, relief='solid', background='white').pack()
        self._reposition(event)

    def _reposition(self, event):
        if self.tip_win:
            x = self.tree.winfo_rootx() + event.x + 18
            y = self.tree.winfo_rooty() + event.y + 18
            self.tip_win.wm_geometry(f'+{x}+{y}')

    def _hide(self):
        if self.tip_win:
            try:
                self.tip_win.destroy()
            except tk.TclError:
                pass
            self.tip_win = None
        self.tip_image = None


def get_candidate_qualifier(cand_el):
    """
    makepri's dump XML tags each Candidate's language/scale/etc. info, but the
    actual attribute name varies by makepri version - some use 'qualifiers',
    others use 'locale'. Check both rather than assuming one.
    """
    return cand_el.get('qualifiers') or cand_el.get('locale') or ''


def strip_publisher_prefix(name):
    """
    Raw package Identity Names commonly look like 'Publisher.AppName' (e.g.
    'Contoso.MyGreatApp', '51045ContosoInc.MyApp'). This strips just the
    first dot-separated segment, leaving the rest as-is. Only meant to be
    applied to raw identity names, never to a genuinely resolved display
    name (from the manifest or resources.pri) which is already the real
    app name.
    """
    if not name or '.' not in name:
        return name
    prefix, rest = name.split('.', 1)
    return rest if rest else name


def parse_appx_manifest(data):
    root = ET.fromstring(data)
    identity = find_child(root, 'Identity')
    name = identity.get('Name') if identity is not None else None
    version = identity.get('Version') if identity is not None else None

    display_name = None
    ms_resource_ref = None
    props = find_child(root, 'Properties')
    if props is not None:
        dn_el = find_child(props, 'DisplayName')
        if dn_el is not None and dn_el.text:
            if dn_el.text.lower().startswith('ms-resource'):
                ms_resource_ref = dn_el.text
            else:
                display_name = dn_el.text

    # identity_based: True when we're falling back to the raw package
    # Identity Name rather than a real resolved display name - only these
    # are eligible for publisher-prefix stripping.
    return {
        'name': display_name or name,
        'version': version,
        'ms_resource_ref': ms_resource_ref,
        'identity_based': display_name is None,
    }


def parse_appxbundle_manifest(data):
    # Bundle manifests generally only carry the package identity, not a
    # resolved friendly display name (that lives in resources.pri of the
    # inner packages), so we fall back to the Identity Name.
    root = ET.fromstring(data)
    identity = find_child(root, 'Identity')
    name = identity.get('Name') if identity is not None else None
    version = identity.get('Version') if identity is not None else None
    return {'name': name, 'version': version, 'identity_based': True}


def parse_wmapp_manifest(data):
    root = ET.fromstring(data)
    app = find_child(root, 'App')
    title = app.get('Title') if app is not None else None
    version = app.get('Version') if app is not None else None
    product_id = app.get('ProductID') if app is not None else None
    wp_resource_ref = None
    if title and title.startswith('@'):
        # Indirect string reference, e.g. "@AppResources.dll,-131492872",
        # pointing at a string-table resource in a DLL bundled in the .xap.
        wp_resource_ref = title
        title = None
    return {
        'name': title or product_id,
        'version': version,
        'identity_based': title is None,
        'wp_resource_ref': wp_resource_ref,
    }


def resolve_wp_resource_dll(zip_path, ref_value):
    """
    Best-effort resolution of a Windows Phone indirect string reference like
    '@AppResources.dll,-101' (same convention as Windows' "@shell32.dll,-1216"
    style indirect strings): finds the named DLL inside the .xap, loads it as
    a data file (no code execution, works regardless of the DLL's original
    CPU architecture), and reads the string resource via the Win32 API.
    Windows-only; returns None anywhere this can't be done.
    """
    if not ref_value or not ref_value.startswith('@') or os.name != 'nt':
        return None
    m = re.match(r'^@([^,]+),\s*(-?\d+)\s*$', ref_value)
    if not m:
        return None
    dll_name, id_str = m.group(1).strip(), m.group(2)
    res_id = abs(int(id_str))

    tmpdir = tempfile.mkdtemp(prefix='wpres_')
    try:
        with zipfile.ZipFile(zip_path) as z:
            entries = [n for n in z.namelist() if os.path.basename(n).lower() == dll_name.lower()]
            if not entries:
                return None
            dll_path = os.path.join(tmpdir, os.path.basename(dll_name))
            with open(dll_path, 'wb') as f:
                f.write(z.read(entries[0]))

        import ctypes
        import ctypes.wintypes as wintypes
        LOAD_LIBRARY_AS_DATAFILE = 0x00000002
        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        user32 = ctypes.WinDLL('user32', use_last_error=True)

        # Must set these explicitly: LoadLibraryExW returns a 64-bit HMODULE,
        # and ctypes defaults foreign-function return types to a 32-bit int,
        # which silently truncates the handle on 64-bit Python and makes a
        # perfectly successful load look like a failure.
        kernel32.LoadLibraryExW.restype = wintypes.HMODULE
        kernel32.LoadLibraryExW.argtypes = [wintypes.LPCWSTR, wintypes.HANDLE, wintypes.DWORD]
        kernel32.FreeLibrary.restype = wintypes.BOOL
        kernel32.FreeLibrary.argtypes = [wintypes.HMODULE]
        user32.LoadStringW.restype = ctypes.c_int
        user32.LoadStringW.argtypes = [wintypes.HINSTANCE, wintypes.UINT, wintypes.LPWSTR, ctypes.c_int]

        handle = kernel32.LoadLibraryExW(dll_path, None, LOAD_LIBRARY_AS_DATAFILE)
        if not handle:
            return None
        try:
            buf = ctypes.create_unicode_buffer(1024)
            n = user32.LoadStringW(handle, res_id, buf, len(buf))
            return buf.value if n > 0 else None
        finally:
            kernel32.FreeLibrary(handle)
    except Exception:
        return None
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


@functools.lru_cache(maxsize=1)
def find_makepri():
    """Locate makepri.exe (ships with the Windows 10/11 SDK). Windows-only, best effort."""
    exe = shutil.which('makepri.exe') or shutil.which('makepri')
    if exe:
        return exe
    candidates = glob.glob(r"C:\Program Files (x86)\Windows Kits\10\bin\*\x64\makepri.exe")
    candidates += glob.glob(r"C:\Program Files (x86)\Windows Kits\10\bin\*\x86\makepri.exe")
    return sorted(candidates)[-1] if candidates else None


def resolve_ms_resource(zip_path, ms_resource_ref):
    """
    Best-effort resolution of an 'ms-resource:XYZ' reference using a
    resources.pri file found inside the package. Requires makepri.exe
    (Windows SDK) to be available on this machine. Returns the resolved
    string, or None if it can't be resolved here.
    """
    if not ms_resource_ref:
        return None
    key = ms_resource_ref.split(':', 1)[-1].strip('/').lower()
    makepri = find_makepri()
    if not makepri:
        return None

    tmpdir = tempfile.mkdtemp(prefix='pri_')
    try:
        with zipfile.ZipFile(zip_path) as z:
            pri_entries = [n for n in z.namelist() if n.lower().endswith('resources.pri')]
            if not pri_entries:
                return None
            pri_path = os.path.join(tmpdir, 'resources.pri')
            with open(pri_path, 'wb') as f:
                f.write(z.read(pri_entries[0]))

        dump_path = os.path.join(tmpdir, 'dump.xml')
        subprocess.run(
            [makepri, 'dump', '/if', pri_path, '/of', dump_path, '/dt', 'Detailed', '/o'],
            capture_output=True, timeout=30, check=False
        )
        if not os.path.exists(dump_path):
            return None

        tree = ET.parse(dump_path)
        for el in tree.getroot().iter():
            if local(el.tag) == 'NamedResource':
                res_name = (el.get('name') or '').split('/')[-1].lower()
                if res_name != key:
                    continue
                candidates = []
                for cand in el:
                    if local(cand.tag) != 'Candidate':
                        continue
                    val_el = find_child(cand, 'Value')
                    if val_el is not None and val_el.text:
                        candidates.append((get_candidate_qualifier(cand), val_el.text))
                if not candidates:
                    continue
                # Prefer an English-qualified candidate (e.g. "Language-EN-US")
                # over whatever happens to be listed first.
                for qualifiers, value in candidates:
                    if re.search(r'(?:^|[-_])en(?:[-_]|$)', qualifiers, re.IGNORECASE):
                        return value
                return candidates[0][1]
        return None
    except Exception:
        return None
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def extract_info(path):
    """Return {'name':..., 'version':...} or {'error':...} for one package file."""
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()

            m = find_manifest(names, 'appxmanifest.xml')
            if m:
                info = parse_appx_manifest(z.read(m))
                ref = info.pop('ms_resource_ref', None)
                if ref:
                    resolved = resolve_ms_resource(path, ref)
                    if resolved:
                        info['name'] = resolved
                        info['resolved_from_pri'] = True
                        info['identity_based'] = False
                return info

            m = find_manifest(names, 'appxmetadata/appxbundlemanifest.xml')
            if m:
                return parse_appxbundle_manifest(z.read(m))

            m = find_manifest(names, 'wmappmanifest.xml')
            if m:
                info = parse_wmapp_manifest(z.read(m))
                ref = info.pop('wp_resource_ref', None)
                if ref:
                    resolved = resolve_wp_resource_dll(path, ref)
                    if resolved:
                        info['name'] = resolved
                        info['resolved_from_dll'] = True
                        info['identity_based'] = False
                return info

        return {'error': 'No recognized manifest found inside the archive'}
    except zipfile.BadZipFile:
        return {'error': 'Not a valid zip/package file'}
    except ET.ParseError:
        return {'error': 'Could not parse manifest XML'}
    except Exception as e:
        return {'error': str(e)}


def get_language_candidates(path):
    """
    For a resource-based display name, return (results, error_message) where
    results is a list of (qualifiers, value) candidates found for that
    resource key, or None with an explanatory message if none apply/found.
    """
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()

            m = find_manifest(names, 'appxmanifest.xml')
            if m:
                raw = parse_appx_manifest(z.read(m))
                ref = raw.get('ms_resource_ref')
                if not ref:
                    return None, "This file's name is a literal string in the manifest (not an ms-resource reference) - there's nothing to look up."
                key = ref.split(':', 1)[-1].strip('/').lower()
                makepri = find_makepri()
                if not makepri:
                    return None, "makepri.exe was not found on this system - can't inspect resources.pri."
                pri_entries = [n for n in names if n.lower().endswith('resources.pri')]
                if not pri_entries:
                    return None, "No resources.pri file found inside this package."

                tmpdir = tempfile.mkdtemp(prefix='prilang_')
                try:
                    pri_path = os.path.join(tmpdir, 'resources.pri')
                    with open(pri_path, 'wb') as f:
                        f.write(z.read(pri_entries[0]))
                    dump_path = os.path.join(tmpdir, 'dump.xml')
                    subprocess.run(
                        [makepri, 'dump', '/if', pri_path, '/of', dump_path, '/dt', 'Detailed', '/o'],
                        capture_output=True, timeout=30, check=False
                    )
                    if not os.path.exists(dump_path):
                        return None, "makepri dump failed - couldn't read resources.pri on this system."

                    results = []
                    tree = ET.parse(dump_path)
                    for el in tree.getroot().iter():
                        if local(el.tag) != 'NamedResource':
                            continue
                        res_name = (el.get('name') or '').split('/')[-1].lower()
                        if res_name != key:
                            continue
                        for cand in el:
                            if local(cand.tag) != 'Candidate':
                                continue
                            val_el = find_child(cand, 'Value')
                            if val_el is not None and val_el.text:
                                qualifiers = get_candidate_qualifier(cand) or '(neutral / no qualifiers)'
                                results.append((qualifiers, val_el.text))
                    if not results:
                        return None, f"No candidates found for resource key '{key}' in resources.pri."
                    return results, None
                finally:
                    shutil.rmtree(tmpdir, ignore_errors=True)

            m = find_manifest(names, 'appxmetadata/appxbundlemanifest.xml')
            if m:
                return None, "This is an .appxbundle - at this level it only carries the raw package identity name, not a resource-based display name to look up."

            m = find_manifest(names, 'wmappmanifest.xml')
            if m:
                raw = parse_wmapp_manifest(z.read(m))
                ref = raw.get('wp_resource_ref')
                if not ref:
                    return None, "This file's title is a literal string (not an indirect resource reference) - there's nothing to look up."
                return None, (
                    f"This is a .xap indirect string reference ({ref}) into a resource DLL - it's a single "
                    "string per DLL/ID pair, not a list of language candidates like resources.pri has."
                )

            return None, "No recognized manifest found inside this file."
    except zipfile.BadZipFile:
        return None, "Not a valid zip/package file."
    except Exception as e:
        return None, f"Error while inspecting file: {e}"


def get_debug_info(path):
    """Human-readable dump of everything the script can determine about a file, for troubleshooting."""
    lines = [f"File: {path}", ""]
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            lines.append(f"Archive contains {len(names)} entr{'y' if len(names) == 1 else 'ies'}.")

            m = find_manifest(names, 'appxmanifest.xml')
            m2 = find_manifest(names, 'appxmetadata/appxbundlemanifest.xml')
            m3 = find_manifest(names, 'wmappmanifest.xml')

            if m:
                lines.append(f"\nManifest found: {m} (AppxManifest.xml)")
                raw = parse_appx_manifest(z.read(m))
                lines.append(f"  Identity Name (raw): {raw.get('name')}")
                lines.append(f"  Version: {raw.get('version')}")
                lines.append(f"  ms-resource reference: {raw.get('ms_resource_ref')}")
                lines.append(f"  identity_based (no real display name found): {raw.get('identity_based')}")
                pri_present = any(n.lower().endswith('resources.pri') for n in names)
                lines.append(f"  resources.pri present in package: {pri_present}")
                makepri_path = find_makepri()
                lines.append(f"  makepri.exe found on this system: {makepri_path or 'NO'}")

                lines.append("\nHow the name was determined:")
                ref = raw.get('ms_resource_ref')
                if not raw.get('identity_based'):
                    lines.append(f"  1. Manifest's Properties/DisplayName was a literal string: '{raw.get('name')}'.")
                    lines.append("  2. Used directly - no resource lookup needed.")
                elif ref:
                    lines.append(f"  1. Properties/DisplayName was an indirect reference: '{ref}'.")
                    if not pri_present:
                        lines.append("  2. No resources.pri found inside the package - can't resolve it.")
                        lines.append(f"  3. Fell back to the raw Identity Name: '{raw.get('name')}'.")
                    elif not makepri_path:
                        lines.append("  2. resources.pri is present, but makepri.exe isn't installed on this system.")
                        lines.append(f"  3. Fell back to the raw Identity Name: '{raw.get('name')}'.")
                    else:
                        resolved = resolve_ms_resource(path, ref)
                        if resolved:
                            lines.append(f"  2. Ran makepri dump on resources.pri and found a match: '{resolved}'.")
                            lines.append("  3. Used the resolved string (preferring an English-qualified candidate if more than one language was found).")
                        else:
                            lines.append("  2. Ran makepri dump on resources.pri but found no matching candidate for that key.")
                            lines.append(f"  3. Fell back to the raw Identity Name: '{raw.get('name')}'.")
                else:
                    lines.append("  1. No Properties/DisplayName element at all in the manifest.")
                    lines.append(f"  2. Fell back to the raw Identity Name: '{raw.get('name')}'.")

            if m2:
                lines.append(f"\nManifest found: {m2} (AppxBundleManifest.xml)")
                raw = parse_appxbundle_manifest(z.read(m2))
                lines.append(f"  Identity Name (raw): {raw.get('name')}")
                lines.append(f"  Version: {raw.get('version')}")
                inner_appx = [n for n in names if n.lower().endswith('.appx')]
                lines.append(f"  Inner .appx packages found: {len(inner_appx)}")
                for n in inner_appx[:15]:
                    lines.append(f"    - {n}")
                lines.append("\nHow the name was determined:")
                lines.append("  1. Bundle manifests only ever carry the raw package Identity Name at this level -")
                lines.append("     the real display name (if any) lives inside the inner .appx packages listed above,")
                lines.append("     which this script does not currently open individually.")
                lines.append(f"  2. Used the raw Identity Name: '{raw.get('name')}'.")

            if m3:
                lines.append(f"\nManifest found: {m3} (WMAppManifest.xml / .xap)")
                raw = parse_wmapp_manifest(z.read(m3))
                lines.append(f"  Resolved literal title (if any): {raw.get('name')}")
                lines.append(f"  Version: {raw.get('version')}")
                ref = raw.get('wp_resource_ref')
                lines.append(f"  Indirect resource reference: {ref}")
                fallback_guid = raw.get('name')

                lines.append("\nHow the name was determined:")
                if not raw.get('identity_based'):
                    lines.append(f"  1. App/@Title in the manifest was a literal string: '{raw.get('name')}'.")
                    lines.append("  2. Used directly - no resource lookup needed.")
                elif ref:
                    lines.append(f"  1. App/@Title was an indirect reference: '{ref}'.")
                    match = re.match(r'^@([^,]+),\s*(-?\d+)\s*$', ref)
                    if not match:
                        lines.append("  2. Reference format did not match the expected '@Dll.dll,-ID' pattern.")
                        lines.append(f"  3. Fell back to ProductID (GUID): '{fallback_guid}'.")
                    else:
                        dll_name = match.group(1)
                        present = any(os.path.basename(n).lower() == dll_name.lower() for n in names)
                        lines.append(f"  2. Referenced DLL '{dll_name}' - present in package: {present}.")
                        if not present:
                            lines.append(f"  3. Can't resolve without the DLL. Fell back to ProductID (GUID): '{fallback_guid}'.")
                        elif os.name != 'nt':
                            lines.append("  3. Resolving this requires the Windows LoadLibraryExW/LoadStringW APIs, not available on this OS.")
                            lines.append(f"  4. Fell back to ProductID (GUID): '{fallback_guid}'.")
                        else:
                            resolved = resolve_wp_resource_dll(path, ref)
                            if resolved:
                                lines.append(f"  3. Loaded '{dll_name}' as a data file and read the string via LoadStringW: '{resolved}'.")
                                lines.append("  4. Used the resolved string.")
                            else:
                                lines.append(f"  3. Attempted to load '{dll_name}' and read the string via LoadStringW, but it failed.")
                                lines.append(f"  4. Fell back to ProductID (GUID): '{fallback_guid}'.")
                else:
                    lines.append("  1. No usable App/@Title attribute at all in the manifest.")
                    lines.append(f"  2. Fell back to ProductID (GUID): '{fallback_guid}'.")

            if not (m or m2 or m3):
                lines.append("\nNo recognized manifest (AppxManifest.xml / AppxBundleManifest.xml / WMAppManifest.xml) found.")
                lines.append("First entries in the archive:")
                for n in names[:20]:
                    lines.append(f"  - {n}")
    except zipfile.BadZipFile:
        lines.append("\nERROR: not a valid zip/package file.")
    except Exception as e:
        lines.append(f"\nERROR while inspecting: {e}")
    return "\n".join(lines)


def get_archive_listing(path):
    """Full file listing inside the package, for browsing what it actually contains."""
    try:
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
            lines = [f"{len(infos)} entr{'y' if len(infos) == 1 else 'ies'} in '{os.path.basename(path)}':", ""]
            for info in infos:
                if info.is_dir():
                    lines.append(f"{'<dir>':>12}   {info.filename}")
                else:
                    lines.append(f"{info.file_size:>10,} B   {info.filename}")
            return "\n".join(lines)
    except zipfile.BadZipFile:
        return "Not a valid zip/package file."
    except Exception as e:
        return f"Error reading archive: {e}"


def _dump_manifest_element(el, depth, lines):
    tag = local(el.tag)
    prefix = '  ' * depth
    if el.attrib:
        lines.append(f"{prefix}{tag}:")
        for key, value in el.attrib.items():
            lines.append(f"{prefix}  {local(key)}: {value}")
    else:
        lines.append(f"{prefix}{tag}")
    if el.text and el.text.strip():
        lines.append(f"{prefix}  (text): {el.text.strip()}")
    for child in el:
        _dump_manifest_element(child, depth + 1, lines)


def get_manifest_fields(path):
    """
    Every field (element + attribute) actually present in this package's
    manifest, indented to show structure - Title/Version plus whatever else
    the manifest happens to carry (Publisher, Capabilities, IconPath, etc.).
    """
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            m = (
                find_manifest(names, 'appxmanifest.xml')
                or find_manifest(names, 'appxmetadata/appxbundlemanifest.xml')
                or find_manifest(names, 'wmappmanifest.xml')
            )
            if not m:
                return "No recognized manifest found inside this file."
            root = ET.fromstring(z.read(m))
            lines = [f"Manifest: {m}", ""]
            _dump_manifest_element(root, 0, lines)
            return "\n".join(lines)
    except zipfile.BadZipFile:
        return "Not a valid zip/package file."
    except ET.ParseError:
        return "Could not parse manifest XML."
    except Exception as e:
        return f"Error reading manifest: {e}"


def build_new_name(path, info):
    ext = os.path.splitext(path)[1]
    name = sanitize(info.get('name')) if info.get('name') else None
    version = sanitize(info.get('version')) if info.get('version') else None
    if not name and not version:
        return None
    if name and version:
        return f"{name}_{version}{ext}"
    return f"{name or version}{ext}"


def unique_path(directory, filename):
    """Avoid clobbering an existing file by appending (1), (2), ..."""
    base, ext = os.path.splitext(filename)
    candidate = filename
    counter = 1
    while os.path.exists(os.path.join(directory, candidate)):
        candidate = f"{base} ({counter}){ext}"
        counter += 1
    return candidate


def decode_process_output(raw: bytes) -> str:
    """
    Decode subprocess output that may be UTF-16 (common for Windows console
    tools when their output is redirected/piped, as it is here) or a normal
    single-byte encoding. Naively decoding UTF-16 bytes as CP1252/UTF-8 turns
    "Microsoft" into "M\\x00i\\x00c\\x00..." which looks like garbage or just "M".
    """
    if not raw:
        return ''
    if raw[:2] in (b'\xff\xfe', b'\xfe\xff') or raw.count(b'\x00') > len(raw) // 4:
        try:
            return raw.decode('utf-16')
        except UnicodeError:
            pass
    for enc in ('utf-8', 'cp1252', 'latin-1'):
        try:
            return raw.decode(enc)
        except UnicodeError:
            continue
    return raw.decode('utf-8', errors='replace')


def test_makepri():
    """
    Try to locate makepri.exe and confirm it actually runs.
    Returns (success: bool, message: str).
    """
    path = find_makepri()
    if not path:
        return False, (
            "makepri.exe was not found.\n\n"
            "Checked: system PATH, and the default Windows SDK install folder\n"
            r"(C:\Program Files (x86)\Windows Kits\10\bin\*\x64\makepri.exe)." "\n\n"
            "If you installed the SDK somewhere else, it won't be auto-detected."
        )

    try:
        result = subprocess.run(
            [path, '/?'], capture_output=True, timeout=15, check=False
        )
        raw = result.stdout or result.stderr or b''
        output = decode_process_output(raw).strip()
        first_lines = '\n'.join(output.splitlines()[:6]) if output else '(no output captured)'
        return True, f"Found makepri.exe at:\n{path}\n\nIt ran successfully. Output preview:\n{first_lines}"
    except subprocess.TimeoutExpired:
        return False, f"Found makepri.exe at:\n{path}\n\nBut it did not respond within 15 seconds."
    except Exception as e:
        return False, f"Found makepri.exe at:\n{path}\n\nBut running it failed: {e}"


class RenamerApp:
    def __init__(self, root):
        self.root = root
        root.title(f"XAP / APPX Renamer v{SCRIPT_VERSION}")
        root.geometry("1150x640")
        root.minsize(950, 480)

        top = ttk.Frame(root, padding=8)
        top.pack(fill='x')

        ttk.Button(top, text="Select Folder…", command=self.pick_folder).pack(side='left', padx=4)
        ttk.Button(top, text="Select Files…", command=self.pick_files).pack(side='left', padx=4)
        self.clear_list_btn = ttk.Button(top, text="Clear List", command=self.clear_list, state='disabled')
        self.clear_list_btn.pack(side='left', padx=4)

        self.recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top, text="Include subfolders", variable=self.recursive_var, command=self.on_recursive_toggle
        ).pack(side='left', padx=10)

        self.strip_publisher_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top, text="Strip publisher prefix (Publisher.App → App)", variable=self.strip_publisher_var,
            command=self.recompute_names
        ).pack(side='left', padx=10)

        top2 = ttk.Frame(root, padding=(8, 0, 8, 4))
        top2.pack(fill='x')
        ttk.Button(top2, text="Select All", command=self.select_all).pack(side='left', padx=4)
        self.select_all_btn = top2.winfo_children()[-1]
        ttk.Button(top2, text="Unselect All", command=self.unselect_all).pack(side='left', padx=4)
        self.unselect_all_btn = top2.winfo_children()[-1]
        ttk.Button(top2, text="Lock Selected", command=self.lock_selected_rows).pack(side='left', padx=10)
        ttk.Button(top2, text="Unlock Selected", command=self.unlock_selected_rows).pack(side='left', padx=4)
        ttk.Button(top2, text="Test makepri.exe", command=self.on_test_makepri).pack(side='left', padx=10)
        ttk.Button(top2, text="Open Folder", command=self.open_folder_for_selection).pack(side='left', padx=10)
        self.selected_count_var = tk.StringVar(value="Selected: 0")
        ttk.Label(top2, textvariable=self.selected_count_var).pack(side='left', padx=(10, 0))
        self.selected_first_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top2, text="Selected first", variable=self.selected_first_var, command=self._refresh_view
        ).pack(side='left', padx=(10, 0))
        # Unlocked ("unblocked") rows before locked ones. If "Selected first"
        # is also on, checked rows still come first no matter what, then the
        # remaining unblocked rows, then the locked/unreadable ones.
        self.unblocked_first_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top2, text="Unblocked first", variable=self.unblocked_first_var, command=self._refresh_view
        ).pack(side='left', padx=(10, 0))

        path_row = ttk.Frame(root, padding=(8, 0, 8, 4))
        path_row.pack(fill='x')
        ttk.Label(path_row, text="Selected file:").pack(side='left')
        self.selected_path_var = tk.StringVar(value="(click a row below to see its full path)")
        ttk.Entry(path_row, textvariable=self.selected_path_var, state='readonly').pack(
            side='left', fill='x', expand=True, padx=4
        )
        ttk.Button(path_row, text="Open Folder", command=self._open_folder_for_selected_path).pack(side='left')

        hint = ttk.Label(
            root,
            text="Click row(s) in the table (Ctrl/Shift to multi-select), then Lock/Unlock. "
                 "Double-click 'New name' to edit it, or click ✕ to remove a row from the list. "
                 "Use the Filters row to narrow the list, or click a column header to sort by it.",
            justify='left', anchor='w'
        )
        hint.pack(fill='x', padx=8, pady=(0, 6))

        def _rewrap_hint(event, label=hint):
            label.config(wraplength=max(200, event.width - 16))
        root.bind('<Configure>', _rewrap_hint)

        search_bar = ttk.Frame(root, padding=(8, 0, 8, 4))
        search_bar.pack(fill='x')
        ttk.Label(search_bar, text="Search:").pack(side='left')
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *a: self._schedule_refresh())
        ttk.Entry(search_bar, textvariable=self.search_var).pack(side='left', fill='x', expand=True, padx=4)

        filter_bar = WrappingFrame(root, padding=(8, 0, 8, 4))
        filter_bar.pack(fill='x')
        filters_label_row = ttk.Frame(filter_bar)
        ttk.Label(filters_label_row, text="Filters:").pack(side='left')
        filter_bar.add_widget(filters_label_row)
        self.column_filter_vars = {}
        for col_id, label_text, entry_width in (
            ('original', 'Original file', 16),
            ('folder', 'Folder', 18),
            ('new_name', 'New name', 16),
            ('status', 'Status', 16),
        ):
            pair = ttk.Frame(filter_bar)
            ttk.Label(pair, text=f"{label_text}:").pack(side='left', padx=(0, 2))
            var = tk.StringVar()
            var.trace_add('write', lambda *a: self._schedule_refresh())
            ttk.Entry(pair, textvariable=var, width=entry_width).pack(side='left')
            self.column_filter_vars[col_id] = var
            filter_bar.add_widget(pair)
        clear_pair = ttk.Frame(filter_bar)
        ttk.Button(clear_pair, text="Clear Filters", command=self._clear_column_filters).pack(side='left')
        filter_bar.add_widget(clear_pair)

        self.sort_col = None
        self.sort_reverse = False
        self._sortable_headings = {'original': 'Original file', 'folder': 'Folder', 'new_name': 'New name', 'status': 'Status'}

        tree_frame = ttk.Frame(root)
        tree_frame.pack(fill='both', expand=True, padx=8, pady=4)
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        columns = ('select', 'original', 'folder', 'new_name', 'status', 'remove')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', selectmode='extended')
        self.tree.heading('select', text='✓/🔒')
        self.tree.heading('original', text='Original file', command=lambda: self._sort_by('original'))
        self.tree.heading('folder', text='Folder', command=lambda: self._sort_by('folder'))
        self.tree.heading('new_name', text='New name', command=lambda: self._sort_by('new_name'))
        self.tree.heading('status', text='Status', command=lambda: self._sort_by('status'))
        self.tree.heading('remove', text='✕')
        self.tree.column('select', width=60, minwidth=50, anchor='center', stretch=False)
        self.tree.column('original', width=220, minwidth=150)
        self.tree.column('folder', width=260, minwidth=120)
        self.tree.column('new_name', width=280, minwidth=150)
        self.tree.column('status', width=260, minwidth=150)
        self.tree.column('remove', width=40, minwidth=36, anchor='center', stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self.tree.bind('<Button-1>', self.on_click)
        self.tree.bind('<Double-1>', self.on_double_click)
        self.tree.bind('<Button-3>', self.on_right_click)
        self.tree.bind('<<TreeviewSelect>>', self._on_tree_selection_changed)
        self.icon_tooltip = IconTooltip(self.tree, self._get_row_icon_bytes)

        bottom = ttk.Frame(root, padding=8)
        bottom.pack(fill='x')
        ttk.Button(bottom, text="Rename checked", command=self.rename_selected).pack(side='right', padx=4)
        self.revert_btn = ttk.Button(bottom, text="Revert renamed files", command=self.revert_all_renames, state='disabled')
        self.revert_btn.pack(side='right', padx=4)
        self.status_var = tk.StringVar(value="No files loaded.")
        ttk.Label(bottom, textvariable=self.status_var).pack(side='left')
        # Shown only while files are being scanned/read in the background.
        self.load_frame = ttk.Frame(bottom)
        self.load_progress = ttk.Progressbar(self.load_frame, mode='determinate', length=180)
        self.load_progress.pack(side='left', padx=(0, 4))
        ttk.Button(self.load_frame, text="Cancel loading", command=self._cancel_load).pack(side='left')

        # row_id -> {'path':..., 'new_name':..., 'ok': bool}
        self.rows = {}
        self._edit_entry = None
        # row_ids currently locked: either already-correctly-named, or renamed
        # and not yet reverted (checkbox/edit disabled for these)
        self.renamed_rows = set()
        # row_id -> TRUE original path, set only the first time a row is
        # renamed and never overwritten afterwards - so Revert can always
        # restore the original filename even if the row was unlocked and
        # renamed again one or more times since.
        self.original_paths = {}
        # Folders picked via "Select Folder…", so toggling "Include
        # subfolders" can re-scan them with the new setting.
        self.loaded_folders = []

        # ---- Background loading ----
        # Reading each package (zip + manifest, and possibly a makepri.exe
        # run or a resource DLL load) used to happen right here on the Tk
        # thread, freezing the window until every file was done. Now worker
        # threads do the reading and hand results back through a queue that
        # the Tk thread drains in small time-boxed batches.
        self._job_queue = queue.Queue()      # (generation, path) for workers
        self._result_queue = queue.Queue()   # results back to the Tk thread
        self._load_generation = 0            # bumped on cancel/clear; stale results are ignored
        self._pending_paths = set()          # queued/being read, not yet in the table
        self._active_scans = 0               # folder scans still running
        self._load_stats = None
        self._poll_scheduled = False
        self._refresh_after_id = None
        self._last_refresh_time = 0.0
        worker_count = max(2, min(6, os.cpu_count() or 4))
        for _ in range(worker_count):
            threading.Thread(target=self._worker_loop, daemon=True).start()

    # ---------- File loading ----------

    @staticmethod
    def scan_folder(folder, recursive):
        """Return matching file paths under `folder`. Runs on a background
        thread, so it must not touch any Tk variables."""
        paths = []
        if recursive:
            for dirpath, _, filenames in os.walk(folder):
                for fn in filenames:
                    if fn.lower().endswith(SUPPORTED_EXTS):
                        paths.append(os.path.join(dirpath, fn))
        else:
            for fn in os.listdir(folder):
                full = os.path.join(folder, fn)
                if os.path.isfile(full) and fn.lower().endswith(SUPPORTED_EXTS):
                    paths.append(full)
        return paths

    def pick_folder(self):
        folder = filedialog.askdirectory(title="Select a folder containing .xap/.appx/.appxbundle files")
        if not folder:
            return
        if folder not in self.loaded_folders:
            self.loaded_folders.append(folder)
        self._start_scan([folder])

    def pick_files(self):
        paths = filedialog.askopenfilenames(
            title="Select .xap/.appx/.appxbundle files",
            filetypes=[("App packages", "*.xap *.appx *.appxbundle"), ("All files", "*.*")]
        )
        if paths:
            self.add_paths(list(paths))

    def on_recursive_toggle(self):
        # Re-scan every folder loaded via "Select Folder…" with the new
        # setting. This only ever ADDS newly-found files, consistent with
        # how loading works everywhere else - turning subfolders off never
        # removes rows that are already in the list.
        if not self.loaded_folders:
            return
        self._start_scan(list(self.loaded_folders))

    def update_clear_button_state(self):
        self.clear_list_btn.config(state='normal' if self.rows else 'disabled')

    def clear_list(self):
        if not self.rows and not self._is_loading():
            return
        if not messagebox.askyesno("Clear list", "Clear the list? This only clears the table - it does not touch any files on disk."):
            return
        # Delete by row id, not get_children() - a filter can leave rows
        # detached (hidden but not gone), and get_children() only returns
        # currently-attached items, which would leave hidden rows as
        # orphaned tree entries never cleaned up.
        self._cancel_load(quiet=True)
        self.tree.delete(*list(self.rows.keys()))
        self.rows.clear()
        self.renamed_rows.clear()
        self.original_paths.clear()
        self.loaded_folders.clear()
        self.revert_btn.config(state='disabled')
        self.update_bulk_buttons_state()
        self.update_clear_button_state()
        self.status_var.set("List cleared.")
        self.selected_count_var.set("Selected: 0")

    def _clear_column_filters(self):
        for var in self.column_filter_vars.values():
            var.set('')

    def _sort_by(self, col_id):
        if self.sort_col == col_id:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_col = col_id
            self.sort_reverse = False
        for cid, heading in self._sortable_headings.items():
            if cid == self.sort_col:
                arrow = ' ▼' if self.sort_reverse else ' ▲'
                self.tree.heading(cid, text=heading + arrow)
            else:
                self.tree.heading(cid, text=heading)
        self._refresh_view()

    def _schedule_refresh(self, delay=150):
        """Debounced refresh - typing in the search/filter boxes used to
        re-sort and re-order every row on each keystroke."""
        if self._refresh_after_id is not None:
            self.root.after_cancel(self._refresh_after_id)
        self._refresh_after_id = self.root.after(delay, self._run_scheduled_refresh)

    def _run_scheduled_refresh(self):
        self._refresh_after_id = None
        self._refresh_view()

    def _is_unblocked(self, row_id):
        """A row is 'unblocked' when it can still be checked/edited/renamed:
        readable, and not locked (auto-correct, renamed, or manual lock)."""
        row = self.rows.get(row_id)
        return bool(row and row.get('ok') and row_id not in self.renamed_rows)

    def _on_row_state_changed(self):
        """Call after a row's checked/locked state changes: re-orders the
        list if one of the grouping options depends on that state,
        otherwise just updates the counter."""
        if self.selected_first_var.get() or self.unblocked_first_var.get():
            self._refresh_view()
        else:
            self._update_selected_count()

    def _refresh_view(self):
        """Recomputes which rows are visible (per the global search box and
        active column filters) and their order (per current sort column/
        direction and the grouping checkboxes), without losing any row's
        underlying data - detach/move rather than delete/reinsert, since
        rows here live directly in the tree rather than a separate list
        re-rendered each time."""
        if self._refresh_after_id is not None:
            self.root.after_cancel(self._refresh_after_id)
            self._refresh_after_id = None
        self._last_refresh_time = time.perf_counter()

        query = self.search_var.get().strip().lower()
        queries = {col: var.get().strip().lower() for col, var in self.column_filter_vars.items() if var.get().strip()}
        value_cols = ('original', 'folder', 'new_name', 'status')

        # One Tcl call per row to fetch all its cells (instead of one per
        # cell, repeated for filtering, sorting and grouping).
        vals = {}
        visible_ids = []
        for row_id in self.rows:
            v = {c: str(x) for c, x in self.tree.set(row_id).items()}
            vals[row_id] = v
            if query and query not in ' '.join(v.get(c, '') for c in value_cols).lower():
                continue
            if any(needle not in v.get(col, '').lower() for col, needle in queries.items()):
                continue
            visible_ids.append(row_id)

        if self.sort_col:
            visible_ids.sort(key=lambda rid: vals[rid].get(self.sort_col, '').lower(), reverse=self.sort_reverse)

        selected_first = self.selected_first_var.get()
        unblocked_first = self.unblocked_first_var.get()
        if selected_first or unblocked_first:
            # Stable grouping pass - keeps the column-sort order established
            # above WITHIN each group:
            #   0 = checked (only when "Selected first" is on - always wins)
            #   1 = other unblocked rows (or simply "everything else")
            #   2 = locked / unreadable (only when "Unblocked first" is on)
            def group(rid):
                if selected_first and vals[rid].get('select') == '☑':
                    return 0
                if unblocked_first and not self._is_unblocked(rid):
                    return 2
                return 1
            visible_ids.sort(key=group)

        current = list(self.tree.get_children())
        if current != visible_ids:
            visible_set = set(visible_ids)
            to_detach = [rid for rid in current if rid not in visible_set]
            if to_detach:
                self.tree.detach(*to_detach)
            for index, row_id in enumerate(visible_ids):
                self.tree.move(row_id, '', index)
        self._update_selected_count(vals)

    def _update_selected_count(self, vals=None):
        if vals is not None:
            count = sum(1 for v in vals.values() if v.get('select') == '☑')
        else:
            count = sum(1 for row_id in self.rows if self.tree.set(row_id, 'select') == '☑')
        self.selected_count_var.set(f"Selected: {count}")

    # ---------- Background loading ----------

    def _worker_loop(self):
        """Runs on each worker thread: reads packages and posts results.
        Never touches Tk - only the Tk thread may do that."""
        while True:
            gen, path = self._job_queue.get()
            if gen != self._load_generation:
                continue  # cancelled before we got to it
            try:
                info = extract_info(path)
            except Exception as e:
                info = {'error': str(e)}
            self._result_queue.put(('info', gen, path, info))

    def _is_loading(self):
        return bool(self._pending_paths) or self._active_scans > 0

    def _new_load_stats(self):
        return {'total': 0, 'done': 0, 'ready': 0, 'already_ok': 0, 'error': 0, 'skipped': 0}

    def _start_scan(self, folders):
        """Walks folders on a background thread (os.walk over a large tree
        can take a while on its own), then feeds the results to add_paths."""
        if not self._is_loading():
            self._load_stats = self._new_load_stats()
        self._active_scans += 1
        gen = self._load_generation
        recursive = self.recursive_var.get()

        def job():
            paths = []
            for folder in folders:
                try:
                    if os.path.isdir(folder):
                        paths.extend(self.scan_folder(folder, recursive))
                except OSError:
                    pass
            self._result_queue.put(('scan', gen, paths))

        threading.Thread(target=job, daemon=True).start()
        self._set_loading_ui(True)
        self._update_load_progress()
        self._ensure_polling()

    def add_paths(self, paths):
        # Loading never clears what's already there - only "Clear List" or
        # closing the app does. Skip files already present in the table, or
        # already queued by a load that's still running.
        existing_paths = {row['path'] for row in self.rows.values()}
        seen = set()
        new_paths = []
        for p in paths:
            if p in existing_paths or p in self._pending_paths or p in seen:
                continue
            seen.add(p)
            new_paths.append(p)
        skipped = len(paths) - len(new_paths)

        if not self._is_loading():
            if not new_paths:
                if skipped:
                    self.status_var.set(f"No new files added - {skipped} were already in the list.")
                else:
                    self.status_var.set("No matching files found.")
                return
            self._load_stats = self._new_load_stats()

        self._load_stats['skipped'] += skipped
        if not new_paths:
            return
        self._load_stats['total'] += len(new_paths)
        gen = self._load_generation
        for p in new_paths:
            self._pending_paths.add(p)
            self._job_queue.put((gen, p))
        self._set_loading_ui(True)
        self._update_load_progress()
        self._ensure_polling()

    def _ensure_polling(self):
        if not self._poll_scheduled:
            self._poll_scheduled = True
            self.root.after(30, self._poll_results)

    def _poll_results(self):
        """Drains finished results in a time-boxed batch (~40 ms) so the
        window keeps repainting and responding between batches."""
        self._poll_scheduled = False
        if self._load_stats is None:
            return  # cancelled
        deadline = time.perf_counter() + 0.04
        inserted = False
        while time.perf_counter() < deadline:
            try:
                item = self._result_queue.get_nowait()
            except queue.Empty:
                break
            kind, gen = item[0], item[1]
            if gen != self._load_generation:
                continue  # left over from a cancelled load
            if kind == 'scan':
                self.add_paths(item[2])
                self._active_scans -= 1
            else:
                path, info = item[2], item[3]
                if path not in self._pending_paths:
                    continue
                self._pending_paths.discard(path)
                category = self._insert_row(path, info)
                self._load_stats['done'] += 1
                self._load_stats[category] += 1
                inserted = True

        if inserted:
            self.update_bulk_buttons_state()
            self.update_clear_button_state()

        if self._is_loading():
            self._update_load_progress()
            # Re-apply filters/sort/grouping now and then while loading, so
            # new rows don't just pile up unsorted at the bottom until the end.
            if inserted and time.perf_counter() - self._last_refresh_time > 1.0:
                self._refresh_view()
            self._ensure_polling()
        else:
            self._finish_load()

    def _update_load_progress(self):
        stats = self._load_stats
        if not stats:
            return
        total, done = stats['total'], stats['done']
        self.load_progress.configure(maximum=max(total, 1), value=done)
        msg = f"Loading… {done} / {total} file(s)"
        if self._active_scans:
            msg += " (still scanning folders…)"
        self.status_var.set(msg)

    def _set_loading_ui(self, loading):
        if loading:
            if not self.load_frame.winfo_ismapped():
                self.load_frame.pack(side='right', padx=8)
        else:
            self.load_frame.pack_forget()
            self.load_progress.configure(value=0)

    def _finish_load(self):
        stats = self._load_stats
        self._load_stats = None
        self._set_loading_ui(False)
        self.update_bulk_buttons_state()
        self.update_clear_button_state()
        self._refresh_view()
        if not stats:
            return
        if stats['total'] == 0:
            if stats['skipped']:
                msg = f"No new files added - {stats['skipped']} were already in the list."
            else:
                msg = "No matching files found."
        else:
            msg = (f"Added {stats['total']} file(s) — {stats['ready']} ready to rename, "
                   f"{stats['already_ok']} already correctly named")
            if stats['error']:
                msg += f", {stats['error']} could not be read"
            msg += "."
            if stats['skipped']:
                msg += f" ({stats['skipped']} already in the list, skipped.)"
        self.status_var.set(msg)

    def _cancel_load(self, quiet=False):
        if not self._is_loading():
            return
        self._load_generation += 1  # workers/scans skip or ignore everything older
        try:
            while True:
                self._job_queue.get_nowait()
        except queue.Empty:
            pass
        not_loaded = len(self._pending_paths)
        self._pending_paths.clear()
        self._active_scans = 0
        stats = self._load_stats
        self._load_stats = None
        self._set_loading_ui(False)
        if quiet:
            return
        self.update_bulk_buttons_state()
        self.update_clear_button_state()
        self._refresh_view()
        done = stats['done'] if stats else 0
        self.status_var.set(f"Loading cancelled — {done} file(s) added, {not_loaded} not loaded.")

    def _insert_row(self, path, info):
        """Adds one already-read file to the table. Tk thread only.
        Returns 'error', 'already_ok' or 'ready' for the load summary."""
        strip_publisher = self.strip_publisher_var.get()
        if 'error' in info:
            row_id = self.tree.insert('', 'end', values=('', os.path.basename(path), os.path.dirname(path), '(could not read)', info['error'], '✕'))
            self.rows[row_id] = {'path': path, 'new_name': None, 'ok': False}
            return 'error'

        # Remember the raw, unstripped fields so toggling "Strip publisher
        # prefix" later can recompute this row without re-opening the file.
        raw_name = info.get('name')
        identity_based = info.get('identity_based', False)
        version = info.get('version')
        resolved_from_pri = bool(info.get('resolved_from_pri'))
        resolved_from_dll = bool(info.get('resolved_from_dll'))

        stripped_publisher = False
        name_for_build = raw_name
        if strip_publisher and identity_based and raw_name:
            candidate = strip_publisher_prefix(raw_name)
            if candidate != raw_name:
                name_for_build = candidate
                stripped_publisher = True

        new_name = build_new_name(path, {'name': name_for_build, 'version': version})
        if not new_name:
            row_id = self.tree.insert('', 'end', values=('', os.path.basename(path), os.path.dirname(path), '(no name/version found)', 'Manifest had no usable Name/Version', '✕'))
            self.rows[row_id] = {'path': path, 'new_name': None, 'ok': False}
            return 'error'

        row_data = {
            'path': path, 'new_name': new_name, 'ok': True,
            'raw_name': raw_name, 'identity_based': identity_based,
            'version': version, 'resolved_from_pri': resolved_from_pri,
            'resolved_from_dll': resolved_from_dll,
            'manually_edited': False,
        }

        original_name = os.path.basename(path)
        if new_name == original_name:
            # Already has the right name - lock it, nothing to do here.
            row_id = self.tree.insert('', 'end', values=('🔒', original_name, os.path.dirname(path), new_name, 'Already correctly named', '✕'))
            row_data['lock_reason'] = 'auto_correct'
            self.rows[row_id] = row_data
            self.renamed_rows.add(row_id)
            return 'already_ok'

        if resolved_from_pri:
            status = 'Ready (name resolved via resources.pri)'
        elif resolved_from_dll:
            status = 'Ready (name resolved via resource DLL)'
        elif stripped_publisher:
            status = 'Ready (publisher prefix stripped)'
        else:
            status = 'Ready'
        row_id = self.tree.insert('', 'end', values=('☑', original_name, os.path.dirname(path), new_name, status, '✕'))
        self.rows[row_id] = row_data
        return 'ready'

    def recompute_names(self):
        """
        Re-derive proposed names for every row that isn't locked for a real
        reason (an actual rename, or an explicit manual lock) and hasn't been
        hand-edited - used when "Strip publisher prefix" is toggled.
        """
        if not self.rows:
            return

        strip_publisher = self.strip_publisher_var.get()
        changed = 0
        for row_id, row in list(self.rows.items()):
            if not row.get('ok') or row.get('manually_edited') or 'raw_name' not in row:
                continue
            lock_reason = row.get('lock_reason')
            if row_id in self.renamed_rows and lock_reason in ('renamed', 'manual'):
                continue  # don't touch actual renames or explicit manual locks

            raw_name = row['raw_name']
            name_for_build = raw_name
            stripped_publisher = False
            if strip_publisher and row.get('identity_based') and raw_name:
                candidate = strip_publisher_prefix(raw_name)
                if candidate != raw_name:
                    name_for_build = candidate
                    stripped_publisher = True

            new_name = build_new_name(row['path'], {'name': name_for_build, 'version': row.get('version')})
            if not new_name:
                continue

            original_name = os.path.basename(row['path'])
            if new_name == original_name:
                if row_id not in self.renamed_rows:
                    self.renamed_rows.add(row_id)
                    changed += 1
                row['lock_reason'] = 'auto_correct'
                self.tree.set(row_id, 'select', '🔒')
                self.tree.set(row_id, 'status', 'Already correctly named')
            else:
                if row_id in self.renamed_rows and lock_reason == 'auto_correct':
                    self.renamed_rows.discard(row_id)
                    row.pop('lock_reason', None)
                    self.tree.set(row_id, 'select', '☑')
                    changed += 1
                if row['new_name'] != new_name:
                    changed += 1
                if row.get('resolved_from_pri'):
                    status = 'Ready (name resolved via resources.pri)'
                elif row.get('resolved_from_dll'):
                    status = 'Ready (name resolved via resource DLL)'
                elif stripped_publisher:
                    status = 'Ready (publisher prefix stripped)'
                else:
                    status = 'Ready'
                self.tree.set(row_id, 'status', status)

            row['new_name'] = new_name
            self.tree.set(row_id, 'new_name', new_name)

        self.update_bulk_buttons_state()
        self.update_clear_button_state()
        if changed:
            self.status_var.set(f"Updated proposed names for {changed} file(s).")
        self._refresh_view()

    # ---------- Lock / unlock ----------

    def lock_selected_rows(self):
        selected = [r for r in self.tree.selection() if r in self.rows]
        if not selected:
            messagebox.showinfo("No rows selected", "Click one or more rows in the table first (Ctrl/Shift to multi-select), then Lock Selected.")
            return

        lockable = [r for r in selected if self.rows[r]['ok'] and r not in self.renamed_rows]
        if not lockable:
            messagebox.showinfo("Nothing to lock", "The selected row(s) are either unreadable or already locked.")
            return

        if not messagebox.askyesno(
            "Lock file(s)",
            f"Lock {len(lockable)} file(s)? Locked files can't be checked, edited, or renamed until you unlock them again. "
            "This does not touch anything on disk."
        ):
            return

        for row_id in lockable:
            self.renamed_rows.add(row_id)
            self.rows[row_id]['lock_reason'] = 'manual'
            self.tree.set(row_id, 'select', '🔒')
            self.tree.set(row_id, 'status', 'Locked manually')

        self.update_bulk_buttons_state()
        self.status_var.set(f"Locked {len(lockable)} file(s).")
        self._on_row_state_changed()

    def unlock_selected_rows(self):
        selected = [r for r in self.tree.selection() if r in self.rows]
        if not selected:
            messagebox.showinfo("No rows selected", "Click one or more rows in the table first (Ctrl/Shift to multi-select), then Unlock Selected.")
            return

        unlockable = [r for r in selected if r in self.renamed_rows]
        if not unlockable:
            messagebox.showinfo("Nothing to unlock", "None of the selected row(s) are currently locked.")
            return

        renamed_among_them = [r for r in unlockable if r in self.original_paths]
        warning = (
            f"Unlock {len(unlockable)} file(s)? They'll become checkable and editable again."
        )
        if renamed_among_them:
            warning += (
                f"\n\n{len(renamed_among_them)} of them were actually renamed on disk. Unlocking is safe - "
                "their original name stays remembered for 'Revert renamed files' - but if you rename them "
                "again before reverting, only their true original name will be restored, not this intermediate one."
            )
        if not messagebox.askyesno("Unlock file(s)", warning):
            return

        for row_id in unlockable:
            self.unlock_row(row_id)

        self.status_var.set(f"Unlocked {len(unlockable)} file(s).")
        self._on_row_state_changed()

    # ---------- Selection helpers ----------

    def update_bulk_buttons_state(self):
        # Enabled purely based on whether there is at least one row that can
        # still be checked/unchecked right now. Deliberately NOT tied to
        # whether a revert is pending elsewhere in the list - since loading
        # is additive, older renamed rows must never block newly added,
        # still-actionable ones.
        any_actionable = any(
            row['ok'] and row_id not in self.renamed_rows
            for row_id, row in self.rows.items()
        )
        state = 'normal' if any_actionable else 'disabled'
        self.select_all_btn.config(state=state)
        self.unselect_all_btn.config(state=state)

    def select_all(self):
        for row_id in self.tree.get_children():
            row = self.rows.get(row_id)
            if row and row['ok'] and row_id not in self.renamed_rows:
                self.tree.set(row_id, 'select', '☑')
        self._on_row_state_changed()

    def unselect_all(self):
        for row_id in self.tree.get_children():
            row = self.rows.get(row_id)
            if row and row['ok'] and row_id not in self.renamed_rows:
                self.tree.set(row_id, 'select', '☐')
        self._on_row_state_changed()

    def on_test_makepri(self):
        success, message = test_makepri()
        if success:
            messagebox.showinfo("makepri.exe check", message)
        else:
            messagebox.showwarning("makepri.exe check", message)

    # ---------- Interaction ----------

    def on_click(self, event):
        region = self.tree.identify('region', event.x, event.y)
        col = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if region != 'cell' or row_id not in self.rows:
            return
        if col == '#1':
            row = self.rows[row_id]
            if not row['ok'] or row_id in self.renamed_rows:
                return
            current = self.tree.set(row_id, 'select')
            self.tree.set(row_id, 'select', '☐' if current == '☑' else '☑')
            self._on_row_state_changed()
        elif col == '#6':
            self.remove_row(row_id)

    def _get_row_icon_bytes(self, row_id):
        row = self.rows.get(row_id)
        if not row or not row.get('path'):
            return None
        return find_manifest_icon_bytes(row['path'])

    def _on_tree_selection_changed(self, event=None):
        sel = self.tree.selection()
        if not sel:
            self.selected_path_var.set("(click a row below to see its full path)")
        elif len(sel) == 1:
            row = self.rows.get(sel[0])
            self.selected_path_var.set(row['path'] if row and row.get('path') else "(no path recorded for this row)")
        else:
            self.selected_path_var.set(f"({len(sel)} files selected - use \"Open Folder\" above to open them all)")

    def _open_folder_for_selected_path(self):
        path = self.selected_path_var.get()
        if not path or path.startswith('('):
            messagebox.showinfo("Nothing selected", "Click a row in the table first.")
            return
        folder = os.path.dirname(path)
        try:
            os.startfile(folder)
        except OSError as e:
            messagebox.showerror("Could not open folder", f"{folder}\n\n{e}")

    def open_folder_for_selection(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Nothing selected", "Select one or more rows first (click a row, or Ctrl/Shift-click for several).")
            return
        folders = set()
        for row_id in sel:
            row = self.rows.get(row_id)
            if row and row.get('path'):
                folders.add(os.path.dirname(row['path']))
        for folder in folders:
            try:
                os.startfile(folder)
            except OSError as e:
                messagebox.showerror("Could not open folder", f"{folder}\n\n{e}")

    def show_in_explorer(self, row_id):
        row = self.rows.get(row_id)
        if not row or not row.get('path'):
            return
        path = row['path']
        if not os.path.isfile(path):
            messagebox.showwarning("File not found", f"This file no longer exists at its recorded path:\n{path}")
            return
        try:
            # Uses row['path'], which is kept up to date after a rename
            # (not the original name), so this points at wherever the file
            # actually is right now. Goes through the actual Windows Shell
            # API rather than shelling out to explorer.exe's own
            # "/select," command-line syntax, which turned out to be
            # unreliable via subprocess regardless of how it was quoted.
            select_in_explorer(path)
        except OSError as e:
            # Fall back to just opening the containing folder - still
            # useful even without the specific file highlighted.
            try:
                os.startfile(os.path.dirname(path))
            except OSError:
                messagebox.showerror("Could not open Explorer", str(e))

    def on_right_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id not in self.rows:
            return
        self.tree.selection_set(row_id)

        menu = tk.Menu(self.tree, tearoff=0)
        menu.add_command(label="Show in Explorer (select this file)", command=lambda: self.show_in_explorer(row_id))
        menu.add_separator()
        menu.add_command(label="Show available languages…", command=lambda: self.show_languages(row_id))
        menu.add_command(label="Show raw debug info…", command=lambda: self.show_debug_info(row_id))
        menu.add_command(label="Show all manifest fields…", command=lambda: self.show_manifest_fields(row_id))
        menu.add_command(label="Show all archive contents…", command=lambda: self.show_archive_contents(row_id))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def show_text_dialog(self, title, text):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("700x450")
        win.minsize(400, 250)

        # Pack the button FIRST so it reserves its slice of the bottom edge -
        # packing it after an expand=True widget leaves it squeezed into
        # whatever sliver of space happens to remain, in the wrong spot.
        ttk.Button(win, text="Close", command=win.destroy).pack(side='bottom', pady=6)

        text_frame = ttk.Frame(win)
        text_frame.pack(side='top', fill='both', expand=True, padx=8, pady=(8, 0))
        text_widget = tk.Text(text_frame, wrap='word')
        vsb = ttk.Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
        text_widget.configure(yscrollcommand=vsb.set)
        text_widget.pack(side='left', fill='both', expand=True)
        vsb.pack(side='left', fill='y')
        text_widget.insert('1.0', text)
        text_widget.configure(state='disabled')

    def show_languages(self, row_id):
        row = self.rows.get(row_id)
        if not row:
            return
        display_name = os.path.basename(row['path'])
        results, message = get_language_candidates(row['path'])
        if results is None:
            self.show_text_dialog(f"Languages — {display_name}", message)
            return
        self.show_language_picker(row_id, display_name, results)

    def show_language_picker(self, row_id, display_name, results):
        win = tk.Toplevel(self.root)
        win.title(f"Languages — {display_name}")
        win.geometry("620x420")
        win.minsize(420, 300)

        ttk.Label(
            win,
            text=f"Found {len(results)} candidate(s). Select one and click 'Use this name' "
                 "to set it as this file's proposed name.",
            wraplength=580, justify='left'
        ).pack(fill='x', padx=8, pady=(8, 4))

        list_frame = ttk.Frame(win)
        list_frame.pack(fill='both', expand=True, padx=8, pady=4)
        listbox = tk.Listbox(list_frame, activestyle='dotbox')
        vsb = ttk.Scrollbar(list_frame, orient='vertical', command=listbox.yview)
        listbox.configure(yscrollcommand=vsb.set)
        for qualifiers, value in results:
            listbox.insert('end', f"[{qualifiers}]  {value}")
        listbox.pack(side='left', fill='both', expand=True)
        vsb.pack(side='left', fill='y')

        btn_frame = ttk.Frame(win)
        btn_frame.pack(side='bottom', fill='x', pady=6)

        def use_selected():
            sel = listbox.curselection()
            if not sel:
                messagebox.showinfo("No selection", "Select a candidate from the list first.", parent=win)
                return
            _, value = results[sel[0]]
            self.apply_manual_name(row_id, value)
            win.destroy()

        ttk.Button(btn_frame, text="Use this name", command=use_selected).pack(side='left', padx=8)
        ttk.Button(btn_frame, text="Close", command=win.destroy).pack(side='right', padx=8)

    def apply_manual_name(self, row_id, name):
        row = self.rows.get(row_id)
        if not row:
            return
        new_name = build_new_name(row['path'], {'name': name, 'version': row.get('version')})
        if not new_name:
            messagebox.showwarning("Could not apply", "Couldn't build a filename from that value.")
            return

        if row_id in self.renamed_rows:
            self.unlock_row(row_id)

        row['new_name'] = new_name
        row['manually_edited'] = True

        original_name = os.path.basename(row['path'])
        if new_name == original_name:
            self.renamed_rows.add(row_id)
            row['lock_reason'] = 'auto_correct'
            self.tree.set(row_id, 'select', '🔒')
            self.tree.set(row_id, 'status', 'Already correctly named')
        else:
            self.tree.set(row_id, 'select', '☑')
            self.tree.set(row_id, 'status', 'Ready (name chosen from language list)')
        self.tree.set(row_id, 'new_name', new_name)

        self.update_bulk_buttons_state()
        self.update_clear_button_state()
        self.status_var.set(f"Applied chosen name for '{original_name}'.")
        self._on_row_state_changed()

    def show_debug_info(self, row_id):
        row = self.rows.get(row_id)
        if not row:
            return
        display_name = os.path.basename(row['path'])
        info_text = get_debug_info(row['path'])
        self.show_text_dialog(f"Debug info — {display_name}", info_text)

    def show_archive_contents(self, row_id):
        row = self.rows.get(row_id)
        if not row:
            return
        display_name = os.path.basename(row['path'])
        listing_text = get_archive_listing(row['path'])
        self.show_text_dialog(f"Archive contents — {display_name}", listing_text)

    def show_manifest_fields(self, row_id):
        row = self.rows.get(row_id)
        if not row:
            return
        display_name = os.path.basename(row['path'])
        fields_text = get_manifest_fields(row['path'])
        self.show_text_dialog(f"Manifest fields — {display_name}", fields_text)

    def remove_row(self, row_id):
        if row_id not in self.rows:
            return
        row = self.rows[row_id]
        display_name = os.path.basename(row['path']) if row.get('path') else '(unknown file)'

        warning = f"Remove '{display_name}' from the list? This only removes it from the table - the file itself is left untouched on disk."
        if row_id in self.original_paths:
            warning += (
                "\n\nThis file was actually renamed and its original name is still remembered for "
                "'Revert renamed files'. Removing it here means you won't be able to revert it from "
                "this list anymore - you'd need to rename it back manually."
            )
        if not messagebox.askyesno("Remove file", warning):
            return

        self.tree.delete(row_id)
        del self.rows[row_id]
        self.renamed_rows.discard(row_id)
        self.original_paths.pop(row_id, None)

        self.revert_btn.config(state='normal' if self.original_paths else 'disabled')
        self.update_bulk_buttons_state()
        self.update_clear_button_state()
        self.status_var.set(f"Removed '{display_name}' from the list.")
        self._update_selected_count()

    def unlock_row(self, row_id):
        self.renamed_rows.discard(row_id)
        if row_id in self.rows:
            self.rows[row_id].pop('lock_reason', None)
        # Deliberately NOT touching self.original_paths here: if this row was
        # actually renamed before, we keep remembering its true original name
        # so "Revert renamed files" can still restore it later, even if it
        # gets edited and renamed again in the meantime.
        self.tree.set(row_id, 'select', '☑')
        self.tree.set(row_id, 'status', 'Unlocked for editing')
        self.update_bulk_buttons_state()
        # Callers re-count/re-order afterwards (once, not per row).

    def on_double_click(self, event):
        col = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if col != '#4' or row_id not in self.rows:
            return
        if row_id in self.renamed_rows:
            if not messagebox.askyesno(
                "Unlock file",
                "This file is locked. Unlock it so you can edit the name and rename it again?"
            ):
                return
            self.unlock_row(row_id)
            self._on_row_state_changed()
            self.tree.see(row_id)
        self.start_edit(row_id, col)

    def start_edit(self, row_id, col):
        if self._edit_entry is not None:
            try:
                if self._edit_entry.winfo_exists():
                    self._edit_entry.destroy()
            except tk.TclError:
                pass
            self._edit_entry = None

        bbox = self.tree.bbox(row_id, col)
        if not bbox:
            return
        x, y, w, h = bbox
        value = self.tree.set(row_id, 'new_name')
        entry = ttk.Entry(self.tree)
        entry.insert(0, value)
        entry.select_range(0, 'end')
        entry.focus()
        entry.place(x=x, y=y, width=w, height=h)

        def commit(_event=None):
            if not entry.winfo_exists():
                return
            new_val = sanitize(entry.get()) or value
            self.tree.set(row_id, 'new_name', new_val)
            if row_id in self.rows:
                self.rows[row_id]['new_name'] = new_val
                self.rows[row_id]['ok'] = True
                self.rows[row_id]['manually_edited'] = True
                self.tree.set(row_id, 'select', '☑')
                self.tree.set(row_id, 'status', 'Ready (edited)')
            entry.destroy()
            self._edit_entry = None
            self.update_bulk_buttons_state()
            self._on_row_state_changed()

        def cancel(_event=None):
            if entry.winfo_exists():
                entry.destroy()
            self._edit_entry = None

        entry.bind('<Return>', commit)
        entry.bind('<FocusOut>', commit)
        entry.bind('<Escape>', cancel)
        self._edit_entry = entry

    # ---------- Renaming ----------

    def rename_selected(self):
        to_rename = [
            (row_id, row) for row_id, row in self.rows.items()
            if row['ok'] and row_id not in self.renamed_rows and self.tree.set(row_id, 'select') == '☑'
        ]
        if not to_rename:
            messagebox.showinfo("Nothing to do", "No files are checked for renaming.")
            return

        if not messagebox.askyesno("Confirm rename", f"Rename {len(to_rename)} file(s) now? You can undo this afterwards with 'Revert renamed files'."):
            return

        renamed, failed = 0, 0
        for row_id, row in to_rename:
            old_path = row['path']
            directory = os.path.dirname(old_path)
            new_name = row['new_name']
            if os.path.basename(old_path) == new_name:
                self.tree.set(row_id, 'status', 'Already named correctly')
                self.tree.set(row_id, 'select', '🔒')
                self.renamed_rows.add(row_id)
                row['lock_reason'] = 'auto_correct'
                continue
            final_name = unique_path(directory, new_name)
            try:
                new_path = os.path.join(directory, final_name)
                os.rename(old_path, new_path)
                self.tree.set(row_id, 'status', f'Renamed → {final_name}')
                self.tree.set(row_id, 'select', '🔒')
                # Remember the TRUE original only the first time this row is
                # ever renamed - if it gets unlocked and renamed again later,
                # this still points at the very first name it had.
                if row_id not in self.original_paths:
                    self.original_paths[row_id] = old_path
                self.rows[row_id]['path'] = new_path
                self.rows[row_id]['lock_reason'] = 'renamed'
                self.renamed_rows.add(row_id)
                renamed += 1
            except OSError as e:
                self.tree.set(row_id, 'status', f'Failed: {e}')
                failed += 1

        if self.original_paths:
            self.revert_btn.config(state='normal')
        self.update_bulk_buttons_state()
        self.status_var.set(f"Done. Renamed {renamed} file(s), {failed} failed.")
        self._on_row_state_changed()
        self._on_tree_selection_changed()

    def revert_all_renames(self):
        if not self.original_paths:
            messagebox.showinfo("Nothing to revert", "There are no renamed files to revert.")
            return
        if not messagebox.askyesno("Confirm revert", f"Revert {len(self.original_paths)} file(s) back to their original names?"):
            return

        reverted, failed = 0, 0
        for row_id, original_path in list(self.original_paths.items()):
            current_path = self.rows.get(row_id, {}).get('path')
            try:
                if current_path and os.path.exists(current_path) and current_path != original_path:
                    os.rename(current_path, original_path)
                if row_id in self.rows:
                    self.rows[row_id]['path'] = original_path
                    self.rows[row_id].pop('lock_reason', None)
                    self.tree.set(row_id, 'status', 'Reverted')
                    self.tree.set(row_id, 'select', '☑')
                self.renamed_rows.discard(row_id)
                del self.original_paths[row_id]
                reverted += 1
            except OSError as e:
                if row_id in self.rows:
                    self.tree.set(row_id, 'status', f'Revert failed: {e}')
                failed += 1

        self.revert_btn.config(state='normal' if self.original_paths else 'disabled')
        self.update_bulk_buttons_state()
        self.status_var.set(f"Reverted {reverted} file(s), {failed} failed.")
        self._on_row_state_changed()
        self._on_tree_selection_changed()


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if 'vista' in style.theme_names():
            style.theme_use('vista')
        elif 'clam' in style.theme_names():
            style.theme_use('clam')
    except Exception:
        pass
    RenamerApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
