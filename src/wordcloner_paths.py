"""Project-root storage and safe migration from earlier WordCloner folders.

Current data is stored directly beside the application:
    settings.json
    templates.json
    Words/

Earlier locations are copied into the project root without overwriting files.
The originals are retained until the user removes them manually.
"""
from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path


def _application_root() -> Path:
    """Use the app directory, unless the user explicitly selects a data root.

    ``WORDCLONER_HOME`` is useful for a regular/system package installation
    whose site-packages directory is read-only. Without it, source/editable
    installs continue to keep settings and Words beside the application files.
    """
    override = os.environ.get("WORDCLONER_HOME")
    if override and override.strip():
        return Path(override).expanduser().resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_ROOT = _application_root()
# Kept for compatibility with the existing file-picker start-directory code.
DOCUMENTS_DIR = APP_ROOT

# Data now lives directly at the project root.
WORDCLONER_DIR = APP_ROOT
WORDS_DIR = APP_ROOT / "Words"
TEMPLATES_JSON = APP_ROOT / "templates.json"
SETTINGS_JSON = APP_ROOT / "settings.json"

# Both locations may contain data from earlier versions.
PREVIOUS_WORDCLONER_DIR = APP_ROOT / "WordCloner"
LEGACY_WORDCLONER_DIR = Path.home() / "Documents" / "WordCloner"


def _copy_missing_tree(source: Path, destination: Path) -> bool:
    """Copy missing files recursively; never overwrite an existing target."""
    copied_any = False
    if not source.is_dir():
        return False

    destination.mkdir(parents=True, exist_ok=True)
    try:
        children = list(source.iterdir())
    except OSError as exc:
        print(f"[migration] Cannot read {source}: {exc}")
        return False

    for child in children:
        target = destination / child.name
        try:
            if child.is_dir():
                if target.exists() and not target.is_dir():
                    print(f"[migration] Kept existing file; skipped folder: {target}")
                    continue
                copied_any = _copy_missing_tree(child, target) or copied_any
            elif not target.exists():
                shutil.copy2(str(child), str(target))
                copied_any = True
        except Exception as exc:
            print(f"[migration] Could not copy {child} to {target}: {exc}")
    return copied_any


def _relocated_address(address):
    """Map old absolute template paths to the new root-level location."""
    if not isinstance(address, str) or not address.strip():
        return address

    try:
        old_path = Path(address).expanduser().resolve()
    except (OSError, RuntimeError):
        return address

    roots = (PREVIOUS_WORDCLONER_DIR, LEGACY_WORDCLONER_DIR)
    for old_root in roots:
        try:
            old_root_resolved = old_root.resolve()
            relative = old_path.relative_to(old_root_resolved)
        except (OSError, ValueError, RuntimeError):
            continue
        return str(APP_ROOT / relative)

    return address


def _rewrite_template_addresses() -> bool:
    """Merge template records from both legacy folders and relocate paths."""
    templates = []
    changed = False

    if TEMPLATES_JSON.is_file():
        try:
            with TEMPLATES_JSON.open("r", encoding="utf-8") as stream:
                current = json.load(stream)
            if isinstance(current, list):
                templates = [item for item in current if isinstance(item, dict)]
            else:
                print(f"[migration] Invalid template list in {TEMPLATES_JSON}")
                return False
        except Exception as exc:
            print(f"[migration] Could not read {TEMPLATES_JSON}: {exc}")
            return False

    def identity(template):
        name = (template.get("name") or "").strip()
        if name:
            return ("name", name.casefold())
        address = template.get("address")
        return ("address", str(address or "").casefold())

    identities = set()
    for template in templates:
        old_address = template.get("address")
        new_address = _relocated_address(old_address)
        if new_address != old_address:
            template["address"] = new_address
            changed = True
        identities.add(identity(template))

    # Merge any templates not already present. Existing root-level records win.
    for source_root in (PREVIOUS_WORDCLONER_DIR, LEGACY_WORDCLONER_DIR):
        source_file = source_root / "templates.json"
        try:
            if not source_file.is_file() or source_file.resolve() == TEMPLATES_JSON.resolve():
                continue
            with source_file.open("r", encoding="utf-8") as stream:
                source_templates = json.load(stream)
            if not isinstance(source_templates, list):
                continue
            for source_template in source_templates:
                if not isinstance(source_template, dict):
                    continue
                template = dict(source_template)
                old_address = template.get("address")
                new_address = _relocated_address(old_address)
                if new_address != old_address:
                    template["address"] = new_address
                key = identity(template)
                if key in identities:
                    continue
                templates.append(template)
                identities.add(key)
                changed = True
        except Exception as exc:
            print(f"[migration] Could not merge templates from {source_file}: {exc}")

    if changed:
        try:
            with TEMPLATES_JSON.open("w", encoding="utf-8") as stream:
                json.dump(templates, stream, ensure_ascii=False, indent=4)
        except Exception as exc:
            print(f"[migration] Could not write {TEMPLATES_JSON}: {exc}")
            return False
    return changed


def migrate_legacy_data() -> bool:
    """Bring old WordCloner contents directly into the project root.

    Migration sources are processed in order: the prior project-local
    ``WordCloner/`` folder first, then ``Documents/WordCloner``. Existing
    root-level files always win; missing files are copied. Source folders are
    deliberately retained as backups. Old absolute addresses in
    ``templates.json`` are then rewritten to ``<project-root>/Words/...``.
    """
    migrated = False
    sources = (PREVIOUS_WORDCLONER_DIR, LEGACY_WORDCLONER_DIR)
    root_resolved = APP_ROOT.resolve()

    for source in sources:
        try:
            if not source.is_dir() or source.resolve() == root_resolved:
                continue
            if _copy_missing_tree(source, APP_ROOT):
                migrated = True
                print(f"[migration] Copied missing data from: {source}")
        except Exception as exc:
            print(f"[migration] Could not migrate {source}: {exc}")

    migrated = _rewrite_template_addresses() or migrated
    if migrated:
        print(f"[migration] Active data location: {APP_ROOT}")
        print("[migration] Original WordCloner folders were left untouched.")
    return migrated
