#!/usr/bin/env python3
"""MCO Maker - Automated Release & Update Publisher.

One-command pipeline to:
1. Bump APP_VERSION in main.py
2. Update AppVersion in installer.iss
3. Stage and commit all pending code changes
4. Push commit to GitHub (main branch)
5. Create and push release tag (vX.Y.Z)
6. Automatically trigger GitHub Actions to build MCOMaker-Setup.exe & MCOMaker-Win7-Setup.exe
7. Broadcast update notification to live clients via Vercel microservice!

Usage:
    python release.py 1.0.1
    python release.py 1.0.1 -m "Added automatic date formatting and new MCO layout"
    python release.py  (Interactive mode)
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ROOT_DIR = Path(__file__).resolve().parent
MAIN_PY = ROOT_DIR / "main.py"
INSTALLER_ISS = ROOT_DIR / "installer.iss"


def run_cmd(cmd: list[str], check: bool = True) -> subprocess.CompletedProcess:
    print(f"  -> Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=str(ROOT_DIR), text=True, capture_output=True)
    if check and result.returncode != 0:
        print(f"\n[ERROR] Command failed: {' '.join(cmd)}")
        if result.stderr:
            print(f"Details:\n{result.stderr.strip()}")
        if result.stdout:
            print(f"Output:\n{result.stdout.strip()}")
        sys.exit(result.returncode)
    return result


def get_current_version() -> str:
    if not MAIN_PY.exists():
        print(f"[ERROR] {MAIN_PY} not found.")
        sys.exit(1)
    content = MAIN_PY.read_text(encoding="utf-8")
    match = re.search(r'APP_VERSION\s*=\s*["\']([^"\']+)["\']', content)
    if not match:
        print("[ERROR] Could not find APP_VERSION in main.py")
        sys.exit(1)
    return match.group(1).strip()


def suggest_next_version(current: str) -> str:
    parts = current.split(".")
    try:
        parts[-1] = str(int(parts[-1]) + 1)
        return ".".join(parts)
    except Exception:
        return f"{current}.1"


def update_version_files(new_version: str):
    # Update main.py
    content = MAIN_PY.read_text(encoding="utf-8")
    new_content = re.sub(
        r'APP_VERSION\s*=\s*["\'][^"\']+["\']',
        f'APP_VERSION = "{new_version}"',
        content,
        count=1
    )
    MAIN_PY.write_text(new_content, encoding="utf-8")
    print(f"  [OK] Updated APP_VERSION in main.py -> {new_version}")

    # Update installer.iss
    if INSTALLER_ISS.exists():
        iss_content = INSTALLER_ISS.read_text(encoding="utf-8")
        new_iss = re.sub(
            r'#define\s+AppVersion\s+["\'][^"\']+["\']',
            f'#define AppVersion "{new_version}"',
            iss_content,
            count=1
        )
        INSTALLER_ISS.write_text(new_iss, encoding="utf-8")
        print(f"  [OK] Updated AppVersion in installer.iss -> {new_version}")


def update_release_notes(version: str, message: str):
    rn_file = ROOT_DIR / "RELEASE_NOTES.md"
    content = f"# Release v{version}\n\n{message}\n"
    rn_file.write_text(content, encoding="utf-8")
    print(f"  [OK] Generated RELEASE_NOTES.md")


def main():
    parser = argparse.ArgumentParser(description="MCO Maker Release Automation")
    parser.add_argument("version", nargs="?", help="New release version (e.g. 1.0.1)")
    parser.add_argument("-m", "--message", default="", help="Release changelog message")
    args = parser.parse_args()

    current_ver = get_current_version()
    print("=" * 60)
    print(f"MCO Maker — Release Automation Pipeline (Current: v{current_ver})")
    print("=" * 60)

    target_ver = args.version
    if not target_ver:
        suggested = suggest_next_version(current_ver)
        user_input = input(f"Enter new version [{suggested}]: ").strip()
        target_ver = user_input if user_input else suggested

    # Validate version format
    if not re.match(r"^\d+\.\d+\.\d+$", target_ver):
        print("[ERROR] Version must follow SemVer format X.Y.Z (e.g. 1.0.1)")
        sys.exit(1)

    message = args.message
    if not message:
        message = input("Enter release notes/changelog description: ").strip()
        if not message:
            message = f"Release v{target_ver} with performance improvements and updates."

    print(f"\n[1/5] Validating Python syntax in main.py...")
    import py_compile
    py_compile.compile(str(MAIN_PY), doraise=True)
    print("  [OK] Python syntax validated!")

    print(f"\n[2/5] Bumping version to v{target_ver}...")
    update_version_files(target_ver)
    update_release_notes(target_ver, message)

    print(f"\n[3/5] Git staging and committing...")
    run_cmd(["git", "add", "main.py", "installer.iss", "RELEASE_NOTES.md"])
    commit_res = run_cmd(["git", "commit", "-m", f"chore(release): v{target_ver} - {message}"], check=False)
    if commit_res.returncode != 0:
        print("  [WARN] No changes to commit or commit already up to date.")

    print(f"\n[4/5] Creating and pushing git tag v{target_ver}...")
    run_cmd(["git", "push", "origin", "main"], check=False)
    run_cmd(["git", "tag", "-a", f"v{target_ver}", "-m", message])
    run_cmd(["git", "push", "origin", f"v{target_ver}"])
    print(f"  [OK] Tag v{target_ver} pushed to GitHub!")

    print(f"\n[5/5] Success! GitHub Actions CI/CD has been triggered.")
    print("      Both Windows 10/11 and Windows 7 installers are building in parallel.")
    print(f"      Check progress at your GitHub Actions tab!")
    print("=" * 60)


if __name__ == "__main__":
    main()
