#!/usr/bin/env python3
"""MCO Maker - Admin Serial Key Generator.

Generates new subscription serial keys via the Vercel Express API
and automatically copies them to the clipboard to send directly to buyers.

Supports both a Dark-Themed Tkinter GUI and CLI command-line mode.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import requests

DEFAULT_API_URL = os.environ.get("LICENSE_API_URL", "https://mcomaker.vercel.app").rstrip("/")
DEFAULT_ADMIN_SECRET = os.environ.get("ADMIN_SECRET", "")


def get_admin_config_path() -> Path:
    """Return local path to admin_config.json in AppData."""
    base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    app_dir = base / "MCOMaker"
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir / "admin_config.json"


def load_admin_config() -> dict:
    p = get_admin_config_path()
    if p.is_file():
        try:
            return json.loads(p.read_text(encoding="utf-8-sig"))
        except Exception:
            pass
    return {}


def save_admin_config(data: dict):
    p = get_admin_config_path()
    try:
        p.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass


# VS Code Dark+ Palette
BG_DARK = "#1e1e1e"
CARD_BG = "#252526"
BORDER_COLOR = "#383838"
TEXT_WHITE = "#ffffff"
TEXT_MAIN = "#cccccc"
TEXT_MUTED = "#8b949e"
ACCENT_GREEN = "#238636"
ACCENT_GREEN_HOVER = "#2ea043"
INPUT_BG = "#2d2d2d"
BTN_SEC_BG = "#3a3d41"
BTN_SEC_HOVER = "#45494e"
UI_FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
MONO_FONT = "Consolas" if sys.platform == "win32" else "DejaVu Sans Mono"

PLANS = [
    ("6_months", "6 Months  —  Rs. 3,000  (180 Days, 1 Device)", 180, 1),
    ("1_year", "1 Year  —  Rs. 5,000  (365 Days, 1 Device)  [Recommended]", 365, 1),
]


def request_generate_key(api_url: str, admin_secret: str, plan_type: str, max_devices: int = 1) -> dict:
    url = f"{api_url.rstrip('/')}/api/admin/generate"
    headers = {
        "Content-Type": "application/json",
        "x-admin-secret": admin_secret
    }
    payload = {
        "admin_secret": admin_secret,
        "plan_type": plan_type,
        "max_devices": max_devices,
        "count": 1
    }
    resp = requests.post(url, json=payload, headers=headers, timeout=12)
    resp.raise_for_status()
    data = resp.json()
    if not data.get("success"):
        raise RuntimeError(data.get("error", "Failed to generate serial key."))
    if "keys" in data and data["keys"]:
        return data["keys"][0]
    if "serial_key" in data:
        return data
    raise RuntimeError(data.get("error", "No serial key returned by license server."))


class AdminKeyGenApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MCO Maker — Admin Key Generator")
        self.geometry("540x580")
        self.minsize(500, 520)
        self.configure(bg=BG_DARK)

        # Apply dark title bar on Windows
        if sys.platform == "win32":
            try:
                import ctypes
                self.update_idletasks()
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
                value = ctypes.c_int(1)
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))
            except Exception:
                pass

        self.cfg = load_admin_config()
        self.api_url_var = tk.StringVar(value=self.cfg.get("api_url", DEFAULT_API_URL))
        self.secret_var = tk.StringVar(value=self.cfg.get("admin_secret", DEFAULT_ADMIN_SECRET))
        self.plan_var = tk.StringVar(value="1_year")
        self.devices_var = tk.IntVar(value=1)
        self.generated_key_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Ready to generate license keys.")

        self._build_ui()

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=CARD_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        hdr.pack(fill="x", padx=16, pady=(16, 12))
        
        tk.Label(
            hdr, text="🔑 MCO Maker — Admin Key Generator",
            font=(UI_FONT, 13, "bold"), fg=TEXT_WHITE, bg=CARD_BG
        ).pack(anchor="w", padx=14, pady=(10, 2))
        
        tk.Label(
            hdr, text="Direct Supabase Provisioning via Vercel Microservice",
            font=(UI_FONT, 9), fg=TEXT_MUTED, bg=CARD_BG
        ).pack(anchor="w", padx=14, pady=(0, 10))

        # Main Card
        card = tk.Frame(self, bg=CARD_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        card.pack(fill="both", expand=True, padx=16, pady=0)

        # Plan Selection
        tk.Label(
            card, text="Select Subscription Plan:",
            font=(UI_FONT, 10, "bold"), fg=TEXT_WHITE, bg=CARD_BG
        ).pack(anchor="w", padx=16, pady=(12, 6))

        for plan_id, label, days, dev in PLANS:
            rb = tk.Radiobutton(
                card, text=label, variable=self.plan_var, value=plan_id,
                command=lambda d=dev: self.devices_var.set(d),
                font=(UI_FONT, 9), fg=TEXT_MAIN, bg=CARD_BG,
                activebackground=CARD_BG, activeforeground=TEXT_WHITE,
                selectcolor=INPUT_BG, highlightthickness=0
            )
            rb.pack(anchor="w", padx=24, pady=2)

        # Device Override
        dev_frame = tk.Frame(card, bg=CARD_BG)
        dev_frame.pack(fill="x", padx=16, pady=(10, 6))
        tk.Label(
            dev_frame, text="Allowed Devices:",
            font=(UI_FONT, 9), fg=TEXT_MAIN, bg=CARD_BG
        ).pack(side="left")
        
        sp = tk.Spinbox(
            dev_frame, from_=1, to=10, textvariable=self.devices_var, width=5,
            bg=INPUT_BG, fg=TEXT_WHITE, insertbackground=TEXT_WHITE,
            buttonbackground=BTN_SEC_BG, relief="flat", font=(MONO_FONT, 10)
        )
        sp.pack(side="left", padx=10)

        # Generate Button
        self.btn_gen = tk.Button(
            card, text="🔑 Generate Serial Key & Copy",
            font=(UI_FONT, 10, "bold"), fg=TEXT_WHITE, bg=ACCENT_GREEN,
            activebackground=ACCENT_GREEN_HOVER, activeforeground=TEXT_WHITE,
            relief="flat", cursor="hand2", padx=14, pady=8,
            command=self._generate_key
        )
        self.btn_gen.pack(fill="x", padx=16, pady=(14, 8))

        # Result Display Box
        res_frame = tk.Frame(card, bg=INPUT_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        res_frame.pack(fill="x", padx=16, pady=6)

        self.key_entry = tk.Entry(
            res_frame, textvariable=self.generated_key_var, font=(MONO_FONT, 13, "bold"),
            fg="#4ec9b0", bg=INPUT_BG, relief="flat", justify="center"
        )
        self.key_entry.pack(fill="x", padx=10, pady=10)

        # Status Label
        tk.Label(
            card, textvariable=self.status_var,
            font=(UI_FONT, 9), fg=TEXT_MUTED, bg=CARD_BG
        ).pack(anchor="w", padx=16, pady=(4, 12))

        # Bottom Bar: Settings & Quit
        bottom = tk.Frame(self, bg=BG_DARK)
        bottom.pack(fill="x", padx=16, pady=(10, 16))

        tk.Button(
            bottom, text="⚙ Settings", font=(UI_FONT, 9), fg=TEXT_MAIN, bg=BTN_SEC_BG,
            activebackground=BTN_SEC_HOVER, activeforeground=TEXT_WHITE,
            relief="flat", padx=12, pady=4, cursor="hand2",
            command=self._open_settings
        ).pack(side="left")

        tk.Button(
            bottom, text="Exit", font=(UI_FONT, 9), fg=TEXT_MAIN, bg=BTN_SEC_BG,
            activebackground=BTN_SEC_HOVER, activeforeground=TEXT_WHITE,
            relief="flat", padx=12, pady=4, cursor="hand2",
            command=self.destroy
        ).pack(side="right")

    def _generate_key(self):
        url = self.api_url_var.get().strip()
        secret = self.secret_var.get().strip()
        plan = self.plan_var.get()
        dev = self.devices_var.get()

        if not secret:
            self._open_settings(initial_message="Please enter your ADMIN_SECRET before generating keys.")
            return

        self.status_var.set("Connecting to license server...")
        self.btn_gen.configure(state="disabled")
        self.update_idletasks()

        try:
            key_data = request_generate_key(url, secret, plan, dev)
            serial = key_data["serial_key"]
            self.generated_key_var.set(serial)
            
            # Copy to clipboard
            self.clipboard_clear()
            self.clipboard_append(serial)
            self.update()

            self.status_var.set(f"✔ Key generated & copied to clipboard! (Expires in {key_data['duration_days']} days)")
            messagebox.showinfo("Key Generated", f"Serial Key Created Successfully:\n\n{serial}\n\nCopied to your clipboard!")
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            messagebox.showerror("Generation Error", f"Failed to generate serial key:\n\n{e}")
        finally:
            self.btn_gen.configure(state="normal")

    def _open_settings(self, initial_message=None):
        dlg = tk.Toplevel(self)
        dlg.title("Admin Settings")
        dlg.geometry("460x280")
        dlg.configure(bg=CARD_BG)
        dlg.transient(self)
        dlg.grab_set()

        if initial_message:
            tk.Label(dlg, text=initial_message, font=(UI_FONT, 9, "bold"), fg="#e5c07b", bg=CARD_BG, wraplength=420).pack(padx=16, pady=(12, 4))

        tk.Label(dlg, text="License Server API URL:", font=(UI_FONT, 9, "bold"), fg=TEXT_WHITE, bg=CARD_BG).pack(anchor="w", padx=16, pady=(12, 2))
        u_ent = tk.Entry(dlg, textvariable=self.api_url_var, font=(MONO_FONT, 9), bg=INPUT_BG, fg=TEXT_WHITE, insertbackground=TEXT_WHITE, relief="flat")
        u_ent.pack(fill="x", padx=16, pady=2)

        tk.Label(dlg, text="Admin Secret Key:", font=(UI_FONT, 9, "bold"), fg=TEXT_WHITE, bg=CARD_BG).pack(anchor="w", padx=16, pady=(10, 2))
        s_ent = tk.Entry(dlg, textvariable=self.secret_var, font=(MONO_FONT, 9), show="*", bg=INPUT_BG, fg=TEXT_WHITE, insertbackground=TEXT_WHITE, relief="flat")
        s_ent.pack(fill="x", padx=16, pady=2)

        def save():
            self.cfg["api_url"] = self.api_url_var.get().strip()
            self.cfg["admin_secret"] = self.secret_var.get().strip()
            save_admin_config(self.cfg)
            dlg.destroy()

        tk.Button(dlg, text="Save Settings", font=(UI_FONT, 10, "bold"), fg=TEXT_WHITE, bg=ACCENT_GREEN, activebackground=ACCENT_GREEN_HOVER, activeforeground=TEXT_WHITE, relief="flat", padx=12, pady=6, command=save).pack(anchor="e", padx=16, pady=16)


def main():
    parser = argparse.ArgumentParser(description="MCO Maker Admin Key Generator")
    parser.add_argument("--plan", choices=["6_months", "1_year"], help="Subscription plan type")
    parser.add_argument("--devices", type=int, default=1, help="Max devices allowed")
    parser.add_argument("--url", default=DEFAULT_API_URL, help="License API URL")
    parser.add_argument("--secret", default=DEFAULT_ADMIN_SECRET, help="Admin secret password")
    args = parser.parse_args()

    if args.plan:
        if not args.secret:
            print("Error: --secret must be provided in CLI mode.")
            sys.exit(1)
        res = request_generate_key(args.url, args.secret, args.plan, args.devices)
        print(f"Generated Key: {res['serial_key']}")
        print(f"Plan: {res['plan_type']} ({res['duration_days']} days, {res['max_devices']} devices)")
    else:
        app = AdminKeyGenApp()
        app.mainloop()


if __name__ == "__main__":
    main()
