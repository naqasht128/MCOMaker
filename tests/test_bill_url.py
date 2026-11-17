import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import main
from main import (
    DEFAULT_BILL_URL,
    load_app_config,
    save_app_config,
    EditBillUrlDialog,
    BillUrlDialog,
    MCOMakerApp,
)


class TestBillUrlFeature(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.mock_appdata = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_default_bill_url(self):
        self.assertEqual(DEFAULT_BILL_URL, "https://bill.pitc.com.pk/pescobill")

    def test_app_config_persistence(self):
        with patch("main.get_appdata_dir", return_value=self.mock_appdata):
            # Initially no config
            cfg = load_app_config()
            self.assertEqual(cfg, {})

            # Save config
            test_data = {"bill_url": "https://bill.pitc.com.pk/mepcobill"}
            save_app_config(test_data)

            # Reload config
            loaded = load_app_config()
            self.assertEqual(loaded.get("bill_url"), "https://bill.pitc.com.pk/mepcobill")

    def test_edit_bill_url_dialog_save(self):
        with patch.object(main.LicenseManager, "is_locally_active", return_value=True), \
             patch.object(main.MCOMakerApp, "_bg_heartbeat"), \
             patch.object(main.MCOMakerApp, "_bg_update_check"):
            app = MCOMakerApp()
            app.withdraw()

            saved_urls = []
            dlg = EditBillUrlDialog(app, "https://bill.pitc.com.pk/pescobill", on_save=saved_urls.append)
            dlg.withdraw()

            # Initial value
            self.assertEqual(dlg.url_var.get(), "https://bill.pitc.com.pk/pescobill")

            # Change URL and save
            dlg.url_var.set("https://bill.pitc.com.pk/mepcobill")
            dlg._save()
            self.assertEqual(saved_urls, ["https://bill.pitc.com.pk/mepcobill"])

            app.destroy()

    def test_edit_bill_url_dialog_validation(self):
        with patch.object(main.LicenseManager, "is_locally_active", return_value=True), \
             patch.object(main.MCOMakerApp, "_bg_heartbeat"), \
             patch.object(main.MCOMakerApp, "_bg_update_check"):
            app = MCOMakerApp()
            app.withdraw()

            saved = []
            dlg = EditBillUrlDialog(app, DEFAULT_BILL_URL, on_save=saved.append)
            dlg.withdraw()

            with patch("tkinter.messagebox.showerror") as mock_err:
                # Test empty URL
                dlg.url_var.set("   ")
                dlg._save()
                mock_err.assert_called_with("Invalid URL", "Bill portal URL cannot be empty.", parent=dlg)
                self.assertEqual(len(saved), 0)

                # Test invalid scheme
                dlg.url_var.set("ftp://invalid-url.com")
                dlg._save()
                mock_err.assert_called_with("Invalid URL", "URL must start with http:// or https://", parent=dlg)
                self.assertEqual(len(saved), 0)

            app.destroy()

    def test_app_bill_url_changed_callback(self):
        with patch.object(main.LicenseManager, "is_locally_active", return_value=True), \
             patch.object(main.MCOMakerApp, "_bg_heartbeat"), \
             patch.object(main.MCOMakerApp, "_bg_update_check"), \
             patch("main.get_appdata_dir", return_value=self.mock_appdata):
            app = MCOMakerApp()
            app.withdraw()

            self.assertEqual(app.bill_url, DEFAULT_BILL_URL)

            # Change bill URL
            new_url = "https://bill.pitc.com.pk/fescobill"
            app._on_bill_url_changed(new_url)

            self.assertEqual(app.bill_url, new_url)
            self.assertEqual(app.app_config.get("bill_url"), new_url)

            # Check config file written
            saved_cfg = load_app_config()
            self.assertEqual(saved_cfg.get("bill_url"), new_url)

            app.destroy()

    def test_scrape_worker_uses_configured_bill_url(self):
        with patch.object(main.LicenseManager, "is_locally_active", return_value=True), \
             patch.object(main.MCOMakerApp, "_bg_heartbeat"), \
             patch.object(main.MCOMakerApp, "_bg_update_check"):
            app = MCOMakerApp()
            app.withdraw()

            app.bill_url = "https://bill.pitc.com.pk/iescobill"
            app.records = [{"reference": "12345678901234"}]

            with patch("main.scrape_pesco_bill") as mock_scrape, \
                 patch.object(app, "_wait_safety_delay"):
                mock_scrape.return_value = {
                    "account_no": "12345678901234",
                    "status": "Normal",
                    "is_defective": False,
                    "is_replaced": False,
                    "captcha_code": "1234",
                    "captcha_attempts": 1
                }
                app._scraping_worker(0)

                mock_scrape.assert_called()
                called_kwargs = mock_scrape.call_args[1]
                self.assertEqual(called_kwargs.get("base_url"), "https://bill.pitc.com.pk/iescobill")

            app.destroy()


if __name__ == "__main__":
    unittest.main()
