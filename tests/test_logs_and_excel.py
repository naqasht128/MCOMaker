import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
import main
from main import MCOMakerApp


class TestLogsAndExcel(unittest.TestCase):
    def test_save_logs_to_excel_structure_and_formatting(self):
        output_file = Path("test_output_logs.xlsx")
        if output_file.exists():
            output_file.unlink()

        # Mock app without opening GUI window
        with patch.object(main.LicenseManager, "is_locally_active", return_value=True), \
             patch.object(main.MCOMakerApp, "_bg_heartbeat"), \
             patch.object(main.MCOMakerApp, "_bg_update_check"):
            app = MCOMakerApp()
            app.withdraw()

            # Set sample records
            app.records = [
                {
                    "reference": "19269310053300",
                    "name": "HAJI MUHAMMAD",
                    "address": "VILLAGE NOWSHERA",
                    "tariff": "A-1a(01)",
                    "sub_division": "NOWSHERA CANTT-II",
                    "status": "Replaced",
                    "is_replaced": True,
                    "_fetched": True,
                    "_log_status": "Fetched",
                    "_log_remarks": "Already replaced in previous month",
                    "msn": "MSN001",
                    "sim_no": "03001234567",
                    "sim_id": "899201234567890",
                    "date": "12-Nov-2024",
                    "telco": "Jazz",
                    "sco_no": "SCO-101",
                    "present_reading": "1200",
                    "prev_reading": "1100",
                    "units": "100"
                },
                {
                    "reference": "03262320230213",
                    "name": "AHMAD ALI",
                    "address": "MARDAN ROAD",
                    "tariff": "A-1a(01)",
                    "sub_division": "RISALPUR",
                    "status": "Def-Est",
                    "is_defective": True,
                    "_fetched": True,
                    "_log_status": "Fetched",
                    "_log_remarks": "Defective",
                    "msn": "MSN002",
                    "sim_no": "03129876543",
                    "sim_id": "899209876543210",
                    "date": "12-Nov-2024",
                    "telco": "Telenor",
                    "sco_no": "SCO-102",
                    "present_reading": "Defective",
                    "prev_reading": "2500",
                    "units": "150"
                },
                {
                    "reference": "14261234567890",
                    "name": "UNKNOWN",
                    "_not_found": True,
                    "_fetched": True,
                    "_log_status": "Error",
                    "_log_remarks": "Bill Not Found"
                }
            ]

            # Patch filedialog.asksaveasfilename and messagebox.askyesno
            with patch("tkinter.filedialog.asksaveasfilename", return_value=str(output_file)), \
                 patch("tkinter.messagebox.askyesno", return_value=False):
                app._save_logs_to_excel()

            app.destroy()

        self.assertTrue(output_file.exists(), "Excel log file should be saved")

        wb = load_workbook(output_file)
        self.assertIn("Scraping Logs", wb.sheetnames)
        ws = wb["Scraping Logs"]

        # Check headers
        headers = [cell.value for cell in ws[1]]
        expected_headers = [
            "Sr #",
            "Reference No",
            "Scraping Status",
            "Remarks"
        ]
        self.assertEqual(headers, expected_headers)

        # Check Row 2 (Replaced)
        self.assertEqual(ws.cell(row=2, column=1).value, 1)
        self.assertEqual(str(ws.cell(row=2, column=2).value), "19269310053300")
        self.assertEqual(ws.cell(row=2, column=3).value, "Fetched")
        self.assertEqual(ws.cell(row=2, column=4).value, "Already replaced in previous month")

        # Check Row 3 (Defective)
        self.assertEqual(ws.cell(row=3, column=1).value, 2)
        self.assertEqual(str(ws.cell(row=3, column=2).value), "03262320230213")
        self.assertEqual(ws.cell(row=3, column=3).value, "Fetched")
        self.assertEqual(ws.cell(row=3, column=4).value, "Defective")

        # Check Row 4 (Error)
        self.assertEqual(ws.cell(row=4, column=1).value, 3)
        self.assertEqual(str(ws.cell(row=4, column=2).value), "14261234567890")
        self.assertEqual(ws.cell(row=4, column=3).value, "Error")
        self.assertEqual(ws.cell(row=4, column=4).value, "Bill Not Found")

        wb.close()
        output_file.unlink()
        print("All Excel export tests passed successfully!")

    def test_log_window_formatting_and_save_button(self):
        with patch.object(main.LicenseManager, "is_locally_active", return_value=True), \
             patch.object(main.MCOMakerApp, "_bg_heartbeat"), \
             patch.object(main.MCOMakerApp, "_bg_update_check"):
            app = MCOMakerApp()
            app.withdraw()

            # Verify "Save the Logs" button
            self.assertTrue(hasattr(app, "btn_save_logs"))
            self.assertEqual(app.btn_save_logs.cget("text"), "💾 Save the Logs")
            self.assertEqual(app.btn_save_logs.cget("bg"), main.BTN_SEC_BG)

            # Test logging lines
            app._clear_log()
            app._log('1/3: 19269310053300 --- Fetched --- "Already replaced in previous month"', tag="err")
            app._log('2/3: 03262320230213 --- Fetched --- "Defective"', tag="blue")
            app._log('3/3: 14261234567890 --- Error --- "Bill Not Found"', tag="err")
            app._log("Scrapping completed. Click the buttons below for generating MCOs.", tag="green")

            log_content = app.log_text.get("1.0", "end").strip().splitlines()
            expected = [
                '1/3: 19269310053300 --- Fetched --- "Already replaced in previous month"',
                '2/3: 03262320230213 --- Fetched --- "Defective"',
                '3/3: 14261234567890 --- Error --- "Bill Not Found"',
                'Scrapping completed. Click the buttons below for generating MCOs.'
            ]
            self.assertEqual(log_content, expected)

            # Test scrollbar attributes
            self.assertTrue(hasattr(app, "scrollbar"))
            self.assertEqual(app.scrollbar.cget("troughcolor"), "#121212")

            app.destroy()
        print("Log window formatting & button tests passed!")

    def test_mco_buttons_disabled_during_scraping_and_sound(self):
        with patch.object(main.LicenseManager, "is_locally_active", return_value=True), \
             patch.object(main.MCOMakerApp, "_bg_heartbeat"), \
             patch.object(main.MCOMakerApp, "_bg_update_check"):
            app = MCOMakerApp()
            app.withdraw()

            # Set sample records
            app.records = [
                {"reference": "19269310053300", "_fetched": True, "name": "TEST"}
            ]

            # Simulate scraping started
            with patch("threading.Thread"):
                app._start_scraping()

            # Verify MCO generation buttons are disabled during scraping
            self.assertEqual(str(app.btn_sub_div.cget("state")), "disabled")
            self.assertEqual(str(app.btn_ro_office.cget("state")), "disabled")
            self.assertEqual(app.btn_sub_div.cget("bg"), main.BTN_DISABLED_BG)
            self.assertEqual(app.btn_ro_office.cget("bg"), main.BTN_DISABLED_BG)
            self.assertIn("Scraping in progress", app.mco_status_hint.cget("text"))

            # Simulate scraping finished
            with patch("main.play_notification_sound") as mock_sound:
                app._on_scraping_finished()
                mock_sound.assert_called_once()

            # Verify MCO buttons are now re-enabled
            self.assertEqual(str(app.btn_sub_div.cget("state")), "normal")
            self.assertEqual(str(app.btn_ro_office.cget("state")), "normal")
            self.assertEqual(app.btn_sub_div.cget("bg"), main.ACCENT_BLUE)
            self.assertEqual(app.btn_ro_office.cget("bg"), "#6f42c1")

            app.destroy()
        print("MCO buttons disabling and notification sound tests passed!")


if __name__ == "__main__":
    unittest.main()
