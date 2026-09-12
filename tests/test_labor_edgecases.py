import os
import csv
import datetime
import tempfile
import unittest

import labor

class EdgeCaseTests(unittest.TestCase):
    def write_rows(self, rows):
        fd, path = tempfile.mkstemp(prefix="att_", suffix=".csv", dir=os.getcwd())
        os.close(fd)
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id','name','date','checkin','checkout','break_minutes','admin_minutes','lessons_main','lessons_sub'])
            for r in rows:
                writer.writerow(r)
        return path

    def test_overnight_split(self):
        # 22:00 -> 06:00 with 60min break = 7h total, split 2h & 5h approx
        checkin = '2026-05-01T22:00:00'
        checkout = '2026-05-02T06:00:00'
        rows = [[ '1', 'Overnight', '2026-05-01', checkin, checkout, '60', '0', '0', '0' ]]
        path = self.write_rows(rows)
        try:
            am = labor.AttendanceManager(path)
            lm = labor.LaborAgreementManager(am)
            monthly = lm.compute_monthly_summary(2026,5)
            # find Overnight
            self.assertTrue(any(r['name']=='Overnight' for r in monthly))
            # total hours should be 7.0
            rec = [r for r in monthly if r['name']=='Overnight'][0]
            self.assertAlmostEqual(rec['total_hours'], 7.0, places=2)
        finally:
            os.remove(path)

    def test_missing_checkout(self):
        rows = [[ '1', 'NoCheckout', '2026-05-03', '2026-05-03T09:00:00', '', '0', '0', '0', '0' ]]
        path = self.write_rows(rows)
        try:
            am = labor.AttendanceManager(path)
            rows2 = am.read_rows()
            self.assertEqual(labor.AttendanceManager.calculate_actual_hours(rows2[0]), 0.0)
        finally:
            os.remove(path)

    def test_bad_timestamp(self):
        rows = [[ '1', 'BadTime', '2026-05-04', 'not-a-time', 'also-bad', '0', '0', '0', '0' ]]
        path = self.write_rows(rows)
        try:
            am = labor.AttendanceManager(path)
            rows2 = am.read_rows()
            self.assertEqual(labor.AttendanceManager.calculate_actual_hours(rows2[0]), 0.0)
        finally:
            os.remove(path)

    def test_break_start_end_minutes(self):
        rows = [[
            '1', 'BreakTest', '2026-05-05', '2026-05-05T09:00:00', '2026-05-05T18:00:00', '0', '0', '0', '0',
            '2026-05-05T12:00:00', '2026-05-05T12:45:00'
        ]]
        # write rows with extra break_start/break_end fields
        fd, path = tempfile.mkstemp(prefix="att_", suffix=".csv", dir=os.getcwd())
        os.close(fd)
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id','name','date','checkin','checkout','break_minutes','admin_minutes','lessons_main','lessons_sub','break_start','break_end'])
            for r in rows:
                writer.writerow(r)
        try:
            am = labor.AttendanceManager(path)
            rows2 = am.read_rows()
            self.assertAlmostEqual(labor.AttendanceManager.calculate_actual_hours(rows2[0]), 8.25, places=2)
            self.assertAlmostEqual(am.calculate_night_hours(rows2[0]), 0.0, places=2)
        finally:
            os.remove(path)

    def test_monthly_summary_includes_night_and_break(self):
        rows = [[
            '1', 'NightBreak', '2026-05-01', '2026-05-01T22:00:00', '2026-05-02T06:00:00', '30', '0', '0', '0',
            '2026-05-01T23:00:00', '2026-05-02T00:00:00'
        ]]
        fd, path = tempfile.mkstemp(prefix="att_", suffix=".csv", dir=os.getcwd())
        os.close(fd)
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id','name','date','checkin','checkout','break_minutes','admin_minutes','lessons_main','lessons_sub','break_start','break_end'])
            for r in rows:
                writer.writerow(r)
        try:
            am = labor.AttendanceManager(path)
            lm = labor.LaborAgreementManager(am)
            monthly = lm.compute_monthly_summary(2026, 5)
            self.assertEqual(len(monthly), 1)
            rec = monthly[0]
            self.assertAlmostEqual(rec['night_hours'], 7.0, places=2)
            self.assertGreaterEqual(rec['break_violations'], 1)
        finally:
            os.remove(path)

if __name__ == '__main__':
    unittest.main()
