"""labor.py の境界値テスト。

対象:
- 休憩時間の法定義務(6h/8h)の境界
- 深夜時間帯(22:00/5:00)の境界
- 日次残業(8h)・週次残業(40h)の境界
"""
import os
import csv
import datetime
import tempfile
import unittest
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import labor


def write_rows(rows):
    fd, path = tempfile.mkstemp(prefix="att_bd_", suffix=".csv", dir=os.getcwd())
    os.close(fd)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["id", "name", "date", "checkin", "checkout",
                          "break_minutes", "work_type", "admin_minutes",
                          "lessons_main", "lessons_sub"])
        for r in rows:
            writer.writerow(r)
    return path


class BreakComplianceBoundaryTests(unittest.TestCase):
    """休憩時間の法定義務: 6h以下=不要、6h超=45分、8h超=60分。"""

    def _row(self, checkin, checkout, break_minutes):
        return {"checkin": checkin, "checkout": checkout, "break_minutes": str(break_minutes)}

    def test_exactly_6h_no_break_required(self):
        row = self._row("2026-06-01T09:00:00", "2026-06-01T15:00:00", 0)
        am = labor.AttendanceManager()
        ok, required, actual = am.check_break_compliance(row)
        self.assertEqual(required, 0)
        self.assertTrue(ok)

    def test_6h_plus_1min_requires_45(self):
        row = self._row("2026-06-01T09:00:00", "2026-06-01T15:01:00", 0)
        am = labor.AttendanceManager()
        ok, required, actual = am.check_break_compliance(row)
        self.assertEqual(required, 45)
        self.assertFalse(ok)

    def test_exactly_8h_still_requires_45_not_60(self):
        row = self._row("2026-06-01T09:00:00", "2026-06-01T17:00:00", 45)
        am = labor.AttendanceManager()
        ok, required, actual = am.check_break_compliance(row)
        self.assertEqual(required, 45)
        self.assertTrue(ok)

    def test_8h_plus_1min_requires_60(self):
        row = self._row("2026-06-01T09:00:00", "2026-06-01T17:01:00", 45)
        am = labor.AttendanceManager()
        ok, required, actual = am.check_break_compliance(row)
        self.assertEqual(required, 60)
        self.assertFalse(ok)  # 45分しか取っていないため不足

    def test_8h_plus_1min_with_60_is_compliant(self):
        row = self._row("2026-06-01T09:00:00", "2026-06-01T17:01:00", 60)
        am = labor.AttendanceManager()
        ok, required, actual = am.check_break_compliance(row)
        self.assertEqual(required, 60)
        self.assertTrue(ok)


class NightHoursBoundaryTests(unittest.TestCase):
    """深夜割増の対象時間帯: 22:00〜翌5:00 の境界。"""

    def test_ends_exactly_at_22_00_no_night_hours(self):
        row = {"checkin": "2026-06-01T18:00:00", "checkout": "2026-06-01T22:00:00"}
        am = labor.AttendanceManager()
        self.assertAlmostEqual(am.calculate_night_hours(row), 0.0, places=4)

    def test_1min_past_22_00_counts_as_night(self):
        row = {"checkin": "2026-06-01T18:00:00", "checkout": "2026-06-01T22:01:00"}
        am = labor.AttendanceManager()
        self.assertAlmostEqual(am.calculate_night_hours(row), 1 / 60, places=4)

    def test_starts_exactly_at_5_00_no_night_hours(self):
        row = {"checkin": "2026-06-02T05:00:00", "checkout": "2026-06-02T09:00:00"}
        am = labor.AttendanceManager()
        self.assertAlmostEqual(am.calculate_night_hours(row), 0.0, places=4)

    def test_1min_before_5_00_counts_as_night(self):
        row = {"checkin": "2026-06-02T04:59:00", "checkout": "2026-06-02T09:00:00"}
        am = labor.AttendanceManager()
        self.assertAlmostEqual(am.calculate_night_hours(row), 1 / 60, places=4)

    def test_full_night_span_22_to_5_is_7_hours(self):
        row = {"checkin": "2026-06-01T22:00:00", "checkout": "2026-06-02T05:00:00"}
        am = labor.AttendanceManager()
        self.assertAlmostEqual(am.calculate_night_hours(row), 7.0, places=4)


class OvertimeBoundaryTests(unittest.TestCase):
    """36協定監視ロジック(compute_monthly_summary)の日次8h・週次40hの境界。"""

    def tearDown(self):
        for p in getattr(self, "_paths", []):
            try:
                os.remove(p)
            except Exception:
                pass

    def _summary_for(self, rows):
        self._paths = getattr(self, "_paths", [])
        path = write_rows(rows)
        self._paths.append(path)
        am = labor.AttendanceManager(path)
        lm = labor.LaborAgreementManager(am)
        return lm.compute_monthly_summary(2026, 6)

    def test_exactly_8h_per_day_no_daily_overtime(self):
        # 2026-06-01 (月) 09:00-18:00 break60 = 8h ちょうど
        rows = [["1", "Staff8h", "2026-06-01", "2026-06-01T09:00:00",
                  "2026-06-01T18:00:00", "60", "通常出勤", "0", "0", "0"]]
        summary = self._summary_for(rows)
        rec = [r for r in summary if r["name"] == "Staff8h"][0]
        self.assertEqual(rec["daily_overtime"], 0.0)

    def test_8h_plus_1min_creates_daily_overtime(self):
        rows = [["1", "Staff8h1m", "2026-06-01", "2026-06-01T09:00:00",
                  "2026-06-01T18:01:00", "60", "通常出勤", "0", "0", "0"]]
        summary = self._summary_for(rows)
        rec = [r for r in summary if r["name"] == "Staff8h1m"][0]
        self.assertGreater(rec["daily_overtime"], 0.0)
        # 1分 = 1/60h = 0.0167h だが、結果は小数第2位に丸められるため 0.02 になる
        self.assertEqual(rec["daily_overtime"], 0.02)

    def test_exactly_40h_per_week_no_weekly_overtime(self):
        # 月〜金 8h/日 x5 = 40h ちょうど（2026-06-01は月曜）
        rows = []
        for i in range(5):
            d = (datetime.date(2026, 6, 1) + datetime.timedelta(days=i)).isoformat()
            rows.append([str(i + 1), "Staff40h", d, f"{d}T09:00:00",
                         f"{d}T18:00:00", "60", "通常出勤", "0", "0", "0"])
        summary = self._summary_for(rows)
        rec = [r for r in summary if r["name"] == "Staff40h"][0]
        self.assertEqual(rec["weekly_overtime"], 0.0)
        self.assertEqual(rec["overtime_hours"], 0.0)

    def test_40h_plus_1min_in_week_creates_weekly_overtime(self):
        rows = []
        for i in range(5):
            d = (datetime.date(2026, 6, 1) + datetime.timedelta(days=i)).isoformat()
            checkout_min = 1 if i == 4 else 0  # 金曜だけ1分延長
            rows.append([str(i + 1), "Staff40h1m", d, f"{d}T09:00:00",
                         f"{d}T18:0{checkout_min}:00" if checkout_min else f"{d}T18:00:00",
                         "60", "通常出勤", "0", "0", "0"])
        summary = self._summary_for(rows)
        rec = [r for r in summary if r["name"] == "Staff40h1m"][0]
        self.assertGreater(rec["overtime_hours"], 0.0)


if __name__ == "__main__":
    unittest.main()
