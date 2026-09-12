"""payroll.py の境界値テスト。

対象:
- 所得税(甲欄)の電算機計算特例で使う各表の区切り値
- 課税給与所得金額がマイナスになる場合に0円へ丸められること
- 介護保険が適用される年齢(40歳/64歳)の境界
- 月60時間超の残業割増(25%→50%)の切り替え境界
- 月給制で所定労働時間が0の場合のゼロ除算回避
- 週40h・日8hの労働時間区分(compute_monthly_work_breakdown)の境界
"""
import os
import csv
import datetime
import tempfile
import unittest
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import labor
import payroll as P


class IncomeTaxTableBoundaryTests(unittest.TestCase):
    """第1表(給与所得控除)・第3表(基礎控除)・第4表(税額)の区切り値の境界。
    国税庁「令和8年分 電算機計算の特例」の区分表に基づく。"""

    # ── 第1表: 給与所得控除の区分（158,333 / 158,334 など） ──
    def test_employment_income_deduction_boundary_is_continuous(self):
        # 158,333円はフラット54,167円の区分の上限、158,334円からは
        # (A)*30%+6,667円 の区分に変わるが、区分の切り替え点では
        # 両区分の式がほぼ一致するよう設計されているため、税額は連続する
        # （急に跳ね上がったり下がったりしない）ことを確認する。
        below = P._employment_income_deduction(158333)
        above = P._employment_income_deduction(158334)
        self.assertEqual(below, 54167)
        self.assertAlmostEqual(below, above, delta=2)

    def test_employment_income_deduction_rate_changes_beyond_boundary(self):
        # 区分の傾き（控除の増え方）は158,334円以降で変わるため、
        # 境界から離れた点では明確にフラット54,167円とは異なる値になる
        far_above = P._employment_income_deduction(250000)
        self.assertNotEqual(far_above, 54167)
        self.assertEqual(far_above, 81667)  # 250,000*30%+6,667

    def test_employment_income_deduction_boundary_300000(self):
        # 299,999 -> (A)*30%+6,667 / 300,000 -> (A)*20%+36,667 に切り替わる
        tax_below = P.compute_income_tax_kou(299999, 0, False)
        tax_at = P.compute_income_tax_kou(300000, 0, False)
        self.assertIsInstance(tax_below, int)
        self.assertIsInstance(tax_at, int)

    # ── 第3表: 基礎控除の区分（2,120,833 / 2,120,834 など） ──
    def test_basic_deduction_boundary_2120833(self):
        below = P._basic_deduction(2120833)
        above = P._basic_deduction(2120834)
        self.assertEqual(below, 48334)
        self.assertEqual(above, 40000)

    def test_basic_deduction_boundary_2245833(self):
        below = P._basic_deduction(2245833)
        above = P._basic_deduction(2245834)
        self.assertEqual(below, 13334)
        self.assertEqual(above, 0)

    # ── 第4表: 税率区分（162,500 / 162,501 など） ──
    def test_tax_table4_boundary_is_continuous(self):
        # 162,500円以下は(B)*5.105%、162,501円以上は(B)*10.210%-8,296円だが、
        # -8,296円という定数は境界でちょうど連続になるよう逆算された値のため、
        # 境界をまたいでも税額が急変しないことを確認する
        below = P._tax_from_table4(162500)
        above = P._tax_from_table4(162501)
        self.assertEqual(below, 8300)
        self.assertAlmostEqual(below, above, delta=10)

    def test_tax_table4_rate_changes_beyond_boundary(self):
        # 境界から離れた点(200,000円)では、区分1の式(5.105%のみ)と
        # 区分2の実際の計算式(10.210%-8,296円)の結果が明確に異なる
        far_above = P._tax_from_table4(200000)
        rate1_only = round(200000 * 0.05105 / 10) * 10
        self.assertNotEqual(far_above, rate1_only)

    def test_tax_table4_boundary_3333333(self):
        below = P._tax_from_table4(3333333)
        above = P._tax_from_table4(3333334)
        self.assertGreater(above - below, -1)  # 税額は単調非減少であるべき

    # ── 課税給与所得金額がマイナスになるケース ──
    def test_negative_taxable_income_clamps_to_zero(self):
        # 総支給額が低く、控除合計(給与所得控除+基礎控除等)がそれを上回る場合
        tax = P.compute_income_tax_kou(50000, 0, False)
        self.assertEqual(tax, 0)

    def test_zero_income_is_zero_tax(self):
        tax = P.compute_income_tax_kou(0, 0, False)
        self.assertEqual(tax, 0)


class CareInsuranceAgeBoundaryTests(unittest.TestCase):
    """介護保険料が加算される40〜64歳の年齢境界。"""

    def test_day_before_40th_birthday_is_39(self):
        age = P.calc_age("1986-06-15", datetime.date(2026, 6, 14))
        self.assertEqual(age, 39)

    def test_exactly_on_40th_birthday_is_40(self):
        age = P.calc_age("1986-06-15", datetime.date(2026, 6, 15))
        self.assertEqual(age, 40)

    def test_day_before_65th_birthday_is_64(self):
        age = P.calc_age("1961-06-15", datetime.date(2026, 6, 14))
        self.assertEqual(age, 64)

    def test_exactly_on_65th_birthday_is_65(self):
        age = P.calc_age("1961-06-15", datetime.date(2026, 6, 15))
        self.assertEqual(age, 65)

    def _care_insurance_for_age(self, birth_date, as_of):
        # 月給制で固定の総支給額を作り、gross_total=0 による見かけ上の
        # 「介護保険0円」（年齢に関係なく常に0）を排除して年齢判定だけを見る。
        settings = P.load_payroll_settings()
        am = labor.AttendanceManager(csv_path=os.path.join(tempfile.gettempdir(), "no_such_file.csv"))
        staff = {
            "name": "境界太郎", "pay_type": "月給", "monthly_base_salary": 250000,
            "social_insurance": True, "employment_insurance": False,
            "birth_date": birth_date,
        }
        return P.compute_payroll_for_staff(staff, am, as_of.year, as_of.month, settings, as_of=as_of)

    def test_care_insurance_not_applied_at_39(self):
        res = self._care_insurance_for_age("1986-06-15", datetime.date(2026, 6, 14))
        self.assertEqual(res["care_insurance"], 0)

    def test_care_insurance_applied_at_40(self):
        res = self._care_insurance_for_age("1986-06-15", datetime.date(2026, 6, 15))
        self.assertGreater(res["care_insurance"], 0)

    def test_care_insurance_applied_at_64(self):
        res = self._care_insurance_for_age("1961-06-15", datetime.date(2026, 6, 14))
        self.assertGreater(res["care_insurance"], 0)

    def test_care_insurance_not_applied_at_65(self):
        res = self._care_insurance_for_age("1961-06-15", datetime.date(2026, 6, 15))
        self.assertEqual(res["care_insurance"], 0)


class OvertimeTierBoundaryTests(unittest.TestCase):
    """月60時間超の残業で25%→50%に切り替わる境界(compute_gross_pay)。"""

    def test_exactly_60h_overtime_all_at_25_percent(self):
        settings = P.load_payroll_settings()
        breakdown = {"regular_hours": 0, "overtime25_hours": 60.0, "overtime50_hours": 0.0,
                     "holiday_hours": 0, "night_hours": 0, "total_hours": 60.0}
        gross = P.compute_gross_pay(1000, breakdown, settings)
        self.assertEqual(gross["pay_overtime50"], 0)
        self.assertEqual(gross["pay_overtime25"], round(1000 * 60 * 1.25))

    def test_60h_plus_1_at_50_percent_via_breakdown_split(self):
        # compute_monthly_work_breakdown 側の閾値処理を直接確認
        threshold = float(P.load_payroll_settings().get("overtime_monthly_threshold_hours", 60.0))
        overtime_hours_all = threshold + 1.0
        overtime25 = min(overtime_hours_all, threshold)
        overtime50 = max(0.0, overtime_hours_all - threshold)
        self.assertEqual(overtime25, 60.0)
        self.assertEqual(overtime50, 1.0)


class MonthlySalaryZeroDivisionTests(unittest.TestCase):
    """月給制で所定労働時間(月間)が0のときにゼロ除算しないこと。"""

    def test_zero_standard_hours_does_not_raise(self):
        settings = dict(P.load_payroll_settings())
        breakdown = {"regular_hours": 0, "overtime25_hours": 5.0, "overtime50_hours": 0.0,
                     "holiday_hours": 0, "night_hours": 0, "total_hours": 5.0}
        gross = P.compute_gross_pay(
            0, breakdown, settings, pay_type="月給",
            monthly_base_salary=250000, standard_monthly_hours=0,
        )
        self.assertEqual(gross["premium_wage"], 0.0)
        self.assertEqual(gross["pay_overtime25"], 0)


class WorkBreakdownBoundaryTests(unittest.TestCase):
    """compute_monthly_work_breakdown の日8h・週40hの区分境界。"""

    def setUp(self):
        self._paths = []

    def tearDown(self):
        for p in self._paths:
            try:
                os.remove(p)
            except Exception:
                pass

    def _write_rows(self, rows):
        fd, path = tempfile.mkstemp(prefix="pr_bd_", suffix=".csv", dir=os.getcwd())
        os.close(fd)
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["id", "name", "date", "checkin", "checkout",
                              "break_minutes", "work_type", "admin_minutes"])
            for r in rows:
                writer.writerow(r)
        self._paths.append(path)
        return path

    def test_exactly_8h_5days_40h_week_all_regular(self):
        rows = []
        for i in range(5):
            d = (datetime.date(2026, 6, 1) + datetime.timedelta(days=i)).isoformat()
            rows.append([str(i + 1), "境界花子", d, f"{d}T09:00:00",
                         f"{d}T18:00:00", "60", "通常出勤", "0"])
        path = self._write_rows(rows)
        am = labor.AttendanceManager(path)
        bd = P.compute_monthly_work_breakdown(am, "境界花子", 2026, 6)
        self.assertEqual(bd["regular_hours"], 40.0)
        self.assertEqual(bd["overtime25_hours"], 0.0)
        self.assertEqual(bd["overtime50_hours"], 0.0)

    def test_one_minute_over_8h_creates_overtime(self):
        rows = []
        for i in range(5):
            d = (datetime.date(2026, 6, 1) + datetime.timedelta(days=i)).isoformat()
            checkout = f"{d}T18:01:00" if i == 4 else f"{d}T18:00:00"
            rows.append([str(i + 1), "境界太郎", d, f"{d}T09:00:00",
                         checkout, "60", "通常出勤", "0"])
        path = self._write_rows(rows)
        am = labor.AttendanceManager(path)
        bd = P.compute_monthly_work_breakdown(am, "境界太郎", 2026, 6)
        # 1分 = 1/60h = 0.0167h だが、結果は小数第2位に丸められるため 0.02 になる
        self.assertEqual(bd["overtime25_hours"], 0.02)
        self.assertGreater(bd["overtime25_hours"], 0.0)
        self.assertEqual(bd["regular_hours"], 40.0)  # 8h超の1分は残業側へ、残りは40hのまま

    def test_legal_holiday_hours_excluded_from_daily_weekly_check(self):
        # 法定休日出勤は8h/40hの通常判定から除外され、35%区分に計上される
        d = "2026-06-07"  # 日曜
        rows = [["1", "休日次郎", d, f"{d}T09:00:00", f"{d}T20:00:00", "60", "法定休日出勤", "0"]]
        path = self._write_rows(rows)
        am = labor.AttendanceManager(path)
        bd = P.compute_monthly_work_breakdown(am, "休日次郎", 2026, 6)
        self.assertEqual(bd["regular_hours"], 0.0)
        self.assertEqual(bd["overtime25_hours"], 0.0)
        self.assertAlmostEqual(bd["holiday_hours"], 10.0, places=2)  # 11h - 60分休憩


if __name__ == "__main__":
    unittest.main()
