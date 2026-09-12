"""attendance.py の非GUIロジックに対するユニットテスト。

attendance.py は tkinter のGUIアプリだが、次の部分はGUIなしで検証できる:

- モジュールレベルのヘルパー関数
    - 認証系: _hash_password / _load_users / _save_users / _verify_user /
      _create_default_user / _get_user_role
    - スタッフ一覧: _load_staff_list / _save_staff_list
    - 有給休暇: _load_paid_leave / _save_paid_leave / _next_pl_id /
      _calc_paid_leave_balance
- AttendanceGUI の純粋なヘルパーメソッド（フォーマッタ、日時パーサ）
    - self をほぼ使わないため、未束縛メソッドにスタブ self を渡して呼ぶ

ファイルパス定数（USER_FILE / STAFF_FILE / PAID_LEAVE_FILE）は
patch.object で一時ディレクトリへ差し替え、実ファイルを汚さない。
"""
import datetime
import json
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import attendance
except Exception as exc:  # pragma: no cover - 依存(openpyxl/fpdf)未導入の環境向け
    attendance = None
    _IMPORT_ERROR = exc


@unittest.skipIf(attendance is None, "attendance を import できない環境（openpyxl/fpdf 未導入）")
class HashPasswordTests(unittest.TestCase):
    def test_deterministic_sha256(self):
        # sha256("admin") の既知値
        self.assertEqual(
            attendance._hash_password("admin"),
            "8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918",
        )

    def test_distinct_inputs_distinct_hashes(self):
        self.assertNotEqual(
            attendance._hash_password("pass1"),
            attendance._hash_password("pass2"),
        )

    def test_unicode_password(self):
        h = attendance._hash_password("パスワード")
        self.assertEqual(len(h), 64)
        self.assertEqual(h, attendance._hash_password("パスワード"))


@unittest.skipIf(attendance is None, "attendance を import できない環境")
class NextPlIdTests(unittest.TestCase):
    def test_empty_returns_1(self):
        self.assertEqual(attendance._next_pl_id([]), 1)

    def test_returns_max_plus_1(self):
        recs = [{"id": 3}, {"id": 7}, {"id": 5}]
        self.assertEqual(attendance._next_pl_id(recs), 8)

    def test_missing_id_treated_as_zero(self):
        self.assertEqual(attendance._next_pl_id([{}, {}]), 1)

    def test_string_ids_are_coerced(self):
        self.assertEqual(attendance._next_pl_id([{"id": "10"}, {"id": "2"}]), 11)


@unittest.skipIf(attendance is None, "attendance を import できない環境")
class CalcPaidLeaveBalanceTests(unittest.TestCase):
    def _data(self, grants=None, usages=None):
        return {"grants": grants or [], "usages": usages or [], "requests": []}

    def test_basic_granted_minus_used(self):
        data = self._data(
            grants=[{"name": "山田", "days": 10, "expiry_date": "2999-12-31"}],
            usages=[{"name": "山田", "days": 3}],
        )
        remaining, granted, used = attendance._calc_paid_leave_balance("山田", data, as_of="2026-01-01")
        self.assertEqual((remaining, granted, used), (7.0, 10.0, 3.0))

    def test_filters_by_name(self):
        data = self._data(
            grants=[
                {"name": "山田", "days": 10, "expiry_date": "2999-12-31"},
                {"name": "佐藤", "days": 5, "expiry_date": "2999-12-31"},
            ],
            usages=[{"name": "佐藤", "days": 1}],
        )
        self.assertEqual(
            attendance._calc_paid_leave_balance("山田", data, as_of="2026-01-01"),
            (10.0, 10.0, 0.0),
        )

    def test_expired_grant_excluded(self):
        data = self._data(
            grants=[{"name": "山田", "days": 10, "expiry_date": "2025-12-31"}],
        )
        remaining, granted, used = attendance._calc_paid_leave_balance("山田", data, as_of="2026-01-01")
        self.assertEqual((remaining, granted, used), (0.0, 0.0, 0.0))

    def test_grant_expiring_exactly_on_as_of_is_included(self):
        # 条件は expiry_date >= as_of なので当日は有効
        data = self._data(
            grants=[{"name": "山田", "days": 4, "expiry_date": "2026-01-01"}],
        )
        remaining, granted, _ = attendance._calc_paid_leave_balance("山田", data, as_of="2026-01-01")
        self.assertEqual((remaining, granted), (4.0, 4.0))

    def test_grant_without_expiry_never_expires(self):
        data = self._data(grants=[{"name": "山田", "days": 6}])
        remaining, granted, _ = attendance._calc_paid_leave_balance("山田", data, as_of="2026-01-01")
        self.assertEqual((remaining, granted), (6.0, 6.0))

    def test_result_is_rounded_to_one_decimal(self):
        data = self._data(
            grants=[{"name": "山田", "days": 1.333333, "expiry_date": "2999-12-31"}],
            usages=[{"name": "山田", "days": 0.0}],
        )
        remaining, granted, used = attendance._calc_paid_leave_balance("山田", data, as_of="2026-01-01")
        self.assertEqual((remaining, granted, used), (1.3, 1.3, 0.0))

    def test_as_of_defaults_to_today(self):
        yesterday = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
        far = (datetime.date.today() + datetime.timedelta(days=365)).isoformat()
        data = self._data(
            grants=[
                {"name": "山田", "days": 5, "expiry_date": yesterday},   # 期限切れ
                {"name": "山田", "days": 8, "expiry_date": far},          # 有効
            ],
        )
        remaining, granted, _ = attendance._calc_paid_leave_balance("山田", data)
        self.assertEqual((remaining, granted), (8.0, 8.0))


class _TmpFileConstMixin:
    """USER_FILE / STAFF_FILE / PAID_LEAVE_FILE を一時ディレクトリへ差し替える。"""

    def setUp(self):
        self._tmpdir = tempfile.mkdtemp(prefix="att_test_")
        self.addCleanup(self._cleanup_tmpdir)
        patches = {
            "USER_FILE": os.path.join(self._tmpdir, "users.json"),
            "STAFF_FILE": os.path.join(self._tmpdir, "staff_list.json"),
            "PAID_LEAVE_FILE": os.path.join(self._tmpdir, "paid_leave.json"),
        }
        for attr, value in patches.items():
            p = mock.patch.object(attendance, attr, value)
            p.start()
            self.addCleanup(p.stop)
            setattr(self, attr, value)

    def _cleanup_tmpdir(self):
        for root, _dirs, files in os.walk(self._tmpdir, topdown=False):
            for name in files:
                try:
                    os.remove(os.path.join(root, name))
                except OSError:
                    pass
        try:
            os.rmdir(self._tmpdir)
        except OSError:
            pass


@unittest.skipIf(attendance is None, "attendance を import できない環境")
class UserStoreTests(_TmpFileConstMixin, unittest.TestCase):
    def test_load_users_missing_file_returns_empty_dict(self):
        self.assertEqual(attendance._load_users(), {})

    def test_save_then_load_round_trip(self):
        users = {"admin": {"password_hash": "x", "role": "管理者"}}
        attendance._save_users(users)
        self.assertTrue(os.path.exists(self.USER_FILE))
        self.assertEqual(attendance._load_users(), users)

    def test_load_users_corrupt_json_returns_empty_dict(self):
        with open(self.USER_FILE, "w", encoding="utf-8") as f:
            f.write("{ this is not json")
        self.assertEqual(attendance._load_users(), {})

    def test_verify_user_correct_and_incorrect(self):
        attendance._save_users(
            {"taro": {"password_hash": attendance._hash_password("secret"), "role": "ユーザー"}}
        )
        self.assertTrue(attendance._verify_user("taro", "secret"))
        self.assertFalse(attendance._verify_user("taro", "wrong"))

    def test_verify_user_unknown_user(self):
        self.assertFalse(attendance._verify_user("nobody", "whatever"))

    def test_get_user_role_known_and_default(self):
        attendance._save_users({"boss": {"password_hash": "x", "role": "管理者"}})
        self.assertEqual(attendance._get_user_role("boss"), "管理者")
        self.assertEqual(attendance._get_user_role("unknown"), "ユーザー")

    def test_create_default_user_creates_admin(self):
        attendance._create_default_user()
        users = attendance._load_users()
        self.assertIn("admin", users)
        self.assertEqual(users["admin"]["role"], "管理者")
        self.assertTrue(attendance._verify_user("admin", "admin"))

    def test_create_default_user_is_idempotent(self):
        attendance._save_users(
            {"admin": {"password_hash": attendance._hash_password("changed"), "role": "管理者"}}
        )
        attendance._create_default_user()
        # 既存 admin を上書きしない
        self.assertTrue(attendance._verify_user("admin", "changed"))
        self.assertFalse(attendance._verify_user("admin", "admin"))


@unittest.skipIf(attendance is None, "attendance を import できない環境")
class StaffListStoreTests(_TmpFileConstMixin, unittest.TestCase):
    def test_load_missing_file_returns_empty_list(self):
        self.assertEqual(attendance._load_staff_list(), [])

    def test_save_then_load_dict_entries_round_trip(self):
        names = [{"name": "山田", "category": "社員"}, {"name": "佐藤", "category": "アルバイト"}]
        attendance._save_staff_list(names)
        self.assertEqual(attendance._load_staff_list(), names)

    def test_plain_string_list_is_normalized(self):
        with open(self.STAFF_FILE, "w", encoding="utf-8") as f:
            json.dump(["山田", "佐藤"], f, ensure_ascii=False)
        self.assertEqual(
            attendance._load_staff_list(),
            [{"name": "山田", "category": "一般"}, {"name": "佐藤", "category": "一般"}],
        )

    def test_mixed_list_non_dict_items_are_normalized(self):
        with open(self.STAFF_FILE, "w", encoding="utf-8") as f:
            json.dump([{"name": "山田", "category": "社員"}, 123], f, ensure_ascii=False)
        self.assertEqual(
            attendance._load_staff_list(),
            [{"name": "山田", "category": "社員"}, {"name": "123", "category": "一般"}],
        )

    def test_corrupt_json_returns_empty_list(self):
        with open(self.STAFF_FILE, "w", encoding="utf-8") as f:
            f.write("not json at all")
        self.assertEqual(attendance._load_staff_list(), [])

    def test_non_list_json_returns_empty_list(self):
        with open(self.STAFF_FILE, "w", encoding="utf-8") as f:
            json.dump({"name": "山田"}, f, ensure_ascii=False)
        self.assertEqual(attendance._load_staff_list(), [])


@unittest.skipIf(attendance is None, "attendance を import できない環境")
class PaidLeaveStoreTests(_TmpFileConstMixin, unittest.TestCase):
    def test_load_missing_file_returns_default_structure(self):
        self.assertEqual(
            attendance._load_paid_leave(),
            {"grants": [], "usages": [], "requests": []},
        )

    def test_save_then_load_round_trip(self):
        data = {
            "grants": [{"id": 1, "name": "山田", "days": 10}],
            "usages": [{"id": 1, "name": "山田", "days": 2}],
            "requests": [],
        }
        attendance._save_paid_leave(data)
        self.assertEqual(attendance._load_paid_leave(), data)

    def test_partial_dict_gets_missing_keys_filled(self):
        with open(self.PAID_LEAVE_FILE, "w", encoding="utf-8") as f:
            json.dump({"grants": [{"id": 1}]}, f, ensure_ascii=False)
        loaded = attendance._load_paid_leave()
        self.assertEqual(loaded["grants"], [{"id": 1}])
        self.assertEqual(loaded["usages"], [])
        self.assertEqual(loaded["requests"], [])

    def test_corrupt_json_returns_default_structure(self):
        with open(self.PAID_LEAVE_FILE, "w", encoding="utf-8") as f:
            f.write("<broken>")
        self.assertEqual(
            attendance._load_paid_leave(),
            {"grants": [], "usages": [], "requests": []},
        )


class _StubGUI:
    """AttendanceGUI の未束縛メソッドを呼び出すための、最小限のダミー self。

    テスト対象のメソッドが self 経由で使う機能だけを用意すればよい。
    parse_date は「日付欄が空欄になっている状態」を再現するため、常に None を返す。
    """

    def parse_date(self):
        return None


@unittest.skipIf(attendance is None, "attendance を import できない環境")
class FormatterMethodTests(unittest.TestCase):
    def setUp(self):
        self.gui = _StubGUI()
        self.G = attendance.AttendanceGUI

    def test_format_timestamp(self):
        self.assertEqual(
            self.G.format_timestamp(self.gui, "2026-06-01T09:30:15"),
            "2026-06-01 09:30:15",
        )
        self.assertEqual(self.G.format_timestamp(self.gui, ""), "-")
        self.assertEqual(self.G.format_timestamp(self.gui, "ゴミ"), "ゴミ")

    def test_format_time(self):
        self.assertEqual(self.G.format_time(self.gui, "2026-06-01T09:30:15"), "09:30")
        self.assertEqual(self.G.format_time(self.gui, ""), "")
        self.assertEqual(self.G.format_time(self.gui, "bad"), "bad")

    def test_format_lessons(self):
        self.assertEqual(self.G.format_lessons(self.gui, {"lessons_main": 0, "lessons_sub": 0}), "0")
        self.assertEqual(
            self.G.format_lessons(self.gui, {"lessons_main": 2, "lessons_sub": 1}),
            "メイン2 / サブ1",
        )
        # None / 空文字も許容される
        self.assertEqual(self.G.format_lessons(self.gui, {"lessons_main": None, "lessons_sub": ""}), "0")

    def test_format_admin_time(self):
        self.assertEqual(self.G.format_admin_time(self.gui, {"admin_minutes": 0}), "0分")
        self.assertEqual(self.G.format_admin_time(self.gui, {"admin_minutes": 90}), "1時間30分")
        self.assertEqual(self.G.format_admin_time(self.gui, {"admin_minutes": 45}), "0時間45分")

    def test_format_break_time(self):
        self.assertEqual(self.G.format_break_time(self.gui, {"break_minutes": 0}), "0:00")
        self.assertEqual(self.G.format_break_time(self.gui, {"break_minutes": 75}), "1:15")
        self.assertEqual(self.G.format_break_time(self.gui, {"break_minutes": 5}), "0:05")

    def test_format_total_work(self):
        row = {
            "checkin": "2026-06-01T09:00:00",
            "checkout": "2026-06-01T18:00:00",
            "break_minutes": "60",
        }
        # 9時間 - 1時間休憩 = 8時間 -> 末尾ゼロは削られる
        self.assertEqual(self.G.format_total_work(self.gui, row), "8")

    def test_format_total_work_fractional(self):
        row = {
            "checkin": "2026-06-01T09:00:00",
            "checkout": "2026-06-01T18:30:00",
            "break_minutes": "60",
        }
        self.assertEqual(self.G.format_total_work(self.gui, row), "8.5")


@unittest.skipIf(attendance is None, "attendance を import できない環境")
class OvernightRolloverTests(unittest.TestCase):
    def setUp(self):
        self.gui = _StubGUI()
        self.G = attendance.AttendanceGUI

    def test_checkout_before_checkin_rolls_to_next_day(self):
        ci = datetime.datetime(2026, 6, 1, 23, 0, 0)
        co = datetime.datetime(2026, 6, 1, 2, 0, 0)
        rolled = self.G._apply_overnight_rollover(self.gui, ci, co)
        self.assertEqual(rolled, datetime.datetime(2026, 6, 2, 2, 0, 0))

    def test_equal_times_roll_forward(self):
        ci = datetime.datetime(2026, 6, 1, 9, 0, 0)
        co = datetime.datetime(2026, 6, 1, 9, 0, 0)
        rolled = self.G._apply_overnight_rollover(self.gui, ci, co)
        self.assertEqual(rolled, datetime.datetime(2026, 6, 2, 9, 0, 0))

    def test_normal_order_unchanged(self):
        ci = datetime.datetime(2026, 6, 1, 9, 0, 0)
        co = datetime.datetime(2026, 6, 1, 18, 0, 0)
        self.assertEqual(self.G._apply_overnight_rollover(self.gui, ci, co), co)

    def test_none_checkout_unchanged(self):
        ci = datetime.datetime(2026, 6, 1, 9, 0, 0)
        self.assertIsNone(self.G._apply_overnight_rollover(self.gui, ci, None))


@unittest.skipIf(attendance is None, "attendance を import できない環境")
class FlexibleDatetimeParseTests(unittest.TestCase):
    def setUp(self):
        self.gui = _StubGUI()
        self.G = attendance.AttendanceGUI

    def test_time_only_uses_today(self):
        result = self.G._parse_flexible_datetime(self.gui, "2:00")
        self.assertEqual(result.date(), datetime.date.today())
        self.assertEqual((result.hour, result.minute, result.second), (2, 0, 0))

    def test_date_and_time_without_zero_padding(self):
        result = self.G._parse_flexible_datetime(self.gui, "2026-6-2 9:5")
        self.assertEqual(result, datetime.datetime(2026, 6, 2, 9, 5, 0))

    def test_with_seconds_and_t_separator(self):
        result = self.G._parse_flexible_datetime(self.gui, "2026-6-2T9:5:30")
        self.assertEqual(result, datetime.datetime(2026, 6, 2, 9, 5, 30))

    def test_invalid_hour_returns_none(self):
        self.assertIsNone(self.G._parse_flexible_datetime(self.gui, "25:00"))

    def test_non_numeric_returns_none(self):
        self.assertIsNone(self.G._parse_flexible_datetime(self.gui, "ゴミ"))

    def test_single_field_returns_none(self):
        self.assertIsNone(self.G._parse_flexible_datetime(self.gui, "0900"))


@unittest.skipIf(attendance is None, "attendance を import できない環境")
class ParseTimestampTests(unittest.TestCase):
    def setUp(self):
        self.G = attendance.AttendanceGUI
        # parse_timestamp は self.parse_date と self._parse_flexible_datetime を使う
        import types

        stub = _StubGUI()
        stub._parse_flexible_datetime = types.MethodType(self.G._parse_flexible_datetime, stub)
        self.gui = stub

    def test_strict_iso_datetime(self):
        self.assertEqual(
            self.G.parse_timestamp(self.gui, "2026-06-01T09:00:00"),
            datetime.datetime(2026, 6, 1, 9, 0, 0),
        )

    def test_empty_and_dash_return_none(self):
        self.assertIsNone(self.G.parse_timestamp(self.gui, ""))
        self.assertIsNone(self.G.parse_timestamp(self.gui, "-"))

    def test_flexible_time_only_falls_through(self):
        result = self.G.parse_timestamp(self.gui, "9:5")
        self.assertEqual(result.date(), datetime.date.today())
        self.assertEqual((result.hour, result.minute), (9, 5))

    def test_invalid_raises_value_error(self):
        with self.assertRaises(ValueError):
            self.G.parse_timestamp(self.gui, "not-a-timestamp")


if __name__ == "__main__":
    unittest.main()
