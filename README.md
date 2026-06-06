# スポーツクラブ 勤怠管理 CLI

簡単なコマンドラインツールです。出勤・退勤・レッスン数をCSVで保存します。

使い方例:

```bash
# 出勤記録（複数スタッフ同時登録可）
python attendance.py checkin 山田 佐藤

# 日付指定して出勤（YYYY-MM-DD または ISO 形式）
python attendance.py checkin 山田 -d 2026-05-22

# レッスンを追加（複数スタッフ可）
python attendance.py lessons 山田 2
python attendance.py lessons 山田 佐藤 1 -d 2026-05-22

# 退勤記録
python attendance.py checkout 山田

# 最近の記録を表示
python attendance.py list 20

# 集計（スタッフ名を指定するとその人のみ）
python attendance.py report
python attendance.py report 山田
```

データは同じディレクトリの `attendance.csv` に保存されます。

## GUI版の起動

Python3で以下のコマンドを使い、GUIを起動します。

```bash
python attendance_gui.py
```

GUIでは次の操作ができます。
- スタッフ名入力
- 日付指定
- 出勤 / 退勤 / レッスン追加
- 最新レコードの一覧表示
- 名前ごとの勤務時間・レッスン集計

必要に応じて日付指定やUIの追加、データベース保存へ拡張できます。
"# attendance" 
"# attendance" 
