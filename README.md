# سیستم مدیریت پرونده مهاجر

نسخه پایه از پروژه‌ی «سیستم مدیریت پرونده مهاجر» با پایتون، SQLite و رابط گرافیکی (CustomTkinter/ttk fallback).

ساختار پروژه

```
project_root/
├── main.py
├── database.py
├── ui/
│   ├── main_window.py
│   ├── add_record.py
│   ├── search_records.py
│   └── details_window.py
├── files/
│   ├── uploads/
│   └── backup/
└── assets/
    └── icons/
```

نیازمندی‌ها

- Python 3.10+
- بسته‌های موجود در `requirements.txt` (CustomTkinter, ttkbootstrap, tkcalendar, jdatetime)

نصب و اجرا

1. ساخت یک محیط مجازی و نصب وابستگی‌ها:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. اجرای برنامه:

```bash
python main.py
```

توضیحات

- دیتابیس SQLite در `files/cases.db` قرار می‌گیرد.
- هنگام ایجاد پرونده، یک شناسهٔ یکتا بر اساس تاریخ و زمان شمسی تولید می‌شود و فولدری در `files/uploads/<id>/` ساخته می‌شود.
- بعد از هر ذخیره یا حذف، از دیتابیس یک نسخه پشتیبان در `files/backup/` ایجاد می‌شود.
- برای نمایش تقویم از `tkcalendar.DateEntry` استفاده شده و برای تولید تاریخ شمسی از `jdatetime`.

توسعهٔ بیشتر

این پیاده‌سازی یک اسکلت اولیه است؛ موارد پیشنهادی بعدی:
- افزودن آیکن‌ها و بهبود استایل UI
- فیلتر پیشرفته و صفحه‌بندی نتایج
- صادرات/واردات CSV
- محافظت با رمز عبور برای بخش مدیریتی

