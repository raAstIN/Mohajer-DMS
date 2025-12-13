import os
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import jdatetime
from database import get_case_by_id, update_case, delete_case
from ui.add_record import open_edit_record

def get_jalali_day_name(date_str):
    """Convert Jalali date string (YYYY-MM-DD) to day name with Farsi weekday."""
    try:
        jdate = jdatetime.datetime.strptime(date_str, '%Y-%m-%d').date()
        # In Jalali calendar: 0=شنبه, 1=یکشنبه, 2=دوشنبه, 3=سه‌شنبه, 4=چهارشنبه, 5=پنج‌شنبه, 6=جمعه
        weekdays = ['شنبه', 'یکشنبه', 'دوشنبه', 'سه‌شنبه', 'چهارشنبه', 'پنج‌شنبه', 'جمعه']
        day_name = weekdays[jdate.weekday()]
        return f"{date_str} ({day_name})"
    except Exception:
        return date_str


def open_details_window(master, case_id):
    """Open a Details window for a given case_id (function-based)."""
    top = ctk.CTkToplevel(master)
    top.title(f'جزئیات پرونده {case_id} — سیستم مدیریت پرونده مهاجر')
    top.geometry('800x600')

    # Remove single-line header; present fields as label:value rows (RTL)
    info = ctk.CTkFrame(top)
    info.pack(padx=12, pady=6, fill='both', expand=True)

    # Styling: bold for field names, regular for values. Use uniform font sizes.
    field_font = ('vazirmatn', 14, 'bold')
    value_font = ('vazirmatn', 14)

    # Create rows with field label and value that we can hide if empty
    rows_data = []
    row_counter = [0]  # Use list to maintain counter through nested function
    
    def mk_row(field_label_text, data_key):
        """Create a row with field label and value. Returns a dict for later manipulation."""
        row_num = row_counter[0]
        row_counter[0] += 1
        
        # ctk label anchor needs to be mapped. 'e' -> 'e', 'w' -> 'w' works.
        lbl_field = ctk.CTkLabel(info, text=field_label_text, font=field_font, anchor='e', justify='right')
        lbl_value = ctk.CTkLabel(info, text='', font=value_font, anchor='w', justify='right', wraplength=600)
        
        row_dict = {
            'field_label': field_label_text,
            'data_key': data_key,
            'row_num': row_num,
            'lbl_field': lbl_field,
            'lbl_value': lbl_value,
            'visible': False
        }
        rows_data.append(row_dict)
        return row_dict

    # Create all row definitions
    row_id = mk_row('شناسه بایگانی:', 'id')
    row_title = mk_row('عنوان:', 'title')
    row_subject = mk_row('موضوع:', 'subject')
    row_case_type = mk_row('نوع پرونده:', 'case_type')
    row_date = mk_row('تاریخ:', 'date')
    row_duration = mk_row('مدت:', 'duration')
    row_mojer = mk_row('موجر:', 'mojer')
    row_mostajjer = mk_row('مستاجر:', 'mostajjer')
    row_karfarma = mk_row('کارفرما:', 'karfarma')
    row_piman = mk_row('پیمانکار:', 'piman')
    row_amount = mk_row('مبلغ قرارداد:', 'contract_amount')
    row_status = mk_row('وضعیت:', 'status')
    row_description = mk_row('توضیحات:', 'description')

    # make first column expandable for long values
    info.grid_columnconfigure(0, weight=1)

    def relayout_visible_rows():
        """Reorganize rows so visible ones appear consecutively."""
        visible_rows = [r for r in rows_data if r['visible']]
        for new_row_pos, row_dict in enumerate(visible_rows):
            row_dict['lbl_field'].grid(row=new_row_pos, column=1, sticky='e', padx=(6,12), pady=4)
            row_dict['lbl_value'].grid(row=new_row_pos, column=0, sticky='e', padx=(6,12), pady=4)
        # Hide all invisible rows
        for row_dict in rows_data:
            if not row_dict['visible']:
                row_dict['lbl_field'].grid_forget()
                row_dict['lbl_value'].grid_forget()

    frm = ctk.CTkFrame(top)
    frm.pack(pady=8)

    def load():
        data = get_case_by_id(case_id)
        if not data:
            messagebox.showerror('خطا', 'پرونده یافت نشد')
            top.destroy()
            return None
        
        # Set values and determine visibility
        row_id['lbl_value'].configure(text=data.get('id') or '')
        row_id['visible'] = bool(data.get('id'))
        
        row_title['lbl_value'].configure(text=data.get('title') or '')
        row_title['visible'] = bool(data.get('title'))
        
        row_subject['lbl_value'].configure(text=data.get('subject') or '')
        row_subject['visible'] = bool(data.get('subject'))
        
        row_case_type['lbl_value'].configure(text=data.get('case_type') or '')
        row_case_type['visible'] = bool(data.get('case_type'))
        
        row_date['lbl_value'].configure(text=get_jalali_day_name(data.get('date')) if data.get('date') else '')
        row_date['visible'] = bool(data.get('date'))
        
        # Duration with from/to
        dur = data.get('duration') or ''
        dfrom = data.get('duration_from') or ''
        dto = data.get('duration_to') or ''
        if dfrom or dto:
            dfrom_display = get_jalali_day_name(dfrom) if dfrom else dfrom
            dto_display = get_jalali_day_name(dto) if dto else dto
            row_duration['lbl_value'].configure(text=f"{dur}  (از {dfrom_display} تا {dto_display})")
        else:
            row_duration['lbl_value'].configure(text=dur)
        row_duration['visible'] = bool(dur or dfrom or dto)
        
        # Party fields
        row_mojer['lbl_value'].configure(text=data.get('mojer') or '')
        row_mojer['visible'] = bool(data.get('mojer'))
        
        row_mostajjer['lbl_value'].configure(text=data.get('mostajjer') or '')
        row_mostajjer['visible'] = bool(data.get('mostajjer'))
        
        row_karfarma['lbl_value'].configure(text=data.get('karfarma') or '')
        row_karfarma['visible'] = bool(data.get('karfarma'))
        
        row_piman['lbl_value'].configure(text=data.get('piman') or '')
        row_piman['visible'] = bool(data.get('piman'))
        
        # Format contract amount with commas
        amount = data.get('contract_amount')
        if amount:
            try:
                row_amount['lbl_value'].configure(text=f"{int(amount):,} ریال")
            except (ValueError, TypeError):
                row_amount['lbl_value'].configure(text=str(amount) + ' ریال')
        else:
            row_amount['lbl_value'].configure(text='')
        row_amount['visible'] = bool(data.get('contract_amount'))
        
        row_status['lbl_value'].configure(text=data.get('status') or 'در جریان')
        row_status['visible'] = bool(data.get('status'))
        
        row_description['lbl_value'].configure(text=data.get('description') or '')
        row_description['visible'] = bool(data.get('description'))
        
        # Store bank data for modal
        bank_info = {
            'owner_name': data.get('bank_owner_name') or '',
            'account_number': data.get('bank_account_number') or '',
            'shaba_number': data.get('bank_shaba_number') or '',
            'card_number': data.get('bank_card_number') or '',
            'bank_name': data.get('bank_name') or '',
            'bank_branch': data.get('bank_branch') or '',
            'payment_id': data.get('payment_id') or ''
        }
        
        relayout_visible_rows()
        return data, bank_info

    # Initialize data and bank_info
    data = {}
    bank_info = {}
    guarantee_info = {}

    # Load data and populate the window
    loaded_data, loaded_bank_info = load()
    if loaded_data:
        data, bank_info = loaded_data, loaded_bank_info

    def open_bank_info_window():
        """Open a modal window with bank account information."""
        has_bank_info = any([bank_info['owner_name'], bank_info['account_number'], 
                             bank_info['shaba_number'], bank_info['card_number'], bank_info['bank_name'],
                             bank_info['bank_branch'], bank_info['payment_id']])
        if not has_bank_info:
            messagebox.showinfo('اطلاع', 'اطلاعات بانکی برای این پرونده موجود نیست.', parent=top)
            return
        
        bank_win = ctk.CTkToplevel(top)
        bank_win.title('اطلاعات حساب بانکی')
        bank_win.geometry('450x320')
        bank_win.resizable(False, False)
        
        # Configure grid layout for RTL
        bank_win.grid_columnconfigure(0, weight=0)
        bank_win.grid_columnconfigure(1, weight=1)
        
        font_bank = ('vazirmatn', 13, 'bold')
        pad = 10
        
        # نام صاحب حساب
        lbl_owner = ctk.CTkLabel(bank_win, text='نام صاحب حساب:', font=('vazirmatn', 12, 'bold'))
        lbl_owner.grid(row=0, column=1, sticky='e', padx=(pad, pad), pady=(pad, 4))
        value_owner = ctk.CTkLabel(bank_win, text=bank_info['owner_name'], font=font_bank, justify='right')
        value_owner.grid(row=0, column=0, sticky='ew', padx=(pad, pad), pady=(pad, 4))
        
        # شماره حساب
        lbl_account = ctk.CTkLabel(bank_win, text='شماره حساب:', font=('vazirmatn', 12, 'bold'))
        lbl_account.grid(row=1, column=1, sticky='e', padx=(pad, pad), pady=4)
        value_account = ctk.CTkLabel(bank_win, text=bank_info['account_number'], font=font_bank, justify='right')
        value_account.grid(row=1, column=0, sticky='ew', padx=(pad, pad), pady=4)
        
        # شماره شبا
        lbl_shaba = ctk.CTkLabel(bank_win, text='شماره شبا:', font=('vazirmatn', 12, 'bold'))
        lbl_shaba.grid(row=2, column=1, sticky='e', padx=(pad, pad), pady=4)
        value_shaba = ctk.CTkLabel(bank_win, text=bank_info['shaba_number'], font=font_bank, justify='right')
        value_shaba.grid(row=2, column=0, sticky='ew', padx=(pad, pad), pady=4)
        
        # نام بانک
        lbl_bank_name = ctk.CTkLabel(bank_win, text='نام بانک:', font=('vazirmatn', 12, 'bold'))
        lbl_bank_name.grid(row=3, column=1, sticky='e', padx=(pad, pad), pady=4)
        value_bank_name = ctk.CTkLabel(bank_win, text=bank_info['bank_name'], font=font_bank, justify='right')
        value_bank_name.grid(row=3, column=0, sticky='ew', padx=(pad, pad), pady=4)

        # شعبه بانک
        lbl_branch = ctk.CTkLabel(bank_win, text='شعبه بانک:', font=('vazirmatn', 12, 'bold'))
        lbl_branch.grid(row=4, column=1, sticky='e', padx=(pad, pad), pady=4)
        value_branch = ctk.CTkLabel(bank_win, text=bank_info['bank_branch'], font=font_bank, justify='right')
        value_branch.grid(row=4, column=0, sticky='ew', padx=(pad, pad), pady=4)

        # شناسه پرداخت
        lbl_payment_id = ctk.CTkLabel(bank_win, text='شناسه پرداخت:', font=('vazirmatn', 12, 'bold'))
        lbl_payment_id.grid(row=5, column=1, sticky='e', padx=(pad, pad), pady=4)
        value_payment_id = ctk.CTkLabel(bank_win, text=bank_info['payment_id'], font=font_bank, justify='right')
        value_payment_id.grid(row=5, column=0, sticky='ew', padx=(pad, pad), pady=4)

        # شماره کارت
        lbl_card = ctk.CTkLabel(bank_win, text='شماره کارت:', font=('vazirmatn', 12, 'bold'))
        lbl_card.grid(row=6, column=1, sticky='e', padx=(pad, pad), pady=4)
        value_card = ctk.CTkLabel(bank_win, text=bank_info['card_number'], font=font_bank, justify='right')
        value_card.grid(row=6, column=0, sticky='ew', padx=(pad, pad), pady=4)
        
        # Close button
        btn_close = ctk.CTkButton(bank_win, text='بستن', command=bank_win.destroy, font=('vazirmatn', 12, 'bold'))
        btn_close.grid(row=7, column=0, columnspan=2, pady=(pad, pad), padx=pad, sticky='ew')

    def open_guarantee_info_window():
        """Open a modal window with guarantee information."""
        guarantee_info = {
            'amount': data.get('guarantee_amount', ''),
            'type': data.get('guarantee_type', '')
        }
        has_guarantee_info = any(guarantee_info.values())
        if not has_guarantee_info:
            messagebox.showinfo('اطلاع', 'اطلاعات تضمین برای این پرونده موجود نیست.', parent=top)
            return

        guarantee_win = ctk.CTkToplevel(top)
        guarantee_win.title('اطلاعات تضمین قرارداد')
        guarantee_win.geometry('450x150')
        guarantee_win.resizable(False, False)

        guarantee_win.grid_columnconfigure(0, weight=1)
        guarantee_win.grid_columnconfigure(1, weight=0)

        font_guarantee = ('vazirmatn', 13, 'bold')
        pad = 10

        # مبلغ تضمین
        lbl_amount = ctk.CTkLabel(guarantee_win, text='مبلغ تضمین:', font=('vazirmatn', 12, 'bold'))
        lbl_amount.grid(row=0, column=1, sticky='e', padx=(pad, pad), pady=(pad, 4))
        amount_val = guarantee_info.get('amount')
        try:
            amount_text = f"{int(amount_val):,} ریال" if amount_val else ''
        except (ValueError, TypeError):
            amount_text = f"{str(amount_val)} ریال" if amount_val else ''
        value_amount = ctk.CTkLabel(guarantee_win, text=amount_text, font=font_guarantee, justify='right')
        value_amount.grid(row=0, column=0, sticky='ew', padx=(pad, pad), pady=(pad, 4))

        # نوع ضمانت
        lbl_type = ctk.CTkLabel(guarantee_win, text='نوع ضمانت:', font=('vazirmatn', 12, 'bold'))
        lbl_type.grid(row=1, column=1, sticky='e', padx=(pad, pad), pady=4)
        value_type = ctk.CTkLabel(guarantee_win, text=guarantee_info['type'], font=font_guarantee, justify='right')
        value_type.grid(row=1, column=0, sticky='ew', padx=(pad, pad), pady=4)

        # Close button
        btn_close = ctk.CTkButton(guarantee_win, text='بستن', command=guarantee_win.destroy, font=('vazirmatn', 12, 'bold'))
        btn_close.grid(row=2, column=0, columnspan=2, pady=(pad, pad), padx=pad, sticky='ew')


    def open_folder():
        # Always fetch the latest data directly from the DB to avoid stale state
        current_data = get_case_by_id(case_id)
        if not current_data:
            messagebox.showerror('خطا', 'پرونده یافت نشد یا حذف شده است.')
            return
        folder = current_data.get('folder_path')
        if not folder or not os.path.exists(folder):
            messagebox.showwarning('هشدار', 'فولدر پیوست برای این پرونده وجود ندارد یا مسیر آن نامعتبر است.')
            return
        try:
            if sys.platform.startswith('win'):
                os.startfile(folder)
            elif sys.platform == 'darwin':
                subprocess.run(['open', folder])
            else:
                subprocess.run(['xdg-open', folder])
        except Exception as e:
            messagebox.showerror('خطا', f'نشد باز شود: {e}')

    def edit():
        # Open the unified edit/create window prefilled with current case data
        open_edit_record(top, case_id)

    def delete():
        if not messagebox.askyesno('تایید', 'آیا از حذف این پرونده مطمئن هستید؟'):
            return
        try:
            delete_case(case_id)
            messagebox.showinfo('موفق', 'پرونده حذف شد')
            top.destroy()
        except Exception as e:
            messagebox.showerror('خطا در حذف', f'{e}')

    # Determine if bank info button should be enabled
    has_bank_info = any([bank_info['owner_name'], bank_info['account_number'], 
                         bank_info['shaba_number'], bank_info['card_number'], bank_info['bank_name'],
                         bank_info['bank_branch'], bank_info['payment_id']])
    bank_btn_state = "normal" if has_bank_info else "disabled"
    
    # Determine if guarantee info button should be enabled
    has_guarantee_info = any([data.get('guarantee_amount'), data.get('guarantee_type')])
    guarantee_btn_state = "normal" if has_guarantee_info else "disabled"

    btn_open = ctk.CTkButton(frm, text='🗂 باز کردن فولدر پیوست‌ها', command=open_folder, font=('vazirmatn', 13, 'bold'))
    btn_bank = ctk.CTkButton(frm, text='🏦 اطلاعات بانکی', command=open_bank_info_window, font=('vazirmatn', 13, 'bold'), state=bank_btn_state)
    btn_guarantee = ctk.CTkButton(frm, text='🛡️ اطلاعات تضمین', command=open_guarantee_info_window, font=('vazirmatn', 13, 'bold'), state=guarantee_btn_state)
    btn_edit = ctk.CTkButton(frm, text='📝 ویرایش پرونده', command=edit, font=('vazirmatn', 13, 'bold'))
    btn_del = ctk.CTkButton(frm, text='🗑 حذف پرونده', command=delete, font=('vazirmatn', 13, 'bold'), fg_color="red", hover_color="darkred")
    
    btn_open.grid(row=0, column=0, padx=6)
    btn_bank.grid(row=0, column=1, padx=6)
    btn_guarantee.grid(row=0, column=2, padx=6)
    btn_edit.grid(row=0, column=3, padx=6)
    btn_del.grid(row=0, column=4, padx=6)
