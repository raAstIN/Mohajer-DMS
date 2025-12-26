import os
import customtkinter as ctk
from PIL import Image, ImageTk
from tkinter import messagebox

def create_login_window(on_success):
    """Create and run the login window.
    
    on_success: callback function to execute on successful login (e.g. launching main window).
    """
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    app = ctk.CTk()
    app.title('ورود به سیستم — مدیریت پرونده مهاجر')
    app.geometry('300x330')
    app.resizable(False, False)

    # Logo
    try:
        logo_path = os.path.join('assets', 'icons', 'logo1.png')
        if os.path.exists(logo_path):
            my_image = ctk.CTkImage(light_image=Image.open(logo_path),
                                    dark_image=Image.open(logo_path),
                                    size=(100, 100))
            logo_label = ctk.CTkLabel(app, image=my_image, text="")
            logo_label.pack(pady=(30, 10))
    except Exception:
        pass

    lbl_info = ctk.CTkLabel(app, text='کلید برنامه را وارد نمایید', font=('vazirmatn', 14, 'bold'))
    lbl_info.pack(pady=(10, 5))

    entry_key = ctk.CTkEntry(app, font=('vazirmatn', 12), justify='center', width=200)
    entry_key.pack(pady=5)
    entry_key.focus()

    login_success = [False]

    def check_login(event=None):
        key = entry_key.get().strip()
        # Simple hardcoded key for demonstration. In production, hash check or DB check.
        # User requested "security for entry", implies a specific key.
        # Defaulting to "admin" or "1234" as placeholders.
        if key == "1234": 
            login_success[0] = True
            app.quit()
        else:
            messagebox.showerror("خطا", "کلید وارد شده اشتباه است")

    btn_login = ctk.CTkButton(app, text='ورود به برنامه', command=check_login, font=('vazirmatn', 13, 'bold'), width=200)
    btn_login.pack(pady=20)
    
    entry_key.bind('<Return>', check_login)

    # Footer
    lbl_author = ctk.CTkLabel(app, text='نوشته شده توسط راستین علیزاده', font=('vazirmatn', 10), text_color='gray')
    lbl_author.pack(side='bottom', pady=(0, 5))
    lbl_version = ctk.CTkLabel(app, text='v2.1.1', font=('vazirmatn', 10), text_color='gray')
    lbl_version.pack(side='bottom', pady=0)

    app.mainloop()
    
    try:
        app.destroy()
    except Exception:
        pass

    if login_success[0]:
        on_success()
