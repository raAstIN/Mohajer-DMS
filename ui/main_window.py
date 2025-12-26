import os
import customtkinter as ctk
from PIL import Image, ImageTk

from ui.add_record import open_add_record
from ui.search_records import open_search_records


def create_main_window():
    """Create and run the main application window using customtkinter.

    This function creates the root CTk window, places two main buttons and
    starts the mainloop.
    """
    ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
    ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

    app = ctk.CTk()
    app.title('سیستم مدیریت پرونده مهاجر')
    app.geometry('340x455') 

    # Load and display the logo above the header
    try:
        # Use relative path assuming CWD is project root (set in main.py)
        logo_path = os.path.join('assets', 'icons', 'logo1.png')
        
        # Use PIL for better image handling in ctk
        my_image = ctk.CTkImage(light_image=Image.open(logo_path),
                                dark_image=Image.open(logo_path),
                                size=(130, 150))
        
        logo_label = ctk.CTkLabel(app, image=my_image, text="")
        logo_label.pack(pady=(20, 10))
    except Exception as e:
        # Handle case where image file is not found or is not a valid image
        print(f"Logo error: {e}")
        pass

    header = ctk.CTkLabel(app, text='سیستم مدیریت پرونده مهاجر', font=('vazirmatn', 23, 'bold'))
    header.pack(pady=(10, 5))

    # Buttons directly in app (no frame)
    btn_new = ctk.CTkButton(app, text='افزودن پرونده جدید', font=('vazirmatn', 13, 'bold'), width=250, height=50,
                        command=lambda: open_add_record(app))
    btn_search = ctk.CTkButton(app, text='جستجوی پرونده‌ها', font=('vazirmatn', 13, 'bold'), width=250, height=50,
                          command=lambda: open_search_records(app))
    
    btn_new.pack(pady=20)
    btn_search.pack(pady=10)

    # Footer labels
    lbl_author = ctk.CTkLabel(app, text='نوشته شده توسط راستین علیزاده', font=('vazirmatn', 10), text_color='gray')
    lbl_author.pack(side='bottom', pady=(0, 5))
    lbl_version = ctk.CTkLabel(app, text='v2.1.1', font=('vazirmatn', 9), text_color='gray')
    lbl_version.pack(side='bottom')

    app.mainloop()
