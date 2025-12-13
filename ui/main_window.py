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
        logo_path = os.path.join('assets', 'icons', 'logo1.png')
        # Use PIL for better image handling in ctk
        my_image = ctk.CTkImage(light_image=Image.open(logo_path),
                                dark_image=Image.open(logo_path),
                                size=(130, 150)) # Approx size based on subsample(6,6) of original if it was large
        
        # Original code used subsample(6,6). Let's assume original was ~600x600 -> 100x100.
        # If we don't know original size, we can try to guess or just use a reasonable size.
        # Let's stick to a reasonable icon size.
        
        logo_label = ctk.CTkLabel(app, image=my_image, text="")
        logo_label.pack(pady=(20, 10))
    except Exception as e:
        # Handle case where image file is not found or is not a valid image
        print(f"Logo error: {e}")
        pass

    header = ctk.CTkLabel(app, text='سیستم مدیریت پرونده مهاجر', font=('vazirmatn', 23, 'bold'))
    header.pack(pady=(10, 5))

    frm = ctk.CTkFrame(app)
    frm.pack(padx=20, pady=12, fill='both', expand=True)

    # Width 30 chars ~ 300px? 
    # Height 2 chars ~ 50-60px?
    # ctk button height default is usually around 28-32. 
    # Let's use width=250, height=50 for a big button feel.
    
    btn_new = ctk.CTkButton(frm, text='➕  افزودن پرونده جدید', font=('vazirmatn', 13, 'bold'), width=250, height=50,
                        command=lambda: open_add_record(app))
    btn_search = ctk.CTkButton(frm, text='🔍  جستجوی پرونده‌ها', font=('vazirmatn', 13, 'bold'), width=250, height=50,
                          command=lambda: open_search_records(app))
    
    btn_new.pack(pady=10)
    btn_search.pack(pady=10)

    # Footer labels
    lbl_author = ctk.CTkLabel(app, text='نوشته شده توسط راستین علیزاده', font=('vazirmatn', 10), text_color='gray')
    lbl_author.pack(side='bottom', pady=(0, 5))
    lbl_version = ctk.CTkLabel(app, text='v0.9.1', font=('vazirmatn', 9), text_color='gray')
    lbl_version.pack(side='bottom')

    app.mainloop()
