import os
from ui.main_window import create_main_window
from ui.login_window import create_login_window


def main():
    # Ensure current working dir is project root
    os.chdir(os.path.dirname(__file__))
    # Launch login window first. On success, it will call create_main_window
    create_login_window(create_main_window)


if __name__ == '__main__':
    main()
