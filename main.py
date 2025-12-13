import os
from ui.main_window import create_main_window


def main():
    # Ensure current working dir is project root
    os.chdir(os.path.dirname(__file__))
    create_main_window()


if __name__ == '__main__':
    main()
