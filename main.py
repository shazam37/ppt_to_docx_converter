import multiprocessing
import customtkinter as ctk
from ui.main_window import MainWindow


def main():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    app = ctk.CTk()
    MainWindow(app)
    app.mainloop()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
