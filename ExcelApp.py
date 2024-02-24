import customtkinter as ctk
from customtkinter import CTkFrame, CTkButton, CTkSwitch
import tkinter as tk
import time
from PIL import Image, ImageTk
from calculation import ExcelCalculatorApp
from add_data import AddDataFrame
from display_data import DisplayDataApp
import customtkinter
class CustomApp(ctk.CTk):
    def __init__(self, *args, **kwargs):
        ctk.CTk.__init__(self, *args, **kwargs)
        self.title("Excel App")
        self.geometry("700x500")
        self.iconbitmap("ExcelApp_logo.ico")
        self.resizable(False, False)

        # Create the menu frame
        option_frame = CTkFrame(self)
        option_frame.pack(side=tk.LEFT, fill=tk.BOTH)
        option_frame.pack_propagate(False)
        option_frame.configure(width=150, height=500)

        calculation_btn = CTkButton(option_frame, text="Open Calculator", hover_color="lightgreen",
                                    fg_color="transparent", text_color="black", border_color="lightgreen", border_width=2, command=self.toggle_calculation)
        calculation_btn.place(x=10, y=50)

        add_data_btn = CTkButton(option_frame, text="Open Add Data", hover_color="lightgreen",
                                    fg_color="transparent", text_color="black", border_color="lightgreen",
                                    border_width=2, command=self.toggle_add_data)
        add_data_btn.place(x=10, y=100)

        display_data_btn = CTkButton(option_frame, text="Open Display Data", hover_color="lightgreen",
                                     fg_color="transparent", text_color="black", border_color="lightgreen",
                                     border_width=2, command=self.toggle_display_data)
        display_data_btn.place(x=10, y=150)

        # Create the main frame
        main_frame = CTkFrame(self)
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        main_frame.pack_propagate(False)
        main_frame.configure(width=550, height=500)

        # Create the red background on the right side
        right_frame = CTkFrame(main_frame)
        right_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Display time, year, month, day
        time_label = tk.Label(right_frame, text="", font=("Bold", 14), bg="black", fg="white")
        time_label.pack(pady=10)

        date_label = tk.Label(right_frame, text="", font=("Bold", 12), bg="black", fg="white")
        date_label.pack(pady=5)

        # Load and display the app logo
        app_logo_path = "E:/ExcelAssist/src/logo_123.png"  # Change to your actual logo file path
        app_logo = Image.open(app_logo_path)
        app_logo_image = ImageTk.PhotoImage(app_logo)
        self.app_logo_label = tk.Label(right_frame, image=app_logo_image)
        self.app_logo_label.image = app_logo_image
        self.app_logo_label.pack(pady=10)

        # Create the theme switch
        self.switch = CTkSwitch(master=self, text="Mode", command=self.theme_change)
        self.switch.toggle(1)
        self.switch.place(relx=0.05, rely=0.05)

        self.update_time(time_label, date_label)

        # Store references to opened windows
        self.calculation_window = None
        self.add_data_window = None
        self.display_data_window = None

        self.after(1000, self.update_time, time_label, date_label)

    def toggle_calculation(self):
        self.calculation_window = ExcelCalculatorApp(self)

    def toggle_add_data(self):
        self.add_data_window = AddDataFrame()

    def toggle_display_data(self):
        self.display_data_window = DisplayDataApp()

    def update_time(self, time_label, date_label):
        current_time = time.strftime("%H:%M:%S")
        current_date = time.strftime("%B %d, %Y")  # Format: Month Day, Year
        time_label.config(text=f"Time: {current_time}")
        date_label.config(text=f"Date: {current_date}")
        self.after(1000, self.update_time, time_label, date_label)

    def theme_change(self):
        if self.switch.get() == 1:
            customtkinter.set_appearance_mode("dark")
            self.app_logo_label.configure(bg="black")  # Change to the desired dark mode color
        else:
            customtkinter.set_appearance_mode("light")
            self.app_logo_label.configure(bg="#FFFFFF")  # Change to the desired light mode color

if __name__ == "__main__":
    app = CustomApp()
    app.mainloop()
