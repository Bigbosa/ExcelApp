from tkinter import scrolledtext, filedialog, Tk, Button, END
import pandas as pd
from customtkinter import CTkButton, CTkFrame

class CTkScrolledText(scrolledtext.ScrolledText):
    def __init__(self, master=None, **kwargs):
        scrolledtext.ScrolledText.__init__(self, master, **kwargs)
        self.config(bg="darkgreen", fg="white", font=("Courier New", 10))

class DisplayDataApp(Tk):
    def __init__(self):
        Tk.__init__(self)
        self.title("Display Data")
        self.geometry("800x380")
        self.configure(bg="black")
        self.iconbitmap("ExcelApp_logo.ico")
        self.resizable(False, False)

        # Create a custom scrolled text widget with dark green background
        self.text_box = CTkScrolledText(self, wrap="word", width=100, height=20)
        self.text_box.pack()

        # Create a custom button with dark green background, white text, and red color when hovered
        self.button = CTkButton(
            self,
            text="Add Excel File",
            command=self.open_excel_file,
            hover_color="lightgreen",
            fg_color="transparent",
            text_color="white",
            border_color="lightgreen",
            border_width=2
        )
        self.button.pack()

    def open_excel_file(self):
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])

        if file_path:
            try:
                # Read data from the Excel file using pandas
                df = pd.read_excel(file_path, engine='openpyxl')

                # Display the data in the custom scrolled text widget
                self.display_data_in_text_box(df)

            except Exception as e:
                print(f"Error reading Excel file: {e}")

    def display_data_in_text_box(self, df):
        # Clear existing data in the custom scrolled text widget
        self.text_box.delete("1.0", END)

        # Customize font and formatting
        header_font = ("Courier New", 10, "underline")

        # Insert column headers with underline
        headers = " | ".join(df.columns)
        self.text_box.tag_configure("header", font=header_font)
        self.text_box.insert(END, headers + "\n", "header")

        # Insert data into the custom scrolled text widget
        for index, row in df.iterrows():
            data = " | ".join(str(cell) for cell in row)
            self.text_box.insert(END, data + "\n")

        # Scroll to the top
        self.text_box.yview(END)

if __name__ == "__main__":
    app = DisplayDataApp()
    app.mainloop()
