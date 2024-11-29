import os
import sqlite3
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox
import pandas as pd
import xlrd
import pyautogui
import time

from openpyxl.reader.excel import load_workbook


class FiltererApp:
    def __init__(self, root):

        conn = sqlite3.connect("inventory.db")
        cursor = conn.cursor()
        cursor.execute("DROP TABLE IF EXISTS inventory")
        conn.commit()
        conn.close()

        if os.path.exists("filtered.txt"):
            os.remove("filtered.txt")
        self.departments_filtered = 0
        self.zero_article = 0
        self.root = root
        self.root.title("0 Filterer")

        # Initialize dictionary with elements and associated False values
        self.elements = {"Grocery": False, "Meat": False, "Bakery": False, "Baby": False, "Fresh": False, "Home": False}

        # Create a frame to hold the buttons and lights
        self.frame = tk.Frame(root)
        self.frame.pack(pady=20, padx=20)

        # Create buttons and lights for each element
        self.buttons = {}
        self.lights = {}
        for idx, key in enumerate(self.elements):
            # Button
            btn = tk.Label(self.frame, text=key)
            btn.grid(row=idx, column=0, padx=10, pady=5)
            self.buttons[key] = btn

            # Light indicator
            light = tk.Canvas(self.frame, width=20, height=20)
            light.create_oval(2, 2, 18, 18, fill="red", tags="light")
            light.grid(row=idx, column=1)
            self.lights[key] = light

        # Add a button for selecting Excel file
        self.upload_button = tk.Button(root, text="Upload List", command=self.upload)
        self.upload_button.pack(pady=10)

        # Add a button to save green-lighted departments to a file
        self.save_button = tk.Button(root, text="Filter DNOs", command=self.write_to_file)
        self.save_button.pack(pady=10)

        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

        self.paste_button = None

    def send_to_SAP(self):
        confirm = messagebox.askokcancel(
            "Confirm Action",
            f"Make sure the SAP window is in the far left position, ETA: {self.zero_article * 4.5 // 60} mins.",
            parent=self.root  # Attach to the main window
        )

        if confirm:
            with open("filtered.txt", 'r') as file:
                lines = file.readlines()
            entryx = 2570  # example x-coordinate of the entry box
            entryy = 92

            time.sleep(3)

            def process_lines(data):
                if not data:
                    return  # base case: empty line
                line = data.pop(0).strip()  # get the first line and remove leading/trailing spaces
                if line:
                    if line == "Article":
                        pass
                    else:
                        # Move the mouse to the specified entry coordinates
                        pyautogui.moveTo(entryx, entryy)
                        time.sleep(2)
                        # Paste the line
                        pyautogui.write(line)

                        # Press Enter to submit the line
                        pyautogui.press('enter')
                        time.sleep(2)  # wait for the submission to process

                        # Press Enter again to move to the next line (or to clear any previous selections)
                        pyautogui.press('enter')
                        time.sleep(0.5)

                    # Recursively process the next line
                    process_lines(data)

            process_lines(lines)

    def close_app(self):

        self.log_activity()

        # Close the Tkinter application
        self.root.destroy()

    def count_true_departments(self):
        return sum(1 for department, secondary in self.elements.items() if secondary)

    def write_to_file(self):
        try:
            # Connect to the SQLite inventory database
            conn = sqlite3.connect("inventory.db")
            cursor = conn.cursor()
            # Query the inventory table for items with inventory <= 0
            cursor.execute("SELECT article, inventory FROM inventory WHERE inventory <= 0")
            inventory_items = cursor.fetchall()

            conn_dno = sqlite3.connect("dno.db")
            cursor_dno = conn_dno.cursor()

            # Fetch all articles from the DNO database
            cursor_dno.execute("SELECT article FROM dno")
            dno_articles = {row[0] for row in cursor_dno.fetchall()}
            conn_dno.close()

            # Filter out inventory items that are already in the DNO database
            filtered_inventory = [item[0] for item in inventory_items if item[0] not in dno_articles]
            conn.close()

            filtered_df = pd.DataFrame(filtered_inventory, columns=["Article"])

            # Read existing articles from the file if it exists
            try:
                existing_df = pd.read_csv("filtered.txt", delimiter="\t")
                existing_articles = set(existing_df["Article"])
            except (FileNotFoundError, pd.errors.EmptyDataError):
                existing_articles = set()

            # Filter out articles already in the file
            new_articles = filtered_df[~filtered_df["Article"].isin(existing_articles)]
            self.zero_article += len(new_articles)
            # Append only new articles to the file
            if not new_articles.empty:
                new_articles.to_csv("filtered.txt", mode="a", header=not existing_articles, index=False)

            messagebox.showinfo("Success", f"Filtered inventory saved to filtered.txt.")
            self.update_paste_button_text(self.zero_article)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def upload(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_path.endswith('.xls'):
            # For older .xls files
            df = pd.read_excel(file_path, engine='xlrd')
            workbook = xlrd.open_workbook(file_path)
        elif file_path.endswith('.xlsx'):
            # For newer .xlsx files
            df = pd.read_excel(file_path, engine='openpyxl')
            workbook = load_workbook(file_path)
        elif not file_path:
            return

        sheet = workbook.sheet_by_index(0)

        # Read the department name from the first column, second row
        department = sheet.cell_value(1, 0)

        # Update the corresponding light to green
        department = str(department).strip()
        if department in self.elements:
            if not self.elements[department]:
                self.elements[department] = True
            self.lights[department].itemconfig("light", fill="green")

        # Extract only the relevant columns: Department, Article, and Inventory
        columns = ['Department', 'Article', 'Inventory']

        filtered_df = df[columns]

        # Connect to SQLite (or create a new database)
        conn = sqlite3.connect('inventory.db')
        cursor = conn.cursor()

        # Create the table
        cursor.execute('''
                        CREATE TABLE if not exists inventory (department TEXT, article INTEGER, inventory REAL)
                    ''')

        # Insert data into the table
        for _, row in filtered_df.iterrows():
            cursor.execute('''
                            INSERT INTO inventory (department, article, inventory) VALUES (?, ?, ?)
                        ''', (row['Department'], row['Article'], row['Inventory']))

        conn.commit()
        conn.close()

    def update_paste_button_text(self, article_count: int):
        if self.paste_button:
            self.paste_button.destroy()
        self.paste_button = tk.Button(self.root, text=f"Send to SAP", command=self.send_to_SAP)
        self.paste_button.config(text=f"Send to SAP ({article_count} articles)")
        self.paste_button.pack(pady=10)

    def log_activity(self):
        """Log the filtered department count and other information to log.txt."""
        if (self.zero_article > 0):
            log_message = (
                f"Session Ended: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Input {self.zero_article} articles to SAP filtered from {self.count_true_departments()}/7 departments\n"
                "--------------------------------------------\n"
            )

            # right here append to log_message the quanitity of each item from the new dict
            with open("log.txt", "a") as log_file:
                log_file.write(log_message)


# Create the Tkinter window
root = tk.Tk()
app = FiltererApp(root)
root.mainloop()
