import pandas as pd
import sqlite3
from tkinter import Tk, filedialog

# Joe if u want to add DNOs, put them in an excel file, over the first 10 columns, run this file,
# it will input the new DNOs
#
def import_article_numbers():
    # Hide the root Tkinter window
    root = Tk()
    root.withdraw()

    # Open a file dialog for the user to select the Excel file
    excel_file_path = filedialog.askopenfilename(
        title="Select the dno file",
        filetypes=[("Excel Files", "*.xls")]
    )

    if not excel_file_path:
        print("No file selected. Exiting.")
        return

    try:
        # Read the selected Excel file
        df = pd.read_excel(excel_file_path)

        # Connect to SQLite database
        conn = sqlite3.connect('dno.db')
        cursor = conn.cursor()

        # Flatten the DataFrame to get all unique values in the first 9 columns
        unique_articles = pd.Series(df.iloc[:, :10].values.ravel()).dropna().astype(str).unique()

        # Insert the unique articles into the database
        insert_count = 0
        for article in unique_articles:
            cursor.execute('INSERT OR IGNORE INTO dno (article) VALUES (?)', (article,))
            insert_count += 1

        # Commit and close the database connection
        conn.commit()
        conn.close()

        print(f"Successfully imported {insert_count} unique articles into dno.db.")

    except Exception as e:
        print(f"An error occurred: {e}")



if __name__ == "__main__":
    import_article_numbers()
