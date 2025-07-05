# -----------------------------------------------------------
#  Step : 1  (Import the Required Libraries0)
# -----------------------------------------------------------

import pandas as pd
import numpy as np

import seaborn as sns
import matplotlib.pyplot as plt

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import os
from glob import glob

import mysql.connector


# -----------------------------------------------------------
# Step : 2 (Load and merge Excel files)
# -----------------------------------------------------------

def load_and_merge_excels(folder_path):


    try:

        excel_files = glob(os.path.join(folder_path, "*.xlsx"))


        if not excel_files:


            messagebox.showwarning("No Files Found", "No Excel files found in this folder!")

            return pd.DataFrame()
        
        
        dataframes = [pd.read_excel(file) for file in excel_files]

        merged_df = pd.concat(dataframes, ignore_index=True)

        return merged_df
    

    except Exception as e:


        messagebox.showerror("Load Error", f"Error loading files:\n{e}")

        return pd.DataFrame()
    

# -----------------------------------------------------------
#  Step : 3  (Clean data)
# -----------------------------------------------------------

def clean_data(df):


    try:

        df_cleaned = df.drop_duplicates()

        df_cleaned = df_cleaned.dropna()

        return df_cleaned
    

    except Exception as e:

        messagebox.showerror("Clean Error", f"Error cleaning data:\n{e}")

        return df
    

# -----------------------------------------------------------
#  Step : 4  (Save the cleaned dataframe into Excel)
# -----------------------------------------------------------

def save_cleaned_excel(df, cleaned_file_path):


    try:

        df.to_excel(cleaned_file_path, index=False)

        messagebox.showinfo("Saved", f"Cleaned data saved to {cleaned_file_path}")


    except Exception as e:

        messagebox.showerror("Save Error", str(e))


# --------------------------------------------------------------
#  Step : 5 (Save the cleaned dataframe into MySQL Database)
# --------------------------------------------------------------

def save_to_mysql(df, table_name):


    try:

        conn = mysql.connector.connect(

            host="localhost",
            user="root",
            password="<give your mysql password>",
            database="Python_Project"

        )

        cursor = conn.cursor()
        
        # build dynamic table structure based on dataframe columns

        columns_sql = []


        for col, dtype in df.dtypes.items():

            if pd.api.types.is_integer_dtype(dtype):

                sql_type = "INT"

            elif pd.api.types.is_float_dtype(dtype):

                sql_type = "DOUBLE"

            elif pd.api.types.is_datetime64_any_dtype(dtype):

                sql_type = "DATE"

            else:

                sql_type = "VARCHAR(255)"



            columns_sql.append(f"{col} {sql_type}")


        
        create_sql = ", ".join(columns_sql)

        
        cursor.execute(f"DROP TABLE IF EXISTS {table_name}")

        cursor.execute(f"CREATE TABLE {table_name} ({create_sql})")

        
        # insert data :

        for _, row in df.iterrows():

            placeholders = ", ".join(["%s"] * len(row))

            cursor.execute(f"INSERT INTO {table_name} VALUES ({placeholders})", tuple(row))

        
        conn.commit()

        cursor.close()

        conn.close()


        messagebox.showinfo("MySQL", f"Cleaned data inserted into MySQL table {table_name} successfully!")



    except Exception as e:

        messagebox.showerror("MySQL Error", str(e))


# ---------------------------------------------------------------
#   Step : 6  (Plot functions) (To visualize the cleaned Data)
# ---------------------------------------------------------------

# (i) Line Plot :
# -----------------

def plot_line(df):


    try:

        df.plot(kind="line", figsize=(10,6))

        plt.title("Line Chart")

        plt.show()


    except Exception as e:

        messagebox.showerror("Plot Error", str(e))

#------------------------------------------------------------------------------------------------


#  (ii) Bar chart :
# -------------------

def plot_bar(df):


    try:

        df.select_dtypes(include=[np.number]).plot(kind="bar", figsize=(10,6))

        plt.title("Bar Chart")

        plt.show()


    except Exception as e:

        messagebox.showerror("Plot Error", str(e))


#------------------------------------------------------------------------------------------------


#  (iii) Pie Chart :
# --------------------

def plot_pie(df):


    try:

        df.iloc[:,1].value_counts().plot(kind="pie", autopct="%.1f%%", figsize=(8,8))

        plt.title("Pie Chart")

        plt.show()


    except Exception as e:

        messagebox.showerror("Plot Error", str(e))


#------------------------------------------------------------------------------------------------


# (iv) Scatter Plot :
# --------------------

def plot_scatter(df):


    try:


        if df.shape[1] >= 2:


            sns.scatterplot(x=df.columns[0], y=df.columns[1], data=df)

            plt.title("Scatter Plot")

            plt.show()


        else:

            messagebox.showwarning("Scatter Plot", "At least 2 columns are required for scatter plot.")


    except Exception as e:

        messagebox.showerror("Plot Error", str(e))


#------------------------------------------------------------------------------------------------

#  (v) Histogram :
# -----------------

def plot_hist(df):


    try:


        df.select_dtypes(include=[np.number]).hist(figsize=(10,6))

        plt.suptitle("Histograms")

        plt.show()


    except Exception as e:

        messagebox.showerror("Plot Error", str(e))


#------------------------------------------------------------------------------------------------

#  (vi) Heatmap :
# ----------------

def plot_heatmap(df):


    try:

        corr = df.select_dtypes(include=np.number).corr()

        sns.heatmap(corr, annot=True, cmap="coolwarm")

        plt.title("Heatmap")

        plt.show()


    except Exception as e:

        messagebox.showerror("Plot Error", str(e))


#------------------------------------------------------------------------------------------------
#                                          Step : 7 ( GUI )
#------------------------------------------------------------------------------------------------


def main_gui():


    root = tk.Tk()

    root.title("Excel Data Cleaning & Visualization Using Python Modules")

    root.geometry("1300x900")


    style = ttk.Style()

    style.theme_use("clam")

    style.configure("TButton", font=("Arial", 12, "bold"), background="#FFA500", foreground="black")


    notebook = ttk.Notebook(root)

    notebook.pack(fill="both", expand=True)


    # ---------------------------------------------------------------------------------------------------------
    #                                        Tab 1 : ( About Project )
    # ---------------------------------------------------------------------------------------------------------

    tab1 = tk.Frame(notebook, bg="#001F3F")

    notebook.add(tab1, text="About")


    about_frame = tk.Frame(tab1, bg="#001F3F", padx=30, pady=30)

    about_frame.pack(fill="both", expand=True)


    tk.Label(
        about_frame,
        text="Excel Data Cleaning & Visualization By Using Python Modules",
        fg="yellow", bg="#001F3F",
        font=("Arial", 28, "bold", "underline")
    ).pack(pady=20)


    text_frame = tk.Frame(about_frame, bg="#001F3F")

    text_frame.pack(fill="both", expand=True, pady=20)


    text_widget = tk.Text(
        text_frame,
        wrap="word",
        font=("Arial", 14),
        bg="#001F3F",
        fg="white",
        padx=15,
        pady=15,
        relief="flat"
    )

    text_widget.pack(side="left", fill="both", expand=True)


    scrollbar = tk.Scrollbar(text_frame, command=text_widget.yview)

    scrollbar.pack(side="right", fill="y")

    text_widget.config(yscrollcommand=scrollbar.set)

    # Insert text with tags for styling :

    text_widget.insert("end", "Project Overview:\n", "heading")
    text_widget.insert("end", "\n")
    text_widget.insert("end", "This project is a command-line-based Python application that automates the process of:\n\n")
    text_widget.insert("end", "1) Loading multiple Excel files (.xls or .xlsx) from a given folder\n", "subheading")
    text_widget.insert("end", "2) Merging and cleaning the data (removing duplicates)\n", "subheading")
    text_widget.insert("end", "3) Saving the cleaned result to a new Excel file and MySQL Database\n", "subheading")
    text_widget.insert("end", "4) Performing interactive visualizations on selected columns\n", "subheading")
    text_widget.insert("end", "\nIt uses powerful Python libraries like pandas, seaborn, and matplotlib for data processing and visualization.\n\n")

    text_widget.insert("end", "Steps Involved:\n", "heading")

    text_widget.insert("end", "\n")


    steps = [
        ("1. Importing Required Libraries\n", 
         "The application imports essential Python libraries such as pandas, seaborn, matplotlib, tkinter, os, glob, and mysql.connector to handle data processing, visualization, GUI, file operations, and database storage.\n\n"),
        ("2. Get Excel Files from Folder\n",
         "The user uses the GUI to select a folder path. The program automatically collects all .xlsx files from the specified location using the glob module.\n\n"),
        ("3. Load Data from Excel Files\n",
         "Each Excel file is read into a pandas DataFrame. All loaded DataFrames are combined for further processing.\n\n"),
        ("4. Merge and Clean DataFrames\n",
         "All DataFrames are concatenated into a single DataFrame using pd.concat(). Duplicate rows are removed using drop_duplicates(), and missing values are also removed using dropna() to ensure data quality.\n\n"),
        ("5. Save the Final Cleaned File\n",
         "The merged and cleaned DataFrame is saved as a new Excel file via the GUI with pandas.to_excel().\n\n"),
        ("6. Save to MySQL Database\n",
         "The cleaned DataFrame is dynamically mapped to a new MySQL table. The program automatically detects the column types and creates the table with proper column definitions, then inserts the cleaned data into the table.\n\n"),
        ("7. Load Final File for Visualization\n",
         "The user can select the cleaned Excel file again for visualization using the GUI.\n\n"),
        ("8. Visualize the Data\n",
         "The application supports multiple interactive visualizations, including:\n"
         "- Line plot\n- Bar chart\n- Pie chart\n- Scatter plot\n- Histogram\n- Correlation heatmap\n"
         "Users can choose which type of plot to generate with just a click.\n")
    ]

    for title, desc in steps:

        text_widget.insert("end", title, "subheading")

        text_widget.insert("end", desc)

        text_widget.insert("end", "\n") 


    # Add tags :

    text_widget.tag_configure("heading", font=("Arial", 18, "bold"), foreground="lightgreen")

    text_widget.tag_configure("subheading", font=("Arial", 14, "bold"), foreground="cyan")

    text_widget.config(state="disabled")


    # ---------------------------------------------------------------------------------------------------------
    #                                        Tab 2 : ( Data Cleaning )
    # ---------------------------------------------------------------------------------------------------------

    tab2 = tk.Frame(notebook, bg="#003366")

    notebook.add(tab2, text="Data Cleaning")


    clean_frame = tk.Frame(tab2, bg="#003366", padx=30, pady=30)

    clean_frame.pack(fill="both", expand=True)


    ttk.Label(clean_frame, text="Select Folder:", background="#003366", foreground="yellow",
              font=("Arial", 14, "bold")).pack(pady=5)
    

    folder_entry = tk.Entry(clean_frame, width=60, font=("Arial", 12))


    folder_entry.pack(pady=5)


    ttk.Button(clean_frame, text="Browse", command=lambda: folder_entry.insert(0, filedialog.askdirectory())).pack(pady=5)

    ttk.Label(clean_frame, text="Save Cleaned Excel File To:", background="#003366", foreground="yellow",
              font=("Arial", 14, "bold")).pack(pady=5)


    cleaned_entry = tk.Entry(clean_frame, width=60, font=("Arial", 12))

    cleaned_entry.pack(pady=5)


    ttk.Button(clean_frame, text="Browse", command=lambda: cleaned_entry.insert(0, filedialog.asksaveasfilename(defaultextension=".xlsx"))).pack(pady=5)


    ttk.Label(clean_frame, text="Enter MySQL Table Name:", background="#003366", foreground="yellow",
              font=("Arial", 14, "bold")).pack(pady=5)


    table_name_entry = tk.Entry(clean_frame, width=30, font=("Arial", 12))

    table_name_entry.pack(pady=5)


    # Clean & Upload button directly below table name
    ttk.Button(clean_frame, text="Clean & Upload", command=lambda: clean_and_upload()).pack(pady=15)

    # Treeview preview after the button
    preview_frame = tk.Frame(clean_frame, bg="#003366")

    preview_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # add border style to the Treeview
    style = ttk.Style()

    style.configure("Custom.Treeview", font=("Arial", 12), rowheight=30, borderwidth=1, relief="solid")

    style.configure("Custom.Treeview.Heading", font=("Arial", 12, "bold"))


    preview = ttk.Treeview(preview_frame, style="Custom.Treeview", show="headings")

    preview.pack(fill="both", expand=True)


    # Add vertical + horizontal scrollbars
    vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=preview.yview)

    vsb.pack(side="right", fill="y")

    preview.configure(yscrollcommand=vsb.set)

    hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=preview.xview)

    hsb.pack(side="bottom", fill="x")

    preview.configure(xscrollcommand=hsb.set)


    def clean_and_upload():

        folder = folder_entry.get()

        cleaned_file = cleaned_entry.get()

        table_name = table_name_entry.get()


        if not folder or not cleaned_file or not table_name:

            messagebox.showwarning("Missing", "Please fill all fields.")

            return


        df = load_and_merge_excels(folder)


        if df.empty:

            return


        cleaned_df = clean_data(df)

        save_cleaned_excel(cleaned_df, cleaned_file)

        save_to_mysql(cleaned_df, table_name)


        # update preview :

        preview.delete(*preview.get_children())

        preview["columns"] = list(cleaned_df.columns)


        for col in cleaned_df.columns:

            preview.heading(col, text=col, anchor="center")

            preview.column(col, anchor="center", width=150, stretch=True)


        for _, row in cleaned_df.iterrows():

            preview.insert("", "end", values=list(row))



# ---------------------------------------------------------------------------------------------------
#                                     Tab 3: Visualizations
# ---------------------------------------------------------------------------------------------------


    tab3 = tk.Frame(notebook, bg="#003300")

    notebook.add(tab3, text="Visualizations")



    viz_frame = tk.Frame(tab3, bg="#003300", padx=30, pady=30)

    viz_frame.pack(fill="both", expand=True)



    ttk.Label(viz_frame, text="Upload Cleaned Excel for Visualization", background="#003300", foreground="white",
              font=("Arial", 16, "bold")).pack(pady=15)



    cleaned_file_entry = tk.Entry(viz_frame, width=60, font=("Arial", 12))


    cleaned_file_entry.pack(pady=5)


    ttk.Button(viz_frame, text="Browse", command=lambda: cleaned_file_entry.insert(0, filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx")]))).pack(pady=5)


    # --------------------------------------

    def safe_plot(plot_func):


        cleaned_file = cleaned_file_entry.get()


        if not cleaned_file:

            messagebox.showwarning("Missing", "Please select a cleaned Excel file.")

            return
        

        df = pd.read_excel(cleaned_file)

        plot_func(df)



    plot_frame = tk.Frame(viz_frame, bg="#003300")

    plot_frame.pack(pady=20)



    ttk.Button(plot_frame, text="Line Plot", command=lambda: safe_plot(plot_line)).grid(row=0, column=0, padx=15, pady=10)

    ttk.Button(plot_frame, text="Bar Plot", command=lambda: safe_plot(plot_bar)).grid(row=0, column=1, padx=15, pady=10)

    ttk.Button(plot_frame, text="Pie Chart", command=lambda: safe_plot(plot_pie)).grid(row=0, column=2, padx=15, pady=10)

    ttk.Button(plot_frame, text="Scatter Plot", command=lambda: safe_plot(plot_scatter)).grid(row=1, column=0, padx=15, pady=10)

    ttk.Button(plot_frame, text="Histogram", command=lambda: safe_plot(plot_hist)).grid(row=1, column=1, padx=15, pady=10)

    ttk.Button(plot_frame, text="Heatmap", command=lambda: safe_plot(plot_heatmap)).grid(row=1, column=2, padx=15, pady=10)



# ---------------------------------------------------------------------------------------------------------
#                                        Tab 4: Conclusion
# ---------------------------------------------------------------------------------------------------------


    tab4 = tk.Frame(notebook, bg="#4B0000")

    notebook.add(tab4, text="Conclusion")


    conclusion_frame = tk.Frame(tab4, bg="#4B0000", padx=40, pady=40)

    conclusion_frame.pack(fill="both", expand=True)


    # Title
    tk.Label(
        conclusion_frame,
        text="End Results",
        fg="yellow",
        bg="#4B0000",
        font=("Arial", 26, "bold", "underline")
    ).pack(pady=20)


    # Subtitle
    tk.Label(
        conclusion_frame,
        text="Conclusion / End Results:",
        fg="lightgreen",
        bg="#4B0000",
        font=("Arial", 18, "bold")
    ).pack(pady=15)


    # Use a scrollable text area for bullets 

    text_frame = tk.Frame(conclusion_frame, bg="#4B0000")

    text_frame.pack(fill="both", expand=True, pady=20)


    text_widget = tk.Text(
        text_frame,
        wrap="word",
        font=("Arial", 14),
        bg="#4B0000",
        fg="white",
        padx=20,
        pady=20,
        relief="flat",
        spacing3=15  # extra spacing between bullet points
    )

    text_widget.pack(side="left", fill="both", expand=True)

    scrollbar = tk.Scrollbar(text_frame, command=text_widget.yview)

    scrollbar.pack(side="right", fill="y")

    text_widget.config(yscrollcommand=scrollbar.set)

    conclusion_text = (

        "\u2022 This project offers an end-to-end data integration and visualization workflow, all driven from a graphical interface.\n\n"
        "\u2022 It reduces manual errors in data cleaning and ensures a repeatable, consistent process for handling Excel files.\n\n"
        "\u2022 Future upgrades can include advanced statistical testing, handling categorical variables, exporting visualizations, "
        "or deploying the app as a web service with frameworks like Streamlit or Flask.\n\n"
        "\u2022 This tool is a strong portfolio project for demonstrating data engineering, data visualization, and Python GUI development skills."
    )

    text_widget.insert("1.0", conclusion_text)

    text_widget.config(state="disabled")

    root.mainloop()


# --------------------------------------------------------------------------------------------------------------------
#                                                   MAIN
# ---------------------------------------------------------------------------------------------------------------------


if __name__ == "__main__":

    main_gui()



# --------------------------------------------------------------------------------------------------------------------