import os
import tempfile
import tkinter as tk
from tkinter import Canvas, ttk, messagebox, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import matplotlib
import numpy as np
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image

matplotlib.use('Agg')

class LactateTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lactate Test App")
        self.root.geometry("2560x1600")

        self.data = {"lactate": [], "heart_rate": [], "power": []}
        self.results = {"FTP": None, "LT1": None, "LT2": None, "FATmax": None}
        self.old_data = {"lactate": [], "heart_rate": [], "power": []}

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.create_data_input_tab()
        self.create_graph_tab()
        self.create_compare_tab()

        self.root.bind('<Configure>', self.on_resize)
        self.root.bind('<Return>', self.on_enter_key)

        self.editing_index = None
        self.editing_column = None
        self.entry = None

    def create_data_input_tab(self):
        self.data_input_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.data_input_frame, text='Data Input')

        input_frame = ttk.Frame(self.data_input_frame)
        input_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        input_frame.columnconfigure(1, weight=1)

        ttk.Label(input_frame, text="Lactate (mmol/L):").grid(column=0, row=0, padx=5, pady=5, sticky="w")
        self.lactate_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.lactate_var).grid(column=1, row=0, padx=5, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Heart Rate (bpm):").grid(column=0, row=1, padx=5, pady=5, sticky="w")
        self.hr_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.hr_var).grid(column=1, row=1, padx=5, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Power (W):").grid(column=0, row=2, padx=5, pady=5, sticky="w")
        self.power_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.power_var).grid(column=1, row=2, padx=5, pady=5, sticky="ew")

        button_frame = ttk.Frame(self.data_input_frame)
        button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        button_frame.columnconfigure((0, 1, 2, 3, 4), weight=1)

        ttk.Button(button_frame, text="Add Data", command=self.add_data).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ttk.Button(button_frame, text="Plot Data", command=self.plot_data).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(button_frame, text="Upload Excel", command=self.upload_excel).grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        ttk.Button(button_frame, text="Clear Data", command=self.clear_data).grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        ttk.Button(button_frame, text="Export to PDF", command=self.export_to_pdf).grid(row=0, column=4, padx=5, pady=5, sticky="ew")

        # Labels and buttons for FTP, LT1, LT2, and FATmax calculations
        self.ftp_frame = ttk.Frame(self.data_input_frame)
        self.ftp_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        self.ftp_label = ttk.Label(self.ftp_frame, text="FTP: Not Calculated")
        self.ftp_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Button(self.ftp_frame, text="Calculate FTP", command=self.calculate_ftp).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.lt1_frame = ttk.Frame(self.data_input_frame)
        self.lt1_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        self.lt1_label = ttk.Label(self.lt1_frame, text="LT1: Not Calculated")
        self.lt1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Button(self.lt1_frame, text="Calculate LT1", command=self.calculate_lt1).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.lt2_frame = ttk.Frame(self.data_input_frame)
        self.lt2_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
        self.lt2_label = ttk.Label(self.lt2_frame, text="LT2: Not Calculated")
        self.lt2_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Button(self.lt2_frame, text="Calculate LT2", command=self.calculate_lt2).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.fatmax_frame = ttk.Frame(self.data_input_frame)
        self.fatmax_frame.grid(row=5, column=0, sticky="ew", padx=10, pady=5)
        self.fatmax_label = ttk.Label(self.fatmax_frame, text="FATmax: Not Calculated")
        self.fatmax_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Button(self.fatmax_frame, text="Calculate FATmax", command=self.calculate_fatmax).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.tree = ttk.Treeview(self.data_input_frame, columns=("Lactate", "Heart Rate", "Power"), show='headings')
        self.tree.heading("Lactate", text="Lactate (mmol/L)")
        self.tree.heading("Heart Rate", text="Heart Rate (bpm)")
        self.tree.heading("Power", text="Power (W)")
        self.tree.grid(row=6, column=0, sticky="nsew", padx=10, pady=5)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Return>", self.on_enter_key)

        tree_scroll = ttk.Scrollbar(self.data_input_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll.set)
        tree_scroll.grid(row=6, column=1, sticky="ns")

        self.data_input_frame.rowconfigure(10, weight=1)
        self.data_input_frame.columnconfigure(0, weight=1)

        ttk.Button(self.data_input_frame, text="Export Table to CSV", command=self.export_to_csv).grid(row=7, column=0, padx=5, pady=5, sticky="ew")
        ttk.Button(self.data_input_frame, text="Export Table to Excel", command=self.export_to_excel).grid(row=8, column=0, padx=5, pady=5, sticky="ew")

    def create_graph_tab(self):
        # Tab for displaying graphs
        self.graph_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.graph_frame, text='Graphs')

        # Scrollable area for plots
        self.canvas_frame = ttk.Frame(self.graph_frame)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.left_canvas = tk.Canvas(self.canvas_frame)
        self.left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.left_scrollbar = ttk.Scrollbar(self.canvas_frame, orient="vertical", command=self.left_canvas.yview)
        self.left_scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)

        self.right_canvas = tk.Canvas(self.canvas_frame)
        self.right_canvas.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.right_scrollbar = ttk.Scrollbar(self.canvas_frame, orient="vertical", command=self.right_canvas.yview)
        self.right_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.right_canvas.configure(yscrollcommand = self.right_scrollbar.set)
            # Frames inside canvases
        self.plot_frame = ttk.Frame(self.left_canvas)
        self.left_canvas.create_window((0, 0), window=self.plot_frame, anchor="nw")
        self.plot_frame.bind("<Configure>", self.on_frame_configure)

        self.compare_plot_frame = ttk.Frame(self.right_canvas)
        self.right_canvas.create_window((0, 0), window=self.compare_plot_frame, anchor="nw")
        self.compare_plot_frame.bind("<Configure>", self.on_frame_configure)

    def create_compare_tab(self):
        # Tab for comparing tests
        self.compare_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.compare_frame, text='Compare Tests')

        button_frame = ttk.Frame(self.compare_frame)
        button_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        button_frame.columnconfigure(0, weight=1)

        ttk.Button(button_frame, text="Upload Old Test", command=self.upload_old_test).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ttk.Button(button_frame, text="Compare Tests", command=self.compare_tests).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.old_tree = ttk.Treeview(self.compare_frame, columns=("Lactate", "Heart Rate", "Power"), show='headings')
        self.old_tree.heading("Lactate", text="Lactate (mmol/L)")
        self.old_tree.heading("Heart Rate", text="Heart Rate (bpm)")
        self.old_tree.heading("Power", text="Power (W)")
        self.old_tree.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        tree_scroll = ttk.Scrollbar(self.compare_frame, orient="vertical", command=self.old_tree.yview)
        self.old_tree.configure(yscroll=tree_scroll.set)
        tree_scroll.grid(row=1, column=1, sticky="ns")

        self.compare_frame.rowconfigure(1, weight=1)
        self.compare_frame.columnconfigure(0, weight=1)

    def on_frame_configure(self, event):
        # Configure scroll region for canvas
        self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
        self.right_canvas.configure(scrollregion=self.right_canvas.bbox("all"))

    def on_resize(self, event):
        self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
        self.right_canvas.configure(scrollregion=self.right_canvas.bbox("all"))

    def on_enter_key(self, event):
        if self.entry:
            self.update_cell()
        else:
            self.add_data()

    def on_double_click(self, event):
        # Start editing a cell on double-click
        item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        self.editing_index = int(self.tree.index(item))
        self.editing_column = int(column[1:]) - 1
        value = self.tree.item(item, "values")[self.editing_column]

        # Positioning the Entry widget over the cell
        x, y, width, height = self.tree.bbox(item, column)
        self.entry = tk.Entry(self.tree, width=width)
        self.entry.insert(0, value)
        self.entry.place(x=x, y=y, width=width, height=height)
        self.entry.focus()
        self.entry.bind("<Return>", self.update_cell)
        self.entry.bind("<FocusOut>", self.cancel_edit)

    def update_cell(self, event=None):
        # Update cell value and data structure
        if not self.entry:
            return
        new_value = self.entry.get()
        try:
            if self.editing_column == 0:
                new_value = float(new_value)
                self.data["lactate"][self.editing_index] = new_value
            elif self.editing_column == 1:
                new_value = int(new_value)
                self.data["heart_rate"][self.editing_index] = new_value
            elif self.editing_column == 2:
                new_value = int(new_value)
                self.data["power"][self.editing_index] = new_value
            self.tree.item(self.tree.selection()[0], values=(
                self.data["lactate"][self.editing_index],
                self.data["heart_rate"][self.editing_index],
                self.data["power"][self.editing_index]
            ))
        except ValueError:
            messagebox.showerror("Invalid input", "Please enter a valid number.")
        finally:
            self.entry.destroy()
            self.entry = None

    def cancel_edit(self, event=None):
        # Cancel editing and remove the entry widget
        if self.entry:
            self.entry.destroy()
            self.entry = None

    def add_data(self):
        # Add data from input fields to the table and internal data structure
        lactate = self.lactate_var.get()
        heart_rate = self.hr_var.get()
        power = self.power_var.get()

        if lactate and heart_rate and power:
            try:
                lactate = float(lactate)
                heart_rate = int(heart_rate)
                power = int(power)

                self.data["lactate"].append(lactate)
                self.data["heart_rate"].append(heart_rate)
                self.data["power"].append(power)

                self.tree.insert("", "end", values=(lactate, heart_rate, power))

                # Clear input fields
                self.lactate_var.set("")
                self.hr_var.set("")
                self.power_var.set("")
            except ValueError:
                messagebox.showerror("Invalid input", "Please enter valid data for all fields.")

    def upload_excel(self):
        # Upload data from an Excel file
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.load_data_from_dataframe(df)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel file: {e}")

    def load_data_from_dataframe(self, df):
        # Load data from a DataFrame
        for _, row in df.iterrows():
            lactate = row['Lactate']
            heart_rate = row['Heart Rate']
            power = row['Power']
            self.data["lactate"].append(lactate)
            self.data["heart_rate"].append(heart_rate)
            self.data["power"].append(power)
            self.tree.insert("", "end", values=(lactate, heart_rate, power))

    def clear_data(self):
        # Clear all data
        for key in self.data:
            self.data[key].clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.ftp_label.config(text="FTP: Not Calculated")
        self.lt1_label.config(text="LT1: Not Calculated")
        self.lt2_label.config(text="LT2: Not Calculated")
        self.fatmax_label.config(text="FATmax: Not Calculated")

    def calculate_ftp(self):
        ftp, _, _, _ = self.calculate_ftp_lt1_lt2_fatmax()
        if ftp is not None:
            self.ftp_label.config(text=f"FTP: {ftp:.2f} W")

    def calculate_lt1(self):
        _, lt1, _, _ = self.calculate_ftp_lt1_lt2_fatmax()
        if lt1 is not None:
            self.lt1_label.config(text=f"LT1: {lt1:.2f} W")

    def calculate_lt2(self):
        _, _, lt2, _ = self.calculate_ftp_lt1_lt2_fatmax()
        if lt2 is not None:
            self.lt2_label.config(text=f"LT2: {lt2:.2f} W")

    def calculate_fatmax(self):
        _, _, _, fatmax = self.calculate_ftp_lt1_lt2_fatmax()
        if fatmax is not None:
            self.fatmax_label.config(text=f"FATmax: {fatmax:.2f} W")

    def calculate_ftp_lt1_lt2_fatmax(self):
        # Calculate FTP, LT1, LT2, and FATmax based on lactate and power data
        lactate = np.array(self.data["lactate"])
        power = np.array(self.data["power"])

        if len(lactate) < 4:
            messagebox.showerror("Insufficient data", "Please add more data points to calculate FTP, LT1, LT2, and FATmax.")
            return None, None, None, None

        baseline_lactate = lactate[0]

        lt1_index = np.argmax(lactate > (baseline_lactate + 0.5))
        lt1_power = power[lt1_index] if lt1_index > 0 else None

        lt2_index = np.argmax(lactate >= 4)
        lt2_power = power[lt2_index] if lt2_index > 0 else None

        ftp_power = lt2_power * 0.95 if lt2_power else power[-1]

        fatmax_power = None
        if lt1_index > 0:
            fatmax_power = power[:lt1_index].max() if lt1_index > 0 else None

        return ftp_power, lt1_power, lt2_power, fatmax_power

    def plot_data(self):
        for widget in self.plot_frame.winfo_children():
            widget.destroy()

        ftp, lt1, lt2, fatmax = self.calculate_ftp_lt1_lt2_fatmax()

        figure, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(8, 12))
        
        stages = list(range(1, len(self.data["lactate"]) + 1))

        ax1.plot(stages, self.data["lactate"], marker='o', color='blue')
        ax1.set_title('Lactate Levels')
        ax1.set_xlabel('Stage')
        ax1.set_ylabel('Lactate (mmol/L)')

        ax2.plot(stages, self.data["heart_rate"], marker='o', color='red')
        ax2.set_title('Heart Rate')
        ax2.set_xlabel('Stage')
        ax2.set_ylabel('Heart Rate (bpm)')

        ax3.plot(stages, self.data["power"], marker='o', color='green')
        ax3.set_title('Power Output')
        ax3.set_xlabel('Stage')
        ax3.set_ylabel('Power (W)')

        if ftp is not None:
            ax3.axhline(y=ftp, color='blue', linestyle='--', label=f'FTP: {ftp:.2f} W')
        if lt1 is not None:
            ax3.axhline(y=lt1, color='orange', linestyle='--', label=f'LT1: {lt1:.2f} W')
        if lt2 is not None:
            ax3.axhline(y=lt2, color='purple', linestyle='--', label=f'LT2: {lt2:.2f} W')
        if fatmax is not None:
            ax3.axhline(y=fatmax, color='cyan', linestyle='--', label=f'FATmax: {fatmax:.2f} W')

        ax3.legend()

        figure.tight_layout()
        canvas = FigureCanvasTkAgg(figure, self.plot_frame)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        canvas.draw()
        
    def export_to_pdf(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not file_path:
            return

        # Recalculate the results before exporting
        ftp, lt1, lt2, fatmax = self.calculate_ftp_lt1_lt2_fatmax()
        self.results['FTP'] = ftp
        self.results['LT1'] = lt1
        self.results['LT2'] = lt2
        self.results['FATmax'] = fatmax

        pdf_doc = SimpleDocTemplate(file_path, pagesize=letter)
        elements = []

        styles = getSampleStyleSheet()
        title_style = styles['Title']
        normal_style = styles['BodyText']

        elements.append(Paragraph("Lactate Test Results", title_style))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(f"FTP: {ftp:.2f} W" if ftp is not None else "FTP: Not Calculated", normal_style))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"LT1: {lt1:.2f} W" if lt1 is not None else "LT1: Not Calculated", normal_style))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"LT2: {lt2:.2f} W" if lt2 is not None else "LT2: Not Calculated", normal_style))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"FATmax: {fatmax:.2f} W" if fatmax is not None else "FATmax: Not Calculated", normal_style))
        elements.append(Spacer(1, 24))

        # Create the figure and save it as an image
        figure, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(8, 12))
        stages = list(range(1, len(self.data["lactate"]) + 1))

        ax1.plot(stages, self.data["lactate"], marker='o', color='blue')
        ax1.set_title('Lactate Levels')
        ax1.set_xlabel('Stage')
        ax1.set_ylabel('Lactate (mmol/L)')

        ax2.plot(stages, self.data["heart_rate"], marker='o', color='red')
        ax2.set_title('Heart Rate')
        ax2.set_xlabel('Stage')
        ax2.set_ylabel('Heart Rate (bpm)')

        ax3.plot(stages, self.data["power"], marker='o', color='green')
        ax3.set_title('Power Output')
        ax3.set_xlabel('Stage')
        ax3.set_ylabel('Power (W)')

        if ftp is not None:
            ax3.axhline(y=ftp, color='blue', linestyle='--', label=f'FTP: {ftp:.2f} W')
        if lt1 is not None:
            ax3.axhline(y=lt1, color='orange', linestyle='--', label=f'LT1: {lt1:.2f} W')
        if lt2 is not None:
            ax3.axhline(y=lt2, color='purple', linestyle='--', label=f'LT2: {lt2:.2f} W')
        if fatmax is not None:
            ax3.axhline(y=fatmax, color='cyan', linestyle='--', label=f'FATmax: {fatmax:.2f} W')
            ax3.legend()

        figure.tight_layout()

        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmpfile:
            pdf_image_path = tmpfile.name
            figure.savefig(pdf_image_path, bbox_inches='tight')
            plt.close(figure)

        # Add image to PDF
        elements.append(Image(pdf_image_path, width=500, height=600, kind='proportional'))

        # Build the PDF
        pdf_doc.build(elements)

        # Clean up the temporary file
        if os.path.exists(pdf_image_path):
            os.remove(pdf_image_path)

    def export_to_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        df = pd.DataFrame(self.data)
        df.to_csv(file_path, index=False)

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        df = pd.DataFrame(self.data)
        df.to_excel(file_path, index=False)

    def upload_old_test(self):
        # Upload old test data from an Excel file
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.load_old_data_from_dataframe(df)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel file: {e}")

    def load_old_data_from_dataframe(self, df):
        # Load old data from a DataFrame
        self.old_data = {"lactate": [], "heart_rate": [], "power": []}
        for _, row in df.iterrows():
            lactate = row['Lactate']
            heart_rate = row['Heart Rate']
            power = row['Power']
            self.old_data["lactate"].append(lactate)
            self.old_data["heart_rate"].append(heart_rate)
            self.old_data["power"].append(power)
            self.old_tree.insert("", "end", values=(lactate, heart_rate, power))

    def compare_tests(self):
        # Plot comparison of old and new test data
        if not self.old_data["lactate"]:
            messagebox.showerror("Error", "No old test data available for comparison.")
            return

        for widget in self.compare_plot_frame.winfo_children():
            widget.destroy()

        figure, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(8, 12))

        stages_new = list(range(1, len(self.data["lactate"]) + 1))
        stages_old = list(range(1, len(self.old_data["lactate"]) + 1))

        ax1.plot(stages_new, self.data["lactate"], marker='o', color='blue', label='New Test')
        ax1.plot(stages_old, self.old_data["lactate"], marker='x', color='green', label='Old Test')
        ax1.set_title('Lactate Levels')
        ax1.set_xlabel('Stage')
        ax1.set_ylabel('Lactate (mmol/L)')
        ax1.legend()

        ax2.plot(stages_new, self.data["heart_rate"], marker='o', color='red', label='New Test')
        ax2.plot(stages_old, self.old_data["heart_rate"], marker='x', color='orange', label='Old Test')
        ax2.set_title('Heart Rate')
        ax2.set_xlabel('Stage')
        ax2.set_ylabel('Heart Rate (bpm)')
        ax2.legend()

        ax3.plot(stages_new, self.data["power"], marker='o', color='green', label='New Test')
        ax3.plot(stages_old, self.old_data["power"], marker='x', color='purple', label='Old Test')
        ax3.set_title('Power Output')
        ax3.set_xlabel('Stage')
        ax3.set_ylabel('Power (W)')
        ax3.legend()
        
        figure.tight_layout()
        canvas = FigureCanvasTkAgg(figure, self.compare_plot_frame)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        canvas.draw()
        

# THIS RUNS THE PROGRAM        
if __name__ == '__main__':
    root = tk.Tk()
    app = LactateTestApp(root)
    root.mainloop()