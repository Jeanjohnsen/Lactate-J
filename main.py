import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

class LactateTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lactate Test App")

        self.data = {"stage": [], "lactate": [], "heart_rate": [], "power": []}

        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create tabs
        self.create_data_input_tab()
        self.create_graph_tab()

    def create_data_input_tab(self):
        # Tab for data input
        self.data_input_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.data_input_frame, text='Data Input')

        # Input fields
        input_frame = ttk.Frame(self.data_input_frame)
        input_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(input_frame, text="Stage:").grid(column=0, row=0, padx=5, pady=5)
        self.stage_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.stage_var).grid(column=1, row=0, padx=5, pady=5)

        ttk.Label(input_frame, text="Lactate (mmol/L):").grid(column=0, row=1, padx=5, pady=5)
        self.lactate_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.lactate_var).grid(column=1, row=1, padx=5, pady=5)

        ttk.Label(input_frame, text="Heart Rate (bpm):").grid(column=0, row=2, padx=5, pady=5)
        self.hr_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.hr_var).grid(column=1, row=2, padx=5, pady=5)

        ttk.Label(input_frame, text="Power (W):").grid(column=0, row=3, padx=5, pady=5)
        self.power_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.power_var).grid(column=1, row=3, padx=5, pady=5)

        # Buttons for adding data, plotting data, and uploading Excel file
        button_frame = ttk.Frame(self.data_input_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(button_frame, text="Add Data", command=self.add_data).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(button_frame, text="Plot Data", command=self.plot_data).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(button_frame, text="Upload Excel", command=self.upload_excel).pack(side=tk.LEFT, padx=5, pady=5)

        # Labels and buttons for FTP, LT1, LT2, and FATmax calculations
        self.ftp_frame = ttk.Frame(self.data_input_frame)
        self.ftp_frame.pack(fill=tk.X, padx=10, pady=5)
        self.ftp_label = ttk.Label(self.ftp_frame, text="FTP: Not Calculated")
        self.ftp_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(self.ftp_frame, text="Calculate FTP", command=self.calculate_ftp).pack(side=tk.LEFT, padx=5)

        self.lt1_frame = ttk.Frame(self.data_input_frame)
        self.lt1_frame.pack(fill=tk.X, padx=10, pady=5)
        self.lt1_label = ttk.Label(self.lt1_frame, text="LT1: Not Calculated")
        self.lt1_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(self.lt1_frame, text="Calculate LT1", command=self.calculate_lt1).pack(side=tk.LEFT, padx=5)

        self.lt2_frame = ttk.Frame(self.data_input_frame)
        self.lt2_frame.pack(fill=tk.X, padx=10, pady=5)
        self.lt2_label = ttk.Label(self.lt2_frame, text="LT2: Not Calculated")
        self.lt2_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(self.lt2_frame, text="Calculate LT2", command=self.calculate_lt2).pack(side=tk.LEFT, padx=5)

        self.fatmax_frame = ttk.Frame(self.data_input_frame)
        self.fatmax_frame.pack(fill=tk.X, padx=10, pady=5)
        self.fatmax_label = ttk.Label(self.fatmax_frame, text="FATmax: Not Calculated")
        self.fatmax_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(self.fatmax_frame, text="Calculate FATmax", command=self.calculate_fatmax).pack(side=tk.LEFT, padx=5)

        # Data table
        self.tree = ttk.Treeview(self.data_input_frame, columns=("Stage", "Lactate", "Heart Rate", "Power"), show='headings')
        self.tree.heading("Stage", text="Stage")
        self.tree.heading("Lactate", text="Lactate (mmol/L)")
        self.tree.heading("Heart Rate", text="Heart Rate (bpm)")
        self.tree.heading("Power", text="Power (W)")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Scrollbar for the treeview
        tree_scroll = ttk.Scrollbar(self.data_input_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll.set)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def create_graph_tab(self):
        # Tab for displaying graphs
        self.graph_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.graph_frame, text='Graphs')

        # Scrollable area for plots
        self.canvas_frame = ttk.Frame(self.graph_frame)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(self.canvas_frame)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar = ttk.Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Frame inside canvas
        self.plot_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.plot_frame, anchor="nw")

        self.plot_frame.bind("<Configure>", self.on_frame_configure)

    def on_frame_configure(self, event):
        # Configure scroll region for canvas
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def add_data(self):
        # Add data from input fields to the table and internal data structure
        try:
            stage = int(self.stage_var.get())
            lactate = float(self.lactate_var.get())
            heart_rate = int(self.hr_var.get())
            power = int(self.power_var.get())

            self.data["stage"].append(stage)
            self.data["lactate"].append(lactate)
            self.data["heart_rate"].append(heart_rate)
            self.data["power"].append(power)

            self.tree.insert("", "end", values=(stage, lactate, heart_rate, power))

            # Clear input fields
            self.stage_var.set("")
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
                messagebox.showerror("Error", f"Failed to read Excel file: {e}")

    def load_data_from_dataframe(self, df):
        # Load data from a DataFrame into the application
        self.data = {"stage": [], "lactate": [], "heart_rate": [], "power": []}
        self.tree.delete(*self.tree.get_children())  # Clear existing data in the treeview

        # Check if "Stage" column exists, if not, generate stages
        if "Stage" not in df.columns:
            df["Stage"] = range(1, len(df) + 1)

        for index, row in df.iterrows():
            stage = row['Stage']
            lactate = row['Lactate']
            heart_rate = row['Heart Rate']
            power = row['Power']

            self.data["stage"].append(stage)
            self.data["lactate"].append(lactate)
            self.data["heart_rate"].append(heart_rate)
            self.data["power"].append(power)

            self.tree.insert("", "end", values=(stage, lactate, heart_rate, power))


    def calculate_ftp(self):
        # Calculate FTP
        ftp, _, _, _ = self.calculate_ftp_lt1_lt2_fatmax()
        if ftp is not None:
            self.ftp_label.config(text=f"FTP: {ftp:.2f} W")

    def calculate_lt1(self):
        # Calculate LT1
        _, lt1, _, _ = self.calculate_ftp_lt1_lt2_fatmax()
        if lt1 is not None:
            self.lt1_label.config(text=f"LT1: {lt1:.2f} W")

    def calculate_lt2(self):
        # Calculate LT2
        _, _, lt2, _ = self.calculate_ftp_lt1_lt2_fatmax()
        if lt2 is not None:
            self.lt2_label.config(text=f"LT2: {lt2:.2f} W")
    
    def calculate_fatmax(self):
    # Calculate FATmax
        _, _, _, fatmax = self.calculate_ftp_lt1_lt2_fatmax()
        if fatmax is not None:
            self.fatmax_label.config(text=f"FATmax: {fatmax:.2f} W")

    def calculate_ftp_lt1_lt2_fatmax(self):
        # Calculate FTP, LT1, LT2, and FATmax
        lactate = np.array(self.data["lactate"])
        power = np.array(self.data["power"])

        if len(lactate) < 4:
            messagebox.showerror("Insufficient data", "Please add more data points to calculate FTP, LT1, LT2, and FATmax.")
            return None, None, None, None

        # Example calculations (adjust as needed for your specific criteria)
        # LT1: The first significant rise in lactate (e.g., >0.5 mmol/L above baseline)
        baseline_lactate = lactate[0]
        lt1_index = np.argmax(lactate > (baseline_lactate + 0.5))
        lt1_power = power[lt1_index] if lt1_index > 0 else None

        # LT2: The point where lactate rises rapidly (e.g., 4 mmol/L)
        lt2_index = np.argmax(lactate >= 4)
        lt2_power = power[lt2_index] if lt2_index > 0 else None

        # FTP: Often approximated as the power at LT2 or a bit lower
        ftp_power = lt2_power if lt2_power else power[-1]

        # FATmax: Identified at moderate intensities before LT1
        fatmax_power = None
        if lt1_index > 0:
            # Assuming FATmax occurs at the highest power before LT1
            fatmax_power = power[:lt1_index].max() if lt1_index > 0 else None

        return ftp_power, lt1_power, lt2_power, fatmax_power

    def plot_data(self):
        # Plot the data
        for widget in self.plot_frame.winfo_children():
            widget.destroy()

        ftp, lt1, lt2, fatmax = self.calculate_ftp_lt1_lt2_fatmax()

        figure, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(8, 12))

        ax1.plot(self.data["stage"], self.data["lactate"], marker='o', color='blue')
        ax1.set_title('Lactate Levels')
        ax1.set_xlabel('Stage')
        ax1.set_ylabel('Lactate (mmol/L)')

        ax2.plot(self.data["stage"], self.data["heart_rate"], marker='o', color='red')
        ax2.set_title('Heart Rate')
        ax2.set_xlabel('Stage')
        ax2.set_ylabel('Heart Rate (bpm)')

        ax3.plot(self.data["stage"], self.data["power"], marker='o', color='green')
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
    
if __name__ == '__main__':
    root = tk.Tk()
    app = LactateTestApp(root)
    root.mainloop()