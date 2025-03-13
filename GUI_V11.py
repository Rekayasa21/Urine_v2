import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap import Treeview
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import serial
import time
from datetime import datetime
import sqlite3
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import os
import webbrowser
import logging

# Configure logging
logging.basicConfig(level=logging.ERROR, filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')

class UroflowmetryApp:
    """
    UroflowmetryApp is a Python application for uroflowmetry analysis.
    It includes functionalities for serial communication, data plotting,
    database interaction, and PDF report generation.
    """
    def __init__(self):
        """
        Initializes the UroflowmetryApp.
        Sets up the main window, serial port, plot data, UI components,
        database, and starts the plot update loop.
        """
        self.root = ttk.Window(themename="flatly")
        self.root.title("UROFLOWMETRY")
        self.root.geometry("1024x600")

        # Serial Port Configuration
        self.serial_port = None
        self.connect_serial()

        # Plot Data
        self.x_data = []
        self.y1_data = []
        self.y2_data = []
        self.is_running = False

        # Initialize UI Components
        self.create_ui()
        self.initialize_database()
        self.update_history()

        # Start Plot Update Loop
        self.root.after(100, self.update_plot)

    def connect_serial(self):
        """
        Attempts to connect to the serial port.
        Logs and shows an error message if connection fails.
        """
        try:
            self.serial_port = serial.Serial('COM3', 115200, timeout=1)
            logging.info(f"Connected to COM 3")
            time.sleep(2)
        except serial.SerialException as e:
            logging.error(f"Failed to connect to COM 3: {str(e)}")
            Messagebox.show_error("Serial Error", f"Failed to connect to COM3: {str(e)}")
            self.root.destroy()
        except Exception as e:
            logging.error(f"Unexpected error: {str(e)}")
            Messagebox.show_error("Unexpected Error", f"An unexpected error occurred: {str(e)}")
            self.root.destroy()

    def create_ui(self):
        """
        Sets up the user interface components including the notebook and frames.
        """
        # Notebook Setup
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill=BOTH)

        self.home_frame = ttk.Frame(self.notebook)
        self.history_frame = ttk.Frame(self.notebook)
        # self.settings_frame = ttk.Frame(self.notebook)
        self.about_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.home_frame, text='Home')
        self.notebook.add(self.history_frame, text='History')
        # self.notebook.add(self.settings_frame, text='Settings')
        self.notebook.add(self.about_frame, text='About')

        # Home Frame Widgets
        self.create_home_ui()
        # History Frame Widgets
        self.create_history_ui()
        # About Frame Widgets
        self.create_about_ui()

    def create_home_ui(self):
        """
        Sets up the UI components for the Home frame, including plots and buttons.
        """
        # Plot Setup
        self.fig, (self.ax1, self.ax2) = plt.subplots(2, 1, figsize=(3, 3))
        self.ax1.set_xlim(0, 60)
        self.ax1.set_ylim(0, 300)
        self.ax1.set_xlabel('Time (s)')
        self.ax1.set_ylabel('Flowrate (mL/s)')
        self.line1, = self.ax1.plot([], [], label='Flowrate', color='blue')
        self.ax1.legend()

        self.ax2.set_xlim(0, 60)
        self.ax2.set_ylim(0, 300)
        self.ax2.set_xlabel('Time (s)')
        self.ax2.set_ylabel('Volume (mL)')
        self.line2, = self.ax2.plot([], [], label='Volume', color='orange')
        self.ax2.legend()

        self.ax1.grid(True)
        self.ax2.grid(True)
        plt.subplots_adjust(top=0.95, hspace=0.25)

        # Canvas for Plot
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.home_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=True)

        # Labels
        self.y1_label = ttk.Label(self.home_frame, text="Flowrate: 0 mL/s")
        self.y2_label = ttk.Label(self.home_frame, text="Volume: 0 mL")
        self.y1_label.pack(side=LEFT, padx=10)
        self.y2_label.pack(side=LEFT, padx=10)

        # Buttons
        buttons = [
            ("Start", self.start_action),
            ("Stop", self.stop_action),
            ("Clear", self.clear_action),
            ("Save", self.save_data)
        ]
        for text, command in buttons:
            btn = ttk.Button(self.home_frame, text=text, command=command)
            btn.pack(side=LEFT, padx=10)

    def create_history_ui(self):
        """
        Sets up the UI components for the History frame, including the treeview and buttons.
        """
        # Treeview for History
        self.history_treeview = Treeview(
            self.history_frame,
            columns=("id_nomor", "ID", "Name", "Date"),
            show="headings",
            height=25
        )
        self.history_treeview.heading("id_nomor", text="No")
        self.history_treeview.heading("ID", text="Patient ID")
        self.history_treeview.heading("Name", text="Name")
        self.history_treeview.heading("Date", text="Date")
        self.history_treeview.grid(row=0, column=0, sticky="nsew")

        # Center align the columns
        for col in ("id_nomor", "ID", "Name", "Date"):
            self.history_treeview.column(col, anchor='center')  # Center align the column data

        self.history_treeview.grid(row=0, column=0, sticky="nsew")

        # Buttons Frame
        button_frame = ttk.Frame(self.history_frame)
        button_frame.grid(row=0, column=1, sticky="n")

        buttons = [
            ("View PDF", self.open_pdf),
            ("Delete", self.delete_measurement),
            ("Print", self.print_measurement)
        ]
        for text, command in buttons:
            btn = ttk.Button(button_frame, text=text, command=command)
            btn.pack(pady=5, fill=X)

        # Grid Configuration
        self.history_frame.grid_rowconfigure(0, weight=1)
        self.history_frame.grid_columnconfigure(0, weight=1)

    def create_about_ui(self):
        """
        Sets up the UI components for the About frame.
        """
        about_text = "Uroflowmetry App v1.0.0\nCreated by PT Edison Medika Kreasindo"
        about_label = ttk.Label(self.about_frame, text=about_text, font=('Arial', 14))
        about_label.pack(pady=50)

    def initialize_database(self):
        """
        Initializes the SQLite database and creates the patients table if it doesn't exist.
        Logs and shows an error message if initialization fails.
        """
        try:
            with sqlite3.connect('V3_UROFLOWMETRY_DATABASE.db') as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS patients (
                        id_nomor INTEGER PRIMARY KEY AUTOINCREMENT,
                        first_name TEXT NOT NULL,
                        last_name TEXT NOT NULL,
                        age_patient INTEGER,
                        patient_id_num INTEGER NOT NULL,
                        gender_patient TEXT NOT NULL,
                        doctor_first_name TEXT NOT NULL,
                        doctor_last_name TEXT NOT NULL,
                        hospital_name TEXT NOT NULL,
                        hospital_address TEXT NOT NULL,
                        created_at TEXT NOT NULL
                        
                    )
                ''')
                logging.info("Database and table initialized successfully.")
        except sqlite3.Error as e:
            logging.error(f"Failed to initialize database: {str(e)}")
            Messagebox.show_error("Database Error", f"Failed to initialize database: {str(e)}")
            self.root.destroy()
        except Exception as e:
            logging.error(f"Unexpected error: {str(e)}")
            Messagebox.show_error("Unexpected Error", f"An unexpected error occurred: {str(e)}")
            self.root.destroy()

    def update_plot(self):
        """
        Updates the plot with new data from the serial port.
        Logs and shows an error message if an exception occurs.
        """
        if self.is_running and self.serial_port and self.serial_port.is_open:
            try:
                if self.serial_port.in_waiting > 0:
                    data = self.serial_port.readline().decode('utf-8').strip().split(',')
                    if len(data) == 2:
                        y1, y2 = map(float, data)
                        current_time = time.time() - self.start_time

                        self.x_data.append(current_time)
                        self.y1_data.append(y1)
                        self.y2_data.append(y2)

                        self.line1.set_data(self.x_data, self.y1_data)
                        self.line2.set_data(self.x_data, self.y2_data)

                        self.y1_label.config(text=f"Flowrate: {y1:.2f} mL/s")
                        self.y2_label.config(text=f"Volume: {y2:.2f} mL")

                        # Auto-scale x-axis
                        if current_time > 60:
                            self.ax1.set_xlim(current_time - 60, current_time)
                            self.ax2.set_xlim(current_time - 60, current_time)

                        self.canvas.draw()
            except Exception as e:
                logging.error(f"Plot Update Error: {str(e)}")
                print(f"Plot Update Error: {str(e)}")

        self.root.after(100, self.update_plot)

    def start_action(self):
        """
        Initializes plot data and starts serial communication.
        """
        self.initialize_plot_data()
        self.start_serial_communication()

    def initialize_plot_data(self):
        """
        Initializes plot data and resets the plot.
        """
        self.is_running = True
        self.start_time = time.time()
        self.x_data.clear()
        self.y1_data.clear()
        self.y2_data.clear()
        self.line1.set_data([], [])
        self.line2.set_data([], [])
        self.canvas.draw()

    def start_serial_communication(self):
        """
        Sends a start command to the serial port.
        """
        if self.serial_port:
            self.serial_port.write(b'A')  # Start command

    def stop_action(self):
        """
        Stops the data collection process.
        """
        self.is_running = False

    def clear_action(self):
        """
        Clears the plot data and resets the labels.
        """
        self.x_data.clear()
        self.y1_data.clear()
        self.y2_data.clear()
        self.line1.set_data([], [])
        self.line2.set_data([], [])
        self.y1_label.config(text="Flowrate: 0 mL/s")
        self.y2_label.config(text="Volume: 0 mL")
        self.canvas.draw()

    def save_data(self):
        """
        Opens a form to save patient data and generates a PDF report.
        """
        def submit():
            try:
                # Get Form Data
                first = first_name.get()
                last = last_name.get()
                age = age_patient.get()  # Correctly retrieve the age from the Spinbox
                patient_id = patient_id_num.get()  # Correctly retrieve the patient ID
                gender = gender_patient.get()
                doctor_first = doctor_first_name.get()
                doctor_last = doctor_last_name.get()
                hospital = hospital_name.get()
                address = hospital_address.get()

                if not (first and last and age and patient_id and gender and doctor_first and doctor_last and hospital and address):
                    raise ValueError("All fields are required")

                # Save to Database
                with sqlite3.connect('V3_UROFLOWMETRY_DATABASE.db') as conn:
                    cursor = conn.cursor()
                    cursor.execute('''
                        INSERT INTO patients (
                            first_name, 
                            last_name, 
                            age_patient, 
                            patient_id_num, 
                            gender_patient,
                            doctor_first_name, 
                            doctor_last_name,
                            hospital_name, 
                            hospital_address,
                            created_at
                            
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        first,
                        last,
                        age,
                        patient_id,
                        gender,
                        doctor_first,
                        doctor_last,
                        hospital,
                        address,
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    ))

                # Generate PDF
                self.generate_pdf(patient_id, first, last, gender, hospital, address, doctor_first, doctor_last)

                # Update History
                self.update_history()
                entry_window.destroy()

                Messagebox.show_info("Success", "Data saved and PDF created")
            except Exception as e:
                logging.error(f"Error saving data: {str(e)}")
                Messagebox.show_error(f"Error SAVE DATA: {str(e)}", f"Failed to save data: {str(e)}")

        entry_window = ttk.Toplevel(self.root)
        entry_window.title("Save Patient Data")

        # Form Fields
        fields = [
            ("First Name", ttk.Entry, "first_name", first_name := ttk.StringVar()),
            ("Last Name", ttk.Entry, "last_name", last_name := ttk.StringVar()),
            ("Age", ttk.Spinbox, "age_patient", age_patient := ttk.IntVar(value=0)),
            ("Patient ID", ttk.Entry, "patient_id_num", patient_id_num := ttk.StringVar()),
            ("Gender", ttk.Combobox, "gender_patient", gender_patient := ttk.StringVar(value="Male")),
            ("Doctor First Name", ttk.Entry, "doctor_first", doctor_first_name := ttk.StringVar()),
            ("Doctor Last Name", ttk.Entry, "doctor_last", doctor_last_name := ttk.StringVar()),
            ("Hospital Name", ttk.Entry, "hospital_name", hospital_name := ttk.StringVar()),
            ("Address", ttk.Entry, "hospital_address", hospital_address := ttk.StringVar()),
        ]

        form_frame = ttk.Frame(entry_window)
        form_frame.pack(padx=20, pady=20)

        for i, (label_text, widget_type, var_name, var) in enumerate(fields):
            ttk.Label(form_frame, text=label_text).grid(row=i, column=0, sticky=W)
            if widget_type == ttk.Combobox:
                w = widget_type(form_frame, textvariable=var, values=["Male", "Female"])
            elif widget_type == ttk.Spinbox:
                w = widget_type(form_frame, from_=0, to=150, textvariable=var)
            else:
                w = widget_type(form_frame, textvariable=var)
            w.grid(row=i, column=1, pady=5)

        ttk.Button(entry_window, text="Save", command=submit).pack(pady=10)

    def generate_pdf(self, patient_id, first, last, gender, hospital, address, doctor_first, doctor_last):
        """
        Generates a PDF report with patient data and plots.
        Logs and shows an error message if an exception occurs.
        """
        try:
            # Save Plots as Images
            flow_img = "flow.png"
            volume_img = "volume.png"
            self.save_plots(flow_img, volume_img)

            # Create PDF
            filename = f"{patient_id}_{first}_{last}.pdf"
            doc = SimpleDocTemplate(filename, pagesize=A4)
            elements = []

            styles = getSampleStyleSheet()
            elements.append(Paragraph(f"<center>{hospital}</center>", styles['Title']))
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(f"<center>{address}</center>", styles['Title']))
            elements.append(Spacer(1, 6))

            data = [
                ["Date", datetime.now().strftime("%Y-%m-%d %H:%M")],
                ["Patient ID", patient_id],
                ["Name", f"{first} {last}"],
                ["Gender", f"{gender}"],
                ["Doctor", f"Dr. {doctor_first} {doctor_last} Sp. U"],
                ["Volume", f" mL"]
            ]
            table = Table(data, colWidths=200)
            table.setStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ])
            elements.append(table)
            elements.append(Spacer(1, 12))

            elements.append(Paragraph("Flowrate", styles['Heading3']))
            elements.append(Image(flow_img, width=400, height=200))
            elements.append(Paragraph("Volume", styles['Heading3']))
            elements.append(Image(volume_img, width=400, height=200))

            doc.build(elements)
        except Exception as e:
            logging.error(f"PDF Error: {str(e)}")
            print(f"PDF Error: {str(e)}")

    def save_plots(self, flow_filename, volume_filename):
        """
        Saves the flowrate and volume plots as images.
        """
        # Flow Plot
        plt.figure(figsize=(6, 4))
        plt.plot(self.x_data, self.y1_data, color='blue')
        plt.title("Flowrate")
        plt.xlabel("Time (s)"), plt.ylabel("Flowrate (mL/s)")
        plt.grid(True)
        plt.savefig(flow_filename)
        plt.close()

        # Volume Plot
        plt.figure(figsize=(6, 4))
        plt.plot(self.x_data, self.y2_data, color='orange')
        plt.title("Volume")
        plt.xlabel("Time (s)"), plt.ylabel("Volume (mL)")
        plt.grid(True)
        plt.savefig(volume_filename)
        plt.close()

    def update_history(self):
        """
        Updates the history treeview with patient data from the database.
        Logs and shows an error message if an exception occurs.
        """
        try:
            self.history_treeview.delete(*self.history_treeview.get_children())
            with sqlite3.connect('V3_UROFLOWMETRY_DATABASE.db') as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM patients ORDER BY created_at DESC")
                for row in cursor.fetchall():
                    self.history_treeview.insert("", "end", values=(
                        row[0],  # id_nomor
                        row[4],  # patient_id_num
                        f"{row[1]} {row[2]}",  # Nama
                        row[10]  # created_at
                    ))
        except sqlite3.OperationalError as e:
            logging.error(f"Database Error: {str(e)}")
            Messagebox.show_error(f"Database Error: {str(e)}", f"Table 'patients' does not exist: {str(e)}")
        except Exception as e:
            logging.error(f"Error updating history: {str(e)}")
            Messagebox.show_error(f"Error updating history: {str(e)}", f"Failed to update history: {str(e)}")

    def open_pdf(self):
        """
        Opens the PDF report for the selected patient.
        Logs and shows an error message if an exception occurs.
        """
        selected = self.history_treeview.selection()
        if not selected:
            return

        item = self.history_treeview.item(selected)
        patient_id = item['values'][1]  # patient_id_num

        with sqlite3.connect('V3_UROFLOWMETRY_DATABASE.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM patients WHERE patient_id_num=?", (patient_id,))
            data = cursor.fetchone()
            if data:
                filename = f"{data[4]}_{data[1]}_{data[2]}.pdf"  # data[4] = patient_id_num
                if os.path.exists(filename):
                    webbrowser.open(filename)
                else:
                    logging.error("PDF Not Found")
                    Messagebox.show_error("PDF Not Found", "The PDF file does not exist")

    def delete_measurement(self):
        """
        Deletes the selected patient's data from the database and the PDF report.
        Logs and shows an error message if an exception occurs.
        """
        selected = self.history_treeview.selection()
        if not selected:
            return

        item = self.history_treeview.item(selected)
        patient_id = item['values'][1]  # patient_id_num

        confirm = Messagebox.yesno("Confirm Deletion", f"Delete patient {patient_id}?")
        if confirm == "Yes":
            with sqlite3.connect('V3_UROFLOWMETRY_DATABASE.db') as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM patients WHERE patient_id_num=?", (patient_id,))

            # Remove PDF File
            filename = f"{patient_id}_{item['values'][2].replace(' ', '_')}.pdf"
            if os.path.exists(filename):
                os.remove(filename)

            self.update_history()

    def print_measurement(self):
        """
        Prints the selected patient's data.
        """
        selected = self.history_treeview.selection()
        if not selected:
            return

        item = self.history_treeview.item(selected)
        Messagebox.show_info("Print", f"Printing record for {item['values'][2]}")

    def run(self):
        """
        Runs the main application loop and ensures resources are closed properly.
        """
        try:
            self.root.mainloop()
        finally:
            self.close_resources()

    def close_resources(self):
        """
        Closes all resources including the serial port.
        """
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
            logging.info("Serial port closed")

if __name__ == "__main__":
    app = UroflowmetryApp()
    app.run()
