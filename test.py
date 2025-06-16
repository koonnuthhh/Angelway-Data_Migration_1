import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys  
import shutil
import threading
import time
import zipfile
from Migration_to_template_1_1_2_1 import Migration_to_template_1_1_2_1
from Migration_to_template_1_2_2_2 import Migration_to_template_1_2_2_2
from Migration_to_template_1_3_2_3 import Migration_to_template_1_3_2_3

from Migration_to_Template_3_WL import Migration_to_Template_3_WL
from Migration_to_Template_7_WL import Migration_to_Template_7_WL
from Migration_to_Template_9_WL import Template_9_WL
from Migration_to_Template_11_WL import Template_11_WL

from Migration_to_Template_3_WM import Migration_to_Template_3_WM
from Migration_to_Template_6_WM import Migration_to_Template_6_WM
from Migration_to_Template_9_WM import Template_9_WM
from Migration_to_template_12_WM import Migration_to_template_12_WM

class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, s):
        self.text_widget.config(state="normal")
        self.text_widget.insert("end", s)
        self.text_widget.see("end")
        self.text_widget.config(state="disabled")

    def flush(self):
        pass  # Needed for compatibility


# --- Simulated Handlers ---
def handle_template_1_1_2_1(inputs,output_dir, log):
    log("Process starting...")
    source_file = inputs["Source File"][0]
    template_1_path = inputs["Template 1.1 File"][0]
    template_2_path = inputs["Template 2.1 File"][0]
    
    source_file = source_file.rsplit('.', 1)[0] + '.xlsx'

    Migration_to_template_1_1_2_1(source_file,"Sheet1",template_1_path,"Template1_WL", template_2_path,"Template2_WL")
    log("Template 1-1-2-1 done.")


def handle_template_1_2_2_2(inputs, output_dir, log):
    log("Process starting...")
    source_file = inputs["Source File"][0]
    source_sheet = "Sheet1"
    template_1_path = inputs["Template 1.2 File"][0]
    template_1_sheet = "Template1_WM"
    template_2_path = inputs["Template 2.2 File"][0]
    template_2_sheet = "Template2_WM"
    full_name_column="ชื่อ-สกุล ลูกค้า"
    address_column="ที่อยู่"
    Migration_to_template_1_2_2_2(source_file,source_sheet,template_1_path,template_1_sheet,template_2_path,template_2_sheet,full_name_column,address_column)
    log("Template 1-2-2-2 done.")


def handle_template_1_3_2_3(inputs, output_dir, log):
    log("Process starting...")
    source_file = inputs["Source File"][0]
    source_sheet = "Sheet1"
    template_1_path = inputs["Template 1.3 File"][0]
    template_1_sheet = "Template1"
    template_2_path = inputs["Template 2.3 File"][0]
    template_2_sheet = "Template2"

    Migration_to_template_1_3_2_3(source_file, source_sheet, template_1_path,template_1_sheet,template_2_path,template_2_sheet)
    log("Template 1-3-2-3 done.")


def handle_template_3_WL(inputs, output_dir, log):
    source_file = inputs["Source File"][0]
    destination_file = inputs["Template 3"][0]
    source_sheet = "Sheet1"
    destination_sheet = "ข้อมูลหลักประกันรถ"
    Migration_to_Template_3_WL(source_file, destination_file, source_sheet, destination_sheet)
    log("Template 3 WL done.")


def handle_template_7_WL(inputs, output_dir, log):
    source_file_1 = inputs["Source 7.1 File"][0]
    source_sheet_1 = "Sheet1"
    source_file_2 = inputs["Source 7.2 File"][0]
    source_sheet_2 = "Sheet1"
    destination_file = inputs["Template 7 File"][0]
    destination_sheet = "ข้อมูลสัญญาเชื้อ"

    Migration_to_Template_7_WL(source_file_1,source_sheet_1,source_file_2,source_sheet_2,destination_file,destination_sheet)
    log("Template 7 WL done.")


def handle_template_9_WL(inputs, output_dir, log):
    source_file = inputs["Source 9 File"][0]
    b_zad_path = inputs["B_Zad File"][0]
    destination_file = inputs["Template 9 File"][0]
    Template_9_WL(source_file,b_zad_path,destination_file)
    log("Template 9 WL done.")


def handle_template_11_WL(inputs, output_dir, log):
    source_file = inputs["Source 11 WL File"][0]
    destination_file = inputs["Template 11 WL File"][0]
    Template_11_WL(source_file,destination_file)
    log("Template 11 WL done.")


def handle_template_3_WM(inputs, output_dir, log):
    source_file = inputs["Source 3 WM File"][0]
    source_sheet ="Sheet1"
    destination_file = inputs["Template 3 WM File"][0]
    destination_sheet = "ข้อมูลหลักประกันรถ"

    Migration_to_Template_3_WM(source_file,destination_file,source_sheet,destination_sheet) 
    log("Template 3 WM done.")


def handle_template_6_WM(inputs, output_dir, log):
    source_file1 = inputs["Source 6.1 File"][0]
    zfloan_raw  = inputs["Zfloan_raw File"][0]
    source_file3 = inputs["Source 6.3 File"][0]
    source_file4 = inputs["Source 6.4 File"][0]
    destination_file = inputs["Destination 6 WM File"][0]

    Migration_to_Template_6_WM(source_file1,zfloan_raw,source_file3,source_file4,destination_file)
    log("Template 6 WM done.")


def handle_template_9_WM(inputs, output_dir, log):
    source_file = inputs["Source File"][0]
    reference_file = inputs["ประเภทการชำระ File"][0]
    destination_file = inputs["Destination 9 WM File"][0]
    Template_9_WM (source_file,reference_file,destination_file)
    log("Template 9 WM done.")


def handle_template_12_WM(inputs, output_dir, log):
    destination_file = inputs["Destination 12 WM File"][0]

    Migration_to_template_12_WM(destination_file)
    log("Template 12 WM done.")


# --- Function Definitions ---

FUNCTIONS = {
    "Template 1-1-2-1": {
        "inputs": {
            "Source File": {"multiple": False},
            "Template 1.1 File": {"multiple": False},
            "Template 2.1 File": {"multiple": False}
        },
        "handler": handle_template_1_1_2_1
    },
    "Template 1-2-2-2": {
        "inputs": {
            "Source File": {"multiple": False},
            "Template 1.2 File": {"multiple": False},
            "Template 2.2 File": {"multiple": False}
        },
        "handler": handle_template_1_2_2_2
    },
    "Template 1-3-2-3": {
        "inputs": {
            "Source File": {"multiple": False},
            "Template 1.3 File": {"multiple": False},
            "Template 2.3 File": {"multiple": False}
        },
        "handler": handle_template_1_3_2_3
    },

    "Template 3 WL": {
        "inputs": {
            "Source File": {"multiple": False},
            "Template 3": {"multiple": False}
        },
        "handler": handle_template_3_WL
    },
    "Template 7 WL": {
        "inputs": {
            "Source 7.1 File": {"multiple": False},
            "Source 7.2 File": {"multiple": False},
            "Template 7 File": {"multiple": False},
        },
        "handler": handle_template_7_WL
    },
    "Template 9 WL": {
        "inputs": {
          "Source 9 File": {"multiple": False},
          "B_Zad File": {"multiple": False},
          "Template 9 File": {"multiple": False},
        },
        "handler": handle_template_9_WL
    },
    "Template 11 WL": {
        "inputs": {
            "Source 11 WL File": {"multiple": False},
            "Template 11 WL File": {"multiple": False},
        },
        "handler": handle_template_11_WL
    },

    "Template 3 WM": {
        "inputs": {
            "Source 3 WM File": {"multiple": False},
            "Template 3 WM File": {"multiple": False},
        },
        "handler": handle_template_3_WM
    },
    "Template 6 WM": {
        "inputs": {
            "Source 6.1 File": {"multiple": False},
            "Source 6.3 File": {"multiple": False},
            "Source 6.4 File": {"multiple": False},
            "Zfloan_raw File": {"multiple": False},
            "Destination 6 WM File": {"multiple": False},
        },
        "handler": handle_template_6_WM
    },
    "Template 9 WM": {
        "inputs": {
            "Source 9 WM File": {"multiple": False},
            "ประเภทการชำระ File": {"multiple": False},
            "Destination 9 WM File": {"multiple": False},
        },
        "handler": handle_template_9_WM
    },
    "Template 12 WM": {
        "inputs": {
            "Destination 12 WM File": {"multiple": False},
        },
        "handler": handle_template_12_WM
    }
}


# --- Main Application ---

class FileProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Flexible File Processor (Dynamic UI)")
        self.root.geometry("700x600")

        self.selected_function = tk.StringVar()
        self.output_dir = tk.StringVar(value=os.getcwd())
        self.input_widgets = {}

        self.create_widgets()
        
         # Redirect stdout to log window
        sys.stdout = TextRedirector(self.log_text)
        sys.stderr = TextRedirector(self.log_text)  # optional: redirect errors too

    def create_widgets(self):
        # Top control
        frame_top = ttk.Frame(self.root)
        frame_top.pack(fill="x", pady=10)

        ttk.Label(frame_top, text="Select Function:").pack(side="left", padx=5)
        self.function_cb = ttk.Combobox(frame_top, textvariable=self.selected_function, values=list(FUNCTIONS.keys()), state="readonly")
        self.function_cb.pack(side="left", padx=5)
        self.function_cb.bind("<<ComboboxSelected>>", self.build_inputs)

        #ttk.Button(frame_top, text="Select Output Folder", command=self.select_output_folder).pack(side="right", padx=5)

        # Dynamic input area
        self.input_frame = ttk.LabelFrame(self.root, text="Inputs")
        self.input_frame.pack(fill="x", padx=10, pady=10)

        # Start button
        ttk.Button(self.root, text="Start Processing", command=self.start_processing).pack(pady=10)

        # Log window
        ttk.Label(self.root, text="Log:").pack(anchor="w")
        self.log_text = tk.Text(self.root, height=15, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=5)

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir.set(folder)
            self.log(f"Output folder set to {folder}")

    def build_inputs(self, event=None):
        for widget in self.input_frame.winfo_children():
            widget.destroy()
        self.input_widgets = {}

        function_name = self.selected_function.get()
        function_data = FUNCTIONS[function_name]

        for idx, (input_name, cfg) in enumerate(function_data["inputs"].items()):
            row = ttk.Frame(self.input_frame)
            row.pack(fill="x", pady=5)

            ttk.Label(row, text=input_name + ":").pack(side="left", padx=5)

            input_type = cfg.get("type", "file")  # default to file

            if input_type == "text":
                # Create text entry
                text_var = tk.StringVar()
                entry = ttk.Entry(row, textvariable=text_var, width=50)
                entry.pack(side="left", expand=True)
                #, padx=5
                self.input_widgets[input_name] = {
                    "type": "text",
                    "var": text_var
                }
            else:
                # Create file picker (your original code)
                file_var = tk.StringVar()
                entry = ttk.Entry(row, textvariable=file_var, width=50)
                entry.pack(side="left", padx=5, expand=True)

                btn = ttk.Button(row, text="Browse", command=lambda name=input_name, multi=cfg.get("multiple", False): self.select_file(name, multi))
                btn.pack(side="left", padx=5)

                self.input_widgets[input_name] = {
                    "type": "file",
                    "var": file_var,
                    "multiple": cfg.get("multiple", False),
                    "files": []
                }

    def select_file(self, input_name, multiple):
        if multiple:
            files = filedialog.askopenfilenames()
        else:
            file = filedialog.askopenfilename()
            files = [file] if file else []

        if files:
            self.input_widgets[input_name]["files"] = files
            self.input_widgets[input_name]["var"].set("; ".join(files))

    def log(self, text):
        self.log_text.config(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def start_processing(self):
        if not self.selected_function.get():
            messagebox.showwarning("Warning", "Please select a function.")
            return

        # Collect input files
        inputs = {}
        for input_name, widget in self.input_widgets.items():
            if not widget["files"]:
                messagebox.showwarning("Warning", f"Please select files for '{input_name}'")
                return
            inputs[input_name] = widget["files"]

        threading.Thread(target=self.run_processing, args=(inputs,), daemon=True).start()

    def run_processing(self, inputs):
        try:
            handler = FUNCTIONS[self.selected_function.get()]["handler"]
            handler(inputs, self.output_dir.get(), self.log)
        except Exception as e:
            self.log(f"Error: {str(e)}")
        finally:
            self.log("Processing Completed")

# --- Run the application ---

if __name__ == "__main__":
    root = tk.Tk()
    app = FileProcessorApp(root)
    root.mainloop()
