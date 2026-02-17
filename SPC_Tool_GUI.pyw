import sys
import traceback
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import os
import threading
from pathlib import Path

# --- CRASH LOGGER START ---
def log_crash(exctype, value, tb):
    """Writes any unhandled crash to a text file."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("crash_log.txt", "w") as f:
        f.write(f"CRITICAL ERROR at {timestamp}\n")
        f.write("="*40 + "\n")
        traceback.print_exception(exctype, value, tb, file=f)
        
    try:
        import tkinter.messagebox
        root = tk.Tk()
        root.withdraw()
        tk.messagebox.showerror("Critical Error", f"The application crashed.\nSee 'crash_log.txt' for details.\n\nError: {value}")
    except:
        pass

sys.excepthook = log_crash
# --- CRASH LOGGER END ---

# --- INTEGRATION SETUP ---
current_dir = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(current_dir, 'src')
sys.path.append(src_path)

# Try importing the existing logic
try:
    from spc_generator.generator import process_single_file
    from spc_generator.template import main as generate_template_logic
    HAS_BACKEND = True
except ImportError as e:
    HAS_BACKEND = False
    IMPORT_ERROR = str(e)

class SPCInterface:
    def __init__(self, root):
        self.root = root
        self.root.title("SPC Generator Dashboard")
        self.root.geometry("700x650")
        
        # Styles
        style = ttk.Style()
        style.configure("Title.TLabel", font=('Arial', 12, 'bold'))

        # --- HEADER ---
        header_frame = tk.Frame(root, bg="#7030A0", pady=10)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="SPC AUTOMATION TOOL", bg="#7030A0", fg="white", 
                 font=("Arial", 16, "bold")).pack()

        # --- ERROR CHECK ---
        if not HAS_BACKEND:
            tk.Label(root, text=f"CRITICAL ERROR: Could not load backend logic.\n{IMPORT_ERROR}", 
                     fg="red", pady=20).pack()
            return

        # --- MAIN CONTENT ---
        main_frame = tk.Frame(root, padx=15, pady=15)
        main_frame.pack(fill="both", expand=True)

        # 1. TOOLBAR (Template & Scan)
        toolbar = tk.Frame(main_frame)
        toolbar.pack(fill="x", pady=(0, 10))
        
        tk.Button(toolbar, text="📄 Get New Template", command=self.run_template_gen, 
                  bg="#e1e1e1").pack(side="left", padx=(0, 10))
        
        tk.Button(toolbar, text="🔄 Refresh File List", command=self.scan_files, 
                  bg="#e1e1e1").pack(side="left")

        # 2. FILE SELECTION (Treeview)
        ttk.Label(main_frame, text="Select Input Files:", style="Title.TLabel", anchor="w").pack(fill="x")
        
        tree_frame = tk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True, pady=5)
        
        # Scrollbar
        scrollbar = tk.Scrollbar(tree_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.tree = ttk.Treeview(tree_frame, columns=("Status",), selectmode="extended", 
                                 yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.tree.yview)
        
        self.tree.heading("#0", text="Filename", anchor="w")
        self.tree.heading("Status", text="State", anchor="w")
        self.tree.column("#0", stretch=True, width=400)
        self.tree.column("Status", width=150)
        self.tree.pack(side="left", fill="both", expand=True)

        # 3. SETTINGS (PROJECT SELECTION)
        settings_frame = tk.LabelFrame(main_frame, text="Project Configuration", padx=10, pady=10)
        settings_frame.pack(fill="x", pady=10)
        
        tk.Label(settings_frame, text="Select Existing Project OR Type New Name:").pack(anchor="w")
        
        proj_frame = tk.Frame(settings_frame)
        proj_frame.pack(fill="x", pady=(5, 0))
        
        # COMBOBOX FOR PROJECTS
        self.project_var = tk.StringVar()
        self.project_combo = ttk.Combobox(proj_frame, textvariable=self.project_var, height=10)
        self.project_combo.pack(side="left", fill="x", expand=True)
        
        # Set default value
        default_proj = f"Run_{datetime.datetime.now().strftime('%Y%m%d')}"
        self.project_combo.set(default_proj)
        
        # Populate the list immediately
        self.refresh_project_list()

        # 4. RUN BUTTON
        self.btn_run = tk.Button(main_frame, text="▶ RUN ANALYSIS", command=self.start_processing_thread, 
                                 bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), height=2)
        self.btn_run.pack(fill="x", pady=(0, 10))

        # 5. LOG WINDOW
        log_frame = tk.LabelFrame(main_frame, text="Execution Log", padx=5, pady=5)
        log_frame.pack(fill="both", expand=True)
        
        self.log_text = tk.Text(log_frame, height=8, state="disabled", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True)
        
        # Initial Scan
        self.scan_files()

    def log(self, message, color="black"):
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{datetime.datetime.now().strftime('%H:%M:%S')}] ", "grey")
        self.log_text.tag_config("grey", foreground="#888888")
        
        tag_name = f"color_{color}"
        self.log_text.tag_config(tag_name, foreground=color)
        self.log_text.insert("end", message + "\n", tag_name)
        
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def scan_files(self):
        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Find files
        count = 0
        if os.path.exists(current_dir):
            for f in os.listdir(current_dir):
                if f.startswith("SPC-DATA_") and f.endswith(".xlsx") and not f.startswith("~$"):
                    self.tree.insert("", "end", text=f, values=("Ready",))
                    count += 1
        
        if count == 0:
            self.log("No input files found. Please generate a template.", "orange")
        else:
            self.log(f"Found {count} input files.", "blue")
            
        # Also refresh projects while we are at it
        self.refresh_project_list()

    def refresh_project_list(self):
        """Scans the 'output' folder and populates the combobox"""
        output_dir = os.path.join(current_dir, "output")
        projects = []
        
        if os.path.exists(output_dir):
            try:
                # Get all subdirectories in output/
                items = os.listdir(output_dir)
                for item in items:
                    full_path = os.path.join(output_dir, item)
                    if os.path.isdir(full_path):
                        projects.append(item)
            except Exception:
                pass
        
        # Sort most recent first (if they follow naming convention) or alphabetical
        projects.sort(reverse=True)
        self.project_combo['values'] = projects

    def run_template_gen(self):
        try:
            generate_template_logic()
            messagebox.showinfo("Success", "Template created: SPC-DATA_Input_Template.xlsx")
            self.scan_files()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create template: {e}")

    def start_processing_thread(self):
        # Get selected items
        selected_items = self.tree.selection()
        if not selected_items:
            # If nothing selected, select ALL
            selected_items = self.tree.get_children()
            if not selected_items:
                messagebox.showwarning("Warning", "No files available to process.")
                return

        files_to_process = [self.tree.item(i)['text'] for i in selected_items]
        
        # Get Project Name
        project_name = self.project_var.get().strip()
        if not project_name:
            messagebox.showwarning("Warning", "Please enter or select a Project Name.")
            return
            
        # Sanitize project name slightly (prevent traversing up directories)
        project_name = project_name.replace("..", "").replace("/", "_").replace("\\", "_")

        # Disable UI
        self.btn_run.config(state="disabled", text="Processing...")
        
        # Start Thread
        t = threading.Thread(target=self.run_process_logic, args=(files_to_process, project_name))
        t.start()

    def run_process_logic(self, filenames, project_name):
        # Construct full path: current_dir/output/project_name
        output_root = Path(current_dir) / "output" / project_name
        
        # Ensure directory exists
        try:
            output_root.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.log(f"Error creating output folder: {e}", "red")
            self.root.after(0, lambda: self.finish_processing(str(output_root), 1))
            return
        
        self.log(f"Project Folder: {project_name}", "purple")
        
        errors = 0
        
        for fname in filenames:
            # Update Tree Status
            item_id = None
            for child in self.tree.get_children():
                if self.tree.item(child)['text'] == fname:
                    item_id = child
                    break
            
            if item_id:
                self.tree.set(item_id, "Status", "Processing...")
            
            full_path = os.path.join(current_dir, fname)
            
            try:
                # CALL THE EXISTING BACKEND LOGIC
                result = process_single_file(full_path, output_root)
                
                # Parse logs from the backend
                if "critical_error" in result:
                    self.log(f"{fname}: {result['critical_error']}", "red")
                    if item_id: self.tree.set(item_id, "Status", "FAILED")
                    errors += 1
                else:
                    tab_count = len(result['processed'])
                    self.log(f"{fname}: Successfully processed {tab_count} tabs.", "green")
                    if item_id: self.tree.set(item_id, "Status", "Done")
                    
            except Exception as e:
                self.log(f"{fname}: Unexpected Error - {str(e)}", "red")
                if item_id: self.tree.set(item_id, "Status", "Error")
                errors += 1

        self.root.after(0, lambda: self.finish_processing(str(output_root), errors))

    def finish_processing(self, output_loc, error_count):
        self.btn_run.config(state="normal", text="▶ RUN ANALYSIS")
        
        # Refresh the list so the new project appears immediately
        self.refresh_project_list()
        
        if error_count == 0:
            messagebox.showinfo("Complete", f"Processing complete!\nFiles saved to:\n{output_loc}")
        else:
            messagebox.showwarning("Complete", f"Processing finished with {error_count} errors.\nCheck the log for details.")

if __name__ == "__main__":
    root = tk.Tk()
    app = SPCInterface(root)
    root.mainloop()