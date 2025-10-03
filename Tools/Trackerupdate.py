import pandas as pd
from tkinter import Tk, filedialog, Toplevel, Label, OptionMenu, StringVar, Button, messagebox, Text, Scrollbar, Frame, Listbox, MULTIPLE, END, Canvas, BOTH, LEFT, RIGHT, Y, VERTICAL, HORIZONTAL, X, TOP, BOTTOM
from tkinter import ttk
import sys
import traceback
import os
import threading
from datetime import datetime

# === COLOR SCHEME ===
DARK_BG = "#0a0a0a"
PANEL_BG = "#1a1a1a"
ACCENT_BLUE = "#00a8ff"
ACCENT_BLUE_DARK = "#0077cc"
TEXT_COLOR = "#e0e0e0"
SUCCESS_COLOR = "#00ff88"
ERROR_COLOR = "#ff4444"
WARNING_COLOR = "#ffaa00"

class ModernGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("USC News Tracker - File Processor")
        self.root.geometry("1200x800")
        self.root.configure(bg=DARK_BG)
        
        # Make window slightly transparent
        self.root.attributes('-alpha', 0.96)
        
        self.setup_styles()
        self.create_main_interface()
        
    def setup_styles(self):
        """Configure ttk styles for modern dark theme"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure Progressbar
        style.configure("Custom.Horizontal.TProgressbar",
                       troughcolor=PANEL_BG,
                       bordercolor=ACCENT_BLUE,
                       background=ACCENT_BLUE,
                       lightcolor=ACCENT_BLUE,
                       darkcolor=ACCENT_BLUE_DARK)
        
    def create_main_interface(self):
        """Create the main application interface"""
        # Header
        header_frame = Frame(self.root, bg=ACCENT_BLUE, height=80)
        header_frame.pack(fill=X, side=TOP)
        header_frame.pack_propagate(False)
        
        title_label = Label(header_frame, text="📊 USC NEWS TRACKER",
                           font=("Arial", 24, "bold"), bg=ACCENT_BLUE, fg="white")
        title_label.pack(pady=20)
        
        # Main content area
        self.content_frame = Frame(self.root, bg=DARK_BG)
        self.content_frame.pack(fill=BOTH, expand=True, padx=20, pady=20)
        
        # Status panel
        self.create_status_panel()
        
        # File list panel
        self.create_file_list_panel()
        
        # Log panel
        self.create_log_panel()
        
        # Control buttons
        self.create_control_buttons()
        
    def create_status_panel(self):
        """Create status information panel"""
        status_frame = Frame(self.content_frame, bg=PANEL_BG, relief="solid", bd=1)
        status_frame.pack(fill=X, pady=(0, 15))
        
        # Status grid
        status_grid = Frame(status_frame, bg=PANEL_BG)
        status_grid.pack(padx=20, pady=15)
        
        # Files selected
        Label(status_grid, text="📁 Files Selected:", font=("Arial", 11, "bold"),
              bg=PANEL_BG, fg=ACCENT_BLUE).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.files_count_label = Label(status_grid, text="0", font=("Arial", 11),
                                       bg=PANEL_BG, fg=TEXT_COLOR)
        self.files_count_label.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        # Valid files
        Label(status_grid, text="✓ Valid Files:", font=("Arial", 11, "bold"),
              bg=PANEL_BG, fg=SUCCESS_COLOR).grid(row=0, column=2, sticky="w", padx=10, pady=5)
        self.valid_files_label = Label(status_grid, text="0", font=("Arial", 11),
                                       bg=PANEL_BG, fg=TEXT_COLOR)
        self.valid_files_label.grid(row=0, column=3, sticky="w", padx=10, pady=5)
        
        # Rows extracted
        Label(status_grid, text="📊 Rows Extracted:", font=("Arial", 11, "bold"),
              bg=PANEL_BG, fg=ACCENT_BLUE).grid(row=0, column=4, sticky="w", padx=10, pady=5)
        self.rows_label = Label(status_grid, text="0", font=("Arial", 11),
                               bg=PANEL_BG, fg=TEXT_COLOR)
        self.rows_label.grid(row=0, column=5, sticky="w", padx=10, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(status_frame, style="Custom.Horizontal.TProgressbar",
                                       length=400, mode='determinate')
        self.progress.pack(pady=(0, 15), padx=20)
        
        # Status message
        self.status_label = Label(status_frame, text="Ready to process files...",
                                 font=("Arial", 10, "italic"), bg=PANEL_BG, fg=TEXT_COLOR)
        self.status_label.pack(pady=(0, 10))
        
    def create_file_list_panel(self):
        """Create panel showing selected files"""
        file_frame = Frame(self.content_frame, bg=PANEL_BG, relief="solid", bd=1)
        file_frame.pack(fill=BOTH, expand=True, pady=(0, 15))
        
        Label(file_frame, text="📂 SELECTED FILES", font=("Arial", 12, "bold"),
              bg=PANEL_BG, fg=ACCENT_BLUE).pack(anchor="w", padx=15, pady=10)
        
        # Scrollable file list
        list_container = Frame(file_frame, bg=PANEL_BG)
        list_container.pack(fill=BOTH, expand=True, padx=15, pady=(0, 15))
        
        scrollbar = Scrollbar(list_container, bg=PANEL_BG)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.file_listbox = Listbox(list_container, bg="#2a2a2a", fg=TEXT_COLOR,
                                    font=("Courier", 9), selectmode=MULTIPLE,
                                    yscrollcommand=scrollbar.set, bd=0,
                                    highlightthickness=0, selectbackground=ACCENT_BLUE)
        self.file_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
    def create_log_panel(self):
        """Create logging panel"""
        log_frame = Frame(self.content_frame, bg=PANEL_BG, relief="solid", bd=1)
        log_frame.pack(fill=BOTH, expand=True)
        
        Label(log_frame, text="📝 PROCESSING LOG", font=("Arial", 12, "bold"),
              bg=PANEL_BG, fg=ACCENT_BLUE).pack(anchor="w", padx=15, pady=10)
        
        # Scrollable log
        log_container = Frame(log_frame, bg=PANEL_BG)
        log_container.pack(fill=BOTH, expand=True, padx=15, pady=(0, 15))
        
        scrollbar = Scrollbar(log_container, bg=PANEL_BG)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.log_text = Text(log_container, bg="#2a2a2a", fg=TEXT_COLOR,
                            font=("Courier", 9), yscrollcommand=scrollbar.set,
                            bd=0, highlightthickness=0, wrap="word", state="disabled")
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)
        
        # Configure text tags for colored logging
        self.log_text.tag_config("info", foreground=TEXT_COLOR)
        self.log_text.tag_config("success", foreground=SUCCESS_COLOR, font=("Courier", 9, "bold"))
        self.log_text.tag_config("error", foreground=ERROR_COLOR, font=("Courier", 9, "bold"))
        self.log_text.tag_config("warning", foreground=WARNING_COLOR)
        self.log_text.tag_config("header", foreground=ACCENT_BLUE, font=("Courier", 10, "bold"))
        
    def create_control_buttons(self):
        """Create control buttons"""
        button_frame = Frame(self.root, bg=DARK_BG)
        button_frame.pack(fill=X, padx=20, pady=(0, 20))
        
        self.select_btn = Button(button_frame, text="📁 SELECT FILES",
                                command=self.select_files, bg=ACCENT_BLUE, fg="white",
                                font=("Arial", 12, "bold"), padx=30, pady=12,
                                relief="flat", cursor="hand2", activebackground=ACCENT_BLUE_DARK)
        self.select_btn.pack(side=LEFT, padx=5)
        
        self.process_btn = Button(button_frame, text="⚡ PROCESS FILES",
                                  command=self.start_processing, bg=SUCCESS_COLOR, fg="black",
                                  font=("Arial", 12, "bold"), padx=30, pady=12,
                                  relief="flat", cursor="hand2", state="disabled")
        self.process_btn.pack(side=LEFT, padx=5)
        
        self.clear_btn = Button(button_frame, text="🗑️ CLEAR",
                               command=self.clear_all, bg="#444444", fg="white",
                               font=("Arial", 12), padx=30, pady=12,
                               relief="flat", cursor="hand2")
        self.clear_btn.pack(side=LEFT, padx=5)
        
        Button(button_frame, text="✕ EXIT", command=self.root.quit,
               bg=ERROR_COLOR, fg="white", font=("Arial", 12), padx=30, pady=12,
               relief="flat", cursor="hand2").pack(side=RIGHT, padx=5)
        
    def log_message(self, message, level="info"):
        """Add message to log with timestamp"""
        self.log_text.config(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(END, f"[{timestamp}] ", "info")
        self.log_text.insert(END, f"{message}\n", level)
        self.log_text.see(END)
        self.log_text.config(state="disabled")
        self.root.update_idletasks()
        
    def update_status(self, message):
        """Update status label"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
        
    def select_files(self):
        """Open file dialog to select files"""
        file_paths = filedialog.askopenfilenames(
            parent=self.root,
            title="Select Excel Files (Multiple Selection Allowed)",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        
        if not file_paths:
            return
            
        self.file_listbox.delete(0, END)
        self.selected_files = list(file_paths)
        
        for file_path in self.selected_files:
            file_name = os.path.basename(file_path)
            self.file_listbox.insert(END, f"  ➤ {file_name}")
            
        self.files_count_label.config(text=str(len(self.selected_files)))
        self.log_message(f"Selected {len(self.selected_files)} file(s)", "success")
        self.process_btn.config(state="normal")
        
    def clear_all(self):
        """Clear all data"""
        self.file_listbox.delete(0, END)
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, END)
        self.log_text.config(state="disabled")
        self.selected_files = []
        self.files_count_label.config(text="0")
        self.valid_files_label.config(text="0")
        self.rows_label.config(text="0")
        self.progress['value'] = 0
        self.update_status("Ready to process files...")
        self.process_btn.config(state="disabled")
        self.log_message("Cleared all data", "info")
        
    def start_processing(self):
        """Start processing in a separate thread"""
        self.select_btn.config(state="disabled")
        self.process_btn.config(state="disabled")
        self.clear_btn.config(state="disabled")
        
        thread = threading.Thread(target=self.process_files, daemon=True)
        thread.start()
        
    def process_files(self):
        """Process selected files"""
        try:
            self.log_message("="*50, "header")
            self.log_message("STARTING FILE PROCESSING", "header")
            self.log_message("="*50, "header")
            
            # Validate files
            self.update_status("Validating files...")
            valid_files, file_data = self.validate_files()
            
            self.valid_files_label.config(text=str(len(valid_files)))
            
            if not valid_files:
                self.log_message("No valid files found!", "error")
                messagebox.showerror("Error", "No valid files to process!")
                self.reset_buttons()
                return
                
            # Get column mapping for first file
            self.update_status("Waiting for column mapping...")
            column_mapping = self.get_column_mapping(file_data[valid_files[0]], "Primary Column Mapping")
            
            if not column_mapping:
                self.log_message("Column mapping cancelled", "warning")
                self.reset_buttons()
                return
                
            # Process files
            self.process_all_files(valid_files, file_data, column_mapping)
            
        except Exception as e:
            self.log_message(f"ERROR: {str(e)}", "error")
            self.log_message(traceback.format_exc(), "error")
            messagebox.showerror("Error", f"Processing failed:\n{str(e)}")
        finally:
            self.reset_buttons()
            
    def validate_files(self):
        """Validate all selected files"""
        valid_files = []
        file_data = {}
        total = len(self.selected_files)
        
        for i, file_path in enumerate(self.selected_files):
            file_name = os.path.basename(file_path)
            self.progress['value'] = (i / total) * 100
            self.update_status(f"Validating: {file_name}")
            
            self.log_message(f"Validating: {file_name}", "info")
            
            try:
                df = pd.read_excel(file_path)
                valid_files.append(file_path)
                file_data[file_path] = df
                self.log_message(f"  ✓ Valid - {len(df)} rows, {len(df.columns)} columns", "success")
            except Exception as e:
                self.log_message(f"  ✗ Error: {str(e)}", "error")
                
        return valid_files, file_data
        
    def get_column_mapping(self, df, title="Column Mapping"):
        """Show column mapping dialog"""
        target_columns = [
            'Date (DD/MM/YYYY)', 'Link', 'Article Title', 'Priority', 'Contractor',
            'Field / Project', 'Country', 'DA Comment / Action'
        ]
        
        return column_mapping_dialog(df, target_columns, self.root, title)
        
    def process_all_files(self, valid_files, file_data, column_mapping):
        """Process all valid files with retry mechanism for failed files"""
        target_columns = list(column_mapping.keys())
        all_mapped_data = []
        total_rows = 0
        failed_files = []
        
        total = len(valid_files)
        
        # First pass: Process all files with initial mapping
        for i, file_path in enumerate(valid_files):
            file_name = os.path.basename(file_path)
            self.progress['value'] = (i / total) * 100
            self.update_status(f"Processing: {file_name}")
            
            self.log_message(f"Processing: {file_name}", "header")
            
            df = file_data[file_path]
            mapped_data = filter_and_map_data(df, column_mapping, target_columns)
            
            if mapped_data is not None and not mapped_data.empty:
                all_mapped_data.append(mapped_data)
                total_rows += len(mapped_data)
                self.log_message(f"  ✓ Extracted {len(mapped_data)} rows", "success")
            else:
                self.log_message(f"  ⚠ No data extracted - will prompt for remapping", "warning")
                failed_files.append((file_path, file_name, df))
        
        # Second pass: Handle failed files with custom mapping
        if failed_files:
            self.log_message("="*50, "header")
            self.log_message(f"REMAPPING REQUIRED FOR {len(failed_files)} FILE(S)", "header")
            self.log_message("="*50, "header")
            
            retry_result = self.handle_failed_files(failed_files, target_columns, all_mapped_data)
            total_rows += retry_result
        
        # Save results
        if all_mapped_data:
            combined_data = pd.concat(all_mapped_data, ignore_index=True)
            tracker_file_path = r'C:\Office work\Upstream SCRAP news\Tracker\USC News Tracker.xlsx'
            save_to_tracker(combined_data, tracker_file_path)
            
            self.rows_label.config(text=str(total_rows))
            self.progress['value'] = 100
            self.update_status("Processing complete!")
            
            self.log_message("="*50, "header")
            self.log_message(f"✓ SUCCESS: {total_rows} rows added to tracker", "success")
            self.log_message("="*50, "header")
            
            success_msg = f"✓ Processing Complete!\n\n"
            success_msg += f"• {total_rows} rows extracted\n"
            success_msg += f"• From {len(all_mapped_data)} files\n"
            if failed_files:
                success_msg += f"• {len(failed_files)} files required custom mapping\n"
            success_msg += f"• Successfully added to tracker"
            
            messagebox.showinfo("Success", success_msg)
        else:
            self.log_message("No data to save", "warning")
            messagebox.showwarning("No Data", "No rows with comments were found in any files")
    
    def handle_failed_files(self, failed_files, target_columns, all_mapped_data):
        """Show remapping dialog for files that failed to extract data"""
        total_recovered_rows = 0
        
        for file_path, file_name, df in failed_files:
            self.log_message(f"Requesting custom mapping for: {file_name}", "warning")
            self.update_status(f"Custom mapping needed: {file_name}")
            
            # Show error explanation dialog
            result = show_remapping_prompt(self.root, file_name, df)
            
            if result == "remap":
                # Show custom column mapping dialog for this specific file
                custom_mapping = self.get_column_mapping(df, f"Custom Mapping - {file_name}")
                
                if custom_mapping:
                    self.log_message(f"  → Applying custom mapping for {file_name}", "info")
                    mapped_data = filter_and_map_data(df, custom_mapping, target_columns)
                    
                    if mapped_data is not None and not mapped_data.empty:
                        all_mapped_data.append(mapped_data)
                        total_recovered_rows += len(mapped_data)
                        self.log_message(f"  ✓ Extracted {len(mapped_data)} rows with custom mapping", "success")
                    else:
                        self.log_message(f"  ✗ Still no data extracted - skipping file", "error")
                else:
                    self.log_message(f"  ⚠ Custom mapping cancelled - skipping file", "warning")
            else:
                self.log_message(f"  ⚠ User chose to skip file", "warning")
        
        return total_recovered_rows
            
    def reset_buttons(self):
        """Reset button states"""
        self.select_btn.config(state="normal")
        self.process_btn.config(state="normal" if hasattr(self, 'selected_files') and self.selected_files else "disabled")
        self.clear_btn.config(state="normal")

def filter_and_map_data(df, column_mapping, target_columns):
    """Filter data and map columns based on user selection."""
    comment_col_source = column_mapping.get('DA Comment / Action')
    if not comment_col_source or comment_col_source not in df.columns:
        return None

    # Filter rows where the comment column is not empty
    filtered_df = df[df[comment_col_source].notna() & (df[comment_col_source].astype(str).str.strip() != '')].copy()

    if filtered_df.empty:
        return None

    # Create a new DataFrame with only the target columns
    mapped_df = pd.DataFrame()
    for target_col, source_col in column_mapping.items():
        if source_col in filtered_df.columns:
            mapped_df[target_col] = filtered_df[source_col]

    # Ensure all target columns are present, even if not mapped
    for col in target_columns:
        if col not in mapped_df.columns:
            mapped_df[col] = pd.NA

    return mapped_df[target_columns]

# === Utility Dialogs ===

def show_remapping_prompt(root, file_name, df):
    """Show a prompt explaining why remapping is needed"""
    dialog = Toplevel(root)
    dialog.title("⚠ Mapping Issue Detected")
    dialog.geometry("700x550")
    dialog.configure(bg=DARK_BG)
    dialog.attributes('-alpha', 0.96)
    dialog.attributes('-topmost', True)
    
    # Header
    header = Frame(dialog, bg=WARNING_COLOR, height=70)
    header.pack(fill=X)
    header.pack_propagate(False)
    Label(header, text="⚠ No Data Extracted from File",
          font=("Arial", 16, "bold"), bg=WARNING_COLOR, fg="black").pack(pady=20)
    
    # Content frame
    content = Frame(dialog, bg=DARK_BG)
    content.pack(fill=BOTH, expand=True, padx=30, pady=20)
    
    # File info
    info_frame = Frame(content, bg=PANEL_BG, relief="solid", bd=1)
    info_frame.pack(fill=X, pady=(0, 20))
    
    Label(info_frame, text=f"📄 File: {file_name}",
          font=("Arial", 11, "bold"), bg=PANEL_BG, fg=TEXT_COLOR,
          anchor="w").pack(fill=X, padx=15, pady=10)
    
    Label(info_frame, text=f"Rows: {len(df)} | Columns: {len(df.columns)}",
          font=("Arial", 10), bg=PANEL_BG, fg=TEXT_COLOR,
          anchor="w").pack(fill=X, padx=15, pady=(0, 10))
    
    # Explanation
    Label(content, text="Possible Reasons:",
          font=("Arial", 12, "bold"), bg=DARK_BG, fg=ACCENT_BLUE,
          anchor="w").pack(fill=X, pady=(10, 5))
    
    reasons_frame = Frame(content, bg=PANEL_BG, relief="solid", bd=1)
    reasons_frame.pack(fill=BOTH, expand=True, pady=(0, 20))
    
    reasons_text = Text(reasons_frame, bg="#2a2a2a", fg=TEXT_COLOR,
                       font=("Arial", 10), wrap="word", bd=0,
                       highlightthickness=0, height=8)
    reasons_text.pack(fill=BOTH, expand=True, padx=15, pady=15)
    
    reasons = """• The 'DA Comment / Action' column is empty in this file
• The mapped column name doesn't match this file's structure
• The filter criteria didn't match any rows
• Column names differ from other files

You can:
1. Remap columns specifically for this file
2. Skip this file and continue with others"""
    
    reasons_text.insert("1.0", reasons)
    reasons_text.config(state="disabled")
    
    # Available columns
    Label(content, text="Available Columns in This File:",
          font=("Arial", 12, "bold"), bg=DARK_BG, fg=ACCENT_BLUE,
          anchor="w").pack(fill=X, pady=(10, 5))
    
    cols_frame = Frame(content, bg=PANEL_BG, relief="solid", bd=1)
    cols_frame.pack(fill=BOTH, expand=True)
    
    cols_text = Text(cols_frame, bg="#2a2a2a", fg=SUCCESS_COLOR,
                    font=("Courier", 9), wrap="word", bd=0,
                    highlightthickness=0, height=6)
    cols_scrollbar = Scrollbar(cols_frame, command=cols_text.yview)
    cols_text.config(yscrollcommand=cols_scrollbar.set)
    
    cols_scrollbar.pack(side=RIGHT, fill=Y)
    cols_text.pack(side=LEFT, fill=BOTH, expand=True, padx=15, pady=15)
    
    for i, col in enumerate(df.columns, 1):
        cols_text.insert(END, f"{i}. {col}\n")
    cols_text.config(state="disabled")
    
    # Result holder
    result = {"action": "skip"}
    
    def on_remap():
        result["action"] = "remap"
        dialog.destroy()
    
    def on_skip():
        result["action"] = "skip"
        dialog.destroy()
    
    # Buttons
    btn_frame = Frame(dialog, bg=DARK_BG)
    btn_frame.pack(fill=X, padx=30, pady=(0, 30))
    
    Button(btn_frame, text="🔄 REMAP THIS FILE",
           command=on_remap, bg=ACCENT_BLUE, fg="white",
           font=("Arial", 12, "bold"), padx=30, pady=12,
           relief="flat", cursor="hand2").pack(side=LEFT, padx=5)
    
    Button(btn_frame, text="⏭️ SKIP THIS FILE",
           command=on_skip, bg="#666666", fg="white",
           font=("Arial", 12), padx=30, pady=12,
           relief="flat", cursor="hand2").pack(side=LEFT, padx=5)
    
    dialog.transient(root)
    dialog.grab_set()
    root.wait_window(dialog)
    
    return result["action"]

def column_mapping_dialog(df, target_columns, root, title="Column Mapping"):
    """Create a modern column mapping dialog"""
    dialog = Toplevel(root)
    dialog.title(title)
    dialog.geometry("650x550")
    dialog.configure(bg=DARK_BG)
    dialog.attributes('-alpha', 0.96)
    dialog.attributes('-topmost', True)
    
    # Header
    header = Frame(dialog, bg=ACCENT_BLUE, height=70)
    header.pack(fill=X)
    header.pack_propagate(False)
    Label(header, text="🔗 Map Target Columns to Source Columns",
          font=("Arial", 14, "bold"), bg=ACCENT_BLUE, fg="white").pack(pady=20)
    
    # Info label
    info_frame = Frame(dialog, bg=PANEL_BG)
    info_frame.pack(fill=X, padx=20, pady=(10, 0))
    Label(info_frame, text=f"File has {len(df)} rows and {len(df.columns)} columns",
          font=("Arial", 9), bg=PANEL_BG, fg=TEXT_COLOR).pack(pady=5)
    
    # Scrollable content
    canvas = Canvas(dialog, bg=DARK_BG, highlightthickness=0)
    scrollbar = Scrollbar(dialog, orient=VERTICAL, command=canvas.yview)
    scrollable_frame = Frame(canvas, bg=DARK_BG)
    
    scrollable_frame.bind("<Configure>", 
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    canvas.pack(side=LEFT, fill=BOTH, expand=True, padx=20, pady=10)
    scrollbar.pack(side=RIGHT, fill=Y, pady=10)
    
    dropdown_vars = {}
    for i, target_col in enumerate(target_columns):
        row_frame = Frame(scrollable_frame, bg=PANEL_BG, relief="solid", bd=1)
        row_frame.pack(fill=X, pady=5, padx=5)
        
        # Show required indicator for DA Comment
        label_text = f"{target_col}:"
        if target_col == "DA Comment / Action":
            label_text += " *"
            label_fg = ERROR_COLOR
        else:
            label_fg = ACCENT_BLUE
        
        Label(row_frame, text=label_text, font=("Arial", 10, "bold"),
              bg=PANEL_BG, fg=label_fg, width=25, anchor="w").pack(side=LEFT, padx=10, pady=10)
        
        var = StringVar(dialog)
        matching_cols = [col for col in df.columns if col.lower() == target_col.lower()]
        var.set(matching_cols[0] if matching_cols else "Not Mapped")
        
        options = ["Not Mapped"] + list(df.columns)
        dropdown = OptionMenu(row_frame, var, *options)
        dropdown.config(bg="#2a2a2a", fg=TEXT_COLOR, font=("Arial", 9),
                       highlightthickness=0, bd=0, activebackground=ACCENT_BLUE)
        dropdown.pack(side=LEFT, fill=X, expand=True, padx=10, pady=5)
        
        dropdown_vars[target_col] = var
    
    # Note about required field
    note_frame = Frame(dialog, bg=DARK_BG)
    note_frame.pack(fill=X, padx=20)
    Label(note_frame, text="* DA Comment / Action is required",
          font=("Arial", 9, "italic"), bg=DARK_BG, fg=ERROR_COLOR,
          anchor="w").pack(fill=X)
    
    final_mapping = {}
    
    def on_cancel():
        final_mapping.clear()
        dialog.destroy()

    def on_ok():
        temp_mapping = {}
        for target, var in dropdown_vars.items():
            selected = var.get()
            if selected != "Not Mapped":
                temp_mapping[target] = selected
        
        # Check if DA Comment / Action is mapped
        if 'DA Comment / Action' not in temp_mapping:
            messagebox.showerror("Required Field Missing",
                "The 'DA Comment / Action' column must be mapped!\n\n"
                "This field is required to filter rows with comments.")
            return
        
        # Check for duplicate source columns
        source_cols = list(temp_mapping.values())
        if len(source_cols) != len(set(source_cols)):
            messagebox.showerror("Duplicate Mapping",
                "Each source column can only be mapped once!\n\n"
                "Please review your mappings.")
            return
        
        final_mapping.update(temp_mapping)
        dialog.destroy()
    
    # Button frame (ADD EVERYTHING BELOW)
    button_frame = Frame(dialog, bg=DARK_BG)
    button_frame.pack(fill=X, padx=20, pady=20)
    
    Button(button_frame, text="✓ OK", command=on_ok,
           bg=SUCCESS_COLOR, fg="black", font=("Arial", 12, "bold"),
           padx=30, pady=12, relief="flat", cursor="hand2").pack(side=LEFT, padx=5)
    
    Button(button_frame, text="✕ Cancel", command=on_cancel,
           bg="#666666", fg="white", font=("Arial", 12),
           padx=30, pady=12, relief="flat", cursor="hand2").pack(side=LEFT, padx=5)
    
    dialog.transient(root)
    dialog.grab_set()
    root.wait_window(dialog)
    
    return final_mapping if final_mapping else None

def save_to_tracker(data_df, tracker_path):
    """Append data to the main tracker file, creating it if it doesn't exist."""
    if os.path.exists(tracker_path):
        try:
            # Use ExcelWriter in append mode to add data without overwriting
            with pd.ExcelWriter(tracker_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                # Read the existing sheet to find the last row
                try:
                    existing_df = pd.read_excel(tracker_path, sheet_name='Tracker')
                    startrow = len(existing_df) + 1
                except ValueError: # Sheet 'Tracker' does not exist
                    startrow = 0
                
                # Append new data, excluding the header if the sheet already has data
                header = False if startrow > 0 else True
                data_df.to_excel(writer, sheet_name='Tracker', startrow=startrow, index=False, header=header)
        except Exception as e:
            # Fallback for complex cases: read, concat, and overwrite
            existing_df = pd.read_excel(tracker_path, sheet_name='Tracker')
            combined_df = pd.concat([existing_df, data_df], ignore_index=True)
            combined_df.to_excel(tracker_path, sheet_name='Tracker', index=False)
    else:
        # Create a new file if it doesn't exist
        os.makedirs(os.path.dirname(tracker_path), exist_ok=True)
        data_df.to_excel(tracker_path, sheet_name='Tracker', index=False)

if __name__ == "__main__":
    root = Tk()
    app = ModernGUI(root)
    root.mainloop()