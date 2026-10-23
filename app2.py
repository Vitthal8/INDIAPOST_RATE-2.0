import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import time
import threading

# Global Theme Configuration
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class ModernRateLookupApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Settings
        self.title("India Post Bulk Rate Processor v4.0")
        self.geometry("600x650")
        self.resizable(False, False)

        self.user_df = None
        self.file_path = None

        # Load Reference Database Safely
        try:
            self.ref_df = pd.read_csv('Found.csv')
            # Normalize first column header to an internal name
            self.master_pincode_col = self.ref_df.columns[0]
            self.ref_df.rename(columns={self.master_pincode_col: 'Master_PinCode'}, inplace=True)
            self.ref_df['Master_PinCode'] = self.ref_df['Master_PinCode'].astype(str).str.strip()
            self.available_slabs = self.ref_df['SLAB'].unique().tolist()
        except FileNotFoundError:
            messagebox.showerror("System Error",
                                 "Critical Asset Missing:\n'Found.csv' must be in the same folder as this application.")
            self.destroy()
            return

        self.create_ui()

    def create_ui(self):
        # Header Canvas
        self.title_label = ctk.CTkLabel(
            self,
            text="🏣 India Post Rate Matcher",
            font=ctk.CTkFont(family="Arial", size=24, weight="bold")
        )
        self.title_label.pack(pady=(25, 5))

        self.subtitle_label = ctk.CTkLabel(
            self,
            text="Automated Excel Tariff Injection Utility",
            font=ctk.CTkFont(family="Arial", size=13),
            text_color="gray"
        )
        self.subtitle_label.pack(pady=(0, 20))

        # 1. Premium Visual Drop Zone Frame
        self.drop_zone = ctk.CTkFrame(self, width=520, height=160, corner_radius=12)
        self.drop_zone.pack(pady=10, padx=40)
        self.drop_zone.pack_propagate(False)

        self.drop_icon = ctk.CTkLabel(self.drop_zone, text="📊", font=ctk.CTkFont(size=36))
        self.drop_icon.pack(pady=(25, 5))

        self.drop_text = ctk.CTkLabel(
            self.drop_zone,
            text="Click here to browse & select file",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.drop_text.pack(pady=2)

        self.file_details_label = ctk.CTkLabel(
            self.drop_zone,
            text="Supports .xlsx, .xls, and .csv data sets",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        self.file_details_label.pack(pady=(0, 15))

        # Bind click event to the drop zone container and its children
        for widget in (self.drop_zone, self.drop_icon, self.drop_text, self.file_details_label):
            widget.bind("<Button-1>", lambda e: self.browse_file())

        # 2. Configuration Settings Panel
        self.config_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.config_frame.pack(pady=15, fill="x", padx=40)

        # Left Config Column: PinCode Mapping
        self.pincode_container = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        self.pincode_container.pack(side="left", expand=True, fill="x", padx=5)

        ctk.CTkLabel(self.pincode_container, text="Select PinCode Column:",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=2)
        self.dropdown_column = ctk.CTkOptionMenu(self.pincode_container, values=["Select Target Column..."], width=230)
        self.dropdown_column.pack(fill="x")

        # Right Config Column: Weight Slabs
        self.slab_container = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        self.slab_container.pack(side="right", expand=True, fill="x", padx=5)

        ctk.CTkLabel(self.slab_container, text="Select Weight Slab Group:",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=2)
        self.dropdown_slab = ctk.CTkOptionMenu(self.slab_container, values=self.available_slabs, width=230)
        self.dropdown_slab.pack(fill="x")

        # 3. Special Service Toggle Zone
        self.ad_var = ctk.BooleanVar(value=False)
        self.checkbox_ad = ctk.CTkCheckBox(
            self,
            text="Append 'RATE With AD' (Proof of Delivery / Acknowledgement)",
            variable=self.ad_var,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#E67E22"
        )
        self.checkbox_ad.pack(pady=20)

        # 4. Status and Real-Time Progress Infrastructure
        self.status_text = ctk.CTkLabel(self, text="System Ready. Awaiting data source input.", text_color="gray",
                                        font=ctk.CTkFont(size=12))
        self.status_text.pack(pady=(15, 5))

        self.progress_bar = ctk.CTkProgressBar(self, width=520)
        self.progress_bar.pack(pady=5, padx=40)
        self.progress_bar.set(0)  # Hide visually initially

        # 5. Primary Process Trigger Button
        self.btn_process = ctk.CTkButton(
            self,
            text="⚡ Run Analysis & Export Data",
            command=self.start_processing_thread,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#2ECC71",
            hover_color="#27AE60",
            height=45
        )
        self.btn_process.pack(pady=(25, 20), fill="x", padx=40)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if file_path:
            self.file_path = file_path
            filename = os.path.basename(file_path)

            # Visual conversion state change for drop zone
            self.drop_text.configure(text=filename, text_color="#3498DB")
            self.file_details_label.configure(text=f"Location: {os.path.dirname(file_path)[:50]}...")
            self.status_text.configure(text="File verified successfully. Mapping data structures...", text_color="gray")

            try:
                if file_path.endswith('.csv'):
                    self.user_df = pd.read_csv(file_path)
                else:
                    self.user_df = pd.read_excel(file_path)

                # Dynamically load columns into picker
                columns = list(self.user_df.columns)
                self.dropdown_column.configure(values=columns)
                self.dropdown_column.set(columns[0])
                self.status_text.configure(text="Data columns mapped. Select your settings.", text_color="green")

            except Exception as e:
                messagebox.showerror("IO Parsing Error", f"Unable to parse document format:\n{e}")

    def start_processing_thread(self):
        """Dispatches job to separate process thread to prevent UI freezing/lagging."""
        if self.user_df is None or not self.file_path:
            messagebox.showwarning("Incomplete Action", "Please select a valid input spreadsheet dataset first.")
            return

        pincode_col = self.dropdown_column.get()
        if pincode_col == "Select Target Column...":
            messagebox.showwarning("Incomplete Action", "Please designate the column containing address PinCodes.")
            return

        # Start execution thread safely
        threading.Thread(target=self.process_data, daemon=True).start()

    def process_data(self):
        # Freeze primary execution key
        self.btn_process.configure(state="disabled", fg_color="gray")
        pincode_col = self.dropdown_column.get()
        selected_slab = self.dropdown_slab.get()
        target_rate_col = 'RATE With AD' if self.ad_var.get() else 'RATE'

        if target_rate_col not in self.ref_df.columns:
            messagebox.showerror("Execution Aborted",
                                 f"Target field '{target_rate_col}' does not exist inside Found.csv database layout.")
            self.btn_process.configure(state="normal", fg_color="#2ECC71")
            return

        try:
            # 10% - Data Initialization
            self.status_text.configure(text="Cleaning system registers and setting up target indexes...",
                                       text_color="orange")
            self.progress_bar.set(0.1)
            time.sleep(0.3)

            # Sanitize tracking column values strings
            self.user_df[pincode_col] = self.user_df[pincode_col].astype(str).str.strip().str.replace(r'\.0$', '',
                                                                                                      regex=True)

            # 40% - Database Filter Stage
            self.status_text.configure(
                text=f"Isolating master tables tracking under weight structure: {selected_slab}...",
                text_color="orange")
            self.progress_bar.set(0.4)
            time.sleep(0.3)

            slab_ref = self.ref_df[self.ref_df['SLAB'] == selected_slab]

            # 70% - Cross Matrix Processing Lookup
            self.status_text.configure(text=f"Injecting mapping coordinates via index key: '{target_rate_col}'...",
                                       text_color="orange")
            self.progress_bar.set(0.7)
            time.sleep(0.4)

            result_df = pd.merge(
                self.user_df,
                slab_ref[['Master_PinCode', target_rate_col]],
                left_on=pincode_col,
                right_on='Master_PinCode',
                how='left'
            )

            # Housekeeping dropped data frame merges
            result_df.drop(columns=['Master_PinCode'], inplace=True, errors='ignore')

            # 90% - Storage Save Prompt Prep
            self.status_text.configure(text="Processing loops completed. Buffering out save target...",
                                       text_color="orange")
            self.progress_bar.set(0.9)

            original_filename = os.path.basename(self.file_path)
            name_without_ext, _ = os.path.splitext(original_filename)
            suggested_filename = f"{name_without_ext}_Updated.xlsx"

            # Open Save prompt safe dialog
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=suggested_filename,
                filetypes=[("Excel file", "*.xlsx")]
            )

            if save_path:
                # 100% - File Generation Commit
                self.status_text.configure(text="Writing data streams to disk target directory...", text_color="green")
                self.progress_bar.set(1.0)

                result_df.to_excel(save_path, index=False, engine='openpyxl')

                messagebox.showinfo(
                    "Job Completed",
                    f"Spreadsheet successfully localized & exported:\n\nSaved: {os.path.basename(save_path)}\nApplied Column Header: {target_rate_col}"
                )
                self.status_text.configure(text="Job done cleanly. Ready for next data process batch.",
                                           text_color="green")
            else:
                self.status_text.configure(text="Save cancelled by operator loop.", text_color="gray")
                self.progress_bar.set(0)

        except Exception as e:
            messagebox.showerror("Fatal Pipeline Error", f"Engine failed to run lookup operations safely:\n\n{str(e)}")
            self.status_text.configure(text="Process pipeline halted due to error state.", text_color="red")
            self.progress_bar.set(0)

        finally:
            # Restore processing buttons safely
            self.btn_process.configure(state="normal", fg_color="#2ECC71")


if __name__ == "__main__":
    app = ModernRateLookupApp()
    app.mainloop()