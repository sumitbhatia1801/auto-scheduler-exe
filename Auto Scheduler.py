#!/usr/bin/env python
# coding: utf-8

# In[1]:


import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns

# Use a non-interactive matplotlib backend so EXE systems behave reliably
plt.switch_backend("Agg")


class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Spot Scheduler - Advanced GUI (Merged Unique ID)")
        self.root.geometry("1000x650")

        self.file_path = None
        self.df_input = None
        self.aggregated_df = None  # final merged (H2) output

        # Top buttons
        top_frame = tk.Frame(root)
        top_frame.pack(pady=10, fill=tk.X)

        btn_load = tk.Button(top_frame, text="Upload Input Excel", command=self.load_file, width=22, bg="#1976D2", fg="white")
        btn_load.grid(row=0, column=0, padx=6)

        btn_process = tk.Button(top_frame, text="Generate & Merge (Unique ID)", command=self.generate_schedule, width=22, bg="#388E3C", fg="white")
        btn_process.grid(row=0, column=1, padx=6)

        btn_heatmap = tk.Button(top_frame, text="Show Heatmap", command=self.show_heatmap, width=22, bg="#F57C00", fg="white")
        btn_heatmap.grid(row=0, column=2, padx=6)

        btn_export = tk.Button(top_frame, text="Export Excel", command=self.export_excel, width=22, bg="#455A64", fg="white")
        btn_export.grid(row=0, column=3, padx=6)

        # Treeview for preview
        self.tree = ttk.Treeview(root, columns=(), show="headings")
        self.tree.pack(expand=True, fill="both", pady=12)

        # Status bar
        self.status = tk.Label(root, text="Upload an Excel file with columns: Unique ID, Channel, Start Date (ddmmyyyy), End Date (ddmmyyyy), Spots, Rule", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status.pack(fill=tk.X)

    # -----------------------------
    # Load Excel File
    # -----------------------------
    def load_file(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
            if not file_path:
                return

            df = pd.read_excel(file_path, dtype=str)  # read as str to parse start/end in ddmmyyyy reliably

            required_cols = {"Unique ID", "Channel", "Start Date", "End Date", "Spots", "Rule"}
            missing = required_cols - set(df.columns)
            if missing:
                messagebox.showerror("Invalid Input", f"Missing columns: {', '.join(missing)}\nPlease provide all required columns.")
                return

            # Keep original raw column types for display, but convert where needed later
            self.file_path = file_path
            self.df_input = df.copy()
            self.update_table(self.df_input)
            self.status.config(text=f"File loaded: {file_path}")
        except Exception as e:
            messagebox.showerror("Error loading file", str(e))

    # -----------------------------
    # Update Treeview Table
    # -----------------------------
    def update_table(self, df):
        # Clear existing
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)

        for col in df.columns:
            self.tree.heading(col, text=col)
            # set column width heuristically
            self.tree.column(col, width=120, anchor=tk.W)

        for _, row in df.iterrows():
            vals = [self._format_cell(x) for x in list(row)]
            self.tree.insert("", tk.END, values=vals)

    def _format_cell(self, val):
        if pd.isna(val):
            return ""
        return str(val)

    # -----------------------------
    # Generate Schedule & Merge by Unique ID (H2)
    # -----------------------------
    def generate_schedule(self):
        if self.df_input is None:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return

        try:
            # Work on a copy
            df = self.df_input.copy()

            # Parse Start/End dates in ddmmyyyy format (no separators)
            # Accept strings like '01022025' or '01-02-2025' but require ddmmyyyy as primary.
            def parse_ddmmyyyy(x):
                x = str(x).strip()
                # Try exact ddmmyyyy first
                try:
                    return pd.to_datetime(x, format="%d%m%Y").date()
                except Exception:
                    # fallback: try common separators
                    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
                        try:
                            return pd.to_datetime(x, format=fmt).date()
                        except Exception:
                            continue
                    raise ValueError(f"Could not parse date '{x}'. Expected ddmmyyyy (e.g. 01022025).")

            df["Start Date"] = df["Start Date"].apply(parse_ddmmyyyy)
            df["End Date"] = df["End Date"].apply(parse_ddmmyyyy)

            # Validate numeric Spots
            df["Spots"] = df["Spots"].astype(int)

            # Normalize Rule column
            df["Rule"] = (
                df["Rule"]
                .astype(str)
                .str.replace(" ", "")
                .str.replace("\u2013", "-")
                .str.replace("\u2014", "-")
                .str.strip()
                .str.title()
            )

            # Rule mapping
            rule_map = {
                "Mon": ["Monday"], "Tue": ["Tuesday"], "Wed": ["Wednesday"],
                "Thu": ["Thursday"], "Fri": ["Friday"], "Sat": ["Saturday"], "Sun": ["Sunday"],
                "Mon-Fri": ["Monday","Tuesday","Wednesday","Thursday","Friday"],
                "Sat-Sun": ["Saturday","Sunday"],
                "Mon-Sun": ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"],
                "Mon-Sat": ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"],
                "Wed-Thu": ["Wednesday","Thursday"],
                "Wed-Sun": ["Wednesday","Thursday","Friday","Saturday","Sunday"],
                "Wed-Sat": ["Wednesday","Thursday","Friday","Saturday"],
                "Mon-Tue": ["Monday","Tuesday"],
                "Thur-Sun": ["Thursday","Friday","Saturday","Sunday"],
                "Fri-Sat": ["Friday","Saturday"]
            }

            # Collect all date range across all rows
            all_dates = set()
            for _, row in df.iterrows():
                all_dates.update(pd.date_range(row["Start Date"], row["End Date"]).date)

            if not all_dates:
                messagebox.showerror("No Dates", "No valid date ranges found in the input.")
                return

            all_dates = sorted(list(all_dates))
            date_labels = [d.strftime("%d-%m-%Y") for d in all_dates]
            day_names = [d.strftime("%A") for d in all_dates]

            # Scheduling arrays: one row per input row
            output = np.zeros((len(df), len(all_dates)), dtype=int)
            total_spots_per_day = np.zeros(len(all_dates), dtype=int)

            # Balanced Round-Robin assignment (per input row)
            for row_idx, row in df.iterrows():
                spots = int(row["Spots"])
                rule = row["Rule"]
                valid_days = rule_map.get(rule, ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"])
                date_range = pd.date_range(start=row["Start Date"], end=row["End Date"]).date

                valid_indices = [i for i, d in enumerate(all_dates) if d in date_range and day_names[i] in valid_days]
                if not valid_indices or spots <= 0:
                    continue

                remaining_spots = spots
                idx_pointer = 0
                while remaining_spots > 0:
                    min_load = min(total_spots_per_day[i] for i in valid_indices)
                    candidate_indices = [i for i in valid_indices if total_spots_per_day[i] == min_load]
                    current_idx = candidate_indices[idx_pointer % len(candidate_indices)]
                    output[row_idx, current_idx] += 1
                    total_spots_per_day[current_idx] += 1
                    remaining_spots -= 1
                    idx_pointer += 1

            # Create schedule DataFrame per input row (dates horizontal)
            schedule_df = pd.DataFrame(output, columns=date_labels)

            # Combine input + schedule
            combined = pd.concat([df.reset_index(drop=True), schedule_df], axis=1)

            # For H2: Merge rows by Unique ID + Channel (summing date columns).
            # Keep Start Date as min, End Date as max, Spots summed, Rule: first (could be multiple)
            agg_dict = {
                "Start Date": "min",
                "End Date": "max",
                "Spots": "sum",
                "Rule": lambda x: ";".join(sorted(set(x.astype(str)))),
                # date columns aggregated below
            }

            # Add date columns to aggregation dict
            for dcol in date_labels:
                agg_dict[dcol] = "sum"

            # Perform groupby aggregation
            merged = combined.groupby(["Unique ID", "Channel"], as_index=False).agg(agg_dict)

            # Reorder columns: Unique ID, Channel, Start Date, End Date, Spots, Rule, <date columns...>
            final_cols = ["Unique ID", "Channel", "Start Date", "End Date", "Spots", "Rule"] + date_labels
            merged = merged[final_cols]

            # Save to instance for export / heatmap
            self.aggregated_df = merged.copy()

            # Display
            # Convert Start/End back to readable string for preview
            display_df = merged.copy()
            display_df["Start Date"] = display_df["Start Date"].apply(lambda d: d.strftime("%d-%m-%Y") if pd.notna(d) else "")
            display_df["End Date"] = display_df["End Date"].apply(lambda d: d.strftime("%d-%m-%Y") if pd.notna(d) else "")

            self.update_table(display_df)
            self.status.config(text=f"Schedule generated and merged by Unique ID. {len(merged)} unique ID(s) output rows.")
        except Exception as e:
            messagebox.showerror("Error while generating schedule", str(e))

    # -----------------------------
    # Show Heatmap (uses aggregated_df)
    # -----------------------------
    def show_heatmap(self):
        if self.aggregated_df is None:
            messagebox.showwarning("Warning", "Generate the schedule first.")
            return

        try:
            df = self.aggregated_df.copy()

            # Melt date columns
            # detect date cols as those matching dd-mm-yyyy pattern
            date_cols = [c for c in df.columns if self._is_date_col(c)]
            if not date_cols:
                messagebox.showinfo("No dates", "No date columns available for heatmap.")
                return

            long_df = df.melt(id_vars=["Unique ID", "Channel"], value_vars=date_cols, var_name="Date", value_name="Spot Count")

            # Pivot: Unique ID as index (merged per your request)
            pivot_df = long_df.pivot_table(index="Unique ID", columns="Date", values="Spot Count", aggfunc="sum", fill_value=0)

            plt.figure(figsize=(14, max(4, 0.5 * len(pivot_df))))
            sns.heatmap(pivot_df, annot=True, fmt="d", cbar_kws={'label': 'Spots Assigned'})
            plt.title("Spot Distribution Heatmap (per Unique ID)")
            plt.tight_layout()

            # Use plt.show() - on GUI-enabled systems this will open a window
            plt.show()
        except Exception as e:
            messagebox.showerror("Error showing heatmap", str(e))

    def _is_date_col(self, col_name):
        # crude check for dd-mm-yyyy style column names
        try:
            datetime.strptime(col_name, "%d-%m-%Y")
            return True
        except Exception:
            return False

    # -----------------------------
    # Export aggregated result to Excel
    # -----------------------------
    def export_excel(self):
        if self.aggregated_df is None:
            messagebox.showwarning("Warning", "Generate the schedule first.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if save_path:
            try:
                # Convert Start/End to readable format before saving
                out = self.aggregated_df.copy()
                out["Start Date"] = out["Start Date"].apply(lambda d: d.strftime("%d-%m-%Y") if pd.notna(d) else "")
                out["End Date"] = out["End Date"].apply(lambda d: d.strftime("%d-%m-%Y") if pd.notna(d) else "")
                out.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"Excel exported successfully to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Export Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()


