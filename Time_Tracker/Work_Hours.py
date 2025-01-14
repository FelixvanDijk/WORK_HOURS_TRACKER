import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import json
import os
import time
from datetime import datetime, timedelta
from openpyxl import Workbook

DATA_FILE = "time_records.json"


class TimeTrackerApp:
    def __init__(self, master):
        self.master = master
        master.title("Work Hours Tracker")

        # -----------------------------
        #    INTERNAL TIMER STATES
        # -----------------------------
        self.timer_running = False
        self.start_time = None  # raw time.time() when session starts/resumes
        self.start_time_dt = None  # actual datetime when user pressed "Start"
        self.elapsed_time = 0.0  # accumulates total seconds across pause/resume

        # -----------------------------
        #         UI ELEMENTS
        # -----------------------------
        # Big timer display
        self.timer_label = tk.Label(master, text="00:00:00", font=("Helvetica", 32))
        self.timer_label.pack(pady=10)

        # Frame for Start/Pause/Resume/Stop
        button_frame = tk.Frame(master)
        button_frame.pack(pady=10)

        self.start_button = tk.Button(
            button_frame, text="Start", width=10, command=self.start_timer
        )
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.pause_button = tk.Button(
            button_frame, text="Pause", width=10, command=self.pause_timer
        )
        self.pause_button.pack(side=tk.LEFT, padx=5)

        self.resume_button = tk.Button(
            button_frame, text="Resume", width=10, command=self.resume_timer
        )
        self.resume_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = tk.Button(
            button_frame, text="Stop", width=10, command=self.stop_timer
        )
        self.stop_button.pack(side=tk.LEFT, padx=5)

        # Frame for Edit & Export
        actions_frame = tk.Frame(master)
        actions_frame.pack(pady=10)

        self.edit_button = tk.Button(
            actions_frame, text="Edit Entries", width=12, command=self.open_edit_window
        )
        self.edit_button.pack(side=tk.LEFT, padx=5)

        self.export_button = tk.Button(
            actions_frame, text="Export", width=12, command=self.export_data_dialog
        )
        self.export_button.pack(side=tk.LEFT, padx=5)

        # Continuously update the displayed timer
        self.update_timer()

    # -------------------------------------------------
    #                   TIMER LOGIC
    # -------------------------------------------------
    def start_timer(self):
        """Start a new timing session if none is running."""
        if self.timer_running:
            messagebox.showwarning("Warning", "Timer is already running!")
            return

        # Actual datetime for the start
        self.start_time_dt = datetime.now()
        # Store the raw time.time()
        self.start_time = time.time()
        self.timer_running = True
        self.elapsed_time = 0.0  # reset to 0 at the start of a new session

    def pause_timer(self):
        """Pause the timer (accumulate elapsed time)."""
        if not self.timer_running:
            messagebox.showwarning("Warning", "No timer is running to pause.")
            return

        # Accumulate elapsed
        self.elapsed_time += time.time() - self.start_time
        self.timer_running = False

    def resume_timer(self):
        """Resume the timer from a paused state."""
        if self.timer_running:
            messagebox.showwarning("Warning", "Timer is already running!")
            return

        self.start_time = time.time()
        self.timer_running = True

    def stop_timer(self):
        """Stop the timer, prompt for a comment, save record, and reset."""
        if not self.timer_running and self.elapsed_time == 0.0:
            messagebox.showwarning("Warning", "No active or paused session to stop.")
            return

        if self.timer_running:
            # finalize the time accumulation
            self.elapsed_time += time.time() - self.start_time
            self.timer_running = False

        # Prompt user for a comment
        comment = simpledialog.askstring(
            "Comment", "Add a comment for this session (optional):"
        )
        if comment is None:
            comment = ""  # user cancelled or closed dialog

        # End time is "now"
        end_time_dt = datetime.now()

        # Save the record
        self.save_time_record(
            start_time_dt=self.start_time_dt,
            end_time_dt=end_time_dt,
            elapsed_seconds=self.elapsed_time,
            comment=comment,
        )

        # Reset the timer
        self.elapsed_time = 0.0
        self.start_time = None
        self.start_time_dt = None
        self.update_timer_label(0.0)

        messagebox.showinfo("Session Recorded", "Your work session has been saved.")

    def update_timer(self):
        """Continuously update the on-screen timer display."""
        if self.timer_running:
            current_elapsed = self.elapsed_time + (time.time() - self.start_time)
        else:
            current_elapsed = self.elapsed_time

        self.update_timer_label(current_elapsed)
        self.master.after(200, self.update_timer)  # run again in 200ms

    def update_timer_label(self, elapsed_seconds):
        """Convert seconds to H:MM:SS and show on the timer label."""
        hours = int(elapsed_seconds // 3600)
        minutes = int((elapsed_seconds % 3600) // 60)
        seconds = int(elapsed_seconds % 60)
        self.timer_label.config(text=f"{hours:02d}:{minutes:02d}:{seconds:02d}")

    # -------------------------------------------------
    #                DATA PERSISTENCE
    # -------------------------------------------------
    def load_records(self):
        """Load time records from JSON."""
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return []

    def save_records(self, records):
        """Save the entire list of records to JSON."""
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(records, f, indent=4)

    def save_time_record(self, start_time_dt, end_time_dt, elapsed_seconds, comment):
        """Append a new record to the JSON file."""
        records = self.load_records()
        new_record = {
            "start_time": start_time_dt.isoformat(),
            "end_time": end_time_dt.isoformat(),
            "elapsed": elapsed_seconds,
            "comment": comment,
        }
        records.append(new_record)
        self.save_records(records)

    # -------------------------------------------------
    #              EDITING & DELETING
    # -------------------------------------------------
    def open_edit_window(self):
        """
        Show a window with all records in a listbox, so the user
        can edit or delete them. We store records in memory and
        rely on list-index for identifying them (since we removed ID).
        """
        self.edit_window = tk.Toplevel(self.master)
        self.edit_window.title("Edit Time Records")

        records = self.load_records()
        if not records:
            messagebox.showinfo("No Data", "No time records available to edit.")
            self.edit_window.destroy()
            return

        # Keep a reference so we know which record is which
        self.edit_window_records = records

        # Frame for the list
        list_frame = tk.Frame(self.edit_window)
        list_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Scrollbar + Listbox
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox = tk.Listbox(
            list_frame, height=12, width=80, yscrollcommand=scrollbar.set
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)

        # Populate the listbox
        for rec in records:
            start_str = rec.get("start_time", "N/A")
            end_str = rec.get("end_time", "N/A")
            cmt = rec.get("comment", "")
            self.listbox.insert(
                tk.END, f"Start: {start_str} | End: {end_str} | Comment: {cmt}"
            )

        # Button frame
        btn_frame = tk.Frame(self.edit_window)
        btn_frame.pack(pady=5)

        edit_button = tk.Button(
            btn_frame, text="Edit Selected", command=self.edit_selected_record
        )
        edit_button.pack(side=tk.LEFT, padx=5)

        delete_button = tk.Button(
            btn_frame, text="Delete Selected", command=self.delete_selected_record
        )
        delete_button.pack(side=tk.LEFT, padx=5)

    def get_selected_list_index(self):
        """Return the selected index in the listbox (or None if none selected)."""
        selection = self.listbox.curselection()
        if not selection:
            return None
        return selection[0]

    def edit_selected_record(self):
        """Edit the record at the selected list index."""
        index = self.get_selected_list_index()
        if index is None:
            messagebox.showwarning("No Selection", "Please select a record to edit.")
            return

        record = self.edit_window_records[index]

        edit_dialog = tk.Toplevel(self.master)
        edit_dialog.title("Edit Record")

        # Start time
        tk.Label(edit_dialog, text="Start Time (YYYY-MM-DD HH:MM:SS):").pack(pady=2)
        start_var = tk.StringVar(value=self.iso_to_display(record["start_time"]))
        start_entry = tk.Entry(edit_dialog, textvariable=start_var, width=30)
        start_entry.pack()

        # End time
        tk.Label(edit_dialog, text="End Time (YYYY-MM-DD HH:MM:SS):").pack(pady=2)
        end_var = tk.StringVar(value=self.iso_to_display(record["end_time"]))
        end_entry = tk.Entry(edit_dialog, textvariable=end_var, width=30)
        end_entry.pack()

        # Comment
        tk.Label(edit_dialog, text="Comment:").pack(pady=2)
        comment_var = tk.StringVar(value=record.get("comment", ""))
        comment_entry = tk.Entry(edit_dialog, textvariable=comment_var, width=50)
        comment_entry.pack()

        def save_changes():
            new_start = start_var.get()
            new_end = end_var.get()
            new_comment = comment_var.get()

            try:
                dt_start = datetime.strptime(new_start, "%Y-%m-%d %H:%M:%S")
                dt_end = datetime.strptime(new_end, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                messagebox.showerror(
                    "Error", "Invalid date/time. Use YYYY-MM-DD HH:MM:SS format."
                )
                return

            if dt_end < dt_start:
                messagebox.showerror("Error", "End time cannot be before start time.")
                return

            elapsed = (dt_end - dt_start).total_seconds()

            # Update the record in memory
            record["start_time"] = dt_start.isoformat()
            record["end_time"] = dt_end.isoformat()
            record["elapsed"] = elapsed
            record["comment"] = new_comment

            # Save to disk
            self.save_records(self.edit_window_records)

            # Refresh listbox
            self.listbox.delete(0, tk.END)
            for r in self.edit_window_records:
                s_str = r.get("start_time", "N/A")
                e_str = r.get("end_time", "N/A")
                cmt = r.get("comment", "")
                self.listbox.insert(
                    tk.END, f"Start: {s_str} | End: {e_str} | Comment: {cmt}"
                )

            messagebox.showinfo("Success", "Record updated.")
            edit_dialog.destroy()

        save_button = tk.Button(edit_dialog, text="Save", command=save_changes)
        save_button.pack(pady=5)

    def delete_selected_record(self):
        """Delete the record at the selected list index."""
        index = self.get_selected_list_index()
        if index is None:
            messagebox.showwarning("No Selection", "Please select a record to delete.")
            return

        if messagebox.askyesno(
            "Confirm Delete", "Are you sure you want to delete this record?"
        ):
            del self.edit_window_records[index]
            self.save_records(self.edit_window_records)

            # Refresh listbox
            self.listbox.delete(0, tk.END)
            for r in self.edit_window_records:
                s_str = r.get("start_time", "N/A")
                e_str = r.get("end_time", "N/A")
                cmt = r.get("comment", "")
                self.listbox.insert(
                    tk.END, f"Start: {s_str} | End: {e_str} | Comment: {cmt}"
                )

            messagebox.showinfo("Deleted", "Record deleted successfully.")

    @staticmethod
    def iso_to_display(iso_str):
        """Convert ISO datetime (e.g. 2025-01-05T14:30:00) to 'YYYY-MM-DD HH:MM:SS'."""
        try:
            dt = datetime.fromisoformat(iso_str)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except (ValueError, TypeError):
            return iso_str

    # -------------------------------------------------
    #                EXPORT TO EXCEL
    # -------------------------------------------------
    def export_data_dialog(self):
        """
        Prompt for start/end dates (YYYY-MM-DD) and a filename, then export
        matching records to an Excel file. Also sum up the total elapsed
        seconds and hours (to 2 decimal places).
        """
        start_date_str = simpledialog.askstring(
            "Export", "Enter START date (YYYY-MM-DD):"
        )
        if not start_date_str:
            return
        end_date_str = simpledialog.askstring("Export", "Enter END date (YYYY-MM-DD):")
        if not end_date_str:
            return

        export_filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Export As",
        )
        if not export_filename:
            return

        try:
            self.export_data_to_excel(start_date_str, end_date_str, export_filename)
            messagebox.showinfo(
                "Export Successful", f"Data exported to {export_filename}"
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def export_data_to_excel(self, start_date_str, end_date_str, filepath):
        """Export records in [start_date, end_date] to an Excel file, plus a sum row."""
        records = self.load_records()

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()
        except ValueError:
            raise ValueError("Dates must be in YYYY-MM-DD format.")

        filtered = []
        for r in records:
            try:
                dt_start = datetime.fromisoformat(r["start_time"]).date()
            except (ValueError, KeyError):
                continue
            if start_date <= dt_start <= end_date:
                filtered.append(r)

        if not filtered:
            raise ValueError("No records found in the specified date range.")

        # Create the workbook and sheet
        wb = Workbook()
        ws = wb.active
        ws.title = "WorkHours"

        # Header row
        headers = ["Start Time", "End Time", "Elapsed (seconds)", "Comment"]
        ws.append(headers)

        # Populate rows
        for rec in filtered:
            ws.append(
                [
                    rec.get("start_time", ""),
                    rec.get("end_time", ""),
                    rec.get("elapsed", ""),
                    rec.get("comment", ""),
                ]
            )

        # --- Summation row(s) ---
        total_seconds = sum(rec["elapsed"] for rec in filtered)
        total_hours = total_seconds / 3600

        # Let's add a blank row first
        ws.append([""] * len(headers))
        # Then a row for total seconds
        ws.append(["", "", total_seconds, "TOTAL SECONDS"])
        # And a row for total hours
        ws.append(["", "", f"{total_hours:.2f}", "TOTAL HOURS"])

        # Save
        wb.save(filepath)


def main():
    root = tk.Tk()
    app = TimeTrackerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
