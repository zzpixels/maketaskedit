import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import os

class CsvEditor(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.df = pd.DataFrame()
        self.file_path = None

        self.build_menu()
        self.build_table()
        self.pack(fill="both", expand=True)

    def build_menu(self):
        menubar = tk.Menu(self.parent)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Open…", command=self.open_csv)
        file_menu.add_command(label="Save", command=self.save_csv, state="disabled")
        file_menu.add_separator()
        file_menu.add_command(label="Quit", command=self.parent.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        tools_menu = tk.Menu(menubar, tearoff=False)
        tools_menu.add_command(label="Fill Column…", command=self.fill_column, state="disabled")
        tools_menu.add_command(label="Quick Fill…", command=self.quick_fill_dialog, state="disabled")
        menubar.add_cascade(label="Tools", menu=tools_menu)

        self.parent.config(menu=menubar)
        self.file_menu = file_menu
        self.tools_menu = tools_menu

    def build_table(self):
        self.tree = ttk.Treeview(self, show="headings", selectmode="extended")
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self.tree.bind("<Double-1>", self.begin_edit)

    def open_csv(self):
        fp = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not fp:
            return
        try:
            self.df = pd.read_csv(fp, dtype=str).fillna("")
        except Exception as e:
            messagebox.showerror("Open CSV", f"Could not open file:\n{e}")
            return 
        self.file_path = fp
        self.refresh_table()
        self.parent.title(f"Make Task Editor – {os.path.basename(fp)}")
        self.file_menu.entryconfig("Save", state="normal")
        self.tools_menu.entryconfig("Fill Column…", state="normal")
        self.tools_menu.entryconfig("Quick Fill…", state="normal")

    def save_csv(self):
        if self.file_path is None:
            return
        self.finish_edit()
        try:
            self.df.to_csv(self.file_path, index=False)
            messagebox.showinfo("Save", f"Saved to {self.file_path}")
        except Exception as e:
            messagebox.showerror("Save CSV", f"Could not save file:\n{e}")

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())

        cols = ["#"] + list(self.df.columns)
        self.tree["columns"] = cols

        self.tree.heading("#0", text="")
        self.tree.column("#0", width=0, stretch=False)

        self.tree.heading("#", text="#")
        self.tree.column("#", width=40, anchor="center")

        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="w")

        for i, (_, row) in enumerate(self.df.iterrows(), start=1):
            values = [str(i)] + list(row)
            self.tree.insert("", "end", values=values)


    def begin_edit(self, event):
        """Start in-place edit of a cell (except the row-number ‘#’ column)."""
        if self.tree.identify_region(event.x, event.y) != "cell":
            return

        row_id  = self.tree.identify_row(event.y)
        col_id  = self.tree.identify_column(event.x)
        col_idx = int(col_id[1:]) - 1

        if col_idx == 0:
            return
        
        df_col_idx = col_idx - 1

        x, y, w, h = self.tree.bbox(row_id, col_id)
        current    = self.tree.set(row_id, column=col_idx)

        edit_box = tk.Entry(self.tree)
        edit_box.place(x=x, y=y, width=w, height=h)
        edit_box.insert(0, current)
        edit_box.focus_set()

        def save_edit(_=None):
            new_val = edit_box.get()
            self.tree.set(row_id, column=col_idx, value=new_val)

            df_row_idx = self.tree.index(row_id)
            self.df.iat[df_row_idx, df_col_idx] = new_val
            edit_box.destroy()

        edit_box.bind("<Return>",   save_edit)
        edit_box.bind("<FocusOut>", save_edit)

    def finish_edit(self):
        self.tree.focus_set()

    def fill_column(self):
        if self.df.empty:
            return

        class FillDialog(simpledialog.Dialog):
            def body(self, master):
                ttk.Label(master, text="Choose column:").grid(row=0, column=0, sticky="w")
                self.col_var = tk.StringVar(value=self.master.df.columns[0])
                self.col_menu = ttk.OptionMenu(master, self.col_var,
                                            self.master.df.columns[0], *self.master.df.columns)
                self.col_menu.grid(row=0, column=1, sticky="ew")

                ttk.Label(master, text="Value to fill:").grid(row=1, column=0, sticky="w")
                self.value_entry = ttk.Entry(master)
                self.value_entry.grid(row=1, column=1, sticky="ew")

                ttk.Label(master, text="Start row:").grid(row=2, column=0, sticky="w")
                self.start_entry = ttk.Entry(master)
                self.start_entry.insert(0, "1")
                self.start_entry.grid(row=2, column=1, sticky="ew")

                ttk.Label(master, text="End row:").grid(row=3, column=0, sticky="w")
                self.end_entry = ttk.Entry(master)
                self.end_entry.insert(0, str(len(self.master.df)))
                self.end_entry.grid(row=3, column=1, sticky="ew")

                self.all_var = tk.BooleanVar()

                def toggle_range_inputs():
                    state = "disabled" if self.all_var.get() else "normal"
                    self.start_entry.configure(state=state)
                    self.end_entry.configure(state=state)

                ttk.Checkbutton(
                    master,
                    text="All rows",
                    variable=self.all_var,
                    command=toggle_range_inputs
                ).grid(row=4, columnspan=2, sticky="w")

                return self.value_entry

            def apply(self):
                self.selected_column = self.col_var.get()
                self.fill_value = self.value_entry.get()
                if self.all_var.get():
                    self.start_row = 0
                    self.end_row = len(self.master.df) - 1
                else:
                    try:
                        self.start_row = int(self.start_entry.get()) - 1
                        self.end_row = int(self.end_entry.get()) - 1
                        if self.start_row < 0 or self.end_row < self.start_row:
                            raise ValueError
                    except ValueError:
                        messagebox.showerror("Invalid Input",
                                            "Start/end rows must be valid integers and start ≤ end.")
                        self.start_row = None

        dlg = FillDialog(self)
        if not hasattr(dlg, "selected_column") or dlg.start_row is None:
            return
        col = dlg.selected_column
        value = dlg.fill_value
        start = dlg.start_row
        end = dlg.end_row

        if start >= len(self.df) or end >= len(self.df):
            messagebox.showerror("Row Range Error", "Row range exceeds available rows.")
            return

        self.df.loc[start:end, col] = value
        self.refresh_table()

    def toggle_range_inputs():
        state = "disabled" if self.all_var.get() else "normal"
        self.start_entry.configure(state=state)
        self.end_entry.configure(state=state)


    def quick_fill_dialog(self):
        """Open the unified Quick-Fill popup and apply the fill if user clicks OK."""
        if self.df.empty:
            return

        class QuickFillDialog(simpledialog.Dialog):
            def __init__(self, parent, df):
                self.df = df
                super().__init__(parent, title="Quick Fill")

            def body(self, master):
                self.mode_var = tk.StringVar(value="paypal")
                self.type_var = tk.StringVar(value="HAS")

                ttk.Label(master, text="Mode:").grid(row=0, column=0, sticky="w")
                ttk.OptionMenu(master, self.mode_var, "paypal",
                               "paypal", "paypalpopnow").grid(row=0, column=1, sticky="ew")

                ttk.Label(master, text="Type:").grid(row=1, column=0, sticky="w")
                ttk.OptionMenu(master, self.type_var, "HAS",
                               "HAS", "MAC", "BIE").grid(row=1, column=1, sticky="ew")

                ttk.Label(master, text="Start row:").grid(row=2, column=0, sticky="w")
                self.start_entry = ttk.Entry(master, width=6)
                self.start_entry.insert(0, "1")
                self.start_entry.grid(row=2, column=1, sticky="w")

                ttk.Label(master, text="End row:").grid(row=3, column=0, sticky="w")
                self.end_entry = ttk.Entry(master, width=6)
                self.end_entry.insert(0, str(len(self.df)))
                self.end_entry.grid(row=3, column=1, sticky="w")

                self.all_var = tk.BooleanVar()

                def toggle_range_inputs():
                    state = "disabled" if self.all_var.get() else "normal"
                    self.start_entry.configure(state=state)
                    self.end_entry.configure(state=state)

                ttk.Checkbutton(master, text="All rows",
                                variable=self.all_var,
                                command=toggle_range_inputs
                                ).grid(row=4, columnspan=2, sticky="w")

                return self.start_entry

            def apply(self):
                self.mode = self.mode_var.get()
                self.type = self.type_var.get()

                if self.all_var.get():
                    self.start_idx, self.end_idx = 0, len(self.df) - 1
                    return

                try:
                    self.start_idx = int(self.start_entry.get()) - 1
                    self.end_idx   = int(self.end_entry.get()) - 1
                    assert 0 <= self.start_idx <= self.end_idx < len(self.df)
                except Exception:
                    messagebox.showerror("Invalid range",
                                         "Please enter valid integer row numbers (start ≤ end).")
                    self.mode = None

        dlg = QuickFillDialog(self, self.df)
        if dlg.mode is None:
            return

        values_map = {
            "paypal":       {"HAS": "1372", "MAC": "675", "BIE": "2155"},
            "paypalpopnow": {"HAS": "50",   "MAC": "40",  "BIE": "195"}
        }

        input_col = next((c for c in self.df.columns if c.lower() == "input"), None)
        mode_col  = next((c for c in self.df.columns if c.lower() == "mode"),  None)
        if not input_col or not mode_col:
            messagebox.showerror("Missing columns",
                                 "CSV must contain 'Input' and 'Mode' columns.")
            return

        fill_val = values_map[dlg.mode][dlg.type]
        self.df.loc[dlg.start_idx:dlg.end_idx, input_col] = fill_val
        self.df.loc[dlg.start_idx:dlg.end_idx, mode_col]  = dlg.mode

        self.refresh_table()


def main():
    root = tk.Tk()
    root.title("Make Task Editor")
    root.geometry("900x600")
    CsvEditor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
