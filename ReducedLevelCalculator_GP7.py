"""
Reduced Level Calculator - Rise & Fall and Height of Collimation Methods
Run in PyCharm:  right-click the file > Run
Needs Python 3 only. (Importing .xlsx files also needs:  pip install openpyxl)
"""

import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

DP = 3  # decimal places used in levelling work


# STAGE 2: READ AND VALIDATE THE FIELD READINGS
def parse_readings(text):
    """Turn pasted text (BS, IS, FS columns) into a list of dictionaries."""
    rows = []
    for line_no, line in enumerate(text.splitlines(), start=1):
        if not line.strip():
            continue                                  # skip blank lines
        sep = "\t" if "\t" in line else ","           # Excel paste = tabs
        cells = [c.strip() for c in line.split(sep)]
        cells = (cells + ["", "", ""])[:3]            # always exactly 3 cells
        try:
            bs, is_, fs = [float(c) if c else None for c in cells]
        except ValueError:
            raise ValueError(f"Line {line_no}: only numbers are allowed.")
        rows.append({"bs": bs, "is": is_, "fs": fs})
    return rows


def validate_rows(rows):
    """Check that the readings follow the rules of a levelling run."""
    if len(rows) < 2:
        raise ValueError("Enter at least a first row (BS) and a last row (FS).")
    first, last = rows[0], rows[-1]
    if first["bs"] is None or first["is"] is not None or first["fs"] is not None:
        raise ValueError("The FIRST row must contain only a Backsight (BS).")
    if last["fs"] is None or last["bs"] is not None or last["is"] is not None:
        raise ValueError("The LAST row must contain only a Foresight (FS).")
    for i, r in enumerate(rows[1:-1], start=2):
        only_is = r["is"] is not None and r["bs"] is None and r["fs"] is None
        is_cp = r["bs"] is not None and r["fs"] is not None and r["is"] is None
        if not (only_is or is_cp):
            raise ValueError(
                f"Row {i}: a middle row must be either an Intermediate sight "
                f"only, or a Foresight + Backsight pair (change point).")


def reading_of(row):
    """The staff reading taken to the point on this row (IS or FS)."""
    return row["is"] if row["is"] is not None else row["fs"]


# STAGE 3: RISE AND FALL METHOD
def rise_and_fall(rows, first_rl):
    out = [{"rise": None, "fall": None, "rl": round(first_rl, DP)}]
    rl = first_rl
    prev = rows[0]["bs"]                       # previous staff reading
    for row in rows[1:]:
        cur = reading_of(row)
        diff = round(prev - cur, DP)           # + = rise, - = fall
        rl = round(rl + diff, DP)
        out.append({"rise": diff if diff >= 0 else None,
                    "fall": -diff if diff < 0 else None,
                    "rl": rl})
        # at a change point the next comparison starts from the new BS
        prev = row["bs"] if row["bs"] is not None else cur
    return out


# STAGE 4: HEIGHT OF COLLIMATION METHOD
def height_of_collimation(rows, first_rl):
    hi = round(first_rl + rows[0]["bs"], DP)
    out = [{"hi": hi, "rl": round(first_rl, DP)}]
    for row in rows[1:]:
        rl = round(hi - reading_of(row), DP)
        if row["bs"] is not None:              # change point: new HI
            hi = round(rl + row["bs"], DP)
        out.append({"hi": hi, "rl": rl})
    return out


# STAGE 5: CORRECTIONS (misclosure shared over the instrument setups)
def apply_corrections(rows, initial_rls, final_rl):
    n_setups = sum(1 for r in rows if r["bs"] is not None)   # = no. of BS
    misclosure = round(final_rl - initial_rls[-1], DP)
    per_setup = misclosure / n_setups
    corrections = [0.0]                        # first benchmark: no correction
    setup = 1
    for row in rows[1:]:
        corrections.append(round(per_setup * setup, DP))
        if row["bs"] is not None:              # a CP closes this setup
            setup += 1
    adjusted = [round(rl + c, DP) for rl, c in zip(initial_rls, corrections)]
    return corrections, adjusted, misclosure, n_setups


# STAGE 6: REMARKS - label points and identify change points
def point_name(n):
    """0->A, 1->B ... 25->Z, 26->AA ..."""
    name = ""
    n += 1
    while n:
        n, rem = divmod(n - 1, 26)
        name = chr(65 + rem) + name
    return name


def make_remarks(rows, first_rl, final_rl):
    remarks = [f"TBM {first_rl:.{DP}f}"]
    for i, row in enumerate(rows[1:-1]):
        tag = point_name(i)
        remarks.append(f"{tag} (CP)" if row["bs"] is not None else tag)
    remarks.append(f"TBM {final_rl:.{DP}f}")
    return remarks


# STAGE 7: ARITHMETIC CHECKS
def arithmetic_checks(rows, rf, hoc, first_rl):
    sum_bs = round(sum(r["bs"] for r in rows if r["bs"] is not None), DP)
    sum_fs = round(sum(r["fs"] for r in rows if r["fs"] is not None), DP)
    sum_rise = round(sum(r["rise"] for r in rf if r["rise"] is not None), DP)
    sum_fall = round(sum(r["fall"] for r in rf if r["fall"] is not None), DP)
    last_minus_first_rf = round(rf[-1]["rl"] - first_rl, DP)
    last_minus_first_hoc = round(hoc[-1]["rl"] - first_rl, DP)
    return {
        "sum_bs": sum_bs, "sum_fs": sum_fs,
        "sum_rise": sum_rise, "sum_fall": sum_fall,
        "bs_fs": round(sum_bs - sum_fs, DP),
        "rise_fall": round(sum_rise - sum_fall, DP),
        "last_first_rf": last_minus_first_rf,
        "last_first_hoc": last_minus_first_hoc,
        "ok_rf": round(sum_bs - sum_fs, DP) == round(sum_rise - sum_fall, DP)
                 == last_minus_first_rf,
        "ok_hoc": round(sum_bs - sum_fs, DP) == last_minus_first_hoc,
    }


# STAGE 8: BUILD THE FINAL TABLES + DIFFERENCE BETWEEN METHODS
def f(v):
    """Format a number for display; blank for None."""
    return "" if v is None else f"{v:.{DP}f}"


RF_HEADERS = ["BS", "IS", "FS", "Rise", "Fall", "Initial RL",
              "Adjustment", "Adjusted RL", "Remarks"]
HOC_HEADERS = ["BS", "IS", "FS", "HI", "Initial RL",
               "Adjustment", "Adjusted RL", "Remarks"]
DIFF_HEADERS = ["Remarks", "Rise&Fall RL", "HoC RL", "Difference"]


def calculate_all(text, first_rl, final_rl):
    rows = parse_readings(text)
    validate_rows(rows)
    rf = rise_and_fall(rows, first_rl)
    hoc = height_of_collimation(rows, first_rl)
    remarks = make_remarks(rows, first_rl, final_rl)

    rf_corr, rf_adj, misclosure, n_setups = apply_corrections(
        rows, [r["rl"] for r in rf], final_rl)
    hoc_corr, hoc_adj, _, _ = apply_corrections(
        rows, [r["rl"] for r in hoc], final_rl)

    rf_table, hoc_table, diff_table = [], [], []
    for i, row in enumerate(rows):
        rf_table.append([f(row["bs"]), f(row["is"]), f(row["fs"]),
                         f(rf[i]["rise"]), f(rf[i]["fall"]), f(rf[i]["rl"]),
                         f(rf_corr[i]), f(rf_adj[i]), remarks[i]])
        hoc_table.append([f(row["bs"]), f(row["is"]), f(row["fs"]),
                          f(hoc[i]["hi"]), f(hoc[i]["rl"]),
                          f(hoc_corr[i]), f(hoc_adj[i]), remarks[i]])
        diff_table.append([remarks[i], f(rf_adj[i]), f(hoc_adj[i]),
                           f(round(rf_adj[i] - hoc_adj[i], DP))])

    chk = arithmetic_checks(rows, rf, hoc, first_rl)
    summary = [
        "ARITHMETIC CHECKS",
        f"Sum BS = {chk['sum_bs']:.3f}    Sum FS = {chk['sum_fs']:.3f}"
        f"    BS - FS = {chk['bs_fs']:.3f}",
        f"Sum Rise = {chk['sum_rise']:.3f}    Sum Fall = {chk['sum_fall']:.3f}"
        f"    Rise - Fall = {chk['rise_fall']:.3f}",
        f"Last RL - First RL (Rise&Fall) = {chk['last_first_rf']:.3f}",
        f"Last RL - First RL (HoC)       = {chk['last_first_hoc']:.3f}",
        f"Rise & Fall check: {'OK' if chk['ok_rf'] else 'FAILED'}",
        f"HoC check:         {'OK' if chk['ok_hoc'] else 'FAILED'}",
        "",
        "MISCLOSURE",
        f"Computed last RL = {rf[-1]['rl']:.3f}   Known final RL = {final_rl:.3f}",
        f"Misclosure (final - computed) = {misclosure:.3f}",
        f"Instrument setups = {n_setups}   "
        f"Correction per setup = {misclosure / n_setups:.4f}",
    ]
    return rf_table, hoc_table, diff_table, summary


# STAGE 9: EXPORT TO CSV (opens in Microsoft Excel)
def write_csv(path, headers, table):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(headers)
        writer.writerows(table)


# STAGE 10: IMPORT READINGS FROM A FILE (.csv, .txt, .xlsx)
def read_file_cells(path):
    """Read any supported file into a list of rows; each row = list of strings."""
    lower = path.lower()
    if lower.endswith((".xlsx", ".xlsm")):
        try:
            from openpyxl import load_workbook
        except ImportError:
            raise ValueError(
                "Reading .xlsx files needs the 'openpyxl' package.\n"
                "In PyCharm: Terminal > pip install openpyxl\n"
                "(or save the sheet as .csv and import that instead).")
        sheet = load_workbook(path, data_only=True).active
        return [["" if c is None else str(c).strip() for c in row]
                for row in sheet.iter_rows(values_only=True)]
    with open(path, "r", encoding="utf-8-sig") as fh:          # csv / txt
        rows = []
        for line in fh.read().splitlines():
            sep = "\t" if "\t" in line else ","
            rows.append([c.strip() for c in line.split(sep)])
        return rows


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def cells_to_readings(cell_rows):
    """
    Convert raw file rows into (text_for_the_box, initial_rl, final_rl).
    - Header rows (e.g. 'BS, IS, FS') are skipped.
    - Rows starting with 'Initial' / 'Final' followed by a number
      (e.g.  'Initial RL, 49.873') fill the RL boxes.
    - Only the first three columns of the remaining rows are used.
    """
    lines, first_rl, final_rl = [], None, None
    for n, cells in enumerate(cell_rows, start=1):
        if not any(cells):
            continue                                   # empty row
        label = cells[0].lower()
        if label.startswith(("initial", "final")):
            nums = [c for c in cells[1:] if is_number(c)]
            if nums:
                if label.startswith("initial"):
                    first_rl = float(nums[0])
                else:
                    final_rl = float(nums[0])
            continue
        three = (cells + ["", "", ""])[:3]
        filled = [c for c in three if c]
        if not filled:
            continue
        if not any(is_number(c) for c in filled):
            continue                                   # header row
        if not all(is_number(c) for c in filled):
            raise ValueError(f"File row {n}: found text among the readings.")
        lines.append("\t".join(three))
    if not lines:
        raise ValueError("No readings were found in that file.")
    return "\n".join(lines), first_rl, final_rl


def import_file(path):
    return cells_to_readings(read_file_cells(path))


# STAGE 11: THE GRAPHICAL INTERFACE
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reduced Level Calculator - Rise & Fall and HoC")
        self.geometry("1000x700")
        self.results = None
        self._build_inputs()
        self._build_outputs()

    # ---------- input area ----------
    def _build_inputs(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="Initial RL (first BM):").grid(row=0, column=0, sticky="w")
        self.first_rl = ttk.Entry(top, width=12)
        self.first_rl.grid(row=0, column=1, padx=5)

        ttk.Label(top, text="Final RL (last BM):").grid(row=0, column=2, sticky="w")
        self.final_rl = ttk.Entry(top, width=12)
        self.final_rl.grid(row=0, column=3, padx=5)

        ttk.Label(
            top,
            text="Paste readings below: one row per line, columns BS | IS | FS "
                 "(tab- or comma-separated; leave blanks where no reading)."
        ).grid(row=1, column=0, columnspan=8, sticky="w", pady=(8, 2))

        self.text = tk.Text(top, height=9, width=60, font=("Consolas", 11))
        self.text.grid(row=2, column=0, columnspan=6, sticky="we")

        buttons = ttk.Frame(top)
        buttons.grid(row=2, column=6, padx=10, sticky="n")
        ttk.Button(buttons, text="Calculate", command=self.on_calculate).pack(fill="x", pady=2)
        ttk.Button(buttons, text="Import file", command=self.on_import).pack(fill="x", pady=2)
        ttk.Button(buttons, text="Export CSV", command=self.on_export).pack(fill="x", pady=2)
        ttk.Button(buttons, text="Load example", command=self.on_example).pack(fill="x", pady=2)
        ttk.Button(buttons, text="Clear", command=self.on_clear).pack(fill="x", pady=2)

    # ---------- output area ----------
    def _build_outputs(self):
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=8)
        self.tree_rf = self._make_tree(self.tabs, RF_HEADERS, "Rise and Fall")
        self.tree_hoc = self._make_tree(self.tabs, HOC_HEADERS, "Height of Collimation")
        self.tree_diff = self._make_tree(self.tabs, DIFF_HEADERS, "Difference")
        frame = ttk.Frame(self.tabs)
        self.summary = tk.Text(frame, font=("Consolas", 11), state="disabled")
        self.summary.pack(fill="both", expand=True)
        self.tabs.add(frame, text="Checks")

    def _make_tree(self, parent, headers, title):
        frame = ttk.Frame(parent)
        tree = ttk.Treeview(frame, columns=headers, show="headings")
        for h in headers:
            tree.heading(h, text=h)
            tree.column(h, width=100, anchor="center")
        tree.pack(fill="both", expand=True)
        parent.add(frame, text=title)
        return tree

    @staticmethod
    def _fill(tree, table):
        tree.delete(*tree.get_children())
        for row in table:
            tree.insert("", "end", values=row)

    # ---------- button actions ----------
    def on_calculate(self):
        try:
            first = float(self.first_rl.get())
            final = float(self.final_rl.get())
            text = self.text.get("1.0", "end")
            rf, hoc, diff, summary = calculate_all(text, first, final)
        except ValueError as err:
            messagebox.showerror("Input error", str(err) or "Check the RL values.")
            return
        self.results = (rf, hoc, diff)
        self._fill(self.tree_rf, rf)
        self._fill(self.tree_hoc, hoc)
        self._fill(self.tree_diff, diff)
        self.summary.config(state="normal")
        self.summary.delete("1.0", "end")
        self.summary.insert("1.0", "\n".join(summary))
        self.summary.config(state="disabled")

    def on_export(self):
        if not self.results:
            messagebox.showinfo("Nothing to export", "Click Calculate first.")
            return
        base = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV file", "*.csv")],
            title="Choose a base file name")
        if not base:
            return
        base = base[:-4] if base.lower().endswith(".csv") else base
        rf, hoc, diff = self.results
        write_csv(base + "_rise_fall.csv", RF_HEADERS, rf)
        write_csv(base + "_hoc.csv", HOC_HEADERS, hoc)
        write_csv(base + "_difference.csv", DIFF_HEADERS, diff)
        messagebox.showinfo("Exported", "Three CSV files saved next to:\n" + base)

    def on_import(self):
        path = filedialog.askopenfilename(
            title="Import readings",
            filetypes=[("Data files", "*.csv *.txt *.xlsx *.xlsm"),
                       ("CSV", "*.csv"), ("Text", "*.txt"),
                       ("Excel", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if not path:
            return
        try:
            text, first, final = import_file(path)
        except (ValueError, OSError) as err:
            messagebox.showerror("Import error", str(err))
            return
        self.text.delete("1.0", "end")
        self.text.insert("1.0", text)
        if first is not None:
            self.first_rl.delete(0, "end")
            self.first_rl.insert(0, str(first))
        if final is not None:
            self.final_rl.delete(0, "end")
            self.final_rl.insert(0, str(final))
        note = "Readings imported."
        if first is None or final is None:
            note += "\nNow type the missing Initial/Final RL, then click Calculate."
        else:
            note += "\nInitial and Final RL were also found. Click Calculate."
        messagebox.showinfo("Import", note)

    def on_example(self):
        self.on_clear()
        self.first_rl.insert(0, "49.873")
        self.final_rl.insert(0, "48.710")
        self.text.insert("1.0",
            "2.191\t\t\n"
            "\t2.505\t\n"
            "\t2.325\t\n"
            "3.019\t\t1.496\n"
            "\t2.513\t\n"
            "1.752\t\t2.811\n"
            "\t\t3.824\n")

    def on_clear(self):
        self.first_rl.delete(0, "end")
        self.final_rl.delete(0, "end")
        self.text.delete("1.0", "end")


if __name__ == "__main__":
    App().mainloop()
