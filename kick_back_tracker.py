#!/usr/bin/env python3
import os
import re
import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# Import your deals logic
from deals import run_deals_reports

DB_NAME = "deals.db"

# ----------------------------------------
# 1) DATABASE & HELPER FUNCTIONS
# ----------------------------------------
def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    # deals table
    c.execute("""
    CREATE TABLE IF NOT EXISTS deals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        brand TEXT NOT NULL,
        description TEXT,
        kickback_amount REAL,
        is_applied INTEGER DEFAULT 0,
        invoice_path TEXT
    )
    """)
    conn.commit()

    # Add 'credit' column if missing
    try:
        c.execute("ALTER TABLE deals ADD COLUMN credit REAL DEFAULT 0")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    # credit_history
    c.execute("""
    CREATE TABLE IF NOT EXISTS credit_history (
        history_id INTEGER PRIMARY KEY AUTOINCREMENT,
        deal_id INTEGER NOT NULL,
        old_credit REAL,
        new_credit REAL,
        change_date TEXT,
        FOREIGN KEY(deal_id) REFERENCES deals(id)
    )
    """)
    conn.commit()

    # deal_runs to avoid double-credit
    c.execute("""
    CREATE TABLE IF NOT EXISTS deal_runs (
        run_id INTEGER PRIMARY KEY AUTOINCREMENT,
        brand TEXT NOT NULL,
        start_date TEXT NOT NULL,
        end_date TEXT NOT NULL,
        total_owed REAL,
        UNIQUE(brand, start_date, end_date) ON CONFLICT IGNORE
    )
    """)
    conn.commit()

    # Insert sample data if empty
    c.execute("SELECT COUNT(*) FROM deals")
    count = c.fetchone()[0]
    if count == 0:
        sample_data = [
            ('Hashish', '50% off concentrate, 25% KB', 100.0, 0, None, 20.0),
            ('Jeeter', '50% off pre-roll, 20% KB', 250.0, 0, None, 0.0),
            ('Kiva',   '50% off Monday special, 25% KB', 150.0, 0, None, 45.0),
        ]
        for row in sample_data:
            c.execute("""
                INSERT INTO deals (brand, description, kickback_amount, is_applied, invoice_path, credit)
                VALUES (?, ?, ?, ?, ?, ?)
            """, row)
        conn.commit()

    conn.close()

def get_deals():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT * FROM deals")
    rows = c.fetchall()
    conn.close()
    return rows

def get_current_credit_in_conn(c, deal_id):
    c.execute("SELECT credit FROM deals WHERE id=?", (deal_id,))
    row = c.fetchone()
    return row[0] if row else 0.0

def log_credit_change_in_conn(c, deal_id, old_credit, new_credit):
    c.execute("""
        INSERT INTO credit_history (deal_id, old_credit, new_credit, change_date)
        VALUES (?, ?, ?, ?)
    """, (deal_id, old_credit, new_credit, datetime.now().isoformat(timespec='seconds')))

def update_deal_applied(deal_id, is_applied):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("UPDATE deals SET is_applied=? WHERE id=?", (is_applied, deal_id))
    conn.commit()
    conn.close()

def update_invoice_path(deal_id, invoice_path):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("UPDATE deals SET invoice_path=? WHERE id=?", (invoice_path, deal_id))
    conn.commit()
    conn.close()

def get_credit_history(deal_id):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("""
        SELECT history_id, old_credit, new_credit, change_date
        FROM credit_history
        WHERE deal_id=?
        ORDER BY history_id ASC
    """, (deal_id,))
    rows = c.fetchall()
    conn.close()
    return rows

# ----------------------------------------
# 2) TAB VIEW: We'll have MV tab & LM tab
# ----------------------------------------
class StoreTab(ttk.Frame):
    """
    Represents a single store tab (e.g., MV or LM).
    Displays a Treeview of deals, plus store-specific run button (if desired).
    """
    def __init__(self, parent, store_name):
        super().__init__(parent)
        self.store_name = store_name

        # Tree + scrollbar
        self.tree = ttk.Treeview(self, columns=(
            "brand", "description", "kickback_amount", "is_applied", "invoice_path", "credit"
        ), show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree.heading("brand", text="Brand")
        self.tree.heading("description", text="Deal Description")
        self.tree.heading("kickback_amount", text="Kickback $")
        self.tree.heading("is_applied", text="Applied?")
        self.tree.heading("invoice_path", text="Invoice Path")
        self.tree.heading("credit", text="Credit Owed")

        self.tree.column("brand", width=120, anchor=tk.W)
        self.tree.column("description", width=200, anchor=tk.W)
        self.tree.column("kickback_amount", width=100, anchor=tk.E)
        self.tree.column("is_applied", width=80, anchor=tk.CENTER)
        self.tree.column("invoice_path", width=140, anchor=tk.W)
        self.tree.column("credit", width=100, anchor=tk.E)

        self.scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscroll=self.scrollbar.set)

        # We'll store the parent reference so we can call shared methods
        self.main_app = None

    def refresh_data(self, all_deals):
        """
        Filter deals for this store if needed, or display them all.
        For now, we simply display all deals in each tab, or 
        you can decide to only show brand data relevant to that store, etc.
        """
        for item in self.tree.get_children():
            self.tree.delete(item)

        # If you want to filter by store, you might do:
        #   if self.store_name == "MV": show only brands that have "MV" or so
        #   if self.store_name == "LM": ...
        # But in this example, we just show them all.

        for row in all_deals:
            deal_id = row[0]
            brand   = row[1]
            desc    = row[2]
            kb_amt  = row[3]
            is_applied = row[4]
            invoice_p  = row[5] or ""
            cred_val   = row[6] if len(row) > 6 else 0.0

            is_applied_str = "Yes" if is_applied == 1 else "No"
            self.tree.insert(
                "", tk.END, iid=str(deal_id),
                values=(brand, desc, kb_amt, is_applied_str, invoice_p, cred_val)
            )

    def get_selected_deal_id(self):
        selection = self.tree.selection()
        if not selection:
            return None
        return int(selection[0])

# ----------------------------------------
# 3) MAIN APP with Notebook/Tabs
# ----------------------------------------
class KickbackTrackerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Kickback/Deals Tracker - Multi-Store")
        self.geometry("1200x600")

        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        # Create a Notebook (tab system)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Two tabs: MV and LM
        self.mv_tab = StoreTab(self.notebook, "MV")
        self.lm_tab = StoreTab(self.notebook, "LM")

        # Let them each have a reference back to the main app
        self.mv_tab.main_app = self
        self.lm_tab.main_app = self

        # Add to notebook
        self.notebook.add(self.mv_tab, text="MV Store")
        self.notebook.add(self.lm_tab, text="LM Store")

        # A frame for global controls (Run Deals, Refresh, etc.)
        self.btn_frame = ttk.Frame(self)
        self.btn_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=5)

        self.run_deals_btn = ttk.Button(self.btn_frame, text="Run Deals", command=self.run_deals_ui)
        self.run_deals_btn.pack(side=tk.LEFT, padx=5)

        self.attach_invoice_btn = ttk.Button(self.btn_frame, text="Attach Invoice", command=self.attach_invoice)
        self.attach_invoice_btn.pack(side=tk.LEFT, padx=5)

        self.mark_as_applied_btn = ttk.Button(self.btn_frame, text="Mark/Unmark Applied", command=self.mark_deal_applied)
        self.mark_as_applied_btn.pack(side=tk.LEFT, padx=5)

        self.refresh_btn = ttk.Button(self.btn_frame, text="Refresh", command=self.refresh_data)
        self.refresh_btn.pack(side=tk.LEFT, padx=5)

        self.credit_btn = ttk.Button(self.btn_frame, text="Add/Update Credit", command=self.add_update_credit)
        self.credit_btn.pack(side=tk.LEFT, padx=5)

        self.history_btn = ttk.Button(self.btn_frame, text="View Credit History", command=self.view_credit_history)
        self.history_btn.pack(side=tk.LEFT, padx=5)

        self.refresh_data()

    def get_current_tab(self):
        """Return the currently selected StoreTab instance."""
        current_tab_id = self.notebook.select()
        current_tab = self.nametowidget(current_tab_id)
        return current_tab

    def refresh_data(self):
        """Refresh both tabs with the latest deals from DB."""
        all_deals = get_deals()
        self.mv_tab.refresh_data(all_deals)
        self.lm_tab.refresh_data(all_deals)

    def run_deals_ui(self):
        """Run the deals logic, update DB, no double-credit."""
        try:
            deals_data = run_deals_reports()  # returns list of dict
        except Exception as e:
            messagebox.showerror("Deals Error", f"Error running deals: {e}")
            return

        if not deals_data:
            messagebox.showinfo("No Data", "Deals returned no brand data.")
            return

        conn = sqlite3.connect(DB_NAME, timeout=10)
        c = conn.cursor()

        for row in deals_data:
            brand = row["brand"]
            owed  = row["owed"]
            sdate = row["start"]
            edate = row["end"]

            # Insert or ignore
            c.execute("""
                INSERT INTO deal_runs (brand, start_date, end_date, total_owed)
                VALUES (?, ?, ?, ?)
            """, (brand, sdate, edate, owed))
            inserted = c.rowcount

            if inserted == 1:
                # brand is newly added for that date range
                c.execute("SELECT id FROM deals WHERE brand=?", (brand,))
                found = c.fetchone()
                if found is None:
                    c.execute("""
                        INSERT INTO deals (brand, description, kickback_amount, is_applied, invoice_path, credit)
                        VALUES (?, ?, 0, 0, NULL, 0)
                    """, (brand, f"Auto from deals run."))
                    conn.commit()
                    deal_id = c.lastrowid
                else:
                    deal_id = found[0]

                old_credit = get_current_credit_in_conn(c, deal_id)
                new_credit = old_credit + owed

                if new_credit != old_credit:
                    log_credit_change_in_conn(c, deal_id, old_credit, new_credit)

                c.execute("UPDATE deals SET credit=? WHERE id=?", (new_credit, deal_id))

        conn.commit()
        conn.close()

        self.refresh_data()
        messagebox.showinfo("Success", "Deals processed successfully!")

    def attach_invoice(self):
        """Attach an invoice for the selected row in the current tab."""
        tab = self.get_current_tab()
        deal_id = tab.get_selected_deal_id()
        if not deal_id:
            messagebox.showwarning("No selection", "Please select a row first.")
            return

        f_path = filedialog.askopenfilename(title="Select Invoice File", filetypes=[("All Files", "*.*")])
        if not f_path:
            return

        update_invoice_path(deal_id, f_path)
        self.refresh_data()
        messagebox.showinfo("Success", f"Invoice attached to deal #{deal_id}")

    def mark_deal_applied(self):
        """Toggle the is_applied flag on the selected row in the current tab."""
        tab = self.get_current_tab()
        deal_id = tab.get_selected_deal_id()
        if not deal_id:
            messagebox.showwarning("No selection", "Please select a row first.")
            return

        # We can read the Treeview values
        vals = tab.tree.item(str(deal_id), "values")
        current_str = vals[3]  # is_applied display
        new_applied = 0 if current_str == "Yes" else 1
        update_deal_applied(deal_id, new_applied)

        self.refresh_data()
        msg = "Deal marked as Applied!" if new_applied == 1 else "Deal marked as Not Applied."
        messagebox.showinfo("Updated", msg)

    def add_update_credit(self):
        """Prompt for a new credit, update the selected brand's credit."""
        tab = self.get_current_tab()
        deal_id = tab.get_selected_deal_id()
        if not deal_id:
            messagebox.showwarning("No selection", "Please select a row first.")
            return

        from tkinter.simpledialog import askstring
        ans = askstring("Credit Input", "Enter new credit amount (numbers only):")
        if ans is None:
            return
        try:
            new_credit = float(ans)
        except ValueError:
            messagebox.showerror("Invalid", "Enter a valid number.")
            return

        # Single connection for this update
        conn = sqlite3.connect(DB_NAME, timeout=10)
        c = conn.cursor()

        old_credit = get_current_credit_in_conn(c, deal_id)
        if new_credit != old_credit:
            log_credit_change_in_conn(c, deal_id, old_credit, new_credit)
        c.execute("UPDATE deals SET credit=? WHERE id=?", (new_credit, deal_id))

        conn.commit()
        conn.close()

        self.refresh_data()
        messagebox.showinfo("Success", f"Credit updated to {new_credit} for deal #{deal_id}")

    def view_credit_history(self):
        """Show the credit history for the selected row in the current tab."""
        tab = self.get_current_tab()
        deal_id = tab.get_selected_deal_id()
        if not deal_id:
            messagebox.showwarning("No selection", "Please select a row first.")
            return

        hist_win = tk.Toplevel(self)
        hist_win.title(f"Credit History for Deal #{deal_id}")
        hist_win.geometry("500x300")

        frame = ttk.Frame(hist_win)
        frame.pack(fill=tk.BOTH, expand=True)

        cols = ("old_credit", "new_credit", "change_date")
        tree_hist = ttk.Treeview(frame, columns=cols, show='headings')
        tree_hist.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tree_hist.heading("old_credit", text="Old Credit")
        tree_hist.heading("new_credit", text="New Credit")
        tree_hist.heading("change_date", text="Date Changed")

        tree_hist.column("old_credit", width=80, anchor=tk.E)
        tree_hist.column("new_credit", width=80, anchor=tk.E)
        tree_hist.column("change_date", width=200, anchor=tk.W)

        scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree_hist.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        tree_hist.configure(yscroll=scroll.set)

        rows = get_credit_history(deal_id)
        for r in rows:
            # (history_id, old_credit, new_credit, change_date)
            tree_hist.insert("", tk.END, values=(r[1], r[2], r[3]))

# ----------------------------------------
# 4) MAIN LAUNCH
# ----------------------------------------
if __name__ == "__main__":
    init_db()
    app = KickbackTrackerApp()
    app.mainloop()
