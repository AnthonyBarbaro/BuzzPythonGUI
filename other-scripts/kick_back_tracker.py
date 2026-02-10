#!/usr/bin/env python3
import os
import re
import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime


# If your deals logic returns {brand, owed, start, end, location, ...}, import that function:
from deals import run_deals_for_store  # This should return rows like {"brand":..., "owed":..., "start":..., "end":..., "location":"MV" or "LM"}

DB_NAME = "deals.db"

# ----------------------------------------
# 1) DATABASE & HELPER FUNCTIONS
# ----------------------------------------
def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    # 1) deals table with store column
    c.execute("""
    CREATE TABLE IF NOT EXISTS deals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        store TEXT NOT NULL DEFAULT 'MV', 
        brand TEXT NOT NULL,
        description TEXT,
        kickback_amount REAL,
        is_applied INTEGER DEFAULT 0,
        invoice_path TEXT,
        credit REAL DEFAULT 0
    )
    """)
    conn.commit()

    # If older DB didn't have store, add it:
    try:
        c.execute("ALTER TABLE deals ADD COLUMN store TEXT NOT NULL DEFAULT 'MV'")
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

    # deal_runs with store
    c.execute("""
    CREATE TABLE IF NOT EXISTS deal_runs (
        run_id INTEGER PRIMARY KEY AUTOINCREMENT,
        store TEXT NOT NULL DEFAULT 'MV',
        brand TEXT NOT NULL,
        start_date TEXT NOT NULL,
        end_date TEXT NOT NULL,
        total_owed REAL,
        UNIQUE(store, brand, start_date, end_date) ON CONFLICT IGNORE
    )
    """)
    conn.commit()

    # Insert sample data if empty
    c.execute("SELECT COUNT(*) FROM deals")
    count = c.fetchone()[0]
    if count == 0:
        sample_data = [

        ]
        for row in sample_data:
            c.execute("""
                INSERT INTO deals (store, brand, description, kickback_amount, is_applied, invoice_path, credit)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, row)
        conn.commit()

    conn.close()

def get_deals(store=None):
    """
    Return deals for a specific store, or all if store=None.
    Each row => (id, store, brand, description, kickback_amount, is_applied, invoice_path, credit)
    """
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    if store:
        c.execute("SELECT * FROM deals WHERE store=?", (store,))
    else:
        c.execute("SELECT * FROM deals")
    rows = c.fetchall()
    conn.close()
    return rows

def get_current_credit_in_conn(c, deal_id):
    c.execute("SELECT credit FROM deals WHERE id=?", (deal_id,))
    row = c.fetchone()
    return row[0] if row else 0.0

def log_credit_change_in_conn(c, deal_id, old_credit, new_credit):
    """Insert a record into credit_history to track changes over time."""
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
    """Returns (history_id, old_credit, new_credit, change_date) rows."""
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
# 2) TAB VIEW
# ----------------------------------------
class StoreTab(ttk.Frame):
    """A single tab for either 'MV' or 'LM' store."""
    def __init__(self, parent, store_name):
        super().__init__(parent)
        self.store_name = store_name

        self.tree = ttk.Treeview(
            self,
            columns=("store","brand","description","kickback_amount","is_applied","invoice_path","credit"),
            show='headings'
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree.heading("store", text="Store")
        self.tree.heading("brand", text="Brand")
        self.tree.heading("description", text="Description")
        self.tree.heading("kickback_amount", text="Kickback $")
        self.tree.heading("is_applied", text="Applied?")
        self.tree.heading("invoice_path", text="Invoice Path")
        self.tree.heading("credit", text="Credit Owed")

        self.tree.column("store", width=60, anchor=tk.CENTER)
        self.tree.column("brand", width=120, anchor=tk.W)
        self.tree.column("description", width=200, anchor=tk.W)
        self.tree.column("kickback_amount", width=80, anchor=tk.E)
        self.tree.column("is_applied", width=80, anchor=tk.CENTER)
        self.tree.column("invoice_path", width=140, anchor=tk.W)
        self.tree.column("credit", width=80, anchor=tk.E)

        self.scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscroll=self.scrollbar.set)

    def refresh_data(self):
        """Load deals for self.store_name, display in the Treeview."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        rows = get_deals(store=self.store_name)  # Filter by store
        for row in rows:
            # row => (id, store, brand, desc, kb_amount, is_applied, invoice_path, credit)
            deal_id  = row[0]
            store    = row[1]
            brand    = row[2]
            desc     = row[3] or ""
            kb_amt   = row[4] or 0.0
            applied  = row[5]
            inv_path = row[6] or ""
            cred_val = row[7] or 0.0

            applied_str = "Yes" if applied else "No"

            self.tree.insert(
                "", tk.END, iid=str(deal_id),
                values=(store, brand, desc, kb_amt, applied_str, inv_path, cred_val)
            )

    def get_selected_deal_id(self):
        selection = self.tree.selection()
        if not selection:
            return None
        return int(selection[0])

# ----------------------------------------
# 3) MAIN APP
# ----------------------------------------
class KickbackTrackerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Kickback/Deals Tracker - Multi-Store")
        self.geometry("1200x600")

        # Initialize DB (ensures 'store' column is present)
        init_db()

        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Two tabs for MV & LM
        self.mv_tab = StoreTab(self.notebook, "MV")
        self.lm_tab = StoreTab(self.notebook, "LM")

        self.notebook.add(self.mv_tab, text="MV Store")
        self.notebook.add(self.lm_tab, text="LM Store")

        # Frame for global buttons
        self.btn_frame = ttk.Frame(self)
        self.btn_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=5)

        ttk.Button(self.btn_frame, text="Run Deals", command=self.run_deals_ui).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="Attach Invoice", command=self.attach_invoice).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="Mark/Unmark Applied", command=self.mark_deal_applied).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="Refresh", command=self.refresh_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="Add/Update Credit", command=self.add_update_credit).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="View Credit History", command=self.view_credit_history).pack(side=tk.LEFT, padx=5)

        self.refresh_data()

    def get_current_tab(self):
        """Return the StoreTab instance that is currently selected."""
        tab_id = self.notebook.select()
        return self.nametowidget(tab_id)

    def refresh_data(self):
        """Refresh MV & LM tabs from DB."""
        self.mv_tab.refresh_data()
        self.lm_tab.refresh_data()

    def run_deals_ui(self):
        """Process deals for the current tab."""
        # Get the current tab (MV or LM)
        current_tab = self.get_current_tab()
        store = current_tab.store_name  # "MV" or "LM"

        try:
            # Call the run_deals_for_store function for the selected store
            print(f"DEBUG: Running deals for store='{store}'")
            deals_data = run_deals_for_store(store)  # Call the function with the store argument
            print(f"DEBUG: Deals processed for store='{store}':")
            for deal in deals_data:
                print(deal)
        except Exception as e:
            messagebox.showerror("Deals Error", f"Error processing deals for store '{store}': {e}")
            return

        # Insert the processed deals into the database
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        for row in deals_data:
            brand = row["brand"]
            owed = row["kickback"]
            sdate = row["start"]
            edate = row["end"]

            # Debugging output
            print(f"DEBUG: Inserting/Updating deal for store='{store}', brand='{brand}', owed={owed}")

            # Insert or ignore into deal_runs
            c.execute("""
                INSERT INTO deal_runs (store, brand, start_date, end_date, total_owed)
                VALUES (?, ?, ?, ?, ?)
            """, (store, brand, sdate, edate, owed))
            inserted = c.rowcount
            if inserted == 1:
                print(f"DEBUG: New entry added to deal_runs for store='{store}', brand='{brand}'")

            # Check if store+brand exists in deals
            c.execute("SELECT id, credit FROM deals WHERE store=? AND brand=?", (store, brand))
            found = c.fetchone()
            if not found:
                # Insert new row for this store+brand
                c.execute("""
                    INSERT INTO deals (store, brand, description, kickback_amount, is_applied, invoice_path, credit)
                    VALUES (?, ?, 'Auto-generated from run_deals_ui', 0, 0, NULL, ?)
                """, (store, brand, owed))
                print(f"DEBUG: New deal inserted for store='{store}', brand='{brand}', owed={owed}")
            else:
                # Update credit for existing row
                deal_id, old_credit = found
                new_credit = old_credit + owed
                c.execute("UPDATE deals SET credit=? WHERE id=?", (new_credit, deal_id))
                print(f"DEBUG: Updated deal for store='{store}', brand='{brand}', new_credit={new_credit}")

        conn.commit()
        conn.close()

        # Refresh the tabs
        self.refresh_data()
        messagebox.showinfo("Success", f"Deals processed for store '{store}'")




    def attach_invoice(self):
        tab = self.get_current_tab()
        deal_id = tab.get_selected_deal_id()
        if not deal_id:
            messagebox.showwarning("No selection", "Please select a row first.")
            return
        fpath = filedialog.askopenfilename(title="Select Invoice", filetypes=[("All Files","*.*")])
        if not fpath:
            return
        update_invoice_path(deal_id, fpath)
        self.refresh_data()
        messagebox.showinfo("Success", f"Invoice attached to deal {deal_id}.")

    def mark_deal_applied(self):
        tab = self.get_current_tab()
        deal_id = tab.get_selected_deal_id()
        if not deal_id:
            messagebox.showwarning("No selection", "Select a row first.")
            return

        vals = tab.tree.item(str(deal_id), "values")
        # (store, brand, desc, kb_amt, "Yes"/"No", invoice_path, credit)
        is_applied_str = vals[4]  # index 4 => "Applied?" column
        new_applied = 0 if is_applied_str == "Yes" else 1
        update_deal_applied(deal_id, new_applied)
        self.refresh_data()
        messagebox.showinfo("Updated", f"Deal marked as {'Applied' if new_applied else 'Not Applied'}.")

    def add_update_credit(self):
        tab = self.get_current_tab()
        deal_id = tab.get_selected_deal_id()
        if not deal_id:
            messagebox.showwarning("No selection", "Please select a row first.")
            return

        from tkinter.simpledialog import askstring
        ans = askstring("Credit Input", "Enter new credit amount (numbers only):")
        if not ans:
            return
        try:
            new_credit = float(ans)
        except ValueError:
            messagebox.showerror("Invalid", "Must be a numeric value.")
            return

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        old_credit = get_current_credit_in_conn(c, deal_id)
        if new_credit != old_credit:
            log_credit_change_in_conn(c, deal_id, old_credit, new_credit)
        c.execute("UPDATE deals SET credit=? WHERE id=?", (new_credit, deal_id))
        conn.commit()
        conn.close()

        self.refresh_data()
        messagebox.showinfo("Success", f"Credit updated to {new_credit} for deal {deal_id}.")

    def view_credit_history(self):
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
            # (history_id, old_cr, new_cr, dt)
            tree_hist.insert("", tk.END, values=(r[1], r[2], r[3]))

# ----------------------------------------
# 4) MAIN LAUNCH
# ----------------------------------------
if __name__ == "__main__":
    KickbackTrackerApp().mainloop()
