#!/usr/bin/env python3
import os
import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from enhanced_deals import update_deal_kickbacks, attach_invoice, update_applied_status
from deals import run_deals_for_store  # Function to fetch data for MV

DB_NAME = "deals.db"

# ----------------------------------------
# 1) DATABASE INITIALIZATION
# ----------------------------------------

from enhanced_deals import init_enhancements

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    # Create or update the deals table
    c.execute("""
    CREATE TABLE IF NOT EXISTS deals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        store TEXT NOT NULL DEFAULT 'MV',
        brand TEXT NOT NULL,
        description TEXT,
        kickback_amount REAL,
        is_applied INTEGER DEFAULT 0,
        applied_notes TEXT,
        invoice_number TEXT,
        invoice_path TEXT,
        credit REAL DEFAULT 0
    )
    """)
    conn.commit()

    # Other table initializations remain unchanged
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

    # Initialize enhancements
    init_enhancements()  # Ensures all schema updates are applied

    conn.close()

# ----------------------------------------
# 2) HELPER FUNCTIONS
# ----------------------------------------

def get_deals(store=None):
    """
    Return deals for a specific store, or all if store=None.
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

# ----------------------------------------
# 3) STORE TAB
# ----------------------------------------

class StoreTab(ttk.Frame):
    """A single tab for either 'MV' or 'LM' store."""
    def __init__(self, parent, store_name):
        super().__init__(parent)
        self.store_name = store_name

        self.tree = ttk.Treeview(
            self,
            columns=("store", "brand", "description", "kickback_amount", "is_applied", "applied_notes", "invoice_number", "invoice_path", "credit"),
            show='headings'
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree.heading("store", text="Store")
        self.tree.heading("brand", text="Brand")
        self.tree.heading("description", text="Description")
        self.tree.heading("kickback_amount", text="Kickback $")
        self.tree.heading("is_applied", text="Applied?")
        self.tree.heading("applied_notes", text="Notes")
        self.tree.heading("invoice_number", text="Invoice #")
        self.tree.heading("invoice_path", text="Invoice Path")
        self.tree.heading("credit", text="Credit Owed")

        self.tree.column("store", width=60, anchor=tk.CENTER)
        self.tree.column("brand", width=120, anchor=tk.W)
        self.tree.column("description", width=200, anchor=tk.W)
        self.tree.column("kickback_amount", width=80, anchor=tk.E)
        self.tree.column("is_applied", width=80, anchor=tk.CENTER)
        self.tree.column("applied_notes", width=140, anchor=tk.W)
        self.tree.column("invoice_number", width=100, anchor=tk.W)
        self.tree.column("invoice_path", width=140, anchor=tk.W)
        self.tree.column("credit", width=80, anchor=tk.E)

        self.scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscroll=self.scrollbar.set)

    def refresh_data(self):
        """Load deals for self.store_name, display in the Treeview."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        rows = get_deals(store=self.store_name)
        for row in rows:
            # row => (id, store, brand, description, kickback_amount, is_applied, applied_notes, invoice_number, invoice_path, credit)
            deal_id = row[0]
            store = row[1]
            brand = row[2]
            desc = row[3] or ""
            kb_amt = row[4] or 0.0
            applied = row[5]
            notes = row[6] or ""
            inv_num = row[7] or ""
            inv_path = row[8] or ""
            cred_val = row[9] or 0.0

            applied_str = "Yes" if applied else "No"

            self.tree.insert(
                "", tk.END, iid=str(deal_id),
                values=(store, brand, desc, kb_amt, applied_str, notes, inv_num, inv_path, cred_val)
            )

    def get_selected_deal_id(self):
        selection = self.tree.selection()
        if not selection:
            return None
        return int(selection[0])

# ----------------------------------------
# 4) MAIN APP
# ----------------------------------------

class KickbackTrackerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Kickback/Deals Tracker - Multi-Store")
        self.geometry("1200x600")

        # Initialize the database
        init_db()

        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create tabs for MV and LM
        self.mv_tab = StoreTab(self.notebook, "MV")
        self.lm_tab = StoreTab(self.notebook, "LM")

        self.notebook.add(self.mv_tab, text="MV Store")
        self.notebook.add(self.lm_tab, text="LM Store")

        # Frame for buttons
        self.btn_frame = ttk.Frame(self)
        self.btn_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=5)

        ttk.Button(self.btn_frame, text="Run Deals", command=self.run_deals_ui).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="Attach Invoice", command=self.attach_invoice).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="Mark/Unmark Applied", command=self.mark_deal_applied).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="Refresh", command=self.refresh_data).pack(side=tk.LEFT, padx=5)

        self.refresh_data()

    def get_current_tab(self):
        """Return the StoreTab instance that is currently selected."""
        current_tab_id = self.notebook.select()
        return self.nametowidget(current_tab_id)

    def refresh_data(self):
        """Refresh both MV and LM tabs."""
        self.mv_tab.refresh_data()
        self.lm_tab.refresh_data()

    def run_deals_ui(self):
        """Process deals for the current tab."""
        current_tab = self.get_current_tab()
        store = current_tab.store_name  # "MV" or "LM"

        try:
            print(f"DEBUG: Running deals for store='{store}'")
            deals_data = run_deals_for_store(store)
            update_deal_kickbacks(deals_data)
            self.refresh_data()
            messagebox.showinfo("Success", f"Deals processed for store '{store}'")
        except Exception as e:
            messagebox.showerror("Deals Error", f"Error processing deals for store '{store}': {e}")

    def attach_invoice(self):
        """Attach an invoice to the selected deal."""
        current_tab = self.get_current_tab()
        deal_id = current_tab.get_selected_deal_id()
        if not deal_id:
            messagebox.showwarning("No selection", "Please select a deal to attach an invoice.")
            return

        file_path = filedialog.askopenfilename(title="Select Invoice File", filetypes=[("All Files", "*.*")])
        if not file_path:
            return

        try:
            attach_invoice(deal_id, file_path)
            self.refresh_data()
            messagebox.showinfo("Success", "Invoice attached successfully.")
        except FileNotFoundError:
            messagebox.showerror("Error", "Invoice file not found.")

    def mark_deal_applied(self):
        """Mark the selected deal as applied with optional notes."""
        current_tab = self.get_current_tab()
        deal_id = current_tab.get_selected_deal_id()
        if not deal_id:
            messagebox.showwarning("No selection", "Please select a deal to mark as applied.")
            return

        from tkinter.simpledialog import askstring
        notes = askstring("Notes", "Enter notes for this deal (optional):")

        try:
            update_applied_status(deal_id, notes)
            self.refresh_data()
            messagebox.showinfo("Success", "Deal marked as applied.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to mark deal as applied: {e}")

# ----------------------------------------
# 5) MAIN LAUNCH
# ----------------------------------------

if __name__ == "__main__":
    KickbackTrackerApp().mainloop()
