import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import datetime
import sqlite3
import re
import traceback
import csv
import os

#Database Connection


class ProjectExpenditureTracker:
    def __init__(self, master):
        self.conn = sqlite3.connect("project_expenditure.db") # Ensures avoidance of conn error
        self.master = master
        self.master.title("Project Expenditure Tracker")
        self.master.geometry("1200x800")

        self.conn = sqlite3.connect('project_expenditure.db')
        self.search_active = False

        self.style = ttk.Style()
        # Create GUI widgets        
        self.create_tables()        
        self.insert_initial_metadata()
        self.create_widgets()
        # Load initial data
        self.load_data()

        # Define color scheme
        self.bg_color = "#f0f0f0"
        self.accent_color = "#4a90e2"
        self.text_color = "#333333"

        self.master.configure(bg=self.bg_color)

        # Configure styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TLabel", background=self.bg_color, foreground=self.text_color)
        self.style.configure("TButton", background=self.accent_color, foreground="white")
        self.style.map("TButton", background=[('active', self.accent_color)])
        self.style.configure("TEntry", fieldbackground="white")
        self.style.configure("TCombobox", fieldbackground="white")
        self.style.configure("Treeview", background="white", fieldbackground="white", foreground=self.text_color)
        self.style.configure("Treeview.Heading", background=self.accent_color, foreground="white")

        try:
            # Initialize database connection and setup
            
            
            self.add_fund_source_column()  # Ensure fund_source column exists
            
            
        except Exception as e:
            messagebox.showerror("Initialization Error", f"An error occurred while initializing the application: {str(e)}")
            print(f"Error details: {traceback.format_exc()}")

            self.search_active = False  # Flag to check if search has been performed

    def create_tables(self):
        cursor = self.conn.cursor()
        
        # Create main expenditures table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS expenditures (
                id INTEGER PRIMARY KEY,
                date TEXT,
                partner TEXT,
                project TEXT,
                year INTEGER,
                quarter INTEGER,
                invoice_number TEXT,
                amount REAL,
                category TEXT,
                fund_source TEXT
            )
        ''')
        
        # Create metadata table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS metadata (
                id INTEGER PRIMARY KEY,
                type TEXT,
                value TEXT
            )
        ''')
        
        # Create entry log table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS entry_log (
                id INTEGER PRIMARY KEY,
                expenditure_id INTEGER,
                timestamp TEXT,
                user TEXT,
                FOREIGN KEY (expenditure_id) REFERENCES expenditures (id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS edit_delete_log (
                id INTEGER PRIMARY KEY,
                action TEXT,
                expenditure_id INTEGER,
                old_data TEXT,
                new_data TEXT,
                timestamp TEXT,
                user TEXT
            )
        ''')
        
        self.conn.commit()
        
        # Create project-specific tables
        projects = self.get_metadata("project")
        for project in projects:
            self.create_project_table(project)

    def add_fund_source_column(self):
        cursor = self.conn.cursor()
        
        # Check if fund_source column exists in expenditures table
        cursor.execute("PRAGMA table_info(expenditures)")
        columns = [column[1] for column in cursor.fetchall()]
        
        if 'fund_source' not in columns:
            cursor.execute("ALTER TABLE expenditures ADD COLUMN fund_source TEXT")
        
        # Check and add fund_source column to all project-specific tables
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'project_%'")
        project_tables = cursor.fetchall()
        
        for table in project_tables:
            table_name = table[0]
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns = [column[1] for column in cursor.fetchall()]
            
            if 'fund_source' not in columns:
                cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN fund_source TEXT")
        
        self.conn.commit()

    def create_project_table(self, project_name):
        cursor = self.conn.cursor()
        table_name = f"project_{self.sanitize_table_name(project_name)}"
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {table_name} (
                id INTEGER PRIMARY KEY,
                date TEXT,
                partner TEXT,
                year INTEGER,
                quarter INTEGER,
                invoice_number TEXT,
                amount REAL,
                category TEXT,
                fund_source TEXT
            )
        ''')
        self.conn.commit()

        safe_table_name = self.sanitize_table_name(project_name)
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS project_{safe_table_name} (
                id INTEGER PRIMARY KEY,
                date TEXT,
                partner TEXT,
                year INTEGER,
                quarter INTEGER,
                invoice_number TEXT,
                amount REAL,
                category TEXT,
                fund_source TEXT
            )
        ''')
        self.conn.commit()

    def insert_initial_metadata(self):
        cursor = self.conn.cursor()
        initial_data = [
            ('partner', 'Partner A'),
            ('partner', 'Partner B'),
            ('project', 'Project X'),
            ('project', 'Project Y'),
            ('category', 'Equipment'),
            ('category', 'Services'),
            ('fund_source', 'Source A'),
            ('fund_source', 'Source B'),
        ]
        cursor.execute("SELECT COUNT(*) FROM metadata")
        if cursor.fetchone()[0] == 0:
            cursor.executemany("INSERT INTO metadata (type, value) VALUES (?, ?)", initial_data)
            self.conn.commit()

    def create_widgets(self):
        main_frame = ttk.Frame(self.master, padding="20")
        #main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_rowconfigure(0, weight=1)

        # Create a custom style for bold LabelFrame text
        self.style.configure("Bold.TLabelframe.Label", font=("TkDefaultFont", 9, "bold"))
        
        # Title
        title_label = ttk.Label(main_frame, text="Project Expenditure Tracker", 
                                font=("Helvetica", 24, "bold"))
        title_label.pack(pady=(0, 20))

        # Data Entry Frame
        entry_frame = ttk.LabelFrame(main_frame, text="Data Entry", padding="10", style="Bold.TLabelframe")
        entry_frame.pack(fill=tk.X, pady=(0, 20))
        #entry_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)

        # Create a grid layout for entry fields
        labels = ["Date:", "Partner:", "Project:", "Year:", "Quarter:", "Invoice #:", "Amount:", "Category:", "Fund Source:"]
        self.entries = {}
        for i, label in enumerate(labels):
            ttk.Label(entry_frame, text=label).grid(row=i//2, column=i%2*2, sticky=tk.E, padx=5, pady=5)
            if label in ["Partner:", "Project:", "Quarter:", "Category:", "Fund Source:"]:
                entry = ttk.Combobox(entry_frame)
            else:
                entry = ttk.Entry(entry_frame)
            entry.grid(row=i//2, column=i%2*2+1, sticky=(tk.W, tk.E), padx=5, pady=5)
            self.entries[label[:-1].lower()] = entry
        
        
        for i in range(5):  # Increased to 5 for the new Fund Source field
            entry_frame.columnconfigure(i, weight=1)

        # Date Entry
        ttk.Label(entry_frame, text="Date:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.date_entry = ttk.Entry(entry_frame)
        self.date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))

        # Partner Combobox
        ttk.Label(entry_frame, text="Partner:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.partner_combobox = ttk.Combobox(entry_frame, values=self.get_metadata("partner"))
        self.partner_combobox.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        # Project Combobox
        ttk.Label(entry_frame, text="Project:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.project_combobox = ttk.Combobox(entry_frame, values=self.get_metadata("project"))
        self.project_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Year Entry
        ttk.Label(entry_frame, text="Year:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.year_entry = ttk.Entry(entry_frame)
        self.year_entry.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
        self.year_entry.insert(0, datetime.now().year)

        # Quarter Combobox
        ttk.Label(entry_frame, text="Quarter:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.quarter_combobox = ttk.Combobox(entry_frame, values=[1, 2, 3, 4])
        self.quarter_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Invoice Entry
        ttk.Label(entry_frame, text="Invoice #:").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.invoice_entry = ttk.Entry(entry_frame)
        self.invoice_entry.grid(row=2, column=3, padx=5, pady=5, sticky="ew")

        # Amount Entry
        ttk.Label(entry_frame, text="Amount:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.amount_entry = ttk.Entry(entry_frame)
        self.amount_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        # Category Combobox
        ttk.Label(entry_frame, text="Category:").grid(row=3, column=2, padx=5, pady=5, sticky="e")
        self.category_combobox = ttk.Combobox(entry_frame, values=self.get_metadata("category"))
        self.category_combobox.grid(row=3, column=3, padx=5, pady=5, sticky="ew")

        # Fund Source Combobox
        ttk.Label(entry_frame, text="Fund Source:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.fund_source_combobox = ttk.Combobox(entry_frame, values=self.get_metadata("fund_source"))
        self.fund_source_combobox.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

        # Save Button
        self.save_button = ttk.Button(entry_frame, text="Save", command=self.save_record)
        #self.save_button.grid(row=5, column=1, columnspan=2, pady=10)
        self.save_button.grid(row=len(labels), column=0, columnspan=2, pady=10)

        # Search and Filter Frame
        self.search_frame = ttk.LabelFrame(main_frame, text="Search and Filter", padding="10")
        self.search_frame.pack(fill=tk.X, pady=(0, 20))
        #self.search_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)

        # Create a grid layout for search fields        
        for i in range(4):
            self.search_frame.columnconfigure(i, weight=1)

        # Project Search
        ttk.Label(self.search_frame, text="Project:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.search_project = ttk.Combobox(self.search_frame, values=["All"] + self.get_metadata("project"))
        self.search_project.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.search_project.set("All")

        # Category Search
        ttk.Label(self.search_frame, text="Category:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.search_category = ttk.Combobox(self.search_frame, values=["All"] + self.get_metadata("category"))
        self.search_category.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        self.search_category.set("All")

        # Partner Search
        ttk.Label(self.search_frame, text="Partner:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.search_partner = ttk.Combobox(self.search_frame, values=["All"] + self.get_metadata("partner"))
        self.search_partner.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.search_partner.set("All")

        # Fund Source Search
        ttk.Label(self.search_frame, text="Fund Source:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.search_fund_source = ttk.Combobox(self.search_frame, values=["All"] + self.get_metadata("fund_source"))
        self.search_fund_source.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
        self.search_fund_source.set("All")

        # Date Range Search
        ttk.Label(self.search_frame, text="Start Date:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.search_start_date = ttk.Entry(self.search_frame)
        self.search_start_date.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.search_start_date.insert(0, "YYYY-MM-DD")

        ttk.Label(self.search_frame, text="End Date:").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.search_end_date = ttk.Entry(self.search_frame)
        self.search_end_date.grid(row=2, column=3, padx=5, pady=5, sticky="ew")
        self.search_end_date.insert(0, "YYYY-MM-DD")

        # Search, Reset, and Export Buttons
        button_frame = ttk.Frame(self.search_frame)
        button_frame.grid(row=3, column=0, columnspan=4, pady=10)

        ttk.Button(button_frame, text="Search", command=self.search_records).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset", command=self.reset_search).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Export Data", command=self.export_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="View Entry Log", command=self.view_entry_log).pack(side=tk.LEFT, padx=5)

        # Add Edit and Delete buttons
       # ttk.Button(button_frame, text="Edit Selected", command=self.edit_record).pack(side=tk.LEFT, padx=5)
        #ttk.Button(button_frame, text="Delete Selected", command=self.delete_record).pack(side=tk.LEFT, padx=5)
        # Add View Edit/Delete Log button
        ttk.Button(button_frame, text="View Edit/Delete Log", command=self.view_edit_delete_log).pack(side=tk.LEFT, padx=5)
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        #self.notebook.grid(row=3, column=0, columnspan=4, sticky="nsew", padx=10, pady=10)

        # Master tab
        self.master_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.master_frame, text="Master Record")
        # Create master treeview
        self.master_tree = self.create_treeview(self.master_frame)

        # Treeview for displaying records
        self.tree = ttk.Treeview(self.master_frame, columns=("Date", "Partner", "Project", "Year", "Quarter", "Invoice#", "Amount", "Category", "Fund Source"), show="headings")
        #self.tree.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)

        # Set up headings
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)  # Adjust width as needed

        # Scrollbar for Treeview
        scrollbar = ttk.Scrollbar(self.master_frame, orient="vertical", command=self.tree.yview)
        #scrollbar.grid(row=2, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Create project-specific tabs
        self.project_trees = {}
        for project in self.get_metadata("project"):
            project_frame = ttk.Frame(self.notebook)
            self.notebook.add(project_frame, text=project)
            self.project_trees[project] = self.create_treeview(project_frame)
    
        # Modify the Edit and Delete buttons to be initially disabled
        self.edit_button = ttk.Button(button_frame, text="Edit Selected", command=self.edit_record, state='disabled')
        self.edit_button.pack(side=tk.LEFT, padx=5)
        self.delete_button = ttk.Button(button_frame, text="Delete Selected", command=self.delete_record, state='disabled')
        self.delete_button.pack(side=tk.LEFT, padx=5)
    
    
    def edit_record(self):
        if not self.search_active:
            messagebox.showwarning("Search Required", "Please perform a search before editing.")
            return        
        
        selected_item = self.master_tree.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a record to edit.")
            return

        item = self.master_tree.item(selected_item)
        values = item['values']

        # Create a new window for editing
        edit_window = tk.Toplevel(self.master)
        edit_window.title("Edit Record")

        fields = ["Date", "Partner", "Project", "Year", "Quarter", "Invoice#", "Amount", "Category", "Fund Source"]
        entries = []

        for i, field in enumerate(fields):
            ttk.Label(edit_window, text=field).grid(row=i, column=0, padx=10, pady=15)
            entry = ttk.Entry(edit_window)
            entry.insert(0, values[i])
            entry.grid(row=i, column=1, padx=5, pady=5)
            entries.append(entry)

        def save_changes():
            new_values = [entry.get() for entry in entries]
            self.update_record(values, new_values, selected_item)
            edit_window.destroy()

        ttk.Button(edit_window, text="Save Changes", command=save_changes).grid(row=len(fields), column=0, columnspan=2, pady=10)

    def update_record(self, old_values, new_values, item_id):
        try:
            cursor = self.conn.cursor()

            # Update main expenditures table
            cursor.execute('''
                UPDATE expenditures
                SET date=?, partner=?, project=?, year=?, quarter=?, invoice_number=?, amount=?, category=?, fund_source=?
                WHERE date=? AND partner=? AND project=? AND year=? AND quarter=? AND invoice_number=? AND amount=? AND category=? AND fund_source=?
            ''', new_values + old_values)

            # Handle project-specific tables
            old_project = old_values[2]
            new_project = new_values[2]
            old_table_name = f"project_{self.sanitize_table_name(old_project)}"
            new_table_name = f"project_{self.sanitize_table_name(new_project)}"

            if old_project != new_project:
                # Delete from old project table
                cursor.execute(f'''
                    DELETE FROM {old_table_name}
                    WHERE date=? AND partner=? AND year=? AND quarter=? AND invoice_number=? AND amount=? AND category=? AND fund_source=?
                ''', [old_values[i] for i in [0, 1, 3, 4, 5, 6, 7, 8]])

                # Ensure new project table exists
                self.ensure_project_table(new_project)

            # Insert or update in new project table
            cursor.execute(f'''
                INSERT OR REPLACE INTO {new_table_name}
                (date, partner, year, quarter, invoice_number, amount, category, fund_source)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', [new_values[i] for i in [0, 1, 3, 4, 5, 6, 7, 8]])

            self.conn.commit()
           # self.load_data()
           # messagebox.showinfo("Success", "Record updated successfully!")

           # Update master treeview
            self.master_tree.item(item_id, values=new_values)

            # Update project-specific treeview
            old_project = old_values[2]
            new_project = new_values[2]
            if old_project in self.project_trees:
                for item in self.project_trees[old_project].get_children():
                    if self.project_trees[old_project].item(item)['values'] == old_values:
                        if old_project == new_project:
                            self.project_trees[old_project].item(item, values=new_values)
                    else:
                        self.project_trees[old_project].delete(item)
                    break
            if new_project != old_project and new_project in self.project_trees:
                self.project_trees[new_project].insert('', 'end', values=new_values)

            messagebox.showinfo("Success", "Record updated successfully!")

            # Log the edit action
            self.log_edit_delete("edit", old_values, new_values)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while updating the record: {str(e)}")
            print(f"Update error details: {traceback.format_exc()}")
            self.conn.rollback()  # Rollback changes in case of error
            

    def delete_record(self):
        if not self.search_active:
            messagebox.showwarning("Search Required", "Please perform a search before deleting.")
            return        
        
        selected_item = self.master_tree.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a record to delete.")
            return

        if messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete this record?"):
            item = self.master_tree.item(selected_item)
            values = item['values']

            try:
                cursor = self.conn.cursor()

                # Delete from main expenditures table
                cursor.execute('''
                    DELETE FROM expenditures
                    WHERE date=? AND partner=? AND project=? AND year=? AND quarter=? AND invoice_number=? AND amount=? AND category=? AND fund_source=?
                ''', values)

                # Delete from project-specific table
                project = values[2]
                table_name = f"project_{self.sanitize_table_name(project)}"
                cursor.execute(f'''
                    DELETE FROM {table_name}
                    WHERE date=? AND partner=? AND year=? AND quarter=? AND invoice_number=? AND amount=? AND category=? AND fund_source=?
                ''', [values[i] for i in [0, 1, 3, 4, 5, 6, 7, 8]])

                self.conn.commit()

                # Remove from master treeview
                self.master_tree.delete(selected_item)

                # Remove from project-specific treeview
                if project in self.project_trees:
                    for item in self.project_trees[project].get_children():
                        if self.project_trees[project].item(item)['values'] == values:
                            self.project_trees[project].delete(item)
                            break

                # Log the delete action
                self.log_edit_delete("delete", values, None)

               # self.load_data()  # Refresh all treeviews to ensure consistency
                messagebox.showinfo("Success", "Record deleted successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while deleting the record: {str(e)}")
                print(f"Delete error details: {traceback.format_exc()}")
                self.conn.rollback()  # Rollback changes in case of error    
        
    def log_edit_delete(self, action, old_data, new_data):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        user = os.getenv('USERNAME', 'Unknown')
        
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO edit_delete_log (action, expenditure_id, old_data, new_data, timestamp, user)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (action, None, str(old_data), str(new_data), timestamp, user))

    def view_edit_delete_log(self):
        log_window = tk.Toplevel(self.master)
        log_window.title("Edit/Delete Log")
        log_window.geometry("1000x600")

        log_tree = ttk.Treeview(log_window, columns=("ID", "Action", "Timestamp", "User", "Old Data", "New Data"), show="headings")
        log_tree.heading("ID", text="Log ID")
        log_tree.heading("Action", text="Action")
        log_tree.heading("Timestamp", text="Timestamp")
        log_tree.heading("User", text="User")
        log_tree.heading("Old Data", text="Old Data")
        log_tree.heading("New Data", text="New Data")
        log_tree.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_window, orient="vertical", command=log_tree.yview)
        scrollbar.pack(side="right", fill="y")
        log_tree.configure(yscrollcommand=scrollbar.set)

        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT id, action, timestamp, user, old_data, new_data
            FROM edit_delete_log
            ORDER BY timestamp DESC
        ''')
        for row in cursor.fetchall():
            log_tree.insert('', 'end', values=row)

    
    
    def create_treeview(self, parent):
        columns = ("Date", "Partner", "Project", "Year", "Quarter", "Invoice#", "Amount", "Category", "Fund Source")
        tree = ttk.Treeview(parent, columns=columns, show="headings")
        tree.heading("Date", text="Date")
        tree.heading("Partner", text="Partner")
        tree.heading("Project", text="Project")
        tree.heading("Year", text="Year")
        tree.heading("Quarter", text="Quarter")
        tree.heading("Invoice#", text="Invoice#")
        tree.heading("Amount", text="Amount")
        tree.heading("Category", text="Category")
        tree.heading("Fund Source", text="Fund Source")

        # Set column widths
        tree.column("Date", width=100)
        tree.column("Partner", width=150)
        tree.column("Project", width=150)
        tree.column("Year", width=60)
        tree.column("Quarter", width=60)
        tree.column("Invoice#", width=100)
        tree.column("Amount", width=100)
        tree.column("Category", width=150)
        tree.column("Fund Source", width=150)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)  # Adjust width as needed   
        
        tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        scrollbar.pack(side="right", fill="y")
        tree.configure(yscrollcommand=scrollbar.set)

        return tree

    def save_record(self):
        try:
            date = self.date_entry.get()
            partner = self.partner_combobox.get()
            project = self.project_combobox.get()
            year = int(self.year_entry.get())
            quarter = int(self.quarter_combobox.get())
            invoice = self.invoice_entry.get()
            amount = float(self.amount_entry.get())
            category = self.category_combobox.get()
            fund_source = self.fund_source_combobox.get()

            # Check and add new metadata if necessary
            self.add_new_metadata("partner", partner)
            self.add_new_metadata("project", project)
            self.add_new_metadata("category", category)
            self.add_new_metadata("fund_source", fund_source)

            cursor = self.conn.cursor()
            
            # Save to main expenditures table
            cursor.execute('''
                INSERT INTO expenditures (date, partner, project, year, quarter, invoice_number, amount, category, fund_source)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (date, partner, project, year, quarter, invoice, amount, category, fund_source))
            
            expenditure_id = cursor.lastrowid

            # Log the entry
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            user = os.getenv('USERNAME', 'Unknown')
            cursor.execute('''
                INSERT INTO entry_log (expenditure_id, timestamp, user)
                VALUES (?, ?, ?)
            ''', (expenditure_id, timestamp, user))

            # Save to project-specific table
            self.ensure_project_table(project)
            table_name = f"project_{self.sanitize_table_name(project)}"
            cursor.execute(f'''
                INSERT INTO {table_name} (date, partner, year, quarter, invoice_number, amount, category, fund_source)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (date, partner, year, quarter, invoice, amount, category, fund_source))
            
            self.conn.commit()

            # Update master treeview
            self.master_tree.insert('', 'end', values=(date, partner, project, year, quarter, invoice, amount, category, fund_source))

            # Update or create project-specific treeview
            if project not in self.project_trees:
                self.add_project_tab(project)
            self.project_trees[project].insert('', 'end', values=(date, partner, project, year, quarter, invoice, amount, category, fund_source))

            self.clear_entries()
            self.load_data()
            self.update_comboboxes()
            messagebox.showinfo("Success", "Record saved successfully!")
        except ValueError:
            messagebox.showerror("Error", "Invalid input. Please check your entries.")
        except sqlite3.OperationalError as e:
            messagebox.showerror("Error", f"Database error: {str(e)}")

    def add_new_metadata(self, metadata_type, value):
        if value and value not in self.get_metadata(metadata_type):
            cursor = self.conn.cursor()
            cursor.execute("INSERT INTO metadata (type, value) VALUES (?, ?)", (metadata_type, value))
            self.conn.commit()

    def ensure_project_table(self, project_name):
        table_name = f"project_{self.sanitize_table_name(project_name)}"
        cursor = self.conn.cursor()
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {table_name} (
                id INTEGER PRIMARY KEY,
                date TEXT,
                partner TEXT,
                year INTEGER,
                quarter INTEGER,
                invoice_number TEXT,
                amount REAL,
                category TEXT,
                fund_source TEXT
            )
        ''')
        self.conn.commit()

    def add_project_tab(self, project):
        project_frame = ttk.Frame(self.notebook)
        self.notebook.add(project_frame, text=project)
        self.project_trees[project] = self.create_treeview(project_frame)

    def search_records(self):
        project = self.search_project.get()
        category = self.search_category.get()
        partner = self.search_partner.get()
        fund_source = self.search_fund_source.get()
        start_date = self.search_start_date.get()
        end_date = self.search_end_date.get()

        # Clear all treeviews
        self.master_tree.delete(*self.master_tree.get_children())
        for tree in self.project_trees.values():
            tree.delete(*tree.get_children())

        cursor = self.conn.cursor()

        # Base query for master data
        master_query = """
        SELECT date, partner, project, year, quarter, invoice_number, amount, category, fund_source 
        FROM expenditures WHERE 1=1
        """
        params = []

        # Add filter conditions
        if project != "All":
            master_query += " AND project = ?"
            params.append(project)
        if category != "All":
            master_query += " AND category = ?"
            params.append(category)
        if partner != "All":
            master_query += " AND partner = ?"
            params.append(partner)
        if fund_source != "All":
            master_query += " AND fund_source = ?"
            params.append(fund_source)
        if start_date != "YYYY-MM-DD":
            master_query += " AND date >= ?"
            params.append(start_date)
        if end_date != "YYYY-MM-DD":
            master_query += " AND date <= ?"
            params.append(end_date)

        # Execute master query and update master treeview
        cursor.execute(master_query, tuple(params))
        for row in cursor.fetchall():
            self.master_tree.insert('', 'end', values=row)

        # Search in project-specific tables
        for project_name, tree in self.project_trees.items():
            table_name = f"project_{self.sanitize_table_name(project_name)}"
            project_query = f"""
            SELECT date, partner, ?, year, quarter, invoice_number, amount, category, fund_source 
            FROM {table_name} WHERE 1=1
            """
            project_params = [project_name]

            # Apply filters to project-specific query
            if project != "All" and project != project_name:
                continue  # Skip this project if it's not the selected one
            if category != "All":
                project_query += " AND category = ?"
                project_params.append(category)
            if partner != "All":
                project_query += " AND partner = ?"
                project_params.append(partner)
            if fund_source != "All":
                project_query += " AND fund_source = ?"
                project_params.append(fund_source)
            if start_date != "YYYY-MM-DD":
                project_query += " AND date >= ?"
                project_params.append(start_date)
            if end_date != "YYYY-MM-DD":
                project_query += " AND date <= ?"
                project_params.append(end_date)

            # Execute project-specific query and update project treeview
            cursor.execute(project_query, tuple(project_params))
            for row in cursor.fetchall():
                tree.insert('', 'end', values=row)

        
        # After search is complete:
        self.search_active = True
        self.edit_button['state'] = 'normal'
        self.delete_button['state'] = 'normal'

        # Update status or show a message about the search results
        total_results = len(self.master_tree.get_children())
        messagebox.showinfo("Search Results", f"Found {total_results} matching records across all projects.")

    
    def reset_search(self):
        self.search_project.set("All")
        self.search_category.set("All")
        self.search_partner.set("All")
        self.search_fund_source.set("All")
        self.search_start_date.delete(0, tk.END)
        self.search_start_date.insert(0, "YYYY-MM-DD")
        self.search_end_date.delete(0, tk.END)
        self.search_end_date.insert(0, "YYYY-MM-DD")
        

        # After reset:
        self.search_active = False
        self.edit_button['state'] = 'disabled'
        self.delete_button['state'] = 'disabled'
        self.load_data()

    def clear_entries(self):
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.partner_combobox.set('')
        self.project_combobox.set('')
        self.year_entry.delete(0, tk.END)
        self.year_entry.insert(0, datetime.now().year)
        self.quarter_combobox.set('')
        self.invoice_entry.delete(0, tk.END)
        self.amount_entry.delete(0, tk.END)
        self.category_combobox.set('')
        self.fund_source_combobox.set('')

    def update_comboboxes(self):
        partners = self.get_metadata("partner")
        projects = self.get_metadata("project")
        categories = self.get_metadata("category")
        fund_sources = self.get_metadata("fund_source")

        self.partner_combobox['values'] = partners
        self.project_combobox['values'] = projects
        self.category_combobox['values'] = categories
        self.fund_source_combobox['values'] = fund_sources
        self.search_project['values'] = ["All"] + projects
        self.search_category['values'] = ["All"] + categories
        self.search_partner['values'] = ["All"] + partners
        self.search_fund_source['values'] = ["All"] + fund_sources

    def load_data(self):
        # Clear all treeviews
        self.master_tree.delete(*self.master_tree.get_children())
        for tree in self.project_trees.values():
            tree.delete(*tree.get_children())
        cursor = self.conn.cursor()    

        # Load data into master treeview        
        cursor.execute("SELECT * FROM expenditures")
        #rows = cursor.fetchall()
        #self.tree.delete(*self.tree.get_children())
        # for row in rows:
        #     self.tree.insert("", "end", values=row[1:])  # Exclude ID from display
        for row in cursor.fetchall():
            self.master_tree.insert("", "end", values=row[1:])


        # Load data into project-specific treeviews
        for project in self.project_trees:
            table_name = f"project_{self.sanitize_table_name(project)}"
            cursor.execute(f"SELECT date, partner, ?, year, quarter, invoice_number, amount, category, fund_source FROM {table_name}", (project,))
            for row in cursor.fetchall():
                self.project_trees[project].insert('', 'end', values=row)

    def export_data(self):
        try:
            current_tab = self.notebook.tab(self.notebook.select(), "text")
            current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"{current_tab.lower().replace(' ', '_')}_{current_datetime}.csv"

            file_path = filedialog.asksaveasfilename(
                initialfile=default_filename,
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")]
            )

            if not file_path:
                return

            if current_tab == "Master Record":
                tree = self.master_tree
            else:
                tree = self.project_trees[current_tab]
            
            headers = [tree.heading(col)["text"] for col in tree["columns"]]
            data = []
            for item in tree.get_children():
                values = tree.item(item)["values"]
                data.append(values)

            with open(file_path, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(headers)
                writer.writerows(data)

            messagebox.showinfo("Export Successful", f"Data exported successfully to {file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while exporting: {str(e)}")
            print(f"Export error details: {traceback.format_exc()}")

    def view_entry_log(self):
        log_window = tk.Toplevel(self.master)
        log_window.title("Entry Log")
        log_window.geometry("800x600")

        log_tree = ttk.Treeview(log_window, columns=("ID", "Timestamp", "User", "Project", "Amount"), show="headings")
        log_tree.heading("ID", text="Entry ID")
        log_tree.heading("Timestamp", text="Timestamp")
        log_tree.heading("User", text="User")
        log_tree.heading("Project", text="Project")
        log_tree.heading("Amount", text="Amount")
        log_tree.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_window, orient="vertical", command=log_tree.yview)
        scrollbar.pack(side="right", fill="y")
        log_tree.configure(yscrollcommand=scrollbar.set)

        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT entry_log.id, entry_log.timestamp, entry_log.user, expenditures.project, expenditures.amount
            FROM entry_log
            JOIN expenditures ON entry_log.expenditure_id = expenditures.id
            ORDER BY entry_log.timestamp DESC
        ''')
        for row in cursor.fetchall():
            log_tree.insert('', 'end', values=row)

    def get_metadata(self, metadata_type):
        cursor = self.conn.cursor()
        cursor.execute("SELECT value FROM metadata WHERE type=?", (metadata_type,))
        return [row[0] for row in cursor.fetchall()]

    def sanitize_table_name(self, name):
        sanitized = re.sub(r'[^\w]', '_', name)
        if not sanitized[0].isalpha():
            sanitized = 'p_' + sanitized
        return "".join(c.lower() if c.isalnum() else "_" for c in name)

# Main execution
def main():
    root = tk.Tk()
    app = ProjectExpenditureTracker(root)
    root.mainloop()

if __name__ == "__main__":
    main()
