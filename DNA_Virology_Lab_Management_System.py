import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sqlite3
from datetime import datetime
import csv
import os
from pathlib import Path
import qrcode
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ReportLabImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import tempfile
import shutil
import hashlib
import json
import threading
import random

# Simulated IoT Device Integration
class IoTDevice:
    def __init__(self, device_id):
        self.device_id = device_id
        self.status = "online"
    
    def scan_item(self, item_id):
        if self.status == "online":
            return f"Item {item_id} scanned by device {self.device_id}"
        else:
            return "Device offline"

# Simulated Blockchain Integration
class Blockchain:
    def __init__(self):
        self.chain = []
        self.create_block(proof=1, previous_hash='0')
    
    def create_block(self, proof, previous_hash):
        block = {
            'index': len(self.chain) + 1,
            'timestamp': str(datetime.now()),
            'proof': proof,
            'previous_hash': previous_hash,
            'transactions': []
        }
        self.chain.append(block)
        return block
    
    def add_transaction(self, transaction):
        self.chain[-1]['transactions'].append(transaction)
        return self.chain[-1]

# Simulated AI Assistant
class AIAssistant:
    def __init__(self):
        self.name = "LabAI"
    
    def predict_inventory_needs(self, inventory_data):
        # Basic prediction logic, ignoring equipment
        low_stock_items = [item for item in inventory_data if item['quantity'] < 10 and item['item_type'] != 'equipment']
        return low_stock_items

# Equipment Cover Generator
class EquipmentCoverGenerator:
    def __init__(self):
        self.page_width, self.page_height = A4
        self.margin = 25 * mm

    def create_custom_styles(self):
        """Create custom styles for the PDF"""
        styles = getSampleStyleSheet()
        
        # Equipment name style
        styles.add(ParagraphStyle(
            name='EquipmentName',
            parent=styles['Title'],
            fontSize=24,
            alignment=TA_CENTER,
            spaceAfter=10 * mm,
            textColor=colors.black,
            backColor=colors.lightgrey
        ))
        
        # Lab name style
        styles.add(ParagraphStyle(
            name='LabName',
            parent=styles['Title'],
            fontSize=20,
            alignment=TA_CENTER,
            spaceAfter=5 * mm
        ))
        
        # Section header style
        styles.add(ParagraphStyle(
            name='SectionHeader',
            parent=styles['Heading2'],
            fontSize=14,
            spaceBefore=15,
            spaceAfter=10
        ))
        
        return styles

# Enhanced Lab Inventory System
class LabInventorySystem:
    def __init__(self, root):
        self.root = root
        self.root.title("DNA Virology Lab Management System-ICGEB China RRC")
        self.root.geometry("1200x700")
        
        # Setup directories
        self.base_dir = Path.home() / "DNA_Virology_Lab_System"
        self.setup_directories()
        
        # Initialize database
        self.init_database()
        
        # Initialize IoT Device
        self.iot_device = IoTDevice("IoT-001")
        
        # Initialize Blockchain
        self.blockchain = Blockchain()
        
        # Initialize AI Assistant
        self.ai_assistant = AIAssistant()
        
        # Initialize Equipment Cover Generator
        self.cover_generator = EquipmentCoverGenerator()
        
        # Create main frames
        self.create_frames()
        
        # Create menu bar
        self.create_menu()
        
        # Initialize tabs
        self.create_tabs()

    def setup_directories(self):
        """Setup necessary directories for the application"""
        self.dirs = {
            'data': self.base_dir / "data",
            'exports': self.base_dir / "exports",
            'qrcodes': self.base_dir / "qrcodes",
            'documents': self.base_dir / "documents",
            'temp': self.base_dir / "temp"
        }
        
        for dir_path in self.dirs.values():
            dir_path.mkdir(parents=True, exist_ok=True)

    def init_database(self):
        """Initialize the SQLite database with datetime support"""
        try:
            db_path = self.dirs['data'] / "lab_inventory.db"
            self.conn = sqlite3.connect(str(db_path), detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES)
            self.cursor = self.conn.cursor()

            # Register custom datetime adapters and converters
            sqlite3.register_adapter(datetime, lambda dt: dt.isoformat())
            sqlite3.register_converter("timestamp", lambda dt: datetime.fromisoformat(dt.decode()))

            # Create tables with proper foreign key support
            self.cursor.execute("PRAGMA foreign_keys = ON")
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS items (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    name_cn TEXT,
                    item_type TEXT NOT NULL,
                    category TEXT,
                    location TEXT,
                    quantity INTEGER DEFAULT 0,
                    unit TEXT,
                    manufacturer TEXT,
                    model_number TEXT,
                    serial_number TEXT,
                    purchase_date TEXT,
                    warranty_until TEXT,
                    maintenance_contact TEXT,
                    last_calibration TEXT,
                    next_calibration TEXT,
                    safety_classification TEXT,
                    last_updated TIMESTAMP,
                    notes TEXT
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS usage_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_id TEXT,
                    user TEXT NOT NULL,
                    user_department TEXT,
                    quantity_changed INTEGER,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    purpose TEXT,
                    notes TEXT,
                    supervisor_approval TEXT,
                    return_time TIMESTAMP,
                    FOREIGN KEY (item_id) REFERENCES items (id) ON DELETE CASCADE
                )
            """)
            self.conn.commit()
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to initialize database: {str(e)}")
            raise

    def create_frames(self):
        """Create main application frames"""
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.status_bar = ttk.Label(self.root, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def create_menu(self):
        """Create menu bar"""
        self.menubar = tk.Menu(self.root)
        self.root.config(menu=self.menubar)
        
        # File Menu
        file_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="File 文件", menu=file_menu)
        file_menu.add_command(label="Export Inventory 导出库存", command=self.export_inventory)
        file_menu.add_command(label="Export Usage Log 导出使用记录", command=self.export_usage_log)
        file_menu.add_separator()
        file_menu.add_command(label="Exit 退出", command=self.root.quit)
        
        # Tools Menu
        tools_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Tools 工具", menu=tools_menu)
        tools_menu.add_command(label="Generate Report 生成报告", command=self.generate_report)
        tools_menu.add_command(label="Backup Database 备份数据库", command=self.backup_database)
        tools_menu.add_command(label="AI Predict Inventory Needs AI预测库存需求", command=self.ai_predict_inventory_needs)
        
        # About Menu
        about_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="About 关于", menu=about_menu)
        about_menu.add_command(label="About 关于", command=self.show_about)

    def create_tabs(self):
        """Create main application tabs"""
        self.tab_control = ttk.Notebook(self.main_frame)
        self.tab_control.pack(fill=tk.BOTH, expand=True)
        
        # Equipment Tab
        self.equipment_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.equipment_tab, text="Equipment 设备")
        self.create_inventory_tab(self.equipment_tab, "equipment")
        
        # Chemicals Tab
        self.chemicals_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.chemicals_tab, text="Chemicals 化学品")
        self.create_inventory_tab(self.chemicals_tab, "chemical")
        
        # Consumables Tab
        self.consumables_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.consumables_tab, text="Consumables 消耗品")
        self.create_inventory_tab(self.consumables_tab, "consumable")
        
        # Other Tab
        self.other_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.other_tab, text="Other 其他")
        self.create_inventory_tab(self.other_tab, "other")
        
        # Usage Log Tab
        self.usage_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.usage_tab, text="Usage Log 使用记录")
        self.create_usage_tab()

    def create_inventory_tab(self, parent, item_type):
        """Create inventory management tab with improved layout"""
        # Control Frame
        control_frame = ttk.LabelFrame(parent, text="Controls 控制")
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(control_frame, text="Add Item 添加物品", 
                command=lambda: self.add_item(item_type)).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(control_frame, text="Edit Item 编辑物品",
                command=lambda: self.edit_item(item_type)).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(control_frame, text="Delete Item 删除物品",
                command=lambda: self.delete_item(item_type)).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(control_frame, text="Generate QR Code 生成二维码",
                command=lambda: self.generate_qr_code(item_type)).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(control_frame, text="Generate File Cover 生成文件封面",
                command=lambda: self.generate_file_covers(item_type)).pack(side=tk.LEFT, padx=5, pady=5)
        
        # Search Frame
        search_frame = ttk.LabelFrame(parent, text="Search 搜索")
        search_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(search_frame, text="Search 搜索:").pack(side=tk.LEFT, padx=5, pady=5)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        
        # Treeview
        columns = ("ID", "Name", "Name_CN", "Category", "Location", "Quantity", "Unit", 
                "Manufacturer", "Model Number", "Serial Number", "Purchase Date", 
                "Warranty Until", "Maintenance Contact", "Last Calibration", 
                "Next Calibration", "Safety Classification")
        tree = ttk.Treeview(parent, columns=columns, show='headings')
        tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbars
        y_scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scrollbar = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=tree.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, minwidth=50)
        
        # Store tree reference and bind search
        if item_type == "equipment":
            self.equipment_tree = tree
        elif item_type == "chemical":
            self.chemicals_tree = tree
        elif item_type == "consumable":
            self.consumables_tree = tree
        else:
            self.other_tree = tree
            
        search_entry.bind('<KeyRelease>', lambda e: self.search_items(item_type, tree, search_var))
        
        # Initial data load
        self.refresh_inventory(item_type, tree)

    def generate_file_covers(self, item_type):
        """Generate file covers for selected items"""
        try:
            tree = self.equipment_tree if item_type == "equipment" else self.chemicals_tree if item_type == "chemical" else self.consumables_tree if item_type == "consumable" else self.other_tree
            selected = tree.selection()
            
            if not selected:
                messagebox.showwarning("Warning", "Please select an item to generate a file cover 请选择要生成文件封面的物品")
                return
            
            item_id = tree.item(selected[0])['values'][0]
            
            cover_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                initialdir=self.dirs['documents'],
                title="Save File Cover",
                filetypes=[("PDF files", "*.pdf")]
            )
            
            if not cover_path:
                return

            # Fetch item details
            self.cursor.execute("SELECT * FROM items WHERE id = ?", (item_id,))
            item_data = self.cursor.fetchone()
            
            # Create PDF with custom page setup
            doc = SimpleDocTemplate(
                cover_path,
                pagesize=A4,
                rightMargin=self.cover_generator.margin,
                leftMargin=self.cover_generator.margin,
                topMargin=self.cover_generator.margin,
                bottomMargin=self.cover_generator.margin
            )
            
            styles = self.cover_generator.create_custom_styles()
            elements = []
            
            # Equipment Name Box
            elements.append(Paragraph(
                f"{item_data[1]}",
                styles['EquipmentName']
            ))
            
            # Lab Information
            elements.append(Paragraph(
                "DNA Virology Lab",
                styles['LabName']
            ))
            elements.append(Paragraph(
                "ICGEB China RRC",
                styles['LabName']
            ))
            elements.append(Spacer(1, 20))
            
            # Create information table
            data = [
                ["Information", ""],
                ["Manufacturer", item_data[8]],
                ["Model Number", item_data[9]],
                ["Serial Number", item_data[10]],
                ["Location", item_data[5]],
                ["Purchase Date", item_data[11]],
                ["Warranty Until", item_data[12]],
                ["Maintenance Contact", item_data[13]],
                ["Last Calibration", item_data[14]],
                ["Next Calibration", item_data[15]],
                ["Safety Classification", item_data[16]]
            ]
            
            # Table Style
            table_style = TableStyle([
                # Header row
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 14),
                ('SPAN', (0, 0), (-1, 0)),  # Merge first row
                
                # Content rows
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),  # Left align labels
                ('ALIGN', (1, 1), (1, -1), 'LEFT'),  # Left align values
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 12),
                ('TOPPADDING', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                ('GRID', (0, 1), (-1, -1), 1, colors.black),
                ('BOX', (0, 0), (-1, -1), 2, colors.black),
            ])
            
            # Create and style the table
            table = Table(data, colWidths=[doc.width * 0.4, doc.width * 0.6])
            table.setStyle(table_style)
            elements.append(table)
            
            # Add generation date
            elements.append(Spacer(1, 20))
            elements.append(Paragraph(
                f"Generated on : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                styles['Normal']
            ))
            
            # Build PDF
            doc.build(elements)
            messagebox.showinfo("Success", f"File cover generated successfully!\n文件封面已生成: {cover_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate file cover: {str(e)}")

    def on_closing(self):
        """Clean up database connection when closing"""
        if hasattr(self, 'conn'):
            self.conn.close()

    def generate_qr_code(self, item_type):
        """Generate QR code for selected item"""
        tree = self.equipment_tree if item_type == "equipment" else self.chemicals_tree if item_type == "chemical" else self.consumables_tree if item_type == "consumable" else self.other_tree
        selected = tree.selection()
        
        if not selected:
            messagebox.showwarning("Warning", "Please select an item to generate QR code 请选择要生成二维码的物品")
            return
        
        try:
            item_id = tree.item(selected[0])['values'][0]
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            
            qr_data = f"Item ID: {item_id}\nProperty of DNA Virology Lab-ICGEB China RRC"
            qr.add_data(qr_data)
            qr.make(fit=True)
            
            qr_img = qr.make_image(fill_color="black", back_color="white")
            qr_path = self.dirs['qrcodes'] / f"item_{item_id}_qr.png"
            qr_img.save(qr_path)
            
            messagebox.showinfo("Success", f"QR code generated successfully!\n二维码已生成: {qr_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate QR code: {str(e)}")

    def create_usage_tab(self):
        """Create usage log tab with improved layout"""
        # Control Frame
        control_frame = ttk.Frame(self.usage_tab)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(control_frame, text="Add Usage Log 添加使用记录",
                  command=self.add_usage_log).pack(side=tk.LEFT, padx=5)
        
        # Treeview
        columns = ("ID", "Item", "User", "Department", "Quantity", "Date", "Purpose", "Status")
        self.usage_tree = ttk.Treeview(self.usage_tab, columns=columns, show='headings')
        self.usage_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbars
        y_scrollbar = ttk.Scrollbar(self.usage_tab, orient=tk.VERTICAL, command=self.usage_tree.yview)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scrollbar = ttk.Scrollbar(self.usage_tab, orient=tk.HORIZONTAL, command=self.usage_tree.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.usage_tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        
        # Configure columns
        for col in columns:
            self.usage_tree.heading(col, text=col)
            self.usage_tree.column(col, width=100, minwidth=50)
        
        self.refresh_usage_log()

    def add_usage_log(self):
        """Add a new usage log entry"""
        add_window = tk.Toplevel(self.root)
        add_window.title("Add Usage Log 添加使用记录")
        add_window.geometry("400x400")
        add_window.transient(self.root)
        add_window.grab_set()
        
        # Create scrollable frame
        canvas = tk.Canvas(add_window)
        scrollbar = ttk.Scrollbar(add_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Dropdown for item selection at the top
        ttk.Label(scrollable_frame, text="Select Item 选择物品").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.item_var = tk.StringVar()
        self.item_dropdown = ttk.Combobox(scrollable_frame, textvariable=self.item_var)
        self.item_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.item_dropdown.bind("<<ComboboxSelected>>", self.on_item_select)
        self.item_dropdown.bind("<KeyRelease>", lambda e: self.filter_item_dropdown())
        self.populate_item_dropdown()
        
        # Fields
        fields = {}
        row = 1
        field_configs = [
            ("user", "User 用户 *", True),
            ("user_department", "Department 部门", False),
            ("quantity_changed", "Quantity Changed 数量变化 *", True),
            ("purpose", "Purpose 目的", False),
            ("notes", "Notes 备注", False),
            ("supervisor_approval", "Supervisor Approval 主管批准", False)
        ]
        
        for field, label, required in field_configs:
            ttk.Label(scrollable_frame, text=label).grid(row=row, column=0, padx=5, pady=5, sticky="w")
            if field == "notes":
                fields[field] = tk.Text(scrollable_frame, height=3, width=30)
            else:
                fields[field] = ttk.Entry(scrollable_frame)
            fields[field].grid(row=row, column=1, padx=5, pady=5, sticky="ew")
            row += 1
        
        def validate_and_save():
            # Validation
            selected_item = self.item_var.get()
            if not selected_item:
                messagebox.showerror("Error", "Please select an item")
                return
            
            item_id = selected_item.split(" - ")[0]
            
            if not fields['user'].get().strip():
                messagebox.showerror("Error", "User is required")
                return
            
            try:
                quantity_changed = int(fields['quantity_changed'].get())
                if quantity_changed == 0:
                    raise ValueError("Quantity changed cannot be zero")
            except ValueError as e:
                messagebox.showerror("Error", f"Invalid quantity: {str(e)}")
                return
            
            # Check if the item has enough quantity
            self.cursor.execute("SELECT quantity FROM items WHERE id = ?", (item_id,))
            current_quantity = self.cursor.fetchone()[0]
            
            if current_quantity - quantity_changed < 0:
                messagebox.showerror("Error", "Not enough quantity in stock")
                return
            
            # Save to database
            try:
                # Update the item's quantity
                self.cursor.execute("""
                    UPDATE items
                    SET quantity = quantity - ?
                    WHERE id = ?
                """, (quantity_changed, item_id))
                
                # Insert the usage log entry
                self.cursor.execute("""
                    INSERT INTO usage_log (
                        item_id, user, user_department, quantity_changed,
                        purpose, notes, supervisor_approval
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    item_id,
                    fields['user'].get().strip(),
                    fields['user_department'].get().strip(),
                    quantity_changed,
                    fields['purpose'].get().strip(),
                    fields['notes'].get("1.0", tk.END).strip(),
                    fields['supervisor_approval'].get().strip()
                ))
                
                # Add transaction to blockchain
                transaction = {
                    'item_id': item_id,
                    'user': fields['user'].get().strip(),
                    'quantity_changed': quantity_changed,
                    'timestamp': str(datetime.now())
                }
                self.blockchain.add_transaction(transaction)
                
                self.conn.commit()
                self.refresh_usage_log()
                add_window.destroy()
                messagebox.showinfo("Success", "Usage log added successfully! 使用记录添加成功！")
            
            except Exception as e:
                self.conn.rollback()
                messagebox.showerror("Error", f"Failed to add usage log: {str(e)}")
        
        # Save button
        ttk.Button(scrollable_frame, text="Save 保存", command=validate_and_save).grid(
            row=row, column=0, columnspan=2, pady=20)
        
        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def filter_item_dropdown(self):
        """Filter the dropdown menu based on user input"""
        search_term = self.item_var.get().lower()
        self.item_dropdown['values'] = [item for item in self.item_dropdown['values'] if search_term in item.lower()]

    def on_item_select(self, event):
        """Handle item selection from dropdown"""
        selected_item = self.item_var.get()
        if selected_item:
            item_id = selected_item.split(" - ")[0]

    def populate_item_dropdown(self):
        """Populate the item dropdown with item IDs and names"""
        self.cursor.execute("SELECT id, name FROM items")
        items = self.cursor.fetchall()
        self.item_dropdown['values'] = [f"{item[0]} - {item[1]}" for item in items]

    def add_item(self, item_type):
        """Add a new item to inventory with improved validation"""
        add_window = tk.Toplevel(self.root)
        add_window.title(f"Add {item_type.capitalize()} 添加{item_type}")
        add_window.geometry("400x800")  # Increased height to accommodate more fields
        add_window.transient(self.root)
        add_window.grab_set()
        
        # Create scrollable frame
        canvas = tk.Canvas(add_window)
        scrollbar = ttk.Scrollbar(add_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Fields
        fields = {}
        row = 0
        field_configs = [
            ("name", "Name 名称 *", True),
            ("name_cn", "Chinese Name 中文名称", False),
            ("category", "Category 类别", False),
            ("location", "Location 位置", False),
            ("quantity", "Quantity 数量 *", True),
            ("unit", "Unit 单位 *", True),
            ("manufacturer", "Manufacturer 制造商", False),
            ("model_number", "Model Number 型号", False),
            ("serial_number", "Serial Number 序列号", False),
            ("purchase_date", "Purchase Date 购买日期 (YYYY-MM-DD)", False),
            ("warranty_until", "Warranty Until 保修至 (YYYY-MM-DD)", False),
            ("maintenance_contact", "Maintenance Contact 维护联系人", False),
            ("last_calibration", "Last Calibration 上次校准 (YYYY-MM-DD)", False),
            ("next_calibration", "Next Calibration 下次校准 (YYYY-MM-DD)", False),
            ("safety_classification", "Safety Classification 安全分类", False),
            ("notes", "Notes 备注", False)
        ]
        
        for field, label, required in field_configs:
            ttk.Label(scrollable_frame, text=label).grid(row=row, column=0, padx=5, pady=5, sticky="w")
            if field == "notes":
                fields[field] = tk.Text(scrollable_frame, height=3, width=30)
            else:
                fields[field] = ttk.Entry(scrollable_frame)
            fields[field].grid(row=row, column=1, padx=5, pady=5, sticky="ew")
            row += 1
        
        def validate_and_save():
            # Validation
            if not fields['name'].get().strip():
                messagebox.showerror("Error", "Name is required")
                return
            
            try:
                quantity = int(fields['quantity'].get())
                if quantity < 0:
                    raise ValueError("Quantity must be positive")
            except ValueError as e:
                messagebox.showerror("Error", f"Invalid quantity: {str(e)}")
                return
            
            # Generate item ID
            if item_type == "chemical":
                prefix = "CHE"
            elif item_type == "equipment":
                prefix = "EQ"
            elif item_type == "consumable":
                prefix = "CON"
            else:
                prefix = "OT"
            
            self.cursor.execute("SELECT COUNT(*) FROM items WHERE item_type = ?", (item_type,))
            count = self.cursor.fetchone()[0]
            item_id = f"{prefix}{count + 1:04d}"
            
            # Save to database
            try:
                self.cursor.execute("""
                    INSERT INTO items (
                        id, name, name_cn, item_type, category, location,
                        quantity, unit, manufacturer, model_number,
                        serial_number, purchase_date, warranty_until,
                        maintenance_contact, last_calibration, next_calibration,
                        safety_classification, notes, last_updated
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    item_id,
                    fields["name"].get().strip(),
                    fields["name_cn"].get().strip(),
                    item_type,
                    fields["category"].get().strip(),
                    fields["location"].get().strip(),
                    quantity,
                    fields["unit"].get().strip(),
                    fields["manufacturer"].get().strip(),
                    fields["model_number"].get().strip(),
                    fields["serial_number"].get().strip(),
                    fields["purchase_date"].get().strip(),
                    fields["warranty_until"].get().strip(),
                    fields["maintenance_contact"].get().strip(),
                    fields["last_calibration"].get().strip(),
                    fields["next_calibration"].get().strip(),
                    fields["safety_classification"].get().strip(),
                    fields["notes"].get("1.0", tk.END).strip(),
                    datetime.now()
                ))
                
                self.conn.commit()
                tree = self.equipment_tree if item_type == "equipment" else self.chemicals_tree if item_type == "chemical" else self.consumables_tree if item_type == "consumable" else self.other_tree
                self.refresh_inventory(item_type, tree)
                add_window.destroy()
                messagebox.showinfo("Success", "Item added successfully! 物品添加成功！")
            
            except Exception as e:
                self.conn.rollback()
                messagebox.showerror("Error", f"Failed to add item: {str(e)}")
        
        # Save button
        ttk.Button(scrollable_frame, text="Save 保存", command=validate_and_save).grid(
            row=row, column=0, columnspan=2, pady=20)
        
        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def edit_item(self, item_type):
        """Edit selected item with improved validation and error handling"""
        tree = self.equipment_tree if item_type == "equipment" else self.chemicals_tree if item_type == "chemical" else self.consumables_tree if item_type == "consumable" else self.other_tree
        selected = tree.selection()
        
        if not selected:
            messagebox.showwarning("Warning", "Please select an item to edit 请选择要编辑的物品")
            return
        
        try:
            item_id = tree.item(selected[0])['values'][0]
            
            self.cursor.execute("""
                SELECT name, name_cn, category, location, quantity, unit,
                    manufacturer, model_number, serial_number, purchase_date,
                    warranty_until, maintenance_contact, last_calibration,
                    next_calibration, safety_classification, notes
                FROM items
                WHERE id = ?
            """, (item_id,))
            
            item_data = self.cursor.fetchone()
            if not item_data:
                messagebox.showerror("Error", "Item not found in database")
                return
            
            edit_window = tk.Toplevel(self.root)
            edit_window.title("Edit Item 编辑物品")
            edit_window.geometry("400x800")  # Increased height to accommodate more fields
            edit_window.transient(self.root)
            edit_window.grab_set()
            
            # Create scrollable frame
            canvas = tk.Canvas(edit_window)
            scrollbar = ttk.Scrollbar(edit_window, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            # Fields
            fields = {}
            field_configs = [
                ("name", "Name 名称 *", item_data[0]),
                ("name_cn", "Chinese Name 中文名称", item_data[1]),
                ("category", "Category 类别", item_data[2]),
                ("location", "Location 位置", item_data[3]),
                ("quantity", "Quantity 数量 *", str(item_data[4])),
                ("unit", "Unit 单位", item_data[5]),
                ("manufacturer", "Manufacturer 制造商", item_data[6]),
                ("model_number", "Model Number 型号", item_data[7]),
                ("serial_number", "Serial Number 序列号", item_data[8]),
                ("purchase_date", "Purchase Date 购买日期 (YYYY-MM-DD)", item_data[9]),
                ("warranty_until", "Warranty Until 保修至 (YYYY-MM-DD)", item_data[10]),
                ("maintenance_contact", "Maintenance Contact 维护联系人", item_data[11]),
                ("last_calibration", "Last Calibration 上次校准 (YYYY-MM-DD)", item_data[12]),
                ("next_calibration", "Next Calibration 下次校准 (YYYY-MM-DD)", item_data[13]),
                ("safety_classification", "Safety Classification 安全分类", item_data[14]),
                ("notes", "Notes 备注", item_data[15])
            ]
            
            row = 0
            for field, label, value in field_configs:
                ttk.Label(scrollable_frame, text=label).grid(row=row, column=0, padx=5, pady=5, sticky="w")
                if field == "notes":
                    fields[field] = tk.Text(scrollable_frame, height=3, width=30)
                    fields[field].insert("1.0", value if value else "")
                else:
                    fields[field] = ttk.Entry(scrollable_frame)
                    fields[field].insert(0, value if value else "")
                fields[field].grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                row += 1
            
            def validate_and_save():
                # Validation
                if not fields['name'].get().strip():
                    messagebox.showerror("Error", "Name is required")
                    return
                
                try:
                    quantity = int(fields['quantity'].get())
                    if quantity < 0:
                        raise ValueError("Quantity must be positive")
                except ValueError as e:
                    messagebox.showerror("Error", f"Invalid quantity: {str(e)}")
                    return
                
                # Save to database
                try:
                    self.cursor.execute("""
                        UPDATE items
                        SET name = ?, name_cn = ?, category = ?, location = ?,
                            quantity = ?, unit = ?, manufacturer = ?, model_number = ?,
                            serial_number = ?, purchase_date = ?, warranty_until = ?,
                            maintenance_contact = ?, last_calibration = ?,
                            next_calibration = ?, safety_classification = ?,
                            notes = ?, last_updated = ?
                        WHERE id = ?
                    """, (
                        fields["name"].get().strip(),
                        fields["name_cn"].get().strip(),
                        fields["category"].get().strip(),
                        fields["location"].get().strip(),
                        quantity,
                        fields["unit"].get().strip(),
                        fields["manufacturer"].get().strip(),
                        fields["model_number"].get().strip(),
                        fields["serial_number"].get().strip(),
                        fields["purchase_date"].get().strip(),
                        fields["warranty_until"].get().strip(),
                        fields["maintenance_contact"].get().strip(),
                        fields["last_calibration"].get().strip(),
                        fields["next_calibration"].get().strip(),
                        fields["safety_classification"].get().strip(),
                        fields["notes"].get("1.0", tk.END).strip(),
                        datetime.now(),
                        item_id
                    ))
                    
                    self.conn.commit()
                    self.refresh_inventory(item_type, tree)
                    edit_window.destroy()
                    messagebox.showinfo("Success", "Item updated successfully! 物品更新成功！")
                
                except Exception as e:
                    self.conn.rollback()
                    messagebox.showerror("Error", f"Failed to update item: {str(e)}")
            
            # Save button
            ttk.Button(scrollable_frame, text="Save Changes 保存更改", 
                    command=validate_and_save).grid(row=row, column=0, columnspan=2, pady=20)
            
            # Pack the canvas and scrollbar
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to edit item: {str(e)}")
            
    def delete_item(self, item_type):
        """Delete selected item with improved error handling"""
        tree = self.equipment_tree if item_type == "equipment" else self.chemicals_tree if item_type == "chemical" else self.consumables_tree if item_type == "consumable" else self.other_tree
        selected = tree.selection()
        
        if not selected:
            messagebox.showwarning("Warning", "Please select an item to delete 请选择要删除的物品")
            return
        
        try:
            item_id = tree.item(selected[0])['values'][0]
            item_name = tree.item(selected[0])['values'][1]
            
            if messagebox.askyesno("Confirm Delete", 
                f"Are you sure you want to delete '{item_name}'?\n确定要删除 '{item_name}' 吗？"):
                
                self.cursor.execute("DELETE FROM items WHERE id = ?", (item_id,))
                self.conn.commit()
                
                self.refresh_inventory(item_type, tree)
                messagebox.showinfo("Success", "Item deleted successfully! 物品删除成功！")
        
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Error", f"Failed to delete item: {str(e)}")

    def search_items(self, item_type, tree, search_var):
        """Search items with improved search logic"""
        search_term = search_var.get().lower().strip()
        
        for item in tree.get_children():
            tree.delete(item)
        
        try:
            if search_term:
                self.cursor.execute("""
                    SELECT id, name, name_cn, category, location, quantity, unit
                    FROM items
                    WHERE item_type = ? AND (
                        LOWER(name) LIKE ? OR
                        LOWER(COALESCE(name_cn, '')) LIKE ? OR
                        LOWER(COALESCE(category, '')) LIKE ? OR
                        LOWER(COALESCE(location, '')) LIKE ? OR
                        LOWER(COALESCE(unit, '')) LIKE ?
                    )
                """, (
                    item_type,
                    f"%{search_term}%",
                    f"%{search_term}%",
                    f"%{search_term}%",
                    f"%{search_term}%",
                    f"%{search_term}%"
                ))
            else:
                self.cursor.execute("""
                    SELECT id, name, name_cn, category, location, quantity, unit
                    FROM items
                    WHERE item_type = ?
                """, (item_type,))
            
            for row in self.cursor.fetchall():
                tree.insert("", "end", values=row)
            
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {str(e)}")

    def refresh_inventory(self, item_type, tree):
        """Refresh inventory display with error handling"""
        for item in tree.get_children():
            tree.delete(item)
        
        try:
            # Fetch data from the database
            self.cursor.execute("""
                SELECT id, name, name_cn, category, location, quantity, unit,
                    manufacturer, model_number, serial_number, purchase_date,
                    warranty_until, maintenance_contact, last_calibration,
                    next_calibration, safety_classification
                FROM items
                WHERE item_type = ?
                ORDER BY name
            """, (item_type,))
            
            # Fetch all rows
            rows = self.cursor.fetchall()
            
            # Debug: Print the number of columns in the Treeview
            print(f"Number of Treeview columns: {len(tree['columns'])}")
            
            # Debug: Print the number of values returned by the SQL query
            if rows:
                print(f"Number of values returned by SQL query: {len(rows[0])}")
            else:
                print("No rows returned by SQL query.")
            
            # Check if the number of columns in the Treeview matches the number of values
            columns = tree["columns"]
            if rows and len(columns) != len(rows[0]):
                raise ValueError(f"Number of Treeview columns ({len(columns)}) does not match the number of values ({len(rows[0])})")
            
            # Insert rows into the Treeview
            for row in rows:
                tree.insert("", "end", values=row)
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh inventory: {str(e)}")

    def refresh_usage_log(self):
        """Refresh usage log display with improved error handling"""
        for item in self.usage_tree.get_children():
            self.usage_tree.delete(item)
        
        try:
            self.cursor.execute("""
                SELECT u.id, i.name, u.user, u.user_department,
                       u.quantity_changed, u.timestamp, u.purpose,
                       CASE WHEN u.return_time IS NULL THEN 'Active' 
                            ELSE 'Returned' END as status
                FROM usage_log u
                JOIN items i ON u.item_id = i.id
                ORDER BY u.timestamp DESC
            """)
            
            for row in self.cursor.fetchall():
                # Format timestamp
                row = list(row)
                if isinstance(row[5], str):
                    timestamp = datetime.fromisoformat(row[5])
                else:
                    timestamp = row[5]
                row[5] = timestamp.strftime("%Y-%m-%d %H:%M")
                
                self.usage_tree.insert("", "end", values=row)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh usage log: {str(e)}")

    def generate_report(self):
        """Generate a comprehensive inventory report with usage history"""
        try:
            report_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                initialdir=self.dirs['exports'],
                title="Save Inventory Report",
                filetypes=[("PDF files", "*.pdf")]
            )
            
            if not report_path:
                return
            
            # Create PDF document
            doc = SimpleDocTemplate(report_path, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []
            
            # Title
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Title'],
                fontSize=24,
                spaceAfter=30
            )
            elements.append(Paragraph("DNA Virology Laboratory Inventory Report", title_style))
            
            # Date
            elements.append(Paragraph(
                f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                styles['Normal']
            ))
            elements.append(Paragraph("<br/><br/>", styles['Normal']))
            
            # Inventory Summary
            elements.append(Paragraph("Inventory Summary", styles['Heading1']))
            
            self.cursor.execute("""
                SELECT item_type,
                       COUNT(*) as total_items,
                       SUM(CASE WHEN quantity <= 0 THEN 1 ELSE 0 END) as out_of_stock,
                       SUM(quantity) as total_quantity
                FROM items
                GROUP BY item_type
            """)
            
            summary_data = [['Type', 'Total Items', 'Out of Stock', 'Total Quantity']]
            summary_data.extend(self.cursor.fetchall())
            
            summary_table = Table(summary_data)
            summary_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BOX', (0, 0), (-1, -1), 2, colors.black)
            ]))
            elements.append(summary_table)
            elements.append(Paragraph("<br/><br/>", styles['Normal']))
            
            # Equipment List
            elements.append(Paragraph("Equipment Inventory", styles['Heading1']))
            
            self.cursor.execute("""
                SELECT name, name_cn, category, location, quantity, unit, manufacturer
                FROM items
                WHERE item_type = 'equipment'
                ORDER BY category, name
            """)
            
            equipment_data = [['Name', 'Chinese Name', 'Category', 'Location', 'Quantity', 'Unit', 'Manufacturer']]
            equipment_data.extend(self.cursor.fetchall())
            
            equipment_table = Table(equipment_data)
            equipment_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BOX', (0, 0), (-1, -1), 2, colors.black)
            ]))
            elements.append(equipment_table)
            elements.append(Paragraph("<br/><br/>", styles['Normal']))
            
            # Chemicals List
            elements.append(Paragraph("Chemicals Inventory", styles['Heading1']))
            
            self.cursor.execute("""
                SELECT name, name_cn, category, location, quantity, unit, manufacturer
                FROM items
                WHERE item_type = 'chemical'
                ORDER BY category, name
            """)
            
            chemicals_data = [['Name', 'Chinese Name', 'Category', 'Location', 'Quantity', 'Unit', 'Manufacturer']]
            chemicals_data.extend(self.cursor.fetchall())
            
            chemicals_table = Table(chemicals_data)
            chemicals_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BOX', (0, 0), (-1, -1), 2, colors.black)
            ]))
            elements.append(chemicals_table)
            
            # Consumables List
            elements.append(Paragraph("Consumables Inventory", styles['Heading1']))
            
            self.cursor.execute("""
                SELECT name, name_cn, category, location, quantity, unit, manufacturer
                FROM items
                WHERE item_type = 'consumable'
                ORDER BY category, name
            """)
            
            consumables_data = [['Name', 'Chinese Name', 'Category', 'Location', 'Quantity', 'Unit', 'Manufacturer']]
            consumables_data.extend(self.cursor.fetchall())
            
            consumables_table = Table(consumables_data)
            consumables_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BOX', (0, 0), (-1, -1), 2, colors.black)
            ]))
            elements.append(consumables_table)
            
            # Other Items List
            elements.append(Paragraph("Other Inventory", styles['Heading1']))
            
            self.cursor.execute("""
                SELECT name, name_cn, category, location, quantity, unit, manufacturer
                FROM items
                WHERE item_type = 'other'
                ORDER BY category, name
            """)
            
            other_data = [['Name', 'Chinese Name', 'Category', 'Location', 'Quantity', 'Unit', 'Manufacturer']]
            other_data.extend(self.cursor.fetchall())
            
            other_table = Table(other_data)
            other_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BOX', (0, 0), (-1, -1), 2, colors.black)
            ]))
            elements.append(other_table)
            
            # Build PDF
            doc.build(elements)
            messagebox.showinfo("Success", f"Report generated successfully!\n报告已生成: {report_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {str(e)}")

    def export_inventory(self):
        """Export inventory data to CSV"""
        try:
            export_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                initialdir=self.dirs['exports'],
                title="Export Inventory",
                filetypes=[("CSV files", "*.csv")]
            )
            
            if not export_path:
                return
            
            self.cursor.execute("""
                SELECT id, name, name_cn, category, location, quantity, unit,
                       manufacturer, model_number, serial_number, purchase_date,
                       warranty_until, maintenance_contact, last_calibration,
                       next_calibration, safety_classification, last_updated, notes
                FROM items
                ORDER BY item_type, category, name
            """)
            
            with open(export_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                # Write header
                writer.writerow([
                    'ID', 'Name', 'Chinese Name', 'Category', 'Location',
                    'Quantity', 'Unit', 'Manufacturer', 'Model Number', 
                    'Serial Number', 'Purchase Date', 'Warranty Until', 
                    'Maintenance Contact', 'Last Calibration', 'Next calibration',
                    'Safety Classification', 'Last Updated', 'Notes'
                ])
                # Write data
                writer.writerows(self.cursor.fetchall())
            
            messagebox.showinfo("Success", f"Inventory exported successfully!\n库存已导出: {export_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export inventory: {str(e)}")

    def export_usage_log(self):
        """Export usage log data to CSV"""
        try:
            export_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                initialdir=self.dirs['exports'],
                title="Export Usage Log",
                filetypes=[("CSV files", "*.csv")]
            )
            
            if not export_path:
                return
            
            self.cursor.execute("""
                SELECT u.id, i.name, i.name_cn, u.user, u.user_department,
                       u.quantity_changed, u.timestamp, u.purpose,
                       u.supervisor_approval, u.return_time, u.notes
                FROM usage_log u
                JOIN items i ON u.item_id = i.id
                ORDER BY u.timestamp DESC
            """)
            
            with open(export_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                # Write header
                writer.writerow([
                    'ID', 'Item Name', 'Item Chinese Name', 'User', 'Department',
                    'Quantity Changed', 'Timestamp', 'Purpose',
                    'Supervisor Approval', 'Return Time', 'Notes'
                ])
                # Write data
                writer.writerows(self.cursor.fetchall())
            
            messagebox.showinfo("Success", f"Usage log exported successfully!\n使用记录已导出: {export_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export usage log: {str(e)}")

    def backup_database(self):
        """Create a backup of the database"""
        try:
            backup_dir = self.dirs['data'] / "backups"
            backup_dir.mkdir(exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = backup_dir / f"lab_inventory_backup_{timestamp}.db"
            
            # Close current connection
            self.conn.close()
            
            # Copy database file
            shutil.copy2(
                self.dirs['data'] / "lab_inventory.db",
                backup_path
            )
            
            # Reopen connection
            self.conn = sqlite3.connect(
                str(self.dirs['data'] / "lab_inventory.db"),
                detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES
            )
            self.cursor = self.conn.cursor()
            
            messagebox.showinfo(
                "Success",
                f"Database backed up successfully!\n数据库已备份: {backup_path}"
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to backup database: {str(e)}")
            # Ensure connection is reopened even if backup fails
            self.conn = sqlite3.connect(
                str(self.dirs['data'] / "lab_inventory.db"),
                detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES
            )
            self.cursor = self.conn.cursor()

    def show_about(self):
        """Show about dialog"""
        about_text = """
        DNA Virology Lab Management System
        病毒实验室管理系统
        
        Version 1.0
        
        Developed for DNA Virology Lab-ICGEB China RRC by Kavindu Munugoda
        Kavindu Munugoda为DNA病毒学实验室-ICGEB China RRC开发了这个
        
        © 2024 All Rights Reserved
        """
        messagebox.showinfo("About 关于", about_text)

    def ai_predict_inventory_needs(self):
        """AI Predict Inventory Needs"""
        try:
            self.cursor.execute("SELECT name, quantity, item_type FROM items")
            inventory_data = [{'name': row[0], 'quantity': row[1], 'item_type': row[2]} for row in self.cursor.fetchall()]
            low_stock_items = self.ai_assistant.predict_inventory_needs(inventory_data)
            
            if low_stock_items:
                message = "Low stock items:\n"
                for item in low_stock_items:
                    message += f"{item['name']} - {item['quantity']} left\n"
                messagebox.showinfo("AI Prediction", message)
            else:
                messagebox.showinfo("AI Prediction", "No low stock items detected.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to predict inventory needs: {str(e)}")

    def __del__(self):
        """Cleanup database connection"""
        try:
            if hasattr(self, 'conn'):
                self.conn.close()
        except Exception:
            pass

def main():
    root = tk.Tk()
    app = LabInventorySystem(root)
    root.mainloop()

if __name__ == "__main__":
    main()