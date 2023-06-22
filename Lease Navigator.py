import sqlite3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel, QTreeWidget, QTreeWidgetItem, QDateEdit, QMessageBox
from PyQt5.QtCore import Qt, QDate
import sys

class PropertyManagerApp(QWidget):
    def __init__(self):
        super().__init__()

        self.db_name = 'property_manager.db'
        self.init_db()

        self.init_ui()
        self.load_data()

    def init_db(self):
        conn = sqlite3.connect(self.db_name)
        c = conn.cursor()

        # Create table if not exists
        c.execute('''
            CREATE TABLE IF NOT EXISTS buildings
            (id INTEGER PRIMARY KEY, name TEXT)
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS apartments
            (id INTEGER PRIMARY KEY, building_id INTEGER, name TEXT, email TEXT, lease_start TEXT, lease_end TEXT)
        ''')

        conn.commit()
        conn.close()

    def init_ui(self):
        self.layout = QVBoxLayout()

        self.tree = QTreeWidget()
        self.tree.setColumnCount(5)  # Increase column count to 5
        self.tree.setHeaderLabels(["Building Name", "Apartment Number", "First Name", "Last Name", "Email", "Lease Start", "Lease End"])  # Update header labels
        self.tree.itemClicked.connect(self.handle_item_clicked)
        self.tree.itemChanged.connect(self.handle_item_changed)

        self.building_name_input = QLineEdit()
        self.add_building_button = QPushButton("Add Building")
        self.add_building_button.clicked.connect(self.add_building)
        self.apartment_name_input = QLineEdit()
        self.first_name_input = QLineEdit()
        self.last_name_input = QLineEdit()
        self.tenant_email_input = QLineEdit()
        self.lease_start_input = QDateEdit(QDate.currentDate())
        self.lease_start_input.setDisplayFormat("MM/dd/yyyy")
        self.lease_end_input = QDateEdit(QDate.currentDate())
        self.lease_end_input.setDisplayFormat("MM/dd/yyyy")
        self.add_apartment_button = QPushButton("Add Apartment")
        self.add_apartment_button.clicked.connect(self.add_apartment)

        self.delete_button = QPushButton('Delete Selected')
        self.delete_button.clicked.connect(self.delete_selected)

        self.send_reminder_button = QPushButton('Send Reminder')
        self.send_reminder_button.clicked.connect(self.send_reminder)

        self.layout.addWidget(self.tree)

        building_layout = QHBoxLayout()
        building_layout.addWidget(self.building_name_input)
        building_layout.addWidget(self.add_building_button)

        apartment_layout = QHBoxLayout()
        apartment_layout.addWidget(QLabel("Apartment Number:"))
        apartment_layout.addWidget(self.apartment_name_input)
        apartment_layout.addWidget(QLabel("First Name:"))
        apartment_layout.addWidget(self.first_name_input)
        apartment_layout.addWidget(QLabel("Last Name:"))
        apartment_layout.addWidget(self.last_name_input)
        apartment_layout.addWidget(QLabel("Tenant Email:"))
        apartment_layout.addWidget(self.tenant_email_input)
        apartment_layout.addWidget(self.lease_start_input)
        apartment_layout.addWidget(self.lease_end_input)
        apartment_layout.addWidget(self.add_apartment_button)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.delete_button)
        button_layout.addWidget(self.send_reminder_button)

        self.layout.addLayout(building_layout)
        self.layout.addLayout(apartment_layout)
        self.layout.addLayout(button_layout)

        self.setLayout(self.layout)
        
        self.tree.setColumnWidth(0, 150)  # Building Name
        self.tree.setColumnWidth(1, 150)  # Apartment Number
        self.tree.setColumnWidth(2, 150)  # First Name
        self.tree.setColumnWidth(3, 150)  # Last Name
        self.tree.setColumnWidth(4, 200)  # Email
        self.tree.setColumnWidth(5, 150)  # Lease Start
        self.tree.setColumnWidth(6, 150)  # Lease End

    def add_building(self):
        building_name = self.building_name_input.text()
        building_item = QTreeWidgetItem(self.tree)
        building_item.setFlags(building_item.flags() | Qt.ItemIsEditable)
        building_item.setText(0, building_name)
        self.building_name_input.clear()

    def handle_item_clicked(self, item, column):
        if item.parent() is None:  # the item is a building
            self.add_apartment_button.setEnabled(True)
        else:  # the item is an apartment
            self.add_apartment_button.setEnabled(False)

    def add_apartment(self):
        building_item = self.tree.currentItem()

        if building_item is None or building_item.parent() is not None:
            QMessageBox.warning(self, "Error", "Please select a building to add an apartment.")
            return

        apartment_number = self.apartment_name_input.text()
        first_name = self.first_name_input.text()
        last_name = self.last_name_input.text()
        email = self.tenant_email_input.text()
        lease_start = self.lease_start_input.date().toString("MM/dd/yyyy")
        lease_end = self.lease_end_input.date().toString("MM/dd/yyyy")

        apartment_item = QTreeWidgetItem()
        apartment_item.setFlags(apartment_item.flags() | Qt.ItemIsEditable)
        apartment_item.setText(0, building_item.text(0))  # Set Building Name
        apartment_item.setText(1, apartment_number)
        apartment_item.setText(2, first_name)
        apartment_item.setText(3, last_name)
        apartment_item.setText(4, email)
        apartment_item.setText(5, lease_start)
        apartment_item.setText(6, lease_end)

        building_item.addChild(apartment_item)

        self.apartment_name_input.clear()
        self.first_name_input.clear()
        self.last_name_input.clear()
        self.tenant_email_input.clear()
        self.lease_start_input.setDate(QDate.currentDate())
        self.lease_end_input.setDate(QDate.currentDate())

    def handle_item_changed(self, item, column):
        if item.parent() is None:  # the item is a building
            return
        if column == 1:  # Apartment number field
            text = item.text(1)
            if not text:
                QMessageBox.critical(self, "Error", "Apartment number cannot be empty.")
                item.setText(1, '')

    def delete_selected(self):
        current_item = self.tree.currentItem()
        if current_item:
            parent_item = current_item.parent()
            if parent_item is None:  # the item is a building
                index = self.tree.indexOfTopLevelItem(current_item)
                self.tree.takeTopLevelItem(index)
            else:  # the item is an apartment
                index = parent_item.indexOfChild(current_item)
                parent_item.takeChild(index)

    def save_data(self):
        conn = sqlite3.connect(self.db_name)
        c = conn.cursor()

        # Delete old data
        c.execute('DELETE FROM apartments')
        c.execute('DELETE FROM buildings')

        # Save new data
        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            building_item = root.child(i)
            building_name = building_item.text(0)

            c.execute('INSERT INTO buildings (name) VALUES (?)', (building_name,))
            building_id = c.lastrowid

            for j in range(building_item.childCount()):
                apartment_item = building_item.child(j)
                apartment_number = apartment_item.text(1)
                first_name = apartment_item.text(2)
                last_name = apartment_item.text(3)
                email = apartment_item.text(4)
                lease_start = apartment_item.text(5)
                lease_end = apartment_item.text(6)

                c.execute('INSERT INTO apartments (building_id, name, email, lease_start, lease_end) VALUES (?, ?, ?, ?, ?)',
                          (building_id, f"{apartment_number} - {first_name} {last_name}", email, lease_start, lease_end))

        conn.commit()
        conn.close()

    def load_data(self):
        conn = sqlite3.connect(self.db_name)
        c = conn.cursor()

        # Load data
        c.execute('SELECT * FROM buildings')
        for building_row in c.fetchall():
            building_id, building_name = building_row
            building_item = QTreeWidgetItem(self.tree)
            building_item.setFlags(building_item.flags() | Qt.ItemIsEditable)
            building_item.setText(0, building_name)

            c.execute('SELECT * FROM apartments WHERE building_id = ?', (building_id,))
            for apartment_row in c.fetchall():
                apartment_id, _, apartment_name, email, lease_start, lease_end = apartment_row
                apartment_item = QTreeWidgetItem()
                apartment_item.setFlags(apartment_item.flags() | Qt.ItemIsEditable)
                apartment_item.setText(0, building_name)  # Set Building Name
                apartment_item.setText(1, apartment_name.split(' - ')[0])  # Set Apartment Number
                apartment_item.setText(2, apartment_name.split(' - ')[1])  # Set First Name
                apartment_item.setText(3, "")  # Leave Last Name field empty initially
                apartment_item.setText(4, email)
                apartment_item.setText(5, lease_start)
                apartment_item.setText(6, lease_end)

                building_item.addChild(apartment_item)

        conn.close()

    def send_reminder(self):
        current_item = self.tree.currentItem()
        if current_item is None:
            QMessageBox.warning(self, "Error", "Please select a tenant to send a reminder to.")
            return

        parent_item = current_item.parent()
        if parent_item is None:  # the item is a building
            QMessageBox.warning(self, "Error", "Please select a tenant, not a building.")
            return

        tenant_info = current_item.text(1), current_item.text(2)  # Apartment Number and First Name
        if "" in tenant_info:
            QMessageBox.critical(self, "Error", "Please enter the Apartment Number and First Name for the tenant.")
            return

        apartment_number, tenant_name = tenant_info
        tenant_email = current_item.text(4)
        lease_end = current_item.text(6)

        msg = MIMEMultipart()
        msg['From'] = "youremail@outlook"
        msg['To'] = tenant_email
        msg['Subject'] = "Lease Expiry Reminder"
        body = f"Dear {tenant_name},\n\nYour lease for apartment {apartment_number} is set to expire on {lease_end}. If you're planning to renew your lease, please see the attached file.\n\nBest,\n\nOceana Management"
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login("youremail@outlook, "your_password")
        text = msg.as_string()
        server.sendmail("youremail@outlook", tenant_email, text)
        server.quit()

        QMessageBox.information(self, "Sent", f"Reminder sent to {tenant_name}.")


app = QApplication(sys.argv)
window = PropertyManagerApp()
window.show()

# Save data to database before application exits
app.aboutToQuit.connect(window.save_data)

sys.exit(app.exec_())
