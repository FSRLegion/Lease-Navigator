import sendgrid
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
import sqlite3
import smtplib
import xlsxwriter
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PyQt5.QtWidgets import (
    QApplication, QWidget, QGridLayout, QComboBox, QGroupBox, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
    QLabel, QTreeWidget, QTreeWidgetItem, QDateEdit, QMessageBox, QSplitter, QTextEdit, QFileDialog,
    QListWidget, QListWidgetItem, QAbstractItemView, QDialog, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QDate, QSettings, QStandardPaths
from PyQt5.QtGui import QFont, QIcon

import sys
import os

class LoginDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.email = None
        self.api_key = None

        self.setWindowTitle("Login")
        layout = QVBoxLayout()

        email_label = QLabel("Email:")
        self.email_input = QLineEdit()
        layout.addWidget(email_label)
        layout.addWidget(self.email_input)

        api_key_label = QLabel("API Key:")
        self.api_key_input = QLineEdit()
        layout.addWidget(api_key_label)
        layout.addWidget(self.api_key_input)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def accept(self):
        self.email = self.email_input.text().strip()
        self.api_key = self.api_key_input.text().strip()
        super().accept()


class PropertyManagerApp(QWidget):
    def __init__(self, email=None, api_key=None):
        super().__init__()

        self.db_name = 'property_manager.db'
        self.settings = QSettings("YourOrganization", "YourApplication")
        self.email = email
        self.api_key = api_key
        self.init_db()
        self.resize(1440, 800)

        self.init_ui()
        self.load_data()

        self.apply_styles()
        self.attached_files = []

        try:
            self.restoreGeometry(self.settings.value("geometry"))
        except TypeError:
            pass

        self.apply_styles()
        self.attached_files = []

    def apply_styles(self):
        self.setStyleSheet("""
            QWidget {
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 18px;
                color: #333;
            }
            QTreeWidget {
                background-color: #fff;
                border: none;
            }
            QTreeWidget::item {
                padding: 10px;
            }
            QTreeWidget::item:selected {
                background-color: #e6f3ff;
                color: #333;
            }
            QTreeWidget::item:hover {
                background-color: #f5f5f5;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                color: #333;
                padding: 8px;
            }
            QPushButton {
                background-color: #1f707f;
                border: none;
                color: white;
                padding: 8px 16px;
                text-align: center;
                text-decoration: none;
                font-size: 16px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #3c9d9b;
            }
            QLineEdit {
                padding: 2px;
                border: 1px solid #ccc;
                border-radius: 4px;
                font-size: 18px;
                background-color: #f5f5f5;
            }
            QLineEdit:focus {
                border: 1px solid #4CAF50;
            }
            QGroupBox {
                border: 1px solid #ccc;
                border-radius: 4px;
                background-color: #fff;
                padding: 10px;
                margin-bottom: 10px;
            }
            QLabel {
                font-weight: bold;
            }
            /* Custom Styles */
            QWidget#chartWidget {
                background-color: #f5f5f5;
            }
            QWidget#inputWidget {
                background-color: #f5f5f5;
                padding: 10px;
            }
            QPushButton#deleteButton {
                background-color: #c62828;
            }
            QPushButton#deleteButton:hover {
                background-color: #ef5350;
            }
            QPushButton#sendReminderButton {
                background-color: #2e7d32;
            }
            QPushButton#sendReminderButton:hover {
                background-color: #66bb6a;
            }
            QPushButton#attachFilesButton {
                background-color: #1f707f;
            }
            QPushButton#attachFilesButton:hover {
                background-color: #3c9d9b;
            }
            QLabel#attachedFilesLabel {
                color: #333;
                font-weight: bold;
                margin-left: 10px;
            }
        """)

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

    def save_changes(self):
        self.save_data()
        QMessageBox.information(self, "Saved", "Changes have been saved.")

    def init_ui(self):
        self.layout = QVBoxLayout()

        # Splitter for the chart and input area
        splitter = QSplitter()

        # Create a vertical layout for the chart area
        chart_layout = QVBoxLayout()

        # Create the chart widget
        chart_widget = QWidget()
        chart_widget.setObjectName("chartWidget")

        # Tree widget
        self.tree = QTreeWidget()
        self.tree.setColumnCount(6)
        header = self.tree.header()
        header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        header.setStyleSheet(
            "QHeaderView::section { background-color: #f5f5f5; color: #333; padding: 8px; border: none; }")
        self.tree.setHeaderLabels(
            ["Building Name", "Apartment", "Name", "Email", "Lease Start", "Lease End"])
        self.tree.itemClicked.connect(self.handle_item_clicked)
        self.tree.itemChanged.connect(self.handle_item_changed)

        self.tree.setStyleSheet("""
            QTreeWidget::item {
                padding: 10px;
                border-right: 1px solid #ccc;  /* Add vertical grid line */
            }
        """)

        # Set column widths
        self.tree.setColumnWidth(0, 250)  # Building Name
        self.tree.setColumnWidth(1, 120)  # Apartment Number
        self.tree.setColumnWidth(2, 120)  # Name
        self.tree.setColumnWidth(3, 270)  # Email
        self.tree.setColumnWidth(4, 140)  # Lease Start
        self.tree.setColumnWidth(5, 200)  # Lease End

        self.tree.setSortingEnabled(False)
        self.tree.sortByColumn(1, Qt.AscendingOrder)
        items = self.tree.findItems("", Qt.MatchContains, 1)
        sorted_items = sorted(items,
                              key=lambda item: [int(num) if num.isdigit() else num for num in re.split(r'(\d+)',
                                                                                                      item.text(1))])
        sorted_items.sort(key=lambda item: int(item.text(1).split('-')[0]) if item.text(1).split('-')[
                                                                                  0].isdigit() else float('inf'))
        for index, item in enumerate(sorted_items):
            self.tree.takeTopLevelItem(self.tree.indexOfTopLevelItem(item))
            self.tree.insertTopLevelItem(index, item)
        self.tree.setSortingEnabled(True)

        self.download_button = QPushButton("Download Chart")
        self.download_button.clicked.connect(self.download_chart)
        chart_layout.addWidget(self.download_button)

        chart_layout.addWidget(self.tree)

        # Add the chart widget to the layout
        chart_widget.setLayout(chart_layout)

        # Add the chart widget to the splitter
        splitter.addWidget(chart_widget)

        # Create a vertical layout for the input area
        input_layout = QVBoxLayout()

        # Create the input widget
        input_widget = QWidget()
        input_widget.setObjectName("inputWidget")

        # Building input
        self.building_name_input = QLineEdit()
        self.add_building_button = QPushButton("Add Building")
        self.add_building_button.clicked.connect(self.add_building)
        building_layout = QVBoxLayout()
        building_layout.addWidget(QLabel("Building Name:"))
        building_layout.addWidget(self.building_name_input)
        building_layout.addWidget(self.add_building_button)
        building_group = QGroupBox("Add New Building")
        building_group.setLayout(building_layout)

        # Apartment input
        self.apartment_name_input = QLineEdit()
        self.name_input = QLineEdit()
        self.tenant_email_input = QLineEdit()
        self.lease_start_input = QDateEdit(QDate.currentDate())
        self.lease_start_input.setDisplayFormat("MM/dd/yyyy")
        self.lease_end_input = QDateEdit(QDate.currentDate())
        self.lease_end_input.setDisplayFormat("MM/dd/yyyy")
        self.add_apartment_button = QPushButton("Add Apartment")
        self.add_apartment_button.clicked.connect(self.add_apartment)

        apartment_layout = QGridLayout()
        apartment_layout.addWidget(QLabel("Apartment Number:"), 0, 0)
        apartment_layout.addWidget(self.apartment_name_input, 0, 1)
        apartment_layout.addWidget(QLabel("Name:"), 0, 2)
        apartment_layout.addWidget(self.name_input, 0, 3)
        apartment_layout.addWidget(QLabel("Tenant Email:"), 1, 0)
        apartment_layout.addWidget(self.tenant_email_input, 1, 1, 1, 3)
        apartment_layout.addWidget(QLabel("Lease Start:"), 2, 0)
        apartment_layout.addWidget(self.lease_start_input, 2, 1)
        apartment_layout.addWidget(QLabel("Lease End:"), 2, 2)
        apartment_layout.addWidget(self.lease_end_input, 2, 3)
        apartment_layout.addWidget(self.add_apartment_button, 3, 0, 1, 4)

        apartment_group = QGroupBox("Add New Apartment")
        apartment_group.setLayout(apartment_layout)
        apartment_group.setMaximumHeight(300)  # Adjust the height as needed

        # Email input
        self.email_text_edit = QTextEdit()
        default_message = "Dear {Name},\n\nYour lease is set to expire on {Lease End}. Please contact us if you wish to renew your lease.\n\nSincerely,\nProperty Management"
        self.email_text_edit.setPlainText(default_message)
        self.send_reminder_button = QPushButton("Send Reminder")
        self.send_reminder_button.clicked.connect(self.send_reminder)

        email_layout = QVBoxLayout()
        email_layout.addWidget(QLabel("Email Template:"))
        email_layout.addWidget(self.email_text_edit)
        email_layout.addWidget(self.send_reminder_button)
        email_group = QGroupBox("Send Lease Reminder")
        email_group.setLayout(email_layout)

        # Attached files
        self.attached_files_list = QListWidget()
        self.attach_files_button = QPushButton("Attach Files")
        self.attach_files_button.clicked.connect(self.attach_files)
        self.delete_attached_file_button = QPushButton("Delete File")
        self.delete_attached_file_button.clicked.connect(self.delete_attached_file)
        attached_files_layout = QVBoxLayout()
        attached_files_layout.addWidget(QLabel("Attached Files:"))
        attached_files_layout.addWidget(self.attached_files_list)
        attached_files_layout.addWidget(self.attach_files_button)
        attached_files_layout.addWidget(self.delete_attached_file_button)
        attached_files_group = QGroupBox("Attachments")
        attached_files_group.setLayout(attached_files_layout)

        # Save changes button
        self.save_changes_button = QPushButton("Save Changes")
        self.save_changes_button.clicked.connect(self.save_changes)

        # Add the input widgets to the layout
        input_layout.addWidget(building_group)
        input_layout.addWidget(apartment_group)
        input_layout.addWidget(email_group)
        input_layout.addWidget(attached_files_group)
        input_layout.addWidget(self.save_changes_button)

        # Add the input widget to the splitter
        splitter.addWidget(input_widget)

        # Set the splitter orientation
        splitter.setOrientation(Qt.Vertical)

        self.layout.addWidget(splitter)
        self.setLayout(self.layout)

    def load_data(self):
        self.tree.clear()

        conn = sqlite3.connect(self.db_name)
        c = conn.cursor()

        # Fetch buildings
        c.execute("SELECT * FROM buildings")
        buildings = c.fetchall()

        for building in buildings:
            building_item = QTreeWidgetItem(self.tree, [building[1]])

            # Fetch apartments for the building
            c.execute("SELECT * FROM apartments WHERE building_id=?", (building[0],))
            apartments = c.fetchall()

            for apartment in apartments:
                apartment_item = QTreeWidgetItem(building_item, [
                    apartment[2],  # Apartment Number
                    apartment[3],  # Name
                    apartment[4],  # Email
                    apartment[5],  # Lease Start
                    apartment[6]   # Lease End
                ])
                apartment_item.setFlags(apartment_item.flags() | Qt.ItemIsEditable)

        conn.close()

    def add_building(self):
        building_name = self.building_name_input.text().strip()

        if building_name:
            conn = sqlite3.connect(self.db_name)
            c = conn.cursor()

            # Insert the new building into the database
            c.execute("INSERT INTO buildings (name) VALUES (?)", (building_name,))
            conn.commit()

            conn.close()

            self.building_name_input.clear()
            self.load_data()
        else:
            QMessageBox.warning(self, "Error", "Building name cannot be empty.")

    def add_apartment(self):
        building_item = self.tree.currentItem()

        if building_item:
            building_name = building_item.text(0)
            apartment_name = self.apartment_name_input.text().strip()
            name = self.name_input.text().strip()
            tenant_email = self.tenant_email_input.text().strip()
            lease_start = self.lease_start_input.date().toString("MM/dd/yyyy")
            lease_end = self.lease_end_input.date().toString("MM/dd/yyyy")

            if apartment_name and name and tenant_email:
                conn = sqlite3.connect(self.db_name)
                c = conn.cursor()

                # Fetch the building id from the database
                c.execute("SELECT id FROM buildings WHERE name=?", (building_name,))
                building_id = c.fetchone()[0]

                # Insert the new apartment into the database
                c.execute("INSERT INTO apartments (building_id, name, email, lease_start, lease_end) VALUES (?, ?, ?, ?, ?)",
                          (building_id, apartment_name, name, tenant_email, lease_start, lease_end))
                conn.commit()

                conn.close()

                self.apartment_name_input.clear()
                self.name_input.clear()
                self.tenant_email_input.clear()
                self.load_data()
            else:
                QMessageBox.warning(self, "Error", "Apartment number, name, and tenant email cannot be empty.")
        else:
            QMessageBox.warning(self, "Error", "Please select a building.")

    def handle_item_clicked(self, item, column):
        if item.parent() is not None:
            apartment_name = item.text(0)
            name = item.text(1)
            tenant_email = item.text(2)
            lease_start = item.text(3)
            lease_end = item.text(4)

            self.apartment_name_input.setText(apartment_name)
            self.name_input.setText(name)
            self.tenant_email_input.setText(tenant_email)
            self.lease_start_input.setDate(QDate.fromString(lease_start, "MM/dd/yyyy"))
            self.lease_end_input.setDate(QDate.fromString(lease_end, "MM/dd/yyyy"))

    def handle_item_changed(self, item, column):
        if item.parent() is not None:
            building_name = item.parent().text(0)
            apartment_name = item.text(0)
            name = item.text(1)
            tenant_email = item.text(2)
            lease_start = item.text(3)
            lease_end = item.text(4)

            conn = sqlite3.connect(self.db_name)
            c = conn.cursor()

            # Fetch the building id from the database
            c.execute("SELECT id FROM buildings WHERE name=?", (building_name,))
            building_id = c.fetchone()[0]

            # Update the apartment in the database
            c.execute("UPDATE apartments SET building_id=?, name=?, email=?, lease_start=?, lease_end=? WHERE name=?",
                      (building_id, apartment_name, name, tenant_email, lease_start, lease_end, apartment_name))
            conn.commit()

            conn.close()

            self.load_data()

    def send_reminder(self):
        item = self.tree.currentItem()

        if item is None:
            QMessageBox.warning(self, "Error", "Please select an apartment.")
            return

        apartment_name = item.text(0)
        name = item.text(1)
        tenant_email = item.text(2)
        lease_start = item.text(3)
        lease_end = item.text(4)

        message = self.email_text_edit.toPlainText()
        message = message.replace("{Name}", name)
        message = message.replace("{Lease Start}", lease_start)
        message = message.replace("{Lease End}", lease_end)

        attachments = []
        for index in range(self.attached_files_list.count()):
            file_path = self.attached_files_list.item(index).data(Qt.UserRole)
            attachments.append(file_path)

        # Send the email
        try:
            send_email(tenant_email, "Lease Expiration Reminder", message, attachments)
            QMessageBox.information(self, "Email Sent", "The reminder email has been sent successfully.")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"An error occurred while sending the email: {str(e)}")

    def attach_files(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)

        if file_dialog.exec_():
            file_paths = file_dialog.selectedFiles()

            for file_path in file_paths:
                file_name = os.path.basename(file_path)

                item = QListWidgetItem(file_name)
                item.setData(Qt.UserRole, file_path)
                self.attached_files_list.addItem(item)

    def delete_attached_file(self):
        selected_items = self.attached_files_list.selectedItems()

        if len(selected_items) > 0:
            for item in selected_items:
                self.attached_files_list.takeItem(self.attached_files_list.row(item))

    def save_changes(self):
        reply = QMessageBox.question(self, "Save Changes", "Are you sure you want to save the changes?",
                                     QMessageBox.Yes | QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.save_data()
            QMessageBox.information(self, "Changes Saved", "The changes have been saved successfully.")

    def save_data(self):
        conn = sqlite3.connect(self.db_name)
        c = conn.cursor()

        # Clear the existing data in the database
        c.execute("DELETE FROM buildings")
        c.execute("DELETE FROM apartments")
        conn.commit()

        # Save the current data in the database
        for i in range(self.tree.topLevelItemCount()):
            building_item = self.tree.topLevelItem(i)
            building_name = building_item.text(0)

            # Insert the building into the database
            c.execute("INSERT INTO buildings (name) VALUES (?)", (building_name,))
            conn.commit()

            for j in range(building_item.childCount()):
                apartment_item = building_item.child(j)
                apartment_name = apartment_item.text(0)
                name = apartment_item.text(1)
                tenant_email = apartment_item.text(2)
                lease_start = apartment_item.text(3)
                lease_end = apartment_item.text(4)

                # Fetch the building id from the database
                c.execute("SELECT id FROM buildings WHERE name=?", (building_name,))
                building_id = c.fetchone()[0]

                # Insert the apartment into the database
                c.execute("INSERT INTO apartments (building_id, name, email, lease_start, lease_end) VALUES (?, ?, ?, ?, ?)",
                          (building_id, apartment_name, name, tenant_email, lease_start, lease_end))
                conn.commit()

        conn.close()

    def download_chart(self):
        file_dialog = QFileDialog()
        file_dialog.setAcceptMode(QFileDialog.AcceptSave)
        file_dialog.setFileMode(QFileDialog.AnyFile)
        file_dialog.setDefaultSuffix("png")

        if file_dialog.exec_():
            file_path = file_dialog.selectedFiles()[0]

            # Generate the chart and save it to the selected file path
            self.generate_chart().save(file_path, "PNG")

    def generate_chart(self):
        building_names = []
        apartment_counts = []

        for i in range(self.tree.topLevelItemCount()):
            building_item = self.tree.topLevelItem(i)
            building_names.append(building_item.text(0))
            apartment_counts.append(building_item.childCount())

        # Create a bar chart
        plt.figure(figsize=(12, 6))
        plt.bar(building_names, apartment_counts)
        plt.xlabel("Building")
        plt.ylabel("Number of Apartments")
        plt.title("Apartment Distribution by Building")
        plt.xticks(rotation=45)

        # Convert the chart to a QPixmap for displaying in the GUI
        canvas = plt.gcf().canvas
        canvas.draw()
        data = canvas.buffer_rgba()
        qimage = QImage(data, canvas.get_width_height()[0], canvas.get_width_height()[1], QImage.Format_RGBA8888)
        pixmap = QPixmap.fromImage(qimage)

        plt.close()

        return pixmap

    def closeEvent(self, event):
        reply = QMessageBox.question(self, "Save Changes", "Do you want to save the changes before exiting?",
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)

        if reply == QMessageBox.Yes:
            self.save_data()
            event.accept()
        elif reply == QMessageBox.No:
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ApartmentManagerWindow()
    window.show()
    sys.exit(app.exec_())
