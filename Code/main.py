import sys
import os
import json
import csv
import logging
import traceback
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget, QDialog, QCheckBox, QDialogButtonBox, QFileDialog, QMessageBox, QLabel, QProgressBar, QMenuBar, QAction, QWidget

# Setup logging
logging.basicConfig(filename='user_scanner.log', level=logging.DEBUG, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

class ExportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Export User List")
        self.layout = QVBoxLayout(self)
        
        self.select_all_button = QPushButton("Select All", self)
        self.select_all_button.clicked.connect(self.select_all_options)
        self.layout.addWidget(self.select_all_button)
        
        self.options = {
            "Username": QCheckBox("Username"),
            "Full Name": QCheckBox("Full Name"),
            "Comment": QCheckBox("Comment"),
            "User ID": QCheckBox("User ID"),
            "Primary Group ID": QCheckBox("Primary Group ID"),
            "Last Logon": QCheckBox("Last Logon"),
            "Last Logoff": QCheckBox("Last Logoff"),
            "Password Last Set": QCheckBox("Password Last Set"),
            "Account Expires": QCheckBox("Account Expires"),
            "Number of Logons": QCheckBox("Number of Logons"),
            "Bad Password Count": QCheckBox("Bad Password Count"),
            "Home Directory": QCheckBox("Home Directory"),
            "Script Path": QCheckBox("Script Path"),
            "Profile Path": QCheckBox("Profile Path"),
        }
        
        for option in self.options.values():
            self.layout.addWidget(option)
        
        self.buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel, self)
        self.layout.addWidget(self.buttons)
        
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
    
    def get_selected_options(self):
        return [key for key, checkbox in self.options.items() if checkbox.isChecked()]

    def select_all_options(self):
        for checkbox in self.options.values():
            checkbox.setChecked(True)

class UserScanner(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle("User Scanner")
        self.setGeometry(100, 100, 600, 400)
        
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        
        self.layout = QVBoxLayout(self.central_widget)
        
        # Menu
        self.menu_bar = self.menuBar()
        self.file_menu = self.menu_bar.addMenu("File")
        self.help_menu = self.menu_bar.addMenu("Help")
        
        self.exit_action = QAction("Exit", self)
        self.exit_action.triggered.connect(self.close)
        self.file_menu.addAction(self.exit_action)
        
        self.tutorial_action = QAction("Tutorial", self)
        self.tutorial_action.triggered.connect(self.show_tutorial)
        self.help_menu.addAction(self.tutorial_action)
        
        # Scan Users Button
        self.scan_button = QPushButton("Scan Users", self)
        self.scan_button.clicked.connect(self.scan_users)
        self.layout.addWidget(self.scan_button)
        
        # User List
        self.user_list = QListWidget(self)
        self.layout.addWidget(self.user_list)
        
        # Export User List Button
        self.export_button = QPushButton("Export User List", self)
        self.export_button.clicked.connect(self.open_export_dialog)
        self.layout.addWidget(self.export_button)
        
        # Progress Bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        self.layout.addWidget(self.progress_bar)
        
        # Status Bar
        self.status_bar = self.statusBar()
    
    def show_tutorial(self):
        tutorial_text = (
            "1. Click 'Scan Users' to scan for all users on the system.\n"
            "2. Select the users you want to export from the list.\n"
            "3. Click 'Export User List' to choose export options and save the user list.\n"
            "4. Choose the information you want to include in the export and click 'Save'.\n"
            "5. Select the location and format for the export file."
        )
        QMessageBox.information(self, "Tutorial", tutorial_text)
    
    def scan_users(self):
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            self.status_bar.showMessage("Scanning users...")
            
            users = self.get_users()
            
            self.user_list.clear()
            for user in users:
                self.user_list.addItem(user)
            
            self.progress_bar.setValue(100)
            self.status_bar.showMessage("User scan complete.", 5000)
        except Exception as e:
            logging.error(f"Error scanning users: {e}")
            logging.debug(traceback.format_exc())
            self.show_error("An error occurred while scanning users.")
    
    def get_users(self):
        try:
            import win32net
            import win32netcon

            self.progress_bar.setValue(20)
            users, _, _ = win32net.NetUserEnum(None, 0, win32netcon.FILTER_NORMAL_ACCOUNT)
            user_list = [user['name'] for user in users]
            self.progress_bar.setValue(60)
            logging.info(f"Found users: {user_list}")
            return user_list
        except ImportError as e:
            logging.error(f"Error importing modules: {e}")
            logging.debug(traceback.format_exc())
            self.progress_bar.setValue(0)
            raise RuntimeError("Required modules are not installed.") from e
        except Exception as e:
            logging.error(f"Error getting users: {e}")
            logging.debug(traceback.format_exc())
            self.progress_bar.setValue(0)
            raise RuntimeError("Failed to retrieve user information.") from e

    def open_export_dialog(self):
        try:
            dialog = ExportDialog(self)
            if dialog.exec_():
                selected_options = dialog.get_selected_options()
                self.export_user_list(selected_options)
        except Exception as e:
            logging.error(f"Error opening export dialog: {e}")
            logging.debug(traceback.format_exc())
            self.show_error("An error occurred while opening the export dialog.")
    
    def export_user_list(self, options):
        try:
            file_dialog = QFileDialog(self)
            save_path, _ = file_dialog.getSaveFileName(self, "Save File", "", "JSON Files (*.json);;CSV Files (*.csv);;Text Files (*.txt)")
        
            if save_path:
                users_info = self.get_users_info(options)
                if save_path.endswith(".json"):
                    self.save_as_json(save_path, users_info)
                elif save_path.endswith(".csv"):
                    self.save_as_csv(save_path, users_info)
                else:
                    self.save_as_txt(save_path, users_info)
            
                QMessageBox.information(self, "Success", "User list exported successfully.")
                self.status_bar.showMessage("User list exported successfully.", 5000)
        except Exception as e:
            logging.error(f"Error exporting user list: {e}")
            logging.debug(traceback.format_exc())
            self.show_error("An error occurred while exporting the user list.")
    
    def get_users_info(self, options):
        try:
            import win32net
            import win32netcon

            users = self.get_users()
            users_info = []

            for username in users:
                user_info = {}
                user_details = win32net.NetUserGetInfo(None, username, 2)
            
                if "Username" in options:
                    user_info["Username"] = username
                if "Full Name" in options:
                    user_info["Full Name"] = user_details.get("full_name", "")
                if "Comment" in options:
                    user_info["Comment"] = user_details.get("comment", "")
                if "User ID" in options:
                    user_info["User ID"] = user_details.get("user_id", "")
                if "Primary Group ID" in options:
                    user_info["Primary Group ID"] = user_details.get("primary_group_id", "")
                if "Last Logon" in options:
                    user_info["Last Logon"] = datetime.fromtimestamp(user_details.get("last_logon", 0)).strftime('%Y-%m-%d %H:%M:%S')
                if "Last Logoff" in options:
                    user_info["Last Logoff"] = datetime.fromtimestamp(user_details.get("last_logoff", 0)).strftime('%Y-%m-%d %H:%M:%S')
                if "Password Last Set" in options:
                    user_info["Password Last Set"] = datetime.fromtimestamp(user_details.get("password_age", 0)).strftime('%Y-%m-%d %H:%M:%S')
                if "Account Expires" in options:
                    expires = user_details.get("acct_expires", 0)
                    user_info["Account Expires"] = "Never" if expires == win32netcon.TIMEQ_FOREVER else datetime.fromtimestamp(expires).strftime('%Y-%m-%d %H:%M:%S')
                if "Number of Logons" in options:
                    user_info["Number of Logons"] = user_details.get("num_logons", 0)
                if "Bad Password Count" in options:
                    user_info["Bad Password Count"] = user_details.get("bad_pw_count", 0)
                if "Home Directory" in options:
                    user_info["Home Directory"] = user_details.get("home_dir", "")
                if "Script Path" in options:
                    user_info["Script Path"] = user_details.get("script_path", "")
                if "Profile Path" in options:
                    user_info["Profile Path"] = user_details.get("profile", "")

                users_info.append(user_info)

            logging.info(f"Users info: {users_info}")
            return users_info
        except ImportError as e:
            logging.error(f"Error importing modules: {e}")
            logging.debug(traceback.format_exc())
            raise RuntimeError("Required modules are not installed.") from e
        except Exception as e:
            logging.error(f"Error getting users info: {e}")
            logging.debug(traceback.format_exc())
            raise RuntimeError("Failed to retrieve user details.") from e
        
    def save_as_json(self, path, data):
        try:
            with open(path, 'w') as file:
                json.dump(data, file, indent=4)
            logging.info(f"Data saved as JSON: {path}")
        except Exception as e:
            logging.error(f"Error saving as JSON: {e}")
            logging.debug(traceback.format_exc())
            raise RuntimeError("Failed to save as JSON.") from e
    
    def save_as_csv(self, path, data):
        try:
            keys = data[0].keys()
            with open(path, 'w', newline='') as file:
                dict_writer = csv.DictWriter(file, keys)
                dict_writer.writeheader()
                dict_writer.writerows(data)
            logging.info(f"Data saved as CSV: {path}")
        except Exception as e:
            logging.error(f"Error saving as CSV: {e}")
            logging.debug(traceback.format_exc())
            raise RuntimeError("Failed to save as CSV.") from e
    
    def save_as_txt(self, path, data):
        try:
            with open(path, 'w') as file:
                for entry in data:
                    file.write(f"{entry}\n")
            logging.info(f"Data saved as TXT: {path}")
        except Exception as e:
            logging.error(f"Error saving as TXT: {e}")
            logging.debug(traceback.format_exc())
            raise RuntimeError("Failed to save as TXT.") from e

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)
        self.status_bar.showMessage(message, 5000)

def main():
    app = QApplication(sys.argv)
    scanner = UserScanner()
    scanner.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
