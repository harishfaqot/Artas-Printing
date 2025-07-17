import sys
from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QTableWidgetItem
from PyQt5.QtGui import QBrush, QColor
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
import csv
from lib.EMARK import EMARKPrinter
from lib.IND231 import WeightReader
from lib.PLC import PLCReader
import os
from PyQt5.QtWidgets import QMessageBox
import json
import glob

CONFIG_FILE = "lib/config.json"

class PrintingSystem(QtWidgets.QMainWindow):
    def __init__(self):
        super(PrintingSystem, self).__init__()

        # Load the UI file
        uic.loadUi('lib/home.ui', self)
        
        # Initialize variables
        self.weight_limits_set = False
        self.length_limits_set = False
        self.weight_unit = "lbs"
        self.length_unit = "ft"
        self.weight = 0
        self.length = 0
        self.counter = 0

        self.EMARK = EMARKPrinter()
        self.WEIGHT = WeightReader()
        self.PLC = PLCReader()
        
        # Setup UI elements
        self.setup_ui()

        self.setup_connections()

        self.config = self.load_config()
        self.auto_connect()

        # Create logs directory if it doesn't exist
        if not os.path.exists("logs"):
            os.makedirs("logs")
        
        # Connect signals
        self.connect_signals()
        
    def setup_ui(self):
        # Set initial values
        self.lineEdit_weight.setText("0")
        self.lineEdit_length.setText("0")
        self.spinBox_counter.setValue(0)
        
        # Set default weight and length units
        self.comboBox_weight.setCurrentText("pound (lbs)")
        self.comboBox_length.setCurrentText("foot (ft)")

        self.lineEdit_input.setText("1ST API SPEC 5CT-2221     05-25 PE 7 26.00 K S P 4600 PSI D   {LENGTH} {WEIGHT} HN  241B11000-1  WO 04-0475")
        self.update_output()

        # Setup table columns
        self.setup_table()
        self.load_last_csv()

    def setup_table(self):
        self.tableWidget.setColumnCount(11)
        self.tableWidget.setHorizontalHeaderLabels([
            "Date", "Time", "Weight", "Min", "Max", 
            "Length", "Min", "Max", "Counter", 
            "Printed Text", "Status"
        ])
        
        # Set different resize modes for different columns
        header = self.tableWidget.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)  # Date
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)  # Time
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)  # Weight
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)  # Min (Weight)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)  # Max (Weight)
        header.setSectionResizeMode(5, QtWidgets.QHeaderView.ResizeToContents)  # Length
        header.setSectionResizeMode(6, QtWidgets.QHeaderView.ResizeToContents)  # Min (Length)
        header.setSectionResizeMode(7, QtWidgets.QHeaderView.ResizeToContents)  # Max (Length)
        header.setSectionResizeMode(8, QtWidgets.QHeaderView.ResizeToContents)  # Counter
        header.setSectionResizeMode(9, QtWidgets.QHeaderView.Stretch)          # Printed Text
        header.setSectionResizeMode(10, QtWidgets.QHeaderView.ResizeToContents) # Status
       
        # Set minimum sizes if needed
        header.setMinimumSectionSize(120)  # Minimum width for all columns

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    return json.load(f)
            except:
                return {
                    "printer_port": "",
                    "weight_port": "",
                    "plc_ip": "",
                    "min_weight": 0,
                    "max_weight": 100,
                    "min_length": 0,
                    "max_length": 100
                }
        else:
            return {
                "printer_port": "",
                "weight_port": "",
                "plc_ip": "",
                "min_weight": 0,
                "max_weight": 100,
                "min_length": 0,
                "max_length": 100
            }

    def save_config(self):
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config, f)
        except:
            pass

    def auto_connect(self):
        # Printer
        for i in range(self.comboBox_com_1.count()):
            if self.config["printer_port"] in self.comboBox_com_1.itemText(i):
                self.comboBox_com_1.setCurrentIndex(i)
                self.connect_printer()
                break

        # Weight
        for i in range(self.comboBox_com_2.count()):
            if self.config["weight_port"] in self.comboBox_com_2.itemText(i):
                self.comboBox_com_2.setCurrentIndex(i)
                self.connect_weight()
                break

        # PLC
        self.lineEdit_IP.setText(self.config["plc_ip"])
        if self.config["plc_ip"]:
            self.connect_PLC()

        self.lineEdit_downlimit.setText(str(self.config["min_weight"]))
        self.lineEdit_uplimit.setText(str(self.config["max_weight"]))
        self.lineEdit_downlimit_1.setText(str(self.config["min_length"]))
        self.lineEdit_uplimit_1.setText(str(self.config["max_length"]))

        self.weight_limits_set = True
        self.length_limits_set = True

    def setup_connections(self):
        from serial.tools import list_ports

        ports = list_ports.comports()
        self.comboBox_com_1.clear()
        self.comboBox_com_2.clear()
        for port in ports:
            self.comboBox_com_1.addItem(str(port))
            self.comboBox_com_2.addItem(str(port)) 

    def connect_signals(self):
        # Home tab signals
        self.pushButton_weight.clicked.connect(self.apply_weight_limits)
        self.pushButton_length.clicked.connect(self.apply_length_limits)

        self.pushButton_weight_2.clicked.connect(self.get_weight)
        self.pushButton_length_2.clicked.connect(self.get_length)

        self.pushButton_run.clicked.connect(self.run_printing)
        self.pushButton_run_2.clicked.connect(self.trigger)
        self.pushButton_connect_1.clicked.connect(self.connect_printer)
        self.pushButton_connect_2.clicked.connect(self.connect_weight)
        self.pushButton_connect_3.clicked.connect(self.connect_PLC)
        
        # History tab signals
        self.pushButton_open.clicked.connect(self.open_file)
        self.pushButton_save.clicked.connect(self.save_data)
        self.pushButton_export.clicked.connect(self.export_to_excel)
        
        # Combo box changes
        self.comboBox_weight.currentTextChanged.connect(self.update_weight_unit)
        self.comboBox_length.currentTextChanged.connect(self.update_length_unit)

        self.lineEdit_input.textChanged.connect(self.update_output)
        
    def update_weight_unit(self, text):
        """Update the weight unit based on combo box selection"""
        self.weight_unit = text.split("(")[1].strip(")")
        self.label_weight1.setText(f"{self.weight_unit}   -")
        self.label_weight2.setText(self.weight_unit)
    
    def update_length_unit(self, text):
        """Update the length unit based on combo box selection"""
        self.length_unit = text.split("(")[1].strip(")")
        self.label_length1.setText(f"{self.length_unit}   -")
        self.label_length2.setText(self.length_unit)
            
    def apply_weight_limits(self):
        try:
            lower = float(self.lineEdit_downlimit.text())
            upper = float(self.lineEdit_uplimit.text())
            
            if lower >= upper:
                QMessageBox.warning(self, "Invalid Range", "Lower limit must be less than upper limit")
                return
                
            self.weight_limits_set = True

            self.config["min_weight"] = lower
            self.config["max_weight"] = upper
            self.save_config()
            
            QMessageBox.information(self, "Success", "Weight limits applied successfully")
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter valid numbers for weight limits")
            
    def apply_length_limits(self):
        try:
            lower = float(self.lineEdit_downlimit_1.text())
            upper = float(self.lineEdit_uplimit_1.text())
            
            if lower >= upper:
                QMessageBox.warning(self, "Invalid Range", "Lower limit must be less than upper limit")
                return
                
            self.length_limits_set = True

            self.config["min_length"] = lower
            self.config["max_length"] = upper
            self.save_config()

            QMessageBox.information(self, "Success", "Length limits applied successfully")
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter valid numbers for length limits")
            
    def connect_printer(self):
        if self.EMARK.connected:
            self.EMARK.close()
            self.printer_connected = False
            self.pushButton_connect_1.setText("Connect")
            self.pushButton_connect_1.setStyleSheet("""
                QPushButton{
                    background-color: rgb(255, 170, 0);
                }
                QPushButton:hover{
                    background-color: rgb(255, 190, 0);
                }
                QPushButton:pressed{
                    background-color: rgb(255, 210, 0);
                }
            """)
        else:
            port = self.comboBox_com_1.currentText().split()[0]
            msg = self.EMARK.connect(port= port)
            if self.EMARK.connected:
                self.pushButton_connect_1.setText("Connected")
                self.pushButton_connect_1.setStyleSheet("""
                    QPushButton{
                        background-color: rgb(0, 170, 0);
                    }
                    QPushButton:hover{
                        background-color: rgb(0, 190, 0);
                    }
                    QPushButton:pressed{
                        background-color: rgb(0, 210, 0);
                    }
                """)

                self.config["printer_port"] = port
                self.save_config()
                self.EMARK.clear_text()
            else:
                QMessageBox.critical(self, "Connection Failed", msg)
    
    def connect_weight(self):
        if self.WEIGHT.connected:
            self.WEIGHT.close()
            self.pushButton_connect_2.setText("Connect")
            self.pushButton_connect_2.setStyleSheet("""
                QPushButton{
                    background-color: rgb(255, 170, 0);
                }
                QPushButton:hover{
                    background-color: rgb(255, 190, 0);
                }
                QPushButton:pressed{
                    background-color: rgb(255, 210, 0);
                }
            """)
        else:
            port = self.comboBox_com_2.currentText().split()[0]
            msg = self.WEIGHT.connect(port= port)
            if self.WEIGHT.connected:
                self.pushButton_connect_2.setText("Connected")
                self.pushButton_connect_2.setStyleSheet("""
                    QPushButton{
                        background-color: rgb(0, 170, 0);
                    }
                    QPushButton:hover{
                        background-color: rgb(0, 190, 0);
                    }
                    QPushButton:pressed{
                        background-color: rgb(0, 210, 0);
                    }
                """)
                self.config["weight_port"] = port
                self.save_config()
            else:
                QMessageBox.critical(self, "Connection Failed", msg)

    def connect_PLC(self):
        if self.PLC.connected:
            self.PLC.close()
            self.pushButton_connect_3.setText("Connect")
            self.pushButton_connect_3.setStyleSheet("""
                QPushButton{
                    background-color: rgb(255, 170, 0);
                }
                QPushButton:hover{
                    background-color: rgb(255, 190, 0);
                }
                QPushButton:pressed{
                    background-color: rgb(255, 210, 0);
                }
            """)
        else:
            ip_address = self.lineEdit_IP.text()
            msg = self.PLC.connect(ip= ip_address)
            if self.PLC.connected:
                self.pushButton_connect_3.setText("Connected")
                self.pushButton_connect_3.setStyleSheet("""
                    QPushButton{
                        background-color: rgb(0, 170, 0);
                    }
                    QPushButton:hover{
                        background-color: rgb(0, 190, 0);
                    }
                    QPushButton:pressed{
                        background-color: rgb(0, 210, 0);
                    }
                """)
                self.config["plc_ip"] = ip_address
                self.save_config()
            else:
                QMessageBox.critical(self, "Connection Failed", msg)

    def update_output(self):
        input_text = self.lineEdit_input.text()
        output_text = input_text.replace("{WEIGHT}", f"{self.weight} {self.weight_unit.upper()}")
        output_text = output_text.replace("{LENGTH}", f"{self.length} {self.length_unit.upper()}")
        output_text = output_text.replace("{COUNTER}", f"{self.counter}")
        
        self.lineEdit_output.setText(output_text)

    def run_printing(self):
        if not self.EMARK.connected:
            QMessageBox.warning(self, "Printer Not Connected", "Please connect to the printer first")
            return
        if not self.WEIGHT.connected:
            QMessageBox.warning(self, "Weight Indicator Not Connected", "Please connect to Weight Indicator first")
            return
        if not self.PLC.connected:
            QMessageBox.warning(self, "PLC Not Connected", "Please connect to the PLC first")
            return
            
        if not self.weight_limits_set or not self.length_limits_set:
            QMessageBox.warning(self, "Limits Not Set", "Please set both weight and length limits first")
            return
        

    def get_weight(self):
            weight = round(self.WEIGHT.read_weight(), 2)
            self.lineEdit_weight.setText(f"{weight:.2f}")
    def get_length(self):
            length = round(self.PLC.read_real(db_number=2, start_byte=0), 2)
            self.lineEdit_length.setText(f"{length:.2f}")

    def trigger(self):
        try:
            weight_lower = float(self.lineEdit_downlimit.text())
            weight_upper = float(self.lineEdit_uplimit.text())
            length_lower = float(self.lineEdit_downlimit_1.text())
            length_upper = float(self.lineEdit_uplimit_1.text())

            # self.get_weight()
            # self.get_length()

            weight = round(float(self.lineEdit_weight.text()),2)
            length = round(float(self.lineEdit_length.text()),2)
            
            # weight = round(self.WEIGHT.read_weight(), 2)
            # length = round(self.PLC.read_real(db_number=2, start_byte=0), 2)
            # weight = 1000
            # length = 90.01
            counter = self.spinBox_counter.value()
            
            # Determine status
            status = "OK"
            if weight < weight_lower:
                status = "UNDERWEIGHT"
            elif weight > weight_upper:
                status = "OVERWEIGHT"
            elif length < length_lower:
                status = "UNDERLENGTH"
            elif length > length_upper:
                status = "OVERLENGTH"
            
            # Update display
            self.lineEdit_weight.setText(f"{weight:.2f}")
            self.lineEdit_length.setText(f"{length:.2f}")
            
            # Process input text
            input_text = self.lineEdit_input.text()
            output_text = input_text.replace("{WEIGHT}", f"{weight:.2f} {self.weight_unit.upper()}")
            output_text = output_text.replace("{LENGTH}", f"{length:.2f} {self.length_unit.upper()}")
            output_text = output_text.replace("{COUNTER}", f"{counter}")
            
            self.lineEdit_output.setText(output_text)

            # Only print if status is OK
            if status == "OK":
                print(output_text)
                response = self.EMARK.send_text(output_text, 
                                        template_num=1, 
                                        font_height=0x0C,  # 16 dot matrix
                                        x_pos=0, 
                                        y_pos=0)
                print(f"Response: {response.hex()}")
                
                # Increment counter only if printed
                self.spinBox_counter.setValue(counter + 1)
            else:
                QMessageBox.warning(self, "Quality Alert", f"Pipe rejected: {status}")
            
            # Add to history (whether printed or rejected)
            self.add_to_history(weight, length, counter, output_text, status)
            
        except ValueError:
            QMessageBox.warning(self, "Error", "Invalid limit values")
        except Exception as e:
            print(e)
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
            
    def add_to_history(self, weight, length, counter, output_text, status):
        current_time = datetime.datetime.now()
        date_str = current_time.strftime("%Y-%m-%d")
        time_str = current_time.strftime("%H:%M:%S")
        
        row_position = self.tableWidget.rowCount()
        self.tableWidget.insertRow(row_position)
        
        weight_lower = self.lineEdit_downlimit.text()
        weight_upper = self.lineEdit_uplimit.text()
        length_lower = self.lineEdit_downlimit_1.text()
        length_upper = self.lineEdit_uplimit_1.text()
        
        for col, value in enumerate([
            date_str,
            time_str,
            f"{weight} {self.weight_unit}",
            f"{weight_lower}",
            f"{weight_upper}",
            f"{length} {self.length_unit}",
            f"{length_lower}",
            f"{length_upper}",
            str(counter),
            output_text,
            status
        ]):
            item = QTableWidgetItem(value)
            item.setTextAlignment(Qt.AlignCenter)
            item.setFlags(item.flags() ^ Qt.ItemIsEditable)
            
            # Color code based on status
            if col == 10 and status != "OK":  # Status column is now index 10
                if "WEIGHT" in status:
                    item.setBackground(QBrush(QColor("yellow")))
                elif "LENGTH" in status:
                    item.setBackground(QBrush(QColor("orange")))
            
            self.tableWidget.setItem(row_position, col, item)

        # Save to CSV
        self.save_to_csv(date_str, time_str, weight, length, counter, output_text, status)

    def save_to_csv(self, date, time, weight, length, counter, output_text, status):
        current_time = datetime.datetime.now()
        date_str = current_time.strftime("%Y-%m-%d")
        csv_file = f"logs/{date_str}.csv"
        
        file_exists = os.path.isfile(csv_file)
        
        weight_min = self.lineEdit_downlimit.text()
        weight_max = self.lineEdit_uplimit.text()
        length_min = self.lineEdit_downlimit_1.text()
        length_max = self.lineEdit_uplimit_1.text()
        
        try:
            with open(csv_file, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                if not file_exists:
                    writer.writerow([
                        "Date", "Time", "Weight", "Min", "Max", 
                        "Length", "Min", "Max", "Counter", 
                        "Printed Text", "Status"
                    ])
                
                writer.writerow([
                    date,
                    time,
                    f"{weight} {self.weight_unit}",
                    weight_min,
                    weight_max,
                    f"{length} {self.length_unit}",
                    length_min,
                    length_max,
                    str(counter),
                    output_text,
                    status
                ])
        except Exception as e:
            print(f"Error saving to CSV: {e}")

    def load_last_csv(self):
        try:
            csv_files = glob.glob("logs/*.csv")
            if not csv_files:
                return

            latest_file = max(csv_files, key=os.path.getmtime)
            self.lineEdit_path.setText(latest_file)

            with open(latest_file, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader)  # Skip header row

                self.tableWidget.setRowCount(0)

                for row_data in reader:
                    row_position = self.tableWidget.rowCount()
                    self.tableWidget.insertRow(row_position)
                    for col, value in enumerate(row_data):
                        item = QTableWidgetItem(value)
                        item.setTextAlignment(Qt.AlignCenter)
                        item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                        
                        # Apply color coding for rejected items when loading
                        if col == 10 and value != "OK":  # Status column is now index 10
                            if "WEIGHT" in value:
                                item.setBackground(QBrush(QColor("yellow")))
                            elif "LENGTH" in value:
                                item.setBackground(QBrush(QColor("orange")))
                                
                        self.tableWidget.setItem(row_position, col, item)
        except Exception as e:
            print(f"Failed to load last CSV: {e}")

    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open CSV File",
            "./logs",  # or use absolute path like "/home/user/logs"
            "CSV Files (*.csv)"
        )
        if file_path:
            self.lineEdit_path.setText(file_path)
            try:
                with open(file_path, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    headers = next(reader)

                    self.tableWidget.setRowCount(0)

                    for row_data in reader:
                        row_position = self.tableWidget.rowCount()
                        self.tableWidget.insertRow(row_position)
                        for col, value in enumerate(row_data):
                            item = QTableWidgetItem(value)
                            item.setTextAlignment(Qt.AlignCenter)
                            item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                            
                            # Apply color coding for rejected items
                            if col == 10 and value != "OK":  # Status column
                                if "WEIGHT" in value:
                                    item.setBackground(QBrush(QColor("yellow")))
                                elif "LENGTH" in value:
                                    item.setBackground(QBrush(QColor("orange")))
                                    
                            self.tableWidget.setItem(row_position, col, item)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to open file: {str(e)}")
            
    def save_data(self):
        QMessageBox.information(self, "Data Saved", "Data has been saved successfully")
        
    def export_to_excel(self):
        if self.tableWidget.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "There is no data to export")
            return
            
        file_path, _ = QFileDialog.getSaveFileName(self, "Export to Excel", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
                
            try:
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet.title = "Printing History"
                
                # Write headers
                headers = []
                for col in range(self.tableWidget.columnCount()):
                    headers.append(self.tableWidget.horizontalHeaderItem(col).text())
                sheet.append(headers)
                
                # Make headers bold
                for cell in sheet[1]:
                    cell.font = Font(bold=True)
                
                # Write data with color coding
                for row in range(self.tableWidget.rowCount()):
                    row_data = []
                    for col in range(self.tableWidget.columnCount()):
                        item = self.tableWidget.item(row, col)
                        row_data.append(item.text() if item else "")
                    sheet.append(row_data)
                    
                    # Apply color to entire row based on status
                    status_item = self.tableWidget.item(row, 10)  # Status column is now index 10
                    if status_item and status_item.text() != "OK":
                        fill_color = "FFFF00" if "WEIGHT" in status_item.text() else "FFA500"
                        for cell in sheet[row+2]:  # +2 because headers are row 1 and Excel is 1-based
                            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                
                workbook.save(file_path)
                QMessageBox.information(self, "Export Successful", f"Data exported to {file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "Export Failed", f"Error exporting data: {str(e)}")

if __name__ == "__main__":
    from PyQt5.QtGui import QPixmap
    from PyQt5.QtWidgets import QSplashScreen
    from PyQt5.QtCore import QTimer

    app = QtWidgets.QApplication(sys.argv)

    # Create splash screen
    splash_pix = QPixmap("lib/logo.jpg")  # <-- you can use any image here
    splash = QSplashScreen(splash_pix)
    splash.show()
    app.processEvents()  # Show the splash screen immediately

    window = PrintingSystem()
    window.show()
    splash.finish(window)

    sys.exit(app.exec_())
