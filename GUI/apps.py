import sys
from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMessageBox, QTableWidgetItem
from PyQt5.QtGui import QBrush, QColor
from lib.EMARK import EMARKPrinter
from lib.IND231 import WeightReader
from lib.PLC import PLCReader
from lib.table import setup_table_functionality, add_to_history, open_file, save_data, export_to_excel, load_last_csv
import os
from PyQt5.QtWidgets import QMessageBox
import json
import time
from collections import deque
from PyQt5.QtWidgets import QTableWidgetItem

CONFIG_FILE = "lib/config.json"

color_green = "#00aa00"
color_red = "#D73232"

class PrintingSystem(QtWidgets.QMainWindow):
    def __init__(self):
        super(PrintingSystem, self).__init__()

        # Load the UI file
        uic.loadUi('lib/home.ui', self)
        
        # Initialize variables
        self.weight_unit = "kg"
        self.length_unit = "mm"
        self.weight = 0
        self.length = 0
        self.length_converted = False
        self.weight_converted = False
        self.OD = 0
        self.WT = 0
        self.length_counter = 0
        self.weight_counter = 0
        self.printer_counter = 0
        self.status = "NORMAL"

        self.length_timer = 0
        self.weight_timer = 0
        self.printer_timer = 0
        self.printer_processed = False
        self.length_processed = False
        self.weight_processed = False

        self.pipe_queue = deque()

        self.EMARK = EMARKPrinter()
        self.WEIGHT = WeightReader()
        self.PLC = PLCReader()
   
        self.setup_connections()

        self.config = self.load_config()

        # Create logs directory if it doesn't exist
        if not os.path.exists("logs"):
            os.makedirs("logs")
        
        # Connect signals
        self.connect_signals()
        self.auto_connect()

        # Start polling
        self.sensor_timer = QTimer()
        self.sensor_timer.timeout.connect(self.poll_sensors)
        self.sensor_timer.start(1000)
        

    def setup_table(self):
        self.tableWidget.setColumnCount(6)
        self.tableWidget.setHorizontalHeaderLabels([
            "Date", "Time", "Length", "Weight", "Printed Text", "Status"
        ])
        self.tableWidget.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        
        # Set different resize modes for different columns
        header = self.tableWidget.horizontalHeader()
        # header.setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)  # Date
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)  # Time
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(5, QtWidgets.QHeaderView.ResizeToContents)
       
        # Set minimum sizes if needed
        header.setMinimumSectionSize(120)  # Minimum width for all columns

        self.tableWidget_home.setColumnCount(4)
        self.tableWidget_home.setHorizontalHeaderLabels(["Printing Text", "Length", "Weight", "Status Print"])
        self.tableWidget_home.setRowCount(999)
        self.tableWidget_home.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        header = self.tableWidget_home.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)  # Stretch all columns equally
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setDefaultAlignment(Qt.AlignCenter)
        vheader = self.tableWidget_home.verticalHeader()
        vheader.setDefaultAlignment(Qt.AlignCenter)

        self.tableWidget_input.setColumnCount(5)
        self.tableWidget_input.setHorizontalHeaderLabels(["TEXT 1", "TEXT 2", "HEAT NUMBER", "WORK ORDER", "PIPE NUMBER"])
        self.tableWidget_input.setRowCount(999)
        header = self.tableWidget_input.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)  # Stretch all columns equally
        header.setDefaultAlignment(Qt.AlignCenter)
        vheader = self.tableWidget_input.verticalHeader()
        vheader.setDefaultAlignment(Qt.AlignCenter)

        data = [
            "1ST API SPEC 5CT-2221",
            "05-25 PE 7 26.00 K S P 4600 PSI D",
            "HN 241B11000-1",
            "WO 04-0475",
            ""  # Empty pipe number
        ]

        for i in range(3):
            for col, text in enumerate(data):
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignCenter)
                self.tableWidget_input.setItem(i, col, item)

        # Apply copy-paste functionality
        setup_table_functionality(self, self.tableWidget_input)

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
                    "min_length": 0,
                    "OD": 0,
                    "WT": 0
                }
        else:
            return {
                "printer_port": "",
                "weight_port": "",
                "plc_ip": "",
                "min_length": 0,
                "OD": 0,
                "WT": 0
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
        if self.config["plc_ip"]:
            self.connect_PLC()

    def setup_connections(self):
        from serial.tools import list_ports

        ports = list_ports.comports()
        self.comboBox_com_1.clear()
        self.comboBox_com_2.clear()
        for port in ports:
            self.comboBox_com_1.addItem(str(port))
            self.comboBox_com_2.addItem(str(port)) 

    def poll_sensors(self):
        try:
            self.length = round(self.PLC.read_real(db_number=2, start_byte=0), 2)
            self.weight = round(self.WEIGHT.read_weight(),2)

            if self.length_unit == "ft":
                self.length = round(self.length / 304.8, 2)

            if self.weight_unit == "lbs":
                self.weight = round(self.weight * 2.20462262, 2)

            
            self.lineEdit_weight.setText(f"{self.weight} {self.weight_unit}")
            self.lineEdit_length.setText(f"{self.length} {self.length_unit}")

            now = time.time()

            rowCount = self.tableWidget_home.rowCount()
            for row in range(rowCount):
                length_item = self.tableWidget_home.item(row, 1)
                if not length_item:
                    self.length_counter = row
                    break
            for row in range(rowCount):
                weight_item = self.tableWidget_home.item(row, 2)
                if not weight_item:
                    self.weight_counter = row
                    break
            # print(f"Total Row: {rowCount} | Length Counter: {self.length_counter} | Weight Counter: {self.weight_counter}")

            # LENGTH sensor
            length_on = self.PLC.read_bit(byte=6, bit=5)
            length_on_2 = self.PLC.read_mem(byte=236, bit=0)
            if length_on and length_on_2:
                if self.PLC.connected and not self.length_processed:
                    self.length_status.setText("LENGTH : MEASURING")

                if now - self.length_timer >= 2 and not self.length_processed:
                    if self.length<=0:
                        self.length_status.setText("LENGTH : INVALID")
                    else:
                        current_text = self.tableWidget_home.item(self.length_counter, 0)
                        if current_text:
                            status_length = self.check_length(self.length)
                            length_text = f"{self.length}\n({status_length})"

                            length_item = QTableWidgetItem(length_text)
                            length_item.setTextAlignment(Qt.AlignCenter)
                            if "UNDERLENGTH" in status_length or "OVERLENGTH" in status_length:
                                length_item.setBackground(QBrush(QColor(color_red)))
                                self.status = "REJECT"
                            elif "NORMAL" in status_length:
                                length_item.setBackground(QBrush(QColor(color_green)))
                                self.status = "NORMAL"

                            self.tableWidget_home.setItem(self.length_counter, 1, length_item)
                            self.tableWidget_home.setWordWrap(True)
                            self.tableWidget_home.resizeRowsToContents()

                            current_text = current_text.text().replace("[L]", str(self.length), 1)
                            self.tableWidget_home.setItem(self.length_counter, 0, QTableWidgetItem(current_text))

                            # self.length_counter +=1
                            self.highlight_row_by_counter()

                            self.length_processed = True
                            self.length_timer = now
                            self.length_status.setText("LENGTH : MEASURE DONE")
                        
                        else:
                            self.length_status.setText("LENGTH : ROW EMPTY")
            else:
                if self.PLC.connected:
                    self.length_status.setText("LENGTH : ONLINE")
                    self.length_status.setStyleSheet("""
                        background-color: rgb(0, 170, 0);
                        border-radius: 5px;
                        border: none;
                    """)
                else:
                    self.length_status.setText("LENGTH : DISCONNECTED")
                    self.length_status.setStyleSheet("""
                        background-color: rgb(255, 170, 0);
                        border-radius: 5px;
                        border: none;
                    """)
                self.length_timer = now
                self.length_processed = False

            # WEIGHT sensor
            weight_on = self.PLC.read_bit(byte=6, bit=6)
            if weight_on:
                if self.WEIGHT.connected and not self.weight_processed:
                    self.weight_status.setText("WEIGHT : MEASURING")

                if now - self.weight_timer >= 2  and not self.weight_processed and self.weight_counter<self.length_counter:
                    if self.weight <= 0:
                        self.weight_status.setText("LENGTH : INVALID")
                    else:
                        status_weight = self.check_weight(self.weight)
                        weight_text = f"{self.weight}\n({status_weight})"

                        weight_item = QTableWidgetItem(weight_text)
                        weight_item.setTextAlignment(Qt.AlignCenter)
                        if "UNDERWEIGHT" in status_weight or "OVERWEIGHT" in status_weight or "REJECT" in self.status:
                            weight_item.setBackground(QBrush(QColor(color_red)))

                            printed_item = QTableWidgetItem("WAITING (REJECT)")
                            printed_item.setTextAlignment(Qt.AlignCenter)
                            self.tableWidget_home.setItem(self.weight_counter, 3, printed_item)

                        elif "NORMAL" in status_weight and "NORMAL" in self.status:
                            weight_item.setBackground(QBrush(QColor(color_green)))

                            printed_item = QTableWidgetItem("WAITING (NORMAL)")
                            printed_item.setTextAlignment(Qt.AlignCenter)
                            self.tableWidget_home.setItem(self.weight_counter, 3, printed_item)

                        self.tableWidget_home.setItem(self.weight_counter, 2, weight_item)
                        self.tableWidget_home.setWordWrap(True)
                        self.tableWidget_home.resizeRowsToContents()

                        current_text = self.tableWidget_home.item(self.weight_counter, 0)
                        current_text = current_text.text().replace("[W]", str(self.weight), 1)
                        self.tableWidget_home.setItem(self.weight_counter, 0, QTableWidgetItem(current_text))

                        # self.weight_counter +=1

                        self.weight_processed = True
                        self.weight_timer = now
                        self.weight_status.setText("WEIGHT : MEASURE DONE")

                elif not self.weight_processed and self.weight_counter>=self.length_counter:
                    self.weight_status.setText("WEIGHT : ROW EMPTY")
            else:
                if self.WEIGHT.connected:
                    self.weight_status.setText("WEIGHT : ONLINE")
                    self.weight_status.setStyleSheet("""
                        background-color: rgb(0, 170, 0);
                        border-radius: 5px;
                        border: none;
                    """)
                else:
                    self.weight_status.setText("WEIGHT : DISCONNECTED")
                    self.weight_status.setStyleSheet("""
                        background-color: rgb(255, 170, 0);
                        border-radius: 5px;
                        border: none;
                    """)
                self.weight_timer = now
                self.weight_processed = False
            
            # WEIGHT sensor
            printer_on = self.PLC.read_bit(byte=6, bit=4)
            if printer_on:
                length_text = self.tableWidget_home.item(self.printer_counter, 1)
                weight_text = self.tableWidget_home.item(self.printer_counter, 2)

                if now - self.printer_timer >= 2 and not self.printer_processed and length_text and weight_text:
                    status_length = self.tableWidget_home.item(self.printer_counter, 3)
                    status_weight = self.tableWidget_home.item(self.printer_counter, 3)
                    status_print = self.tableWidget_home.item(self.printer_counter, 3)
                    if status_print:
                        status_print = status_print.text()
                        result = status_print.split('(')[1].strip(')')
                        self.printer(result)

                        if "REJECT" == result:
                            printed_item = QTableWidgetItem("REJECT")
                            printed_item.setTextAlignment(Qt.AlignCenter)
                            printed_item.setBackground(QBrush(QColor(color_red)))
                            self.tableWidget_home.setItem(self.printer_counter, 3, printed_item)
                        elif "NORMAL" == result:
                            printed_item = QTableWidgetItem("NORMAL")
                            printed_item.setTextAlignment(Qt.AlignCenter)
                            printed_item.setBackground(QBrush(QColor(color_green)))
                            self.tableWidget_home.setItem(self.printer_counter, 3, printed_item)

                    self.printer_counter+=1
                    self.printer_processed = True
                    self.printer_timer = now
            else:
                self.printer_timer = now
                self.printer_processed = False

        except Exception as e:
            print(f"Sensor read error: {e}")

    def connect_signals(self):
        # Home tab signals
        self.pushButton_connect_1.clicked.connect(self.connect_printer)
        self.pushButton_connect_2.clicked.connect(self.connect_weight)
        self.pushButton_connect_3.clicked.connect(self.connect_PLC)
        
        # History tab signals
        self.pushButton_open.clicked.connect(lambda: open_file(self))
        self.pushButton_save.clicked.connect(lambda: save_data(self))
        self.pushButton_export.clicked.connect(lambda: export_to_excel(self))
        
        # Combo box changes
        self.comboBox_weight.currentTextChanged.connect(self.update_weight_unit)
        self.comboBox_length.currentTextChanged.connect(self.update_length_unit)

        self.lineEdit_length_min.textChanged.connect(self.update_length_min)
        self.lineEdit_OD.textChanged.connect(self.update_OD)
        self.lineEdit_WT.textChanged.connect(self.update_WT)

        self.tableWidget_input.cellChanged.connect(self.on_cell_changed)

        self.lineEdit_length_min.setText(str(self.config["min_length"]))
        self.lineEdit_OD.setText(str(self.config["OD"]))
        self.lineEdit_WT.setText(str(self.config["WT"]))
        self.lineEdit_IP.setText(self.config["plc_ip"])

        self.lineEdit_weight.setText(f"{self.weight} {self.weight_unit}")
        self.lineEdit_length.setText(f"{self.length} {self.length_unit}")

        # Setup table columns
        self.setup_table()
        load_last_csv(self)

        # self.highlight_row_by_counter()
    
    def update_length_min(self):
        try:
            self.config["min_length"] = float(self.lineEdit_length_min.text())
            print("Updating min length", self.config["min_length"])
            self.save_config()
        except Exception as e:
            print(e)

    def update_OD(self):
        try:
            self.config["OD"] = float(self.lineEdit_OD.text())
            print("Updating OD", self.config["OD"])
            self.save_config()
        except Exception as e:
            print(e)

    def update_WT(self):
        try:
            self.config["WT"] = float(self.lineEdit_WT.text())
            print("Updating WT", self.config["WT"])
            self.save_config()
        except Exception as e:
            print(e)
    
    def highlight_row_by_counter(self):    
        row = self.length_counter

        item = self.tableWidget_home.item(row-1, 0)
        if item:
            item.setTextAlignment(Qt.AlignCenter)
            item.setBackground(QBrush(QColor("white")))

        # Now apply highlight and bold to the desired row
        item = self.tableWidget_home.item(row, 0)
        if item:
            item.setTextAlignment(Qt.AlignCenter)
            item.setBackground(QBrush(QColor(color_green)))  # gold color

    def on_cell_changed(self, row, column):
        # new_value = self.tableWidget_input.item(row, column).text()
        # print(f"Cell ({row}, {column}) changed to: {new_value}")

        values = []
        for col in range(self.tableWidget_input.columnCount()):
            item = self.tableWidget_input.item(row, col)
            text = item.text() if item else ""
            values.append(text)

        combined_text = f"{values[0]}          {values[1]}    [L]{self.length_unit}    [W]{self.weight_unit}   {values[2]}   {values[3]}   {values[4]}"

        # Update the first column (Printing Text) of tableWidget_home
        self.tableWidget_home.setItem(row, 0, QtWidgets.QTableWidgetItem(combined_text))
        item = self.tableWidget_home.item(row, 0)
        if item:
            item.setTextAlignment(Qt.AlignCenter)

    def update_weight_unit(self, text):
        """Update the weight unit based on combo box selection"""
        self.weight_unit = text.split("(")[1].strip(")")
        self.lineEdit_weight.setText(f"{self.weight} {self.weight_unit}")
    
    def update_length_unit(self, text):
        """Update the length unit based on combo box selection"""
        self.length_unit = text.split("(")[1].strip(")")
        self.lineEdit_length.setText(f"{self.length} {self.length_unit}")
            
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
                self.weight_status.setText("STATUS : ONLINE")
                self.weight_status.setStyleSheet("""
                    background-color: rgb(0, 170, 0);
                    border-radius: 5px;
                    border: none;
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
                self.length_status.setText("STATUS : ONLINE")
                self.length_status.setStyleSheet("""
                    background-color: rgb(0, 170, 0);
                    border-radius: 5px;
                    border: none;
                """)

                self.config["plc_ip"] = ip_address
                self.save_config()
            else:
                QMessageBox.critical(self, "Connection Failed", msg)

    def check_setup(self):
        if not self.EMARK.connected:
            QMessageBox.warning(self, "Printer Not Connected", "Please connect to the printer first")
            return
        if not self.WEIGHT.connected:
            QMessageBox.warning(self, "Weight Indicator Not Connected", "Please connect to Weight Indicator first")
            return
        if not self.PLC.connected:
            QMessageBox.warning(self, "PLC Not Connected", "Please connect to the PLC first")
            return
    
    def check_length(self, length):
        # Determine status
        status_length = "NORMAL"
        if length/1000 < self.config["min_length"]:
            status_length = "UNDERLENGTH"
        
        return status_length
    
    def check_weight(self, weight):
        thr_weight = (self.config["OD"] - self.config["WT"]) * self.config["WT"] * self.length/1000 * 0.02466
        self.config["min_weight"] = thr_weight - (thr_weight * 0.035)
        self.config["max_weight"] = thr_weight + (thr_weight * 0.035)
        print(f"WEIGHT = {self.weight} | THR = {thr_weight} | MIN = {self.config['min_weight']} | MAX = {self.config['max_weight']}")

        # Determine status
        status_weight = "NORMAL"
        if weight < self.config["min_weight"]:
            status_weight = "UNDERWEIGHT"
        elif weight > self.config["max_weight"]:
            status_weight = "OVERWEIGHT"

        return status_weight

    def printer(self, status):
        try:
            output_text = self.tableWidget_home.item(self.printer_counter, 0)
            length_text = self.tableWidget_home.item(self.printer_counter, 1)
            weight_text = self.tableWidget_home.item(self.printer_counter, 2)
            if output_text:
                if status == "NORMAL":
                    response = None
                    for i in range(6):
                        response = self.EMARK.send_text(output_text.text(), 
                                                template_num=1, 
                                                font_height=0x0C,  # 16 dot matrix
                                                x_pos=0, 
                                                y_pos=0)
                        if response:
                            print(f"Response: {response}")
                            break
                        else:
                            print(f"Retrying {i} : EMark Printer Not Connected or Error!")

                    if response:
                        print(f"Response: {response}")
                    else:
                        QMessageBox.warning(self, "Hardware Error", f"EMark Printer Not Connected or Error!")
                elif status == "REJECT":
                    self.EMARK.clear_text()

                # Add to history (whether printed or rejected)
                add_to_history(self, length_text.text().replace("\n", " "), weight_text.text().replace("\n", " "), output_text.text(), status)
            
        except Exception as e:
            print(e)
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    from PyQt5.QtGui import QPixmap
    from PyQt5.QtWidgets import QSplashScreen
    from PyQt5.QtCore import QTimer

    app = QtWidgets.QApplication(sys.argv)

    # Create splash screen
    splash_pix = QPixmap("lib/logo - Copy.jpg")  # <-- you can use any image here
    splash = QSplashScreen(splash_pix)
    splash.show()
    app.processEvents()  # Show the splash screen immediately

    window = PrintingSystem()
    window.show()
    splash.finish(window)

    sys.exit(app.exec_())
