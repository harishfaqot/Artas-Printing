import serial
import time
import random

class WeightReader:
    def __init__(self):
        self.connected = False

    def connect(self, port, baudrate=9600):
        try:
            self.serial = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )
            print(f"Connected to {port}")
            self.connected = True
        except serial.SerialException as e:
            print(f"Failed to open serial port {port}: {e}")
            self.serial = None
            self.connected = False
            return f"Could not connect to device.\nPlease check the port or connection.\nError: {e}"

    def read_weight(self):
        if not hasattr(self, 'serial'):
            print("Serial port not available.")
            return 0
            # return 501.33

        try:
            self.serial.flushInput()
            self.serial.write(b'SI\r\n')  # Send SI command
            response = self.serial.readline().decode('utf-8').strip()
            print(f"IND 231 Raw response: {response}")

            if response.startswith("S S") or response.startswith("S D"):
                # Attempt to extract the numeric weight value
                parts = response.split()
                for part in parts:
                    try:
                        weight = float(part)
                        print(f"Weight: {weight}")
                        return weight
                    except ValueError:
                        continue

            elif response.startswith("S I"):
                print("Command understood but not executable at present.")
            elif response.startswith("S +"):
                print("Terminal in overload range.")
            elif response.startswith("S -"):
                print("Terminal in underload range.")
            else:
                print(f"Unknown response: {response}")
                # return 0
                # return 573.03

            # return response

        except serial.SerialException as e:
            print(f"Serial Weight error: {e}")
            self.connected = False
        except Exception as e:
            print(f"Error: {e}")
            self.connected = False


    def close(self):
        if self.serial and self.serial.is_open:
            self.serial.close()
            print("Serial port closed.")
            self.connected = False

# Example usage:
if __name__ == "__main__":
    scale = WeightReader()
    scale.connect(port = 'COM6')
    while True:
        print(scale.read_weight())

    scale.close()
