import serial
import time

class WeightReader:
    def __init__(self):
        self.connected = False

    def connect(self, port, baudrate=9600):
        try:
            self.ser = serial.Serial(
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
            self.ser = None
            self.connected = False
            return f"Could not connect to device.\nPlease check the port or connection.\nError: {e}"

    def read_weight(self):
        if not self.ser or not self.ser.is_open:
            print("Serial port not available.")
            return None

        try:
            self.ser.flushInput()
            self.ser.write(b'SI\r\n')  # Send SI command
            response = self.ser.readline().decode('utf-8').strip()
            print(f"Raw response: {response}")

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

            return response

        except serial.SerialException as e:
            print(f"Serial error: {e}")
        except Exception as e:
            print(f"Error: {e}")


    def close(self):
        if self.ser and self.ser.is_open:
            self.ser.close()
            print("Serial port closed.")
            self.connected = False

# Example usage:
if __name__ == "__main__":
    scale = WeightReader()
    scale.connect(port = 'COM5')
    while True:
        print(scale.read_weight())

    scale.close()
