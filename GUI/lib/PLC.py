import snap7
from snap7.util import get_real

class PLCReader:
    def __init__(self):
        self.connected = False  # <-- default to False

    def connect(self, ip, rack=0, slot=2):
        self.client = snap7.client.Client()
        try:
            self.client.connect(ip, rack, slot)
            if self.client.get_connected():
                print(f"Connected to PLC at {ip}")
                self.connected = True
            else:
                print(f"Failed to connect to PLC at {ip}")
        except Exception as e:
            print(f"Connection error: {e}")
            self.client = None
            self.connected = False
            return f"Could not connect to device.\nPlease check the IP or connection.\nError: {e}"

    def read_real(self, db_number, start_byte):
        if not self.client or not self.client.get_connected():
            print("Not connected to PLC.")
            return None
        try:
            data = self.client.db_read(db_number, start_byte, 4)
            value = get_real(data, 0)
            print(f"Value at DB{db_number}.DBD{start_byte}:", value)
            return value
        except Exception as e:
            print(f"Read error: {e}")
            return None

    def close(self):
        if self.client and self.client.get_connected():
            self.client.disconnect()
            print("Disconnected from PLC.")
            self.connected = False

# Example usage
if __name__ == "__main__":
    plc = PLCReader()
    plc.connect(ip='192.168.0.97')
    plc.read_real(db_number=2, start_byte=0)
    plc.close()
