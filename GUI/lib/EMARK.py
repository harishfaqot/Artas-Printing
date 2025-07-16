import serial
import time
from functools import reduce
import random

CHAR_REMAP = {
    '~': "|",
    '`': '_',
    '@': '"',
    '^': '&',
    '&': "'",
    '*': '(',
    '(': ')',
    '_': '}',
    '-': '+',
    '+': ',',
    '{': '.',
    '[': ';',
    ']': '<',
    '}': '/',
    '\\': '>',
    ';': '@',
    ':': '?',
    "'": '&',
    ',': '\\',
    '.': '`',
    '<': ']',
    '>': '^',
    '/': '*',
    '?': '{'
}

class EMARKPrinter:
    def __init__(self):
        self.dest_addr = 0x01  # Default destination address
        self.src_addr = 0x00   # Default source address (PC)
        self.connected = False
    
    def connect(self, port, baudrate = 57600):
        """
        Initialize the printer connection
        :param port: Serial port name (e.g., 'COM3' or '/dev/ttyUSB0')
        :param baudrate: Baud rate (default 57600)
        """
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
        
        # if self.connected:
        #     self.clear_text()

    def clear_text(self):
        spaces = " " * 255
        for i in range(10):
            response = self.send_text(spaces, 
                                    template_num=1, 
                                    font_height=0x10,  # 16 dot matrix
                                    x_pos=i, 
                                    y_pos=0,
                                    char_spacing=i)
            print(f"Response: {response.hex()}")
        
    def calculate_checksum(self, data):
        """Calculate XOR checksum for the data"""
        return reduce(lambda x, y: x ^ y, data)
    
    def send_command(self, command_word, frame_data):
        """
        Send a command to the printer
        :param command_word: Command byte
        :param frame_data: Data bytes (after frame length)
        :return: Response from printer
        """
        # Build the frame
        frame_length = len(frame_data) + 1
        frame = bytes([
            self.dest_addr,
            self.src_addr,
            command_word,
            (frame_length >> 8) & 0xFF,  # High byte of frame length
            frame_length & 0xFF          # Low byte of frame length
        ]) + frame_data
        
        # Add checksum
        checksum = self.calculate_checksum(frame)
        frame += bytes([checksum])
        
        # Send the frame
        self.serial.write(frame)
        
        # Wait for response (200ms timeout as per protocol)
        time.sleep(0.2)
        response = self.serial.read_all()
        
        return response
    
    
    def remap_special_chars(self, text):
        return ''.join(CHAR_REMAP.get(c, c) for c in text)
    
    def send_text(self, text, template_num=1, font_height=0x10, x_pos=0, y_pos=0, char_spacing=5):
        """
        Send text to be printed (Command D3H - Transmit fixed information)
        :param text: Text to print
        :param template_num: Template number (file number in printer)
        :param font_height: Font height code (0x08=5x6, 0x10=16 dot matrix, etc.)
        :param x_pos: X position in lines (width)
        :param y_pos: Y position (0=top)
        :param char_spacing: Spacing between characters
        """
        # Convert text to ASCII/GB2312 bytes
        try:
            # First try ASCII encoding
            # text = "1ST API 5CT+2221 LOGO 05+25 PE 7 26`00 K S P 4600 PSI D   40`30 FT 1037 LBS HN  241B11000+1  WO 04+0475"
            print("Before:",text)
            text = self.remap_special_chars(text)
            print("After :",text)
            data_bytes = text.encode('ascii')
            print(data_bytes.hex(' '))
        except UnicodeEncodeError:
            print("ERROR")
            try:
                # If ASCII fails, try GB2312 for Chinese characters
                data_bytes = text.encode('gb2312')
            except:
                # Fallback to ASCII with replacement for unsupported chars
                data_bytes = text.encode('ascii', errors='replace')
        
        # Build font attributes (6 bytes)
        font_attrs = bytes([
            font_height,                # Font height
            len(text),                  # Font length (number of characters)
            (x_pos >> 8) & 0xFF,       # X position high byte
            x_pos & 0xFF,               # X position low byte
            y_pos,                      # Y position
            char_spacing                # Character spacing
        ])
        
        # Build the information structure
        info = bytes([template_num]) + font_attrs + data_bytes
        
        # Send the command
        response = self.send_command(0xD3, info)
        
        return response
    
    def reset_current_template(self):
        """Reset template to empty state"""
        empty_template = bytes([
            0x01,       # Template number
            0x10,       # Font height (16-dot)
            0x00,       # Character count
            0x00, 0x00, # X position
            0x00,       # Y position
            0x00        # Spacing
        ]) + b''        # Empty data
        self.send_command(0xD3, empty_template)
    
    def turn_on_printing(self, on=True):
        """Turn printing on/off (Command A5H)"""
        return self.send_command(0xA5, bytes([0x01 if on else 0x00]))
    
    def set_printing_speed(self, speed):
        """Set printing speed (Command A1H)"""
        if speed < 0 or speed > 255:
            raise ValueError("Speed must be between 0 and 255")
        return self.send_command(0xA1, bytes([speed]))
    
    def close(self):
        """Close the serial connection"""
        self.serial.close()
        print("PRINTER COM CLOSED")
        self.connected = False

# Example usage
if __name__ == "__main__":
    printer = EMARKPrinter(port='COM2')
    try:
        # Initialize printer - change port to your actual serial port
        
        
        # Turn on printing
        # printer.turn_on_printing(True)
        
        # Send text to print
        # English text
        
        response = printer.send_text("1ST API 5CT-2221 LOGO 05-25 PE 7 26-00 K S P 4600 PSI D   402.1 FT 1037 LBS HN  241B11000-1  WO 04-0475", 
                                   template_num=1, 
                                   font_height=0x10,  # 16 dot matrix
                                   x_pos=0, 
                                   y_pos=0)
        print(f"Response: {response.hex()}")

        time.sleep(5)

        spaces = " " * 255
        for i in range(10):
            response = printer.send_text(spaces, 
                                    template_num=1, 
                                    font_height=0x10,  # 16 dot matrix
                                    x_pos=i, 
                                    y_pos=0,
                                    char_spacing=i)
            print(f"Response: {response.hex()}")
        
    except Exception as e:
        print(f"Error: {e}")
    finally:
        printer.close()