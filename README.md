# Modbus Registers Reader

Modbus Registers Reader is a Python application that allows users to connect to devices using the Modbus TCP/IP protocol and read data from various types of Modbus registers. The application supports reading coils, discrete inputs, input registers, and holding registers, and provides an intuitive user interface for easy configuration and usage.

## Features

- Connect to Modbus devices via Modbus TCP/IP.
- Read and display data from coils, discrete inputs, input registers, and holding registers.
- Filter and display non-zero register values.
- Export data to an Excel file for detailed analysis.
- User-friendly graphical interface for easy configuration and use.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Dependencies](#dependencies)
- [Project Structure](#project-structure)
- [Acknowledgements](#acknowledgements)
- [Contact](#contact)

## Installation

**Clone the repository:**

  ```sh
git clone https://github.com/alessandro-scalambrino/ModbusTCP_Memory_Scanner
```
**Install dependencies:**
```sh
cd ModbusTCP_Memory_Scanner
pip install -r requirements.txt
```


**Usage
Enter the IP address of the Modbus device.
Enter the port number (default is 502).
Specify the range of registers to read.
Start reading data:

Click on the "Read Registers" button.
View the data in the provided tables.
Export the data to an Excel file if needed.

Dependencies
Python 3.7+
pandas
pymodbus
tabulate
openpyxl
tkinter

Acknowledgements
The application uses the pymodbus library for Modbus communication.
The graphical interface is built using the tkinter library.
Contact
For any questions or suggestions, please contact alessandro.scalambrino.work@gmail.com.
