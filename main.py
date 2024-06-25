import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from tabulate import tabulate
from pymodbus.client.sync import ModbusTcpClient

class ModbusReader:
    def __init__(self, host='localhost', port=502):
        self.client = ModbusTcpClient(host, port)
        self.client.connect()

    def read_nonzero_registers(self, register_type, start_address, num_registers, step_size=127):
        current_address = start_address
        remaining_registers = num_registers
        data = []

        while remaining_registers > 0:
            read_size = min(step_size, remaining_registers)
            if register_type == 'coils':
                response = self.client.read_coils(current_address, read_size)
                values = response.bits
            elif register_type == 'discrete_inputs':
                response = self.client.read_discrete_inputs(current_address, read_size)
                values = response.bits
            elif register_type == 'input_registers':
                response = self.client.read_input_registers(current_address, read_size)
                values = response.registers
            elif register_type == 'holding_registers':
                response = self.client.read_holding_registers(current_address, read_size)
                values = response.registers
            else:
                raise ValueError("Invalid register type")

            if not response.isError():
                for i, value in enumerate(values):
                    if value != 0:
                        data.append((current_address + i, value))
            else:
                print(f"Error reading {register_type} from {current_address} to {current_address + read_size - 1}")
                break

            current_address += read_size
            remaining_registers -= read_size

        return data

    def close_connection(self):
        self.client.close()

class ModbusGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Modbus Registers Reader")

        self.host_label = ttk.Label(root, text="Modbus IP Address:")
        self.host_label.grid(row=0, column=0, padx=5, pady=5)
        self.host_entry = ttk.Entry(root, width=20)
        self.host_entry.insert(0, 'localhost')  # Default value
        self.host_entry.grid(row=0, column=1, padx=5, pady=5)

        self.port_label = ttk.Label(root, text="Modbus Port:")
        self.port_label.grid(row=0, column=2, padx=5, pady=5)
        self.port_entry = ttk.Entry(root, width=8)
        self.port_entry.insert(0, '502')  # Default value
        self.port_entry.grid(row=0, column=3, padx=5, pady=5)

        self.start_label = ttk.Label(root, text="Start Address:")
        self.start_label.grid(row=1, column=0, padx=5, pady=5)
        self.start_entry = ttk.Entry(root, width=10)
        self.start_entry.grid(row=1, column=1, padx=5, pady=5)

        self.num_label = ttk.Label(root, text="Number of Registers:")
        self.num_label.grid(row=1, column=2, padx=5, pady=5)
        self.num_entry = ttk.Entry(root, width=10)
        self.num_entry.grid(row=1, column=3, padx=5, pady=5)

        self.read_button = ttk.Button(root, text="Read Registers", command=self.read_registers)
        self.read_button.grid(row=2, column=0, columnspan=4, padx=5, pady=10)

        self.result_label = ttk.Label(root, text="Results:")
        self.result_label.grid(row=3, column=0, padx=5, pady=5, sticky='w')

        self.result_text = tk.Text(root, height=10, width=60)
        self.result_text.grid(row=4, column=0, columnspan=4, padx=5, pady=5)

        self.export_button = ttk.Button(root, text="Export to Excel", command=self.export_to_excel)
        self.export_button.grid(row=5, column=0, columnspan=4, padx=5, pady=10)

        # Modbus reader instance
        self.modbus_reader = None

        # Data storage
        self.coils_df = None
        self.discrete_inputs_df = None
        self.input_registers_df = None
        self.holding_registers_df = None

    def read_registers(self):
        try:
            host = self.host_entry.get()
            port = int(self.port_entry.get())
            start_address = int(self.start_entry.get())
            num_registers = int(self.num_entry.get())

            # Update Modbus reader instance with new host and port
            if self.modbus_reader:
                self.modbus_reader.close_connection()
            self.modbus_reader = ModbusReader(host, port)

            # Read non-zero registers
            coils_data = self.modbus_reader.read_nonzero_registers('coils', start_address, num_registers)
            discrete_inputs_data = self.modbus_reader.read_nonzero_registers('discrete_inputs', start_address, num_registers)
            input_registers_data = self.modbus_reader.read_nonzero_registers('input_registers', start_address, num_registers)
            holding_registers_data = self.modbus_reader.read_nonzero_registers('holding_registers', start_address, num_registers)

            # Create DataFrames
            self.coils_df = pd.DataFrame(coils_data, columns=['Address', 'Value'])
            self.discrete_inputs_df = pd.DataFrame(discrete_inputs_data, columns=['Address', 'Value'])
            self.input_registers_df = pd.DataFrame(input_registers_data, columns=['Address', 'Value'])
            self.holding_registers_df = pd.DataFrame(holding_registers_data, columns=['Address', 'Value'])

            # Format results as tabular text
            coils_text = f"Coils:\n{tabulate(self.coils_df, headers='keys', tablefmt='grid')}\n\n"
            discrete_inputs_text = f"Discrete Inputs:\n{tabulate(self.discrete_inputs_df, headers='keys', tablefmt='grid')}\n\n"
            input_registers_text = f"Input Registers:\n{tabulate(self.input_registers_df, headers='keys', tablefmt='grid')}\n\n"
            holding_registers_text = f"Holding Registers:\n{tabulate(self.holding_registers_df, headers='keys', tablefmt='grid')}\n\n"

            # Display results in the text box
            self.result_text.delete('1.0', tk.END)
            self.result_text.insert(tk.END, coils_text)
            self.result_text.insert(tk.END, discrete_inputs_text)
            self.result_text.insert(tk.END, input_registers_text)
            self.result_text.insert(tk.END, holding_registers_text)

        except ValueError as ve:
            messagebox.showerror("Error", f"Invalid input: {ve}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def export_to_excel(self):
        try:
            # Check if data is available
            if not self.coils_df.empty or not self.discrete_inputs_df.empty or not self.input_registers_df.empty or not self.holding_registers_df.empty:
                # Export data to Excel file
                with pd.ExcelWriter('modbus_registers.xlsx') as writer:
                    if not self.coils_df.empty:
                        self.coils_df.to_excel(writer, sheet_name='Coils', index=False)
                    if not self.discrete_inputs_df.empty:
                        self.discrete_inputs_df.to_excel(writer, sheet_name='Discrete Inputs', index=False)
                    if not self.input_registers_df.empty:
                        self.input_registers_df.to_excel(writer, sheet_name='Input Registers', index=False)
                    if not self.holding_registers_df.empty:
                        self.holding_registers_df.to_excel(writer, sheet_name='Holding Registers', index=False)

                messagebox.showinfo("Success", "Data has been exported to 'modbus_registers.xlsx'")
            else:
                messagebox.showwarning("Warning", "No data to export")

        except ValueError as ve:
            messagebox.showerror("Error", f"Invalid input: {ve}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def close_connection(self):
        if self.modbus_reader:
            self.modbus_reader.close_connection()

if __name__ == "__main__":
    root = tk.Tk()
    app = ModbusGUI(root)
    root.mainloop()
