# LEFOO Differential Pressure Transmitter(LFM52-6-O-E) data reading via RS485(MODBUS)
import serial
import struct
import crcmod
import time

# make serial connection
ser = serial.Serial(port='COM6', baudrate=19200,
                    parity='E', stopbits=1, bytesize=8)

# except when KeyboardInterrupt
try:

    start_time = time.time()
    avg = []
    while True:
        # Calculate the CRC16 value for the Modbus RTU message
        crc16 = crcmod.predefined.Crc('modbus')

        # initiate Tx codes
        machine_addr = b'\x01'
        function_code = b'\x03'
        data_start_addr = b'\x00\x02'
        data_Nos = b'\x00\x01'
        data_send = machine_addr + function_code + data_start_addr + data_Nos

        # print(data_send)
        crc16.update(data_send)
        crc_bytes = crc16.digest()

        # Reverse the byte order of the CRC16 bytes
        crc_bytes_reversed = crc_bytes[::-1]
        # Append the CRC16 value to the Modbus RTU message
        data = data_send + crc_bytes_reversed
        # Send the Modbus RTU request
        ser.write(data)
        # Wait for a response
        response = ser.read(7)

        # print(f" Tx: {data}, Rx: {response}") # check Hex data of Tx and Rx

        # Extract the relevant fields
        _, _, _, value, crc_code = struct.unpack('>BBBhH', response)
        avg.append(value)
        time.sleep(0.001)

        # Smoothing data by averaging samples taken every 2 seconds
        elapsed_time = time.time() - start_time
        if elapsed_time >= 2:
            avg_val = sum(avg) / len(avg)
            print(f"{avg_val: .1f}")
            start_time = time.time()

# when terminating
except KeyboardInterrupt:
    # serial connection closing
    ser.close()
    print("stopped")
