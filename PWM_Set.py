# for SY-LD213 PWM signal generator frequency and duty set and confim
import time
import serial

# The frequency is not changed frequently,
# so the value is directly inputted into the variable.
freq = b"F25.0"  # range: F001 ~ F1.5.0

# The duty cycle is received from the user as input.
duty_cycle = int(input("Enter duty cycle value (0-100): "))
# Change into form of "DXXX"
duty_str = "D{}".format(str(duty_cycle).zfill(3))
# Make it byte
duty = duty_str.encode("UTF-8")

# Extract only the value of frequency and duty
freq_set = freq.decode().split("F")[1]
duty_set = duty.decode().split("D")[1]


def setSerial(value):
    # Function to send value to device

    # Create a serial port object
    ser = serial.Serial(
        port='/dev/ttyS0',              # Port: '/dev/ttyS0'
        baudrate=9600,                  # Baud rate: 9600
        parity=serial.PARITY_NONE,      # Parity bit: None
        stopbits=serial.STOPBITS_ONE,   # Stop bit: 1
        bytesize=serial.EIGHTBITS,      # Data bit: 8
        timeout=1                       # Timeout: 1 second
    )

    time.sleep(0.1)
    # Send the input value via serial communication
    ser.write(value)
    time.sleep(0.1)
    # Close the serial port connection
    ser.close()


def freqShow(value):
    # Function to change notation of frequency
    split_length = len(value.split('.'))
    if split_length == 1:
        return f"{value} Hz"
    elif split_length == 2:
        return f"{int(value)} kHz"
    else:
        return f"{int(value.replace(".", ""))} kHz"


if __name__ == "__main__":

    # Send duty cycle to device
    setSerial(duty)

    # Check the set value to confirm
    try:
        while True:
            # Initialize the serial connection
            ser = serial.Serial(
                port='/dev/ttyS0',
                baudrate=9600,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                bytesize=serial.EIGHTBITS,
                timeout=1
            )
            # Send the command "READ" to the device to request data
            ser.write(b"READ")
            time.sleep(0.1)
            try:
                # Read 12 bytes of data from the device
                res = ser.read(12)
                # Check if the first byte of the response is equal to 70 (hexadecimal value of "F")
                if res[0] == 70:
                    res = res.decode()
                    # Extract the frequency value from the response string
                    freq_res = res.split("D")[0].split("F")[1].strip()
                    # Extract the duty cycle value from the response string
                    duty_res = res.split(freq_set)[1].split("D")[1].strip()
                    # Print the frequency and duty cycle values that were set and read
                    print(
                        f"Frequency Set: {freqShow(freq_set)}, Read: {freqShow(freq_res)}")
                    print(
                        f"DutyCycle Set: {int(duty_set)} %, Read: {int(duty_res)} %")
                    # Exit the while loop
                    break
            # If response from the device does not contain expected values,
            # pass over this iteration of the while loop
            except IndexError:
                pass
            # Close the serial connection with the device
            ser.close()
    # If a KeyboardInterrupt occurs(Ctrl + C)
    except KeyboardInterrupt:
        # Close the serial connection with the device
        ser.close()
