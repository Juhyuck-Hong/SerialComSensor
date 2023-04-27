# for SY-LD213 PWM signal generator frequency and duty set and confim
import time
import serial

freq = b"F25.0" # range: F001 ~ F1.5.0
duty_cycle = int(input("Enter duty cycle value (0-100): "))
duty_str = "D{}".format(str(duty_cycle).zfill(3))
duty = duty_str.encode("UTF-8")
freq_set = freq.decode().split("F")[1]
duty_set = duty.decode().split("D")[1]

def setSerial(value):
    ser = serial.Serial(
        port='/dev/ttyS0',
        baudrate = 9600,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.EIGHTBITS,
        timeout=1
        )
    time.sleep(0.1)
    ser.write(value)
    time.sleep(0.1)
    ser.close()

setSerial(duty)

try:
    while True:
        ser = serial.Serial(
            port='/dev/ttyS0',
            baudrate = 9600,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            bytesize=serial.EIGHTBITS,
            timeout=1
        )

        ser.write(b"READ")
        time.sleep(0.1)
        try:
            res = ser.read(12)
            if res[0] == 70:
                res = res.decode()
                freq_res = res.split("D")[0].split("F")[1].strip()
                duty_res = res.split(freq_set)[1].split("D")[1].strip()
                print(f"Frequency Set: {freq_set}, Read: {freq_res}")
                print(f"DutyCycle Set: {duty_set}, Read: {duty_res}")
                break
        except IndexError:
            pass
        ser.close()
except KeyboardInterrupt:
    ser.close()