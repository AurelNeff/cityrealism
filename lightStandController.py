# Author: Aurel Neff, small parts of the code come from publicly available online sources
# License: Public Domain
import time
from Adafruit_GPIO import I2C
tca = I2C.get_i2c_device(address=0x70)
import RPi.GPIO as GPIO
import xlsxwriter
import Adafruit_TCS34725
import smbus

kP = 0.055
kI = 0.00005
kD = 0.00005

#GPIO Settings, set PWM output Pins
GPIO.setmode(GPIO.BCM)  # choose BCM or BOARD numbering schemes
GPIO.setup(17, GPIO.OUT)
GPIO.setup(27, GPIO.OUT)
GPIO.setup(22, GPIO.OUT)
GPIO.setup(23, GPIO.OUT)
GPIO.setup(24, GPIO.OUT)
GPIO.setup(18, GPIO.OUT)

led2 = GPIO.PWM(17, 100)    # create object led tile one for PWM on port 17 at 100 H$
led6 = GPIO.PWM(22, 100)
led3 = GPIO.PWM(27, 100)
led7 = GPIO.PWM(23,100)
led5 = GPIO.PWM(24,100)
led4 = GPIO.PWM(18,100)
#start leds
led2.start(0)
led3.start(0)
led4.start(0)
led5.start(0)
led6.start(0)
led7.start(0)
#Name file for data registration
a=input('Name file')
workbook = xlsxwriter.Workbook('Measurement Lightstand V{}.xlsx'.format(a))
worksheet = workbook.add_worksheet()

#using channels 2-7 of the Multiplexer
def tca_select(channel):
    """Select an individual channel."""
    if channel > 7:
        return
    tca.writeRaw8(1 << channel)

def tca_set(mask):
    """Select one or more channels.
           chan =   76543210
           mask = 0b00000000
    """
    if mask > 0xff:
        return
    tca.writeRaw8(mask)
#Use this code to deactivate the LEDs on the sensor. Not needed if Pin is soldered directly to ground.
#while x < 8:
#	tca_select(x)
#       tcs = Adafruit_TCS34725.TCS34725(integration_time=Adafruit_TCS34725.TCS34725_INTEGRATIONTIME_50MS, gain=Adafruit_TCS34725.TCS34725_GAIN_4X)
#       tcs.set_interrupt(True)
#	x=x+1
# Registration of data, initialization:
worksheet.write(0,0,'Timestamp')
worksheet.write(0,1,'Tile 2 Lux output')
worksheet.write(0,7,'Tile 2 measurement')
worksheet.write(0,2,'Tile 3 Lux output')
worksheet.write(0,8,'Tile 3 measurement')
worksheet.write(0,3,'Tile 4 Lux output')
worksheet.write(0,9,'Tile 4 measurement')
worksheet.write(0,4,'Tile 5 Lux output')
worksheet.write(0,10,'Tile 5 measurement')
worksheet.write(0,5,'Tile 6 Lux output')
worksheet.write(0,11,'Tile 6 measurement')
worksheet.write(0,6,'Tile 7 Lux output')
worksheet.write(0,12,'Tile 7 measurement')

lightinput = 0.0
#PID Controller starts after this line
#LEDone=2;two=6;three=3,four=7;five=5;six=4;
count = 0
t = time.time()
t_prev = [t,t,t,t,t,t,t,t]
ui_prev = [0,0,0,0,0,0,0,0]
ew_prev = [0,0,0,0,0,0,0,0]
lightinput_prev = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]
#parameter results from sensor calibration:
deslux = [0,0,138.6,166.6,95.4,204.4,163.7,200.6]
try:
# Stop the program after 10'000 measurements, just for safety.
	while count<10000:
		x = 2
		count = count + 1
#		raw_input('Press enter to continue: ')
#		a=input("Press Enter to continue...")
#		k = int(time.time()
#	        k=datetime.now().strftime("%A, %d. %B %Y %I:%M%p")
		worksheet.write(count,0,count)
		while x < 8:
			tca_select(x)
			tcs = Adafruit_TCS34725.TCS34725(integration_time=Adafruit_TCS34725.TCS34725_INTEGRATIONTIME_50MS, gain=Adafruit_TCS34725.TCS34725_GAIN_4X)
# Read R, G, B, C color data from the sensor.
			r, g, b, c = tcs.get_raw_data()
# Calculate lux out of RGB measurements.
			lux = Adafruit_TCS34725.calculate_lux(r, g, b)
#start of control
			t = time.time()
			timeChange=t-t_prev[x]
			t_prev[x]=t
#Calculate the error
	                ew = deslux[x] - lux
	                ui_white = ui_prev[x] + ew*timeChange
	                ud_white = (ew - ew_prev[x]) / timeChange
	                ew_prev[x] = ew
	                ui_prev[x] = ui_white
	                lightinput = lightinput_prev[x]+kP*ew + kI*ui_white + ud_white*kD
			lightinput_prev[x] = lightinput
#set limits to the output. PWM is limited from 0 to 100, but I capped it at 15. This number is coming from the calibration.
	                if lightinput>15:
	                        lightinput=15.0
	                if lightinput<0:
	                        lightinput=0.0
			if x==2:
	                        led2.ChangeDutyCycle(lightinput)
                        if x==3:
                                led3.ChangeDutyCycle(lightinput)
                        if x==4:
                                led4.ChangeDutyCycle(lightinput)
                        if x==5:
                                led5.ChangeDutyCycle(lightinput)
                        if x==6:
                                led6.ChangeDutyCycle(lightinput)
                        if x==7:
                                led7.ChangeDutyCycle(lightinput)

# Print out the values.
			print(str(count)+",Sensor: "+str(x)+",Error:"+str(ew_prev[x])+",Input: "+str(lightinput))
                        worksheet.write(count,x-1,lux)
                        worksheet.write(count,x+5,lightinput)
			x=x+1
# Close the workbook, stop LEDs
	workbook.close()
	led2.stop()
	led3.stop()
	led4.stop()
	led5.stop()
	led6.stop()
	led7.stop()

except KeyboardInterrupt:
    led2.stop()
    led3.stop()
    led4.stop()
    led5.stop()
    led6.stop()
    led7.stop()
    workbook.close()
    tcs.disable()
