#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import sys
import ttk
import json
import time
import win32print
import win32ui
import win32con
import qrcode 
import tkFileDialog
import Tkinter
import win32gui
import serial
import serial.tools.list_ports
import threading
from PIL import Image, ImageWin,ImageDraw,ImageFont

reload(sys) 
sys.setdefaultencoding('gb18030') 

mcuType = ''
rtAddrInfo = ''
nbAddrInf0 = ''
font1 = ImageFont.truetype('msyh.ttf', 16)
font2 = ImageFont.truetype('msyh.ttf', 20)
font3 = ImageFont.truetype('msyh.ttf', 30)
currentDate = time.strftime("%Y-%m-%d", time.localtime())

lock = threading.Lock()

class MSerialPort:
	message=''
	def __init__(self,port,buand):  
		self.port=serial.Serial(port,buand)  
		if not self.port.isOpen():  
			self.port.open()  
	def port_open(self):  
		if not self.port.isOpen():  
			self.port.open()  
	def port_close(self):  
		self.port.close()  
	def send_data(self,data):  
		number=self.port.write(data)  
		return number  
	def read_data(self):  
		while True:
			data=self.port.read(1)  #readline()
			self.message += data			
			if data == '\r':
				lock.acquire()  
				if len(self.message) < 20 :  # the length of  router's address is 18 bytes
					global rtAddrInfo
					rtAddrInfo = ''
					rtAddrInfo = self.message[:-1]
					self.message = ''
					macAdrInfo.set(rtAddrInfo)
				else :
					global nbAddrInf0
					nbAddrInf0 = ''
					#nbAddrInf0 = self.message[:-1]
					nbAddrInf0 = self.message[:15] + ':' + self.message[-20:-1]
					self.message = ''
					nbAdrInfo.set(nbAddrInf0)
				lock.release()

def printerList():
	printerList = []
	for printerItem in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1):
		flags, desc, name, comment = printerItem
		printerList.append(name.encode('utf-8'))
	return printerList

def serialList():
	serialList = []
	serialArray = list(serial.tools.list_ports.comports())
	for serialItem in serialArray:
	 	serialInfo = list(serialItem)
		serialName = serialInfo[0]
	 	serialList.append(serialName)

	return serialList

def sacnnerChosenEvent(event):
	global scannerObj
	scannerObj=MSerialPort(scannerChosen.get(),115200)  
	mSerialThread = threading.Thread(target=scannerObj.read_data, name='mSerialThread')
	mSerialThread.start()
	return


#devNo :zigbee网关序列号
#rtDevNo :路由器MAC地址
#nbDevNo :NB信息
def print2Printer(devNo,rtDevNo,nbDevNo):
	tempZgStringNo = devNo
	tempRtStringNo = rtDevNo
	tempNbStringNo = nbDevNo	

	#deviceString = '智能网关'.decode('utf-8')
	modelString = '型  号： '.decode('utf-8') + 'HA_1MGWA'
	ipString = 'IP地址：'.decode('utf-8') + '192.168.1.1'
	userString = '用户名：'.decode('utf-8') + 'root'
	powerString = '电源规格：'.decode('utf-8') + '12V--1A'

	zgSerialString= '序列号: '.decode('utf-8')  + tempZgStringNo
	rtSerialString= 'MAC地址: '.decode('utf-8') + tempRtStringNo
	companyString = '深圳市丰利源节能科技有限公司'.decode('utf-8')
	contactString= '联系电话：'.decode('utf-8')+'0755-28195888'
	produceString = '生产日期：'.decode('utf-8')+currentDate

	qr2 = qrcode.QRCode(version=1,  
                 error_correction=qrcode.constants.ERROR_CORRECT_L,  
                 box_size=10,  
                 border=1,  
                 )  
	qr2.add_data(devNo)  
	qr2.make(fit=True)  
	img2 = qr2.make_image() 
	img2 = img2.resize((80,80),Image.ANTIALIAS)

	nbQr = qrcode.QRCode(version=1,  
                 error_correction=qrcode.constants.ERROR_CORRECT_L,  
                 box_size=10,  
                 border=1,  
                 )  
	nbQr.add_data(nbDevNo)  
	nbQr.make(fit=True)  
	nbImg = nbQr.make_image() 
	nbImg = nbImg.resize((80,80),Image.ANTIALIAS)

	newImg1  = Image.new("RGB",(600,280),(255,255,255))

	imgSize1 = img2.size
	box = (0, 0, imgSize1[0], imgSize1[1])
	box1 = (30,200,110,280)
	region = img2.crop(box)
	newImg1.paste(region,box1)

	nbImgSize = nbImg.size
	nbImgbox = (0, 0, nbImgSize[0], nbImgSize[1])
	nbImgbox1 = (140,200,220,280)
	nbImgregion = nbImg.crop(nbImgbox)
	newImg1.paste(nbImgregion,nbImgbox1)


	a2 = ImageDraw.Draw(newImg1)
	a2.ink = 0 + 0 * 256 + 0 * 256 * 256
	a2.text((250,30),"合格证".decode('utf-8'),font=font3)
	#a2.text((150,20),companyString,font=font2)

	#a2.text((20,50),deviceString,font=font2)
	a2.text((20,80),modelString,font=font2)

	a2.text((20,110),ipString,font=font2)
	a2.text((20,140),userString,font=font2)
	a2.text((20,170),powerString,font=font2)

	a2.text((280,80),rtSerialString,font=font2)	
	a2.text((280,110),zgSerialString,font=font2)
	a2.text((280,140),contactString,font=font2)
	a2.text((280,170),'http://www.fly1992.com',font=font2)
	a2.text((280,200),produceString,font=font2)
	a2.text((280,230),companyString,font=font2)

	#newImg1.resize((600,280),Image.ANTIALIAS)
	newImg1.save('gwTest.png')

	hDC = win32ui.CreateDC ()
	hDC.CreatePrinterDC (printerChosen.get())
	hDC.StartDoc ('qrcode')
	hDC.StartPage ()
	hDC.SetMapMode (win32con.MM_TWIPS)

	dib1 = ImageWin.Dib (newImg1)
	dib1.draw (hDC.GetHandleOutput(), (10, 0, 4210, -1960))

	hDC.EndPage ()
	hDC.EndDoc ()
	hDC.DeleteDC ()

	return


def btnChooseFirmWare(event):
	filePath = tkFileDialog.askopenfilename(title=u"选择文件",initialdir=(os.path.expanduser('./')))
	fileInfo.set(filePath)
	root.mainloop()

	return	


def btnPrintDevInfo(event):

	btnState['text'] = 'WAIT'
	btnState['background'] = 'white'
	root.update()

	toolPath  = ".\commander"

	deviceName = 'EFR32' #mcuType

	cmdGetDevInfo = toolPath + '\commander.exe device info --device ' + deviceName
	devinfo=os.popen(cmdGetDevInfo)  	#popen与system可以执行指令,popen可以接受返回对象  
	devinfo=devinfo.read() 				#读取输出
	stateText.delete(0.0, Tkinter.END) 
	stateText.insert(Tkinter.END,devinfo) 

	lock.acquire()
	global rtAddrInfo
	global nbAddrInf0
	rtDevInfo = rtAddrInfo
	nbDevInfo = nbAddrInf0
	
	rtAddrInfo = ''
	nbAddrInf0 = ''
	macAdrInfo.set('')
	nbAdrInfo.set('')
	lock.release()

	if rtDevInfo=='':
		printInfo = 'Warning:please read the address of router.'
		stateText.insert(Tkinter.END,printInfo) 
		root.update()
		return

	if 	nbDevInfo == '':
		printInfo = 'Warning:please read the address of router.'
		stateText.insert(Tkinter.END,printInfo) 
		root.update()
		return


	if 'ERROR' in devinfo:
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.update()
	else :
		btnState['text'] = 'PASS'
		btnState['background'] = 'green'
		root.update()

		devId = devinfo[-22:-6]
		print2Printer(devId.upper(),rtDevInfo,nbDevInfo)

	root.mainloop()

	return


def btnUnlockDev(event):
	btnState['text'] = 'WAIT'
	btnState['background'] = 'white'
	root.update()

	toolPath  = ".\commander"
	#deviceName = mcuChosen.get()
	deviceName = 'EFR32'      #mcuType

	outputInfo = ''
	cmdUnlockDev = toolPath + '\commander.exe device lock --debug disable --device ' + deviceName
	unlockInfo=os.popen(cmdUnlockDev)  	#popen与system可以执行指令,popen可以接受返回对象  
	unlockInfo=unlockInfo.read() 			#读取输出 
	outputInfo = unlockInfo
	stateText.delete(0.0, Tkinter.END)
	stateText.insert(Tkinter.END,unlockInfo)

	if 'ERROR' in outputInfo:
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.update()
	else :
		btnState['text'] = 'PASS'
		btnState['background'] = 'green'
		root.update()

	root.mainloop()

	return	


def btnEraseDev(event):
	btnState['text'] = 'WAIT'
	btnState['background'] = 'white'
	root.update()

	toolPath  = ".\commander"
	
	#deviceName = mcuChosen.get()
	deviceName = 'EFR32'      #mcuType

	outputInfo = ''

	cmdEraseDev = toolPath + '\commander.exe device masserase ' + '--device ' + deviceName
	eraseInfo=os.popen(cmdEraseDev)  	#popen与system可以执行指令,popen可以接受返回对象  
	eraseInfo=eraseInfo.read() 			#读取输出 
	outputInfo = outputInfo+eraseInfo
	stateText.delete(0.0, Tkinter.END)
	stateText.insert(Tkinter.END,eraseInfo)

	if 'ERROR' in outputInfo:
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.update()
	else :
		btnState['text'] = 'PASS'
		btnState['background'] = 'green'
		root.update()

	root.mainloop()

	return	



def btnflashDev(event):
	btnState['text'] = 'WAIT'
	btnState['background'] = 'white'
	root.update()

	toolPath  = ".\commander"
	flashFilePath = firewareEntry.get()
	if flashFilePath[-3:] != 'hex' and flashFilePath[-3:] != 'bin' :
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.mainloop()
		return 

	#deviceName = mcuChosen.get()
	deviceName = 'EFR32'      #mcuType

	outputInfo = ''
	cmdFlashDev = toolPath + '\commander.exe flash ' + flashFilePath + ' --address 0x0 --device ' + deviceName
	flashInfo=os.popen(cmdFlashDev)  	#popen与system可以执行指令,popen可以接受返回对象  
	flashInfo=flashInfo.read() 			#读取输出 
	outputInfo = outputInfo+flashInfo
	stateText.delete(0.0, Tkinter.END)
	stateText.insert(Tkinter.END,flashInfo)


	if 'ERROR' in outputInfo:
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.update()
	else :
		btnState['text'] = 'PASS'
		btnState['background'] = 'green'
		root.update()

	root.mainloop()

	return	


def btnLockDev(event):
	btnState['text'] = 'WAIT'
	btnState['background'] = 'white'
	root.update()

	toolPath  = ".\commander"
	
	#deviceName = mcuChosen.get()
	deviceName = 'EFR32'      #mcuType

	outputInfo = ''
	cmdLockDev = toolPath + '\commander.exe device lock --debug enable --device ' + deviceName
	lockInfo=os.popen(cmdLockDev)  	#popen与system可以执行指令,popen可以接受返回对象  
	lockInfo=lockInfo.read() 			#读取输出 
	outputInfo = outputInfo+lockInfo
	stateText.delete(0.0, Tkinter.END)
	stateText.insert(Tkinter.END,lockInfo)

	if 'ERROR' in outputInfo:
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.update()
	else :
		btnState['text'] = 'PASS'
		btnState['background'] = 'green'
		root.update()

	root.mainloop()

	return	

def closeWindow():
	if scannerChosen.get():
		global scannerObj
		scannerObj.port_close()

	print("exit.")
	os._exit(0)
	#root.destroy()
	#sys.exit(0)

global root
root = Tkinter.Tk()                     # 创建窗口对象的背景色
root.title('条码工具')
root.geometry('800x600+500+300')
root.resizable(False, False)
#root.wm_attributes('-topmost',1)
root.attributes('-topmost',1)

printLabel=Tkinter.Label(root,text='请选择连接的打印机型号：', font=(10))
printLabel.place(x = 50, y = 30)

printerList=printerList()
printerType=Tkinter.StringVar()
printerChosen=ttk.Combobox(root, width=12, textvariable=printerType)
printerChosen['values']=printerList
printerChosen.place(x = 300, y = 30) 
defaultPrintName = win32print.GetDefaultPrinter().encode('utf-8')
printIndex = 0
for item in printerList:	
	if item == defaultPrintName:
		break

	printIndex = printIndex + 1	
printerChosen.current(printIndex)    					 	# 设置下拉列表默认显示的值


scannerLabel=Tkinter.Label(root,text='请选择连接的扫码枪接口：', font=(10))
scannerLabel.place(x = 50, y = 80) 

scannerList = serialList()
scannerType = Tkinter.StringVar()
scannerChosen = ttk.Combobox(root, width=12, textvariable=scannerType)
scannerChosen['values'] = scannerList
scannerChosen.place(x = 300, y = 80)
#scannerChosen.current(0)
scannerChosen.bind("<<ComboboxSelected>>",sacnnerChosenEvent)


macInfoLabel=Tkinter.Label(root,text='路由器MAC地址：',font=(10))
macInfoLabel.place(x = 50, y = 130)
macAdrInfo = Tkinter.StringVar()
macAdrEntry = Tkinter.Entry(root,textvariable=macAdrInfo)
macAdrEntry['state'] = 'readonly'
macAdrEntry.place(x = 180, y = 130)


nbInfoLabel=Tkinter.Label(root,text='NB信息：',font=(10))
nbInfoLabel.place(x = 350, y = 130)
nbAdrInfo = Tkinter.StringVar()
nbAdrEntry = Tkinter.Entry(root,textvariable=nbAdrInfo)
nbAdrEntry['state'] = 'readonly'
nbAdrEntry.place(x = 430, y = 130)

firewareLabel=Tkinter.Label(root,text='网关烧录的固件：',font=(10))
firewareLabel.place(x = 50, y = 180)
fileInfo = Tkinter.StringVar() 
#firewareEntry = Entry(root,text = '')
firewareEntry = Tkinter.Entry(root,textvariable=fileInfo,width = 35)
firewareEntry['state'] = 'readonly'
firewareEntry.place(x = 180, y = 180)
firewareBtn = Tkinter.Button(root,text = '选择固件',height = 1)
firewareBtn.place(x = 450, y = 175)
firewareBtn.bind("<Button-1>", btnChooseFirmWare)

stateText = Tkinter.Text(root, height=17, width=100)
stateText.place(x = 50, y = 230)

btnStateText = Tkinter.StringVar() 
btnState = Tkinter.Button(root,text='WAIT',width = 40 ,height = 6)
btnStateText.set('WAIT')
btnState.place(x = 50,y = 470)

btnPrint = Tkinter.Button(root,text = '打印',width = 15,height = 3)
btnPrint.place(x = 380,y = 475)
btnPrint.bind("<Button-1>", btnPrintDevInfo)

btnUnlock= Tkinter.Button(root,text = '解密',width = 8,height = 1)
btnUnlock.place(x = 530,y = 475)
btnUnlock.bind("<Button-1>", btnUnlockDev)

btnErase = Tkinter.Button(root,text = '擦除',width = 8,height = 1)
btnErase.place(x = 620,y = 475)
btnErase.bind("<Button-1>", btnEraseDev)

btnFlash = Tkinter.Button(root,text = '编程',width = 8,height = 1)
btnFlash.place(x = 530,y = 515)
btnFlash.bind("<Button-1>", btnflashDev)

btnLock = Tkinter.Button(root,text = '加密',width = 8,height = 1)
btnLock.place(x = 620,y = 515)
btnLock.bind("<Button-1>", btnLockDev)

try:
	root.protocol('WM_DELETE_WINDOW', closeWindow)
	root.mainloop()								# 进入消息循环
except Exception as e:
	os._exit(0)

