#coding:utf-8
import httplib
import time
import json
import ctypes
import os
from xlwt import *

class MPInterfaceTest:

	def __init__(self):
		print "完成初始化设置..."
		self.excelRow = 0 
		
	def interfacesTest(self):
		conn = httplib.HTTPConnection("c.miaopai.com")
		param = self.getTestParams()
		
		for i in param:
			print i[1] + "接口测试结果："
			t1 = time.time()
			if(i[2] == ""):
				res = conn.request(i[0],i[3])
				res = conn.getresponse()
				t2 = time.time()
				print "接口响应时间大致为：" + str('%.4f'%(t2 - t1))
				print "接口返回状态：" + str(res.status)
				print "请求头部信息："
				print res.getheaders()
				print "---------------------------------"
				print "\n"
			else:
				res = conn.request(i[0],i[3],i[2])
				res = conn.getresponse()
				rstatus = res.status
				if(rstatus == 200):
					print "服务器响应状态码：" + str(rstatus)
					t2 = time.time()
					duration = '%.4f'%(t2 - t1)
					print "接口响应时间大致为：" + str(duration)
					#print "接口返回状态：" + str(res.status)   #打印应答状态
					#print "请求头部信息："
					#print res.getheaders()      #打印请求头部信息
					print "---------------------------------"
					result = res.read()
					#print result
					hstatus = json.loads(result)['status']
					print "接口响应状态码：" + str(hstatus)
					if(hstatus == 200):
						print i[1] + "---接口正常"
					else:
						print i[1] + "---接口异常"
					if(duration > '%.4f'%1.5):
						print "接口响应慢"
					else:
						print "接口响应不慢"
					print "\n"
					if(self.excelRow == 0):
						self.createExecel()
						self.writeToExecel(i,rstatus,hstatus,duration)
					else:
						self.writeToExecel(i,rstatus,hstatus,duration)
					time.sleep(2)
			
	
	def getTestParams(self):
		params = [
		["GET","热门接口","header","http://api.miaopai.com/m/v6_hot_channel.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&refresh=2&sinaad=1&timestamp=1468311253119&pname=com.yixia.videoeditor&os=android&version=6.5.5&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=taobao&page=1&per=20"],
		["GET","排行接口","header","http://api.miaopai.com/m/v6_hot_channel.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&refresh=2&sinaad=1&timestamp=1468311253119&pname=com.yixia.videoeditor&os=android&version=6.5.5&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=taobao&page=1&per=20"],
		["GET","频道接口","header","http://api.miaopai.com/m/v6_hot_channel.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&refresh=2&sinaad=1&timestamp=1468311253119&pname=com.yixia.videoeditor&os=android&version=6.5.5&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=taobao&page=1&per=20"]
		]
		return params

	def createExecel(self):
		if os.path.exists("interfaceTestResult.xls"):
			print "interfaceTestResult.xls is exists"
		else:
			self.w = Workbook("utf-8")
			self.ws1 = self.w.add_sheet('sheet 1')
			
			font = Font()
			font.bold = True
			style = XFStyle()
			style.font = font
			self.ws1.write(0,0,'请求方法'.decode('gbk'),style)
			self.ws1.write(0,1,'服务器状态码'.decode('gbk'),style)
			self.ws1.write(0,2,'接口状态码'.decode('gbk'),style)
			self.ws1.write(0,3,'请求时间'.decode('gbk'),style)
			self.ws1.write(0,4,'接口描述'.decode('gbk'),style)
			self.ws1.write(0,5,'接口详情'.decode('gbk'),style)
			self.ws1.col(0).width = 3000
			self.ws1.col(1).width = 4000
			self.ws1.col(2).width = 3000
			self.ws1.col(3).width = 3000
			self.ws1.col(4).width = 5000
			self.ws1.col(5).width = 16000
			self.w.save('interfaceTestResult.xls')
			self.excelRow = self.excelRow + 1

	def writeToExecel(self,idata,rstatus,hstatus,duration):
		self.ws1.write(self.excelRow,0,idata[0])
		self.ws1.write(self.excelRow,1,rstatus)
		self.ws1.write(self.excelRow,2,hstatus)
		self.ws1.write(self.excelRow,3,duration)
		self.ws1.write(self.excelRow,4,idata[1].decode('gbk'))
		self.ws1.write(self.excelRow,5,idata[3])
		self.excelRow = self.excelRow + 1
		self.w.save('interfaceTestResult.xls')
				
print MPInterfaceTest().interfacesTest()