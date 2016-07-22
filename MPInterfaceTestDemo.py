#coding:utf-8
import httplib
import time
import json
import ctypes
import os
from xlwt import *

class MPInterfaceTest:

	def __init__(self):
		print "��ɳ�ʼ������..."
		self.excelRow = 0 
		
	def interfacesTest(self):
		conn = httplib.HTTPConnection("c.miaopai.com")
		param = self.getTestParams()
		
		for i in param:
			print i[1] + "�ӿڲ��Խ����"
			t1 = time.time()
			if(i[2] == ""):
				res = conn.request(i[0],i[3])
				res = conn.getresponse()
				t2 = time.time()
				print "�ӿ���Ӧʱ�����Ϊ��" + str('%.4f'%(t2 - t1))
				print "�ӿڷ���״̬��" + str(res.status)
				print "����ͷ����Ϣ��"
				print res.getheaders()
				print "---------------------------------"
				print "\n"
			else:
				res = conn.request(i[0],i[3],i[2])
				res = conn.getresponse()
				rstatus = res.status
				if(rstatus == 200):
					print "��������Ӧ״̬�룺" + str(rstatus)
					t2 = time.time()
					duration = '%.4f'%(t2 - t1)
					print "�ӿ���Ӧʱ�����Ϊ��" + str(duration)
					#print "�ӿڷ���״̬��" + str(res.status)   #��ӡӦ��״̬
					#print "����ͷ����Ϣ��"
					#print res.getheaders()      #��ӡ����ͷ����Ϣ
					print "---------------------------------"
					result = res.read()
					#print result
					hstatus = json.loads(result)['status']
					print "�ӿ���Ӧ״̬�룺" + str(hstatus)
					if(hstatus == 200):
						print i[1] + "---�ӿ�����"
					else:
						print i[1] + "---�ӿ��쳣"
					if(duration > '%.4f'%1.5):
						print "�ӿ���Ӧ��"
					else:
						print "�ӿ���Ӧ����"
					print "\n"
					if(self.excelRow == 0):
						self.createExecel()
						self.writeToExecel(i,rstatus,hstatus,duration)
					else:
						self.writeToExecel(i,rstatus,hstatus,duration)
					time.sleep(2)
			
	
	def getTestParams(self):
		params = [
		["GET","���Žӿ�","header","http://api.miaopai.com/m/v6_hot_channel.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&refresh=2&sinaad=1&timestamp=1468311253119&pname=com.yixia.videoeditor&os=android&version=6.5.5&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=taobao&page=1&per=20"],
		["GET","δ��½״̬��עҳ��ˢ��","header","http://api.miaopai.com/m/recommend_follows_unlogin.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&size=50&token=&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&refresh=0&timestamp=1469156142672&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&os=android"],
		["GET","�Ȱ�ҳ��","header","http://api.miaopai.com/m/ascent_channel.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156241409&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=1&os=android&per=20"],
		["GET","����ҳ��banner�ӿ�","header","http://api.miaopai.com/m/discovery_summary.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156296061&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&os=android"],
		["GET","����ҳ�滰��ӿ�","header","http://api.miaopai.com/m/discovery_topic.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156296106&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=1&os=android&per=20"],
		["GET","����ҳ�����ͽӿ�","header","http://api.miaopai.com/m/getRewardList.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&pageSize=10&pageIndex=1&timestamp=1469156296111&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&orderType=1&channel=yixia&os=android"],
		["GET","׷��ҳ��","header","http://api.miaopai.com/m/cate_talent.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&timestamp=1469156461602&os=android&chanStat=20&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=1&per=20&cateId=124"],
		#["GET","����ҳ��-1","header","http://api.miaopai.com/m/reccategory.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156502587&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=1&os=android&per=4"],
		#["GET","����ҳ��-2","header","http://api.miaopai.com/m/cate_talent.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&timestamp=1469156502970&os=android&chanStat=20&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=1&per=20&cateId=132"],
		#["GET","����ҳ��-3","header","http://api.miaopai.com/m/cate_talent.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&timestamp=1469156503007&os=android&chanStat=20&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=1&per=20&cateId=172"],
		["GET","΢����½֮������ӿ�-1","header","http://api.miaopai.com/m/v4_remind.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156670791&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&os=android"],
		["GET","΢����½֮������ӿ�-2","header","http://api.miaopai.com/m/v2_sinafriends.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156681658&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&os=android"],
		["GET","�������ӿ�","header","http://api.miaopai.com/m/ad.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156823268&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&os=android"],
		["GET","δ֪�ӿ�-1","header","http://api.miaopai.com/m/ver.json?device=android&deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&oem=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&timestamp=1469156825366&os=android&network=WIFI&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&text=versionName%3D6.5.7%2Coem%3Dnull%2Cmodel%3DATH-AL00%2CdeviceId%3D14a9e93cd6d1f62c%2Cuuid%3De3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d%2Cmanufacturer%3DHUAWEI%2CreleaseVersion%3D5.1.1%2Cresolution%3D1080x1776%2CsdkVersion%3D22&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia"],
		["GET","δ֪�ӿ�-2","header","http://api.miaopai.com/m/v6_recommend_theme.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156825556&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&type=28&os=android"],
		["GET","δ֪�ӿ�-3","header","http://api.miaopai.com/m/v5_randfriends.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156826032&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&os=android"],
		["GET","δ֪�ӿ�-4","header","http://api.miaopai.com/m/samecity_v6.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&orderby=bscore&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156826160&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=1&os=android&per=20"],
		["GET","���ѽӿ�","header","http://api.miaopai.com/m/hotword.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469156826180&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&os=android"],
		["GET","�ҵ�ҳ��","header","http://api.miaopai.com/m/v2_userinfo.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469157016794&f_type=v2&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&extend=1&channel=yixia&os=android"],
		["GET","������ҳͷ����Ϣ","header","http://api.miaopai.com/m/space_user_info.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469157047836&f_type=v2&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&suid=hT3I5E1~CMJyBAhMhIy7Ww__&os=android"],
		["GET","������ҳfeed","header","http://api.miaopai.com/m/channel_forward_reward.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&timestamp=1469157048301&live=1&f_type=v2&suid=hT3I5E1~CMJyBAhMhIy7Ww__&os=android&likeStat=1&timeflag=0&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=1&per=20"],
		["GET","���������","header","http://api.miaopai.com/m/search_hint.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&keyword=%E8%BF%9C&timestamp=1469157124193&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&os=android"],
		["GET","�������ɴ��۾�����ҳ��","header","http://api.miaopai.com/m/v2_topic.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&topicname=%E8%90%8C%E8%90%8C%E5%A4%A7%E7%9C%BC%E7%9D%9B&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&timestamp=1469157235205&os=android&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=1&type=2&per=20"],
		["GET","����ҳ��@����","header","http://api.miaopai.com/m/v2relation/follow.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&timestamp=1469157308052&f_type=v2&suid=hT3I5E1~CMJyBAhMhIy7Ww__&os=android&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=2&per=50"],
		["GET","����ҳ�滰��","header","http://api.miaopai.com/m/all-topic.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&timestamp=1469157344721&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&os=android&per=5"],
		["GET","�ҵ�ҳ���޹�����Ƶ�б�","header","http://api.miaopai.com/m/liked_video.json?deviceId=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&vend=miaopai&token=16v1DTbsvqPVNFKPhi3LpDz1tE030SLy&uuid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&timestamp=1469157386828&suid=hT3I5E1~CMJyBAhMhIy7Ww__&os=android&likeStat=1&unique_id=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&version=6.5.7&udid=e3ffe3a4-0dc9-36ba-932d-b9919fbf0c7d&channel=yixia&page=1&per=20"],
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
			self.ws1.write(0,0,'���󷽷�'.decode('gbk'),style)
			self.ws1.write(0,1,'������״̬��'.decode('gbk'),style)
			self.ws1.write(0,2,'�ӿ�״̬��'.decode('gbk'),style)
			self.ws1.write(0,3,'����ʱ��'.decode('gbk'),style)
			self.ws1.write(0,4,'�ӿ�����'.decode('gbk'),style)
			self.ws1.write(0,5,'�ӿ�����'.decode('gbk'),style)
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
				
MPInterfaceTest().interfacesTest()