#!/usr/bin/python
#!_*_coding:utf-8 _*_
import wx
import sys
import recorder
from wx.lib.pubsub import Publisher
import rate
import threading
import os
class myexception(Exception):pass

class mygui(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self,None,title=u'媳妇的记账器')
		self.SetSizeHintsSz((900,400),(900,400))
		self.panel=wx.Panel(self)
		self.panel.SetBackgroundColour(wx.Colour(230,255,255)) 
		self.recorder=recorder.recorder()
		self.Bind(wx.EVT_CLOSE,self.closeaction)
		
		self.userlistselectionindex=0
		self.font = wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD)		

		self.searchheader=wx.StaticText(self.panel,label=u'搜素栏')
		self.searchheader.SetFont(self.font)
		self.searchheader.SetForegroundColour(wx.Colour(255,0,0))
		self.updateheader=wx.StaticText(self.panel,label=u'更新栏')		
		self.updateheader.SetFont(self.font)
		self.updateheader.SetForegroundColour(wx.Colour(255,0,0))

		self.username=wx.TextCtrl(self.panel)
		self.address=wx.TextCtrl(self.panel,style=wx.TE_MULTILINE)
		self.product=wx.TextCtrl(self.panel)
		self.product.Bind(wx.EVT_TEXT,self.valuechangemoreaction)
		self.price=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.price.Bind(wx.EVT_TEXT_ENTER,self.textenteraction)
		self.price.Bind(wx.EVT_TEXT,self.valuechangeaction)
		self.changerate=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.changerate.Bind(wx.EVT_TEXT_ENTER,self.textenteraction)
		self.changerate.Bind(wx.EVT_TEXT,self.valuechangeaction)
		self.counts=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.counts.Bind(wx.EVT_TEXT_ENTER,self.textenteraction)
		self.counts.Bind(wx.EVT_TEXT,self.valuechangeaction)
		self.fee=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.fee.Bind(wx.EVT_TEXT_ENTER,self.textenteraction)
		self.fee.Bind(wx.EVT_TEXT,self.valuechangeaction)
		self.totalkr=wx.TextCtrl(self.panel)
		self.totalkr.SetEditable(False)
		self.totalrmb=wx.TextCtrl(self.panel)
		self.totalrmb.SetEditable(False)
		self.savebutton=wx.Button(self.panel,label=u'保存')
		self.savebutton.Bind(wx.EVT_BUTTON,self.saveaction)
	   	self.clearbutton=wx.Button(self.panel,label=u'清空')
		self.clearbutton.Bind(wx.EVT_BUTTON,self.clearaction)
		self.leftstatus=wx.StaticText(self.panel)		
		self.leftstatus.SetFont(self.font)
		self.leftstatus.SetForegroundColour(wx.RED)

		self.searchusername=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.searchusername.Bind(wx.EVT_TEXT_ENTER,self.searchaction)
		self.searchusername.Bind(wx.EVT_TEXT,self.searchvaluechangeaction)
		self.searchaddress=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.searchaddress.Bind(wx.EVT_TEXT_ENTER,self.searchaction)
		self.searchaddress.Bind(wx.EVT_TEXT,self.searchvaluechangeaction)
		self.searchbutton=wx.Button(self.panel,label=u'搜素')
		self.searchdeletebutton=wx.Button(self.panel,label=u'删除该项')
		self.searchdeletebutton.Enable(False)
		self.searchdeletebutton.Bind(wx.EVT_BUTTON,self.searchdeleteaction)
		self.searchbutton.Bind(wx.EVT_BUTTON,self.searchaction)
		self.userlist=wx.ListBox(self.panel,-1,(100,100),(150,170),[],wx.LB_SINGLE)
		self.userlist.SetBackgroundColour(wx.Colour(255, 255, 255))	
		self.userlist.Bind(wx.EVT_LISTBOX,self.userlistselection)		
	
		self.updateusername=wx.TextCtrl(self.panel)
		self.updateusername.Bind(wx.EVT_TEXT,self.updatetextchangeaction2)
		self.updateaddress=wx.TextCtrl(self.panel,style=wx.TE_MULTILINE)
		self.updateaddress.Bind(wx.EVT_TEXT,self.updatetextchangeaction2)
		self.updateproduct=wx.TextCtrl(self.panel)
		self.updateproduct.Bind(wx.EVT_TEXT,self.updatetextchangeaction2)
		self.updateprice=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.updateprice.Bind(wx.EVT_TEXT_ENTER,self.updatepriceaction)
		self.updateprice.Bind(wx.EVT_TEXT,self.updatetextchangeaction)
		self.updatechangerate=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.updatechangerate.Bind(wx.EVT_TEXT_ENTER,self.updatepriceaction)
		self.updatechangerate.Bind(wx.EVT_TEXT,self.updatetextchangeaction)
		self.updatecounts=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.updatecounts.Bind(wx.EVT_TEXT_ENTER,self.updatepriceaction)
		self.updatecounts.Bind(wx.EVT_TEXT,self.updatetextchangeaction)
		self.updatefee=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.updatefee.Bind(wx.EVT_TEXT_ENTER,self.updatepriceaction)
		self.updatefee.Bind(wx.EVT_TEXT,self.updatetextchangeaction)
		self.updatetotalkr=wx.TextCtrl(self.panel)
		self.updatetotalkr.SetEditable(False)
		self.updatetotalrmb=wx.TextCtrl(self.panel)
		self.updatetotalrmb.SetEditable(False)
		self.updatebutton=wx.Button(self.panel,label=u'更新')
		self.updatebutton.Bind(wx.EVT_BUTTON,self.updateaction)
		self.updatebutton.Disable()
		self.deletebutton=wx.Button(self.panel,label=u'删除此人')
		self.deletebutton.Bind(wx.EVT_BUTTON,self.deleteaction)
		self.deletebutton.Disable()
		self.updatestatus=wx.StaticText(self.panel)
		self.updatestatus.SetFont(self.font)
		self.updatestatus.SetForegroundColour(wx.RED)

		self.lefthbox1=wx.BoxSizer()
		self.lefthbox1.Add(wx.StaticText(self.panel,label=u'用户名:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox1.Add(self.username,proportion=3,flag=wx.EXPAND,border=5)

		self.lefthbox2=wx.BoxSizer()
		self.lefthbox2.Add(wx.StaticText(self.panel,label=u'地址:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox2.Add(self.address,proportion=3,flag=wx.EXPAND,border=5)

		self.lefthbox3=wx.BoxSizer()
		self.lefthbox3.Add(wx.StaticText(self.panel,label=u'商品名:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox3.Add(self.product,proportion=3,flag=wx.EXPAND,border=5)
  	
		self.lefthbox4=wx.BoxSizer()
		self.lefthbox4.Add(wx.StaticText(self.panel,label=u'单价:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox4.Add(self.price,proportion=3,flag=wx.EXPAND,border=5)

		self.lefthbox5=wx.BoxSizer()
		self.lefthbox5.Add(wx.StaticText(self.panel,label=u'数量:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox5.Add(self.counts,proportion=3,flag=wx.EXPAND,border=5)

		self.lefthbox6=wx.BoxSizer()
		self.lefthbox6.Add(wx.StaticText(self.panel,label=u'邮费:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox6.Add(self.fee,proportion=3,flag=wx.EXPAND,border=5)

		self.lefthbox7=wx.BoxSizer()
		self.lefthbox7.Add(wx.StaticText(self.panel,label=u'总数(韩元):'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox7.Add(self.totalkr,proportion=3,flag=wx.EXPAND,border=5)

		self.lefthbox8=wx.BoxSizer()
		self.lefthbox8.Add(wx.StaticText(self.panel,label=u'汇率:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox8.Add(self.changerate,proportion=3,flag=wx.EXPAND,border=5)

		self.lefthbox9=wx.BoxSizer()
		self.lefthbox9.Add(wx.StaticText(self.panel,label=u'总数(人民币):'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox9.Add(self.totalrmb,proportion=3,flag=wx.EXPAND,border=5)
  	
		self.lefthbox10=wx.BoxSizer()
		self.lefthbox10.Add(self.savebutton,proportion=1,flag=wx.EXPAND,border=5)
		self.lefthbox10.Add(self.clearbutton,proportion=1,flag=wx.EXPAND,border=5)
		
		self.lefthbox11=wx.BoxSizer()
		self.lefthbox11.Add(self.leftstatus,proportion=1,flag=wx.EXPAND|wx.ALIGN_LEFT|wx.ALIGN_RIGHT,border=1)

		self.leftvbox=wx.BoxSizer(orient=wx.VERTICAL)
		self.leftvbox.Add(self.lefthbox1)
		self.leftvbox.Add(self.lefthbox2)
		self.leftvbox.Add(self.lefthbox3)
	   	self.leftvbox.Add(self.lefthbox4)
		self.leftvbox.Add(self.lefthbox5)
		self.leftvbox.Add(self.lefthbox6)
		self.leftvbox.Add(self.lefthbox7)
		self.leftvbox.Add(self.lefthbox8)
		self.leftvbox.Add(self.lefthbox9)
		self.leftvbox.Add(self.lefthbox10)
		self.leftvbox.Add(self.lefthbox11)
	
		self.midhbox1=wx.BoxSizer()
		self.midhbox1.Add(self.searchheader,proportion=1,flag=wx.EXPAND,border=1)

		self.midhbox2=wx.BoxSizer()
		self.midhbox2.Add(wx.StaticText(self.panel,label=u'用户名:'),proportion=1,flag=wx.EXPAND,border=0)
		self.midhbox2.Add(self.searchusername,proportion=2,flag=wx.EXPAND,border=0)

		self.midhbox3=wx.BoxSizer()
		self.midhbox3.Add(wx.StaticText(self.panel,label=u'地址:'),proportion=1,flag=wx.EXPAND,border=0)
		self.midhbox3.Add(self.searchaddress,proportion=2,flag=wx.EXPAND,border=0)

		self.midhbox4=wx.BoxSizer()
		self.midhbox4.Add(self.userlist,proportion=3,flag=wx.EXPAND|wx.ALIGN_LEFT|wx.ALIGN_RIGHT,border=0)
  	
		self.midhbox5=wx.BoxSizer()
		self.midhbox5.Add(self.searchbutton,proportion=1,flag=wx.EXPAND,border=0)
		self.midhbox5.Add(self.searchdeletebutton,proportion=1,flag=wx.EXPAND,border=0)
		self.midvbox=wx.BoxSizer(orient=wx.VERTICAL)
  		self.midvbox.Add(self.midhbox1)
		self.midvbox.Add(self.midhbox2)
		self.midvbox.Add(self.midhbox3)
		self.midvbox.Add(self.midhbox4)
		self.midvbox.Add(self.midhbox5)

		self.righthbox1=wx.BoxSizer()
		self.righthbox1.Add(wx.StaticText(self.panel,label=u'用户名:'),proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox1.Add(self.updateusername,proportion=3,flag=wx.EXPAND,border=5)

		self.righthbox2=wx.BoxSizer()
		self.righthbox2.Add(wx.StaticText(self.panel,label=u'地址:'),proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox2.Add(self.updateaddress,proportion=3,flag=wx.EXPAND,border=5)

		self.righthbox3=wx.BoxSizer()
		self.righthbox3.Add(wx.StaticText(self.panel,label=u'商品名:'),proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox3.Add(self.updateproduct,proportion=3,flag=wx.EXPAND,border=5)
  	
		self.righthbox4=wx.BoxSizer()
		self.righthbox4.Add(wx.StaticText(self.panel,label=u'单价:'),proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox4.Add(self.updateprice,proportion=3,flag=wx.EXPAND,border=5)

		self.righthbox5=wx.BoxSizer()
		self.righthbox5.Add(wx.StaticText(self.panel,label=u'数量:'),proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox5.Add(self.updatecounts,proportion=3,flag=wx.EXPAND,border=5)

		self.righthbox6=wx.BoxSizer()
		self.righthbox6.Add(wx.StaticText(self.panel,label=u'邮费:'),proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox6.Add(self.updatefee,proportion=3,flag=wx.EXPAND,border=5)

		self.righthbox7=wx.BoxSizer()
		self.righthbox7.Add(wx.StaticText(self.panel,label=u'总数(韩元):'),proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox7.Add(self.updatetotalkr,proportion=3,flag=wx.EXPAND,border=5)

		self.righthbox8=wx.BoxSizer()
		self.righthbox8.Add(wx.StaticText(self.panel,label=u'汇率:'),proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox8.Add(self.updatechangerate,proportion=3,flag=wx.EXPAND,border=5)

		self.righthbox9=wx.BoxSizer()
		self.righthbox9.Add(wx.StaticText(self.panel,label=u'总数(人民币):'),proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox9.Add(self.updatetotalrmb,proportion=3,flag=wx.EXPAND,border=5)
  	
		self.righthbox10=wx.BoxSizer()
		self.righthbox10.Add(self.updatebutton,proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox10.Add(self.deletebutton,proportion=1,flag=wx.EXPAND,border=5)
	
		self.righthbox11=wx.BoxSizer()
		self.righthbox11.Add(self.updatestatus,proportion=1,flag=wx.EXPAND|wx.LEFT|wx.RIGHT,border=1)
		
		self.righthbox12=wx.BoxSizer()
		self.righthbox12.Add(self.updateheader)

		self.rightvbox=wx.BoxSizer(wx.VERTICAL)
		self.rightvbox.Add(self.righthbox12)		
		self.rightvbox.Add(self.righthbox1)
		self.rightvbox.Add(self.righthbox2)
		self.rightvbox.Add(self.righthbox3)
		self.rightvbox.Add(self.righthbox4)
		self.rightvbox.Add(self.righthbox5)
		self.rightvbox.Add(self.righthbox6)
		self.rightvbox.Add(self.righthbox7)
		self.rightvbox.Add(self.righthbox8)
		self.rightvbox.Add(self.righthbox9)
		self.rightvbox.Add(self.righthbox10)
  		self.rightvbox.Add(self.righthbox11)

		self.vbox=wx.BoxSizer()
		self.vbox.Add((80,0))
		self.vbox.Add(self.leftvbox,border=1)
		self.vbox.Add(self.midvbox,border=1)
		self.vbox.Add(self.rightvbox,border=1)
  		
		self.hbox=wx.BoxSizer(wx.VERTICAL)
		self.filename=wx.TextCtrl(self.panel)
		self.titlebox=wx.BoxSizer()
		self.titlebox.Add((80,0))
		self.titlebox.Add(wx.StaticText(self.panel,label=u'另存为..'),proportion=1)
		self.titlebox.Add(self.filename,proportion=3)
		self.hbox.Add(self.titlebox,proportion=0,border=1)
		self.hbox.Add(self.vbox,proportion=4,border=1)
		self.panel.SetSizer(self.hbox)
		Publisher.subscribe(self.getrateresult,'getrateresult')
		t=getratethread()
		t.start()
	def searchdeleteaction(self,evt):
		pass
		
	def updatetextchangeaction(self,evt):
		self.updatestatus.SetLabel('')
	def updatetextchangeaction2(self,evt):
		self.updatetextchangeaction(self)
		self.updateprice.SetValue('')
		self.updatecounts.SetValue('')
		self.updatetotalkr.SetValue('')
		self.updatetotalrmb.SetValue('')
	def updatepriceaction(self,evt):
		try:
			
			updateprice=float(self.updateprice.GetValue().strip())
			updatecounts=int(self.updatecounts.GetValue().strip())
			updaterate=float(self.updatechangerate.GetValue().strip())
			updatefee=float(self.updatefee.GetValue().strip())
			self.updatetotalkr.SetValue('%.2f'%(updateprice*updatecounts+updatefee))
			self.updatetotalrmb.SetValue('%.2f'%((updateprice*updatecounts+updatefee)*updaterate))
		except Exception,e:
			pass
	def userlistselection(self,evt=None):
		self.userlistselectionindex=evt.GetSelection() if evt!=None else self.userlistselectionindex
		self.updatebutton.Enable()
		self.deletebutton.Enable()
		self.searchdeletebutton.Enable(True)
#		row=self.result['rows'][int(self.userlistselectionindex)]
		row=self.showitems[int(self.userlistselectionindex)][1]		
#		sheetname=self.result['sheetname']
		sheetname=self.showitems[int(self.userlistselectionindex)][0]	
		searchresult=self.recorder.getmoredetail(sheetname=sheetname,row=row)
		print 'row:%d sheetname:%s'%(row,sheetname)
		self.sheetindex=searchresult['sheetindex']
		self.row=row
		self.updateusername.SetValue(searchresult['name'])
		self.updateaddress.SetValue(searchresult['address'])
		self.updateproduct.SetValue(searchresult['product'])
		self.updateprice.SetValue(str(searchresult['price']))
		self.updatecounts.SetValue('%d'%searchresult['counts'])
		self.updatefee.SetValue(str(searchresult['fee']))
		self.updatetotalkr.SetValue(str(searchresult['totalkr']))
		self.updatetotalrmb.SetValue(str(searchresult['totalrmb']))
		self.updatechangerate.SetValue(str(searchresult['rate']))
		
	def searchvaluechangeaction(self,evt):
		self.userlist.Clear()
		self.searchdeletebutton.Enable(False)
	def searchaction(self,evt):
		
		self.userlist.Clear()
		self.searchusernamevalue=self.searchusername.GetValue().strip()
		self.searchaddressvalue=self.searchaddress.GetValue().strip()
		if self.searchusernamevalue=='' and self.searchaddressvalue=='':
			self.showmessage('请先填写搜索姓名或地址')
		else:
			status,self.result= self.recorder.searchexcel(name=u'%s'%self.searchusernamevalue) if self.searchusernamevalue!='' else self.recorder.searchexcel(address=u'%s'%self.searchaddressvalue)
			
			if status:
				showlist=[]
				self.showitems=[]
				for key,value in self.result.items():
					showlist+=value['value']
					for x in value['rows']:
						self.showitems.append((key,x))
				
				self.userlist.SetItems(showlist)
				self.userlist.SetFocus()
				self.userlist.SetSelection(self.userlistselectionindex)
				self.userlistselection()
			else:
				self.showmessage(u'没找到<%s>相应的匹配'%(self.searchusernamevalue if self.searchusernamevalue!='' else self.searchaddressvalue))
	def valuechangemoreaction(self,evt):
		self.valuechangeaction(self)
		self.price.SetValue('')
		self.counts.SetValue('')
	def valuechangeaction(self,evt):
		self.totalkr.SetValue('')
		self.totalrmb.SetValue('')
	def textenteraction(self,evt):
		try:
			self.pricevalue=float(self.price.GetValue())
			self.countsvalue=int(self.counts.GetValue())
			self.feevalue=float(self.fee.GetValue())
			self.ratevalue=0.5 if self.changerate.GetValue()=='' else float(self.changerate.GetValue())
		except Exception,e:
			pass
		else:
			total=self.pricevalue*self.countsvalue+self.feevalue
			self.totalkr.SetValue(str(total))
			self.totalrmb.SetValue('%2.f'%(total/100.0*self.ratevalue))
	def deleteaction(self,evt):
		sheetname=self.result['sheetname']
		status,info=self.recorder.deletesheet(sheetname=sheetname)
		if status:
			self.clearupdateinfo()
			self.userlist.Clear()
		self.updatestatus.SetLabel(info)
		self.showmessage(info)
	def clearupdateinfo(self):
		self.updateusername.SetValue('')
		self.updateaddress.SetValue('')
		self.updateprice.SetValue('')
		self.updatecounts.SetValue('')
		self.updatefee.SetValue('')
		self.updatechangerate.SetValue('')
		self.updatetotalkr.SetValue('')
		self.updatetotalrmb.SetValue('')
		self.updatebutton.Disable()
		self.deletebutton.Disable()
	def saveaction(self,evt):
		#try:
		if self.filename.GetValue()!='':
			self.recorder.setfilename(self.filename.GetValue().strip())
		self.filenamevalue=self.filename.GetValue().strip()
		if self.filenamevalue!='':
			self.filenamevalue=os.path.join('result','%s.xls'%self.filenamevalue.split('.')[0])
		self.usernamevalue=self.username.GetValue().strip()
		self.addressvalue=(' '.join(self.address.GetValue().strip().split('\n'))).strip()
		self.productvalue=self.product.GetValue().strip()
		self.pricevalue=float(self.price.GetValue().strip())
		self.countsvalue=int(self.counts.GetValue().strip())
		self.feevalue=float(self.fee.GetValue().strip())
		self.ratevalue=0.5 if self.changerate.GetValue()=='' else float(self.changerate.GetValue().strip())
		total=self.pricevalue*self.countsvalue+self.feevalue
		if self.totalkr.GetValue()=='':
			self.totalkr.SetValue(str(total))
		if self.totalrmb.GetValue()=='':
			self.totalrmb.SetValue('%.2f'%(total/100.0*self.ratevalue))
		self.recorder.setchangerate(self.ratevalue)
		if self.filenamevalue!='':
			self.recorder.setfilename(self.filenamevalue)
			self.recorder.setfilename(self.filenamevalue)
		self.recorder.writeexcel(name=self.usernamevalue,address=self.addressvalue,product=self.productvalue,price=self.pricevalue,counts=self.countsvalue,fee=self.feevalue)
		self.leftstatus.SetLabel('写入成功!')
		self.showmessage('写入成功!')
#		except Exception,e:
#			raise e
#			self.leftstatus.SetLabel('请先填满所有选项!')
#			self.showmessage('请先填满所有选项!')
#		self.leftstatus.SetLabel('')
#			
	def showmessage(self,msg):
		dlg=wx.MessageDialog(self.panel,msg,caption='注意',style=wx.OK)
		dlg.ShowModal()
		dlg.Destroy()
	def clearaction(self,evt):
		self.filename.SetValue('')
		self.username.SetValue('')
		self.address.SetValue('')
		self.product.SetValue('')
		self.price.SetValue('')
		self.counts.SetValue('')
		self.fee.SetValue('')
		self.totalkr.SetValue('')
		self.totalrmb.SetValue('')
		self.changerate.SetValue('')
		self.leftstatus.SetValue('')
	def closeaction(self,evt):
		sys.exit(0)
	def updateaction(self,evt):
		try:
		
			name=self.updateusername.GetValue()
			address=self.updateaddress.GetValue()
			product=self.updateproduct.GetValue()
			price=float(self.updateprice.GetValue())
			counts=int(self.updatecounts.GetValue())
			fee=float(self.updatefee.GetValue())
			rate=float(self.updatechangerate.GetValue())
			self.updatetotalkr.SetValue('%.2f'%(price*counts+fee))
			self.updatetotalrmb.SetValue('%.2f'%((price*counts+fee)/100.0*rate))
			status,info=self.recorder.updateexcel(name=name,address=address,product=product,counts=counts,fee=fee,rate=rate,price=price,sheetindex=self.sheetindex,row=self.row)
			self.updatestatus.SetLabel(info)
			
			status,self.result= self.recorder.searchexcel(name=u'%s'%self.searchusernamevalue) if self.searchusernamevalue!='' else self.recorder.searchexcel(address=u'%s'%self.searchaddressvalue)
			
			if status:
				showlist=[]
				self.showitems=[]
				for key,value in self.result.items():
					showlist+=value['value']
					for x in value['rows']:
						self.showitems.append((key,x))
				
				self.userlist.SetItems(showlist)
				self.userlist.SetFocus()
				self.userlist.SetSelection(self.userlistselectionindex)
				self.userlistselection()

			self.showmessage(info)
		except Exception,e:
			self.showmessage(u'请输入正确数字格式')
	def getrateresult(self,result):
		self.changerate.SetValue('%.2f'%float(result.data))
class getratethread(threading.Thread):
	def run(self):
		result=rate.getkoreanratechange()
		wx.CallAfter(self.postdata,result)
	def postdata(self,result):
		Publisher.sendMessage('getrateresult',result)
class App(wx.App):
	def __init__(self):
		super(self.__class__,self).__init__()
	def OnInit(self):
		frame=mygui()
		frame.Show(True)
		self.SetTopWindow(frame)
		return True
if __name__=='__main__':
	app=App()
	app.MainLoop()
