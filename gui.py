#!/usr/bin/python
#!_*_coding:utf-8 _*_
import wx
import sys
import recorder
from wx.lib.pubsub import Publisher
import rate
import threading
class myexception(Exception):pass

class mygui(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self,None,title=u'记账器')
		self.SetSizeHintsSz((900,350),(900,350))
		self.panel=wx.Panel(self)
		self.panel.SetBackgroundColour('white') 
		self.recorder=recorder.recorder()
		self.Bind(wx.EVT_CLOSE,self.closeaction)
		
		self.username=wx.TextCtrl(self.panel)
		self.address=wx.TextCtrl(self.panel)
		self.product=wx.TextCtrl(self.panel)
		self.price=wx.TextCtrl(self.panel)
		self.changerate=wx.TextCtrl(self.panel)
		self.counts=wx.TextCtrl(self.panel)
		self.fee=wx.TextCtrl(self.panel)
		self.changerate=wx.TextCtrl(self.panel)
		self.totalkr=wx.TextCtrl(self.panel)
		self.totalrmb=wx.TextCtrl(self.panel)
		self.savebutton=wx.Button(self.panel,label=u'保存')
		self.savebutton.Bind(wx.EVT_BUTTON,self.saveaction)
	   	self.clearbutton=wx.Button(self.panel,label=u'清空')
		self.clearbutton.Bind(wx.EVT_BUTTON,self.clearaction)

		self.searchusername=wx.TextCtrl(self.panel)
		self.searchaddress=wx.TextCtrl(self.panel)
		self.searchbutton=wx.Button(self.panel,label=u'搜素')
		self.userlist=wx.ListBox(self.panel,-1,(100,100),(150,170),[],wx.LB_SINGLE)


		self.updateusername=wx.TextCtrl(self.panel)
		self.updateaddress=wx.TextCtrl(self.panel)
		self.updateproduct=wx.TextCtrl(self.panel)
		self.updateprice=wx.TextCtrl(self.panel)
		self.updatechangerate=wx.TextCtrl(self.panel)
		self.updatecounts=wx.TextCtrl(self.panel)
		self.updatefee=wx.TextCtrl(self.panel)
		self.updatechangerate=wx.TextCtrl(self.panel)
		self.updatetotalkr=wx.TextCtrl(self.panel)
		self.updatetotalrmb=wx.TextCtrl(self.panel)
		self.updatebutton=wx.Button(self.panel,label=u'更新')
		self.updatebutton.Bind(wx.EVT_BUTTON,self.updateaction)
		self.deletebutton=wx.Button(self.panel,label=u'删除')
		self.deletebutton.Bind(wx.EVT_BUTTON,self.deleteaction)

		self.lefthbox1=wx.BoxSizer()
	
		self.lefthbox1.Add(wx.StaticText(self.panel,label=u'用户名:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox1.Add(self.username,proportion=3,flag=wx.EXPAND,border=0)

		self.lefthbox2=wx.BoxSizer()
		self.lefthbox2.Add(wx.StaticText(self.panel,label=u'地址:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox2.Add(self.address,proportion=3,flag=wx.EXPAND,border=0)

		self.lefthbox3=wx.BoxSizer()
		self.lefthbox3.Add(wx.StaticText(self.panel,label=u'商品名:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox3.Add(self.product,proportion=3,flag=wx.EXPAND,border=0)
  	
		self.lefthbox4=wx.BoxSizer()
		self.lefthbox4.Add(wx.StaticText(self.panel,label=u'单价:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox4.Add(self.price,proportion=3,flag=wx.EXPAND,border=0)

		self.lefthbox5=wx.BoxSizer()
		self.lefthbox5.Add(wx.StaticText(self.panel,label=u'数量:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox5.Add(self.counts,proportion=3,flag=wx.EXPAND,border=0)

		self.lefthbox6=wx.BoxSizer()
		self.lefthbox6.Add(wx.StaticText(self.panel,label=u'邮费:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox6.Add(self.fee,proportion=3,flag=wx.EXPAND,border=0)

		self.lefthbox7=wx.BoxSizer()
		self.lefthbox7.Add(wx.StaticText(self.panel,label=u'总数(韩元):'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox7.Add(self.totalkr,proportion=3,flag=wx.EXPAND,border=0)

		self.lefthbox8=wx.BoxSizer()
		self.lefthbox8.Add(wx.StaticText(self.panel,label=u'汇率:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox8.Add(self.changerate,proportion=3,flag=wx.EXPAND,border=0)

		self.lefthbox9=wx.BoxSizer()
		self.lefthbox9.Add(wx.StaticText(self.panel,label=u'总数(人民币):'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox9.Add(self.totalrmb,proportion=3,flag=wx.EXPAND,border=0)
  	
		self.lefthbox10=wx.BoxSizer()
		self.lefthbox10.Add(self.savebutton,proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox10.Add(self.clearbutton,proportion=1,flag=wx.EXPAND,border=0)

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

		self.midhbox1=wx.BoxSizer()
		self.midhbox1.Add(wx.StaticText(self.panel,label=u'搜素....'),proportion=1,flag=wx.EXPAND,border=1)

		self.midhbox2=wx.BoxSizer()
		self.midhbox2.Add(wx.StaticText(self.panel,label=u'用户名:'),proportion=1,flag=wx.EXPAND,border=0)
		self.midhbox2.Add(self.searchusername,proportion=2,flag=wx.EXPAND,border=0)

		self.midhbox3=wx.BoxSizer()
		self.midhbox3.Add(wx.StaticText(self.panel,label=u'地址:'),proportion=1,flag=wx.EXPAND,border=0)
		self.midhbox3.Add(self.searchaddress,proportion=2,flag=wx.EXPAND,border=0)

		self.midhbox4=wx.BoxSizer()
		self.midhbox4.Add(self.userlist,proportion=3,flag=wx.EXPAND|wx.ALIGN_LEFT|wx.ALIGN_RIGHT,border=0)
  	
		self.midhbox5=wx.BoxSizer()
		self.midhbox5.Add(self.searchbutton,proportion=2,flag=wx.EXPAND,border=0)
		self.midvbox=wx.BoxSizer(orient=wx.VERTICAL)
  		self.midvbox.Add(self.midhbox1)
		self.midvbox.Add(self.midhbox2)
		self.midvbox.Add(self.midhbox3)
		self.midvbox.Add(self.midhbox4)
		self.midvbox.Add(self.midhbox5)

		self.righthbox1=wx.BoxSizer()
		self.righthbox1.Add(wx.StaticText(self.panel,label=u'用户名:'),proportion=1,flag=wx.EXPAND,border=0)
		self.righthbox1.Add(self.updateusername,proportion=3,flag=wx.EXPAND,border=0)

		self.righthbox2=wx.BoxSizer()
		self.righthbox2.Add(wx.StaticText(self.panel,label=u'地址:'),proportion=1,flag=wx.EXPAND,border=0)
		self.righthbox2.Add(self.updateaddress,proportion=3,flag=wx.EXPAND,border=0)

		self.righthbox3=wx.BoxSizer()
		self.righthbox3.Add(wx.StaticText(self.panel,label=u'商品名:'),proportion=1,flag=wx.EXPAND,border=0)
		self.righthbox3.Add(self.updateproduct,proportion=3,flag=wx.EXPAND,border=0)
  	
		self.righthbox4=wx.BoxSizer()
		self.righthbox4.Add(wx.StaticText(self.panel,label=u'单价:'),proportion=1,flag=wx.EXPAND,border=0)
		self.righthbox4.Add(self.updateprice,proportion=3,flag=wx.EXPAND,border=0)

		self.righthbox5=wx.BoxSizer()
		self.righthbox5.Add(wx.StaticText(self.panel,label=u'数量:'),proportion=1,flag=wx.EXPAND,border=0)
		self.righthbox5.Add(self.updatecounts,proportion=3,flag=wx.EXPAND,border=0)

		self.righthbox6=wx.BoxSizer()
		self.righthbox6.Add(wx.StaticText(self.panel,label=u'邮费:'),proportion=1,flag=wx.EXPAND,border=0)
		self.righthbox6.Add(self.updatefee,proportion=3,flag=wx.EXPAND,border=0)

		self.righthbox7=wx.BoxSizer()
		self.righthbox7.Add(wx.StaticText(self.panel,label=u'总数(韩元):'),proportion=1,flag=wx.EXPAND,border=0)
		self.righthbox7.Add(self.updatetotalkr,proportion=3,flag=wx.EXPAND,border=0)

		self.righthbox8=wx.BoxSizer()
		self.righthbox8.Add(wx.StaticText(self.panel,label=u'汇率:'),proportion=1,flag=wx.EXPAND,border=0)
		self.righthbox8.Add(self.updatechangerate,proportion=3,flag=wx.EXPAND,border=0)

		self.righthbox9=wx.BoxSizer()
		self.righthbox9.Add(wx.StaticText(self.panel,label=u'总数(人民币):'),proportion=1,flag=wx.EXPAND,border=0)
		self.righthbox9.Add(self.updatetotalrmb,proportion=3,flag=wx.EXPAND,border=0)
  	
		self.righthbox10=wx.BoxSizer()
		self.righthbox10.Add(self.updatebutton,proportion=1,flag=wx.EXPAND,border=0)
		self.righthbox10.Add(self.deletebutton,proportion=1,flag=wx.EXPAND,border=0)

		self.rightvbox=wx.BoxSizer(wx.VERTICAL)
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
	def deleteaction(self,evt):
		pass
	def saveaction(self,evt):
		pass
	def clearaction(self,evt):
		pass
	def closeaction(self,evt):
		sys.exit(0)
	def updateaction(self,evt):
		pass
	def getrateresult(self,result):
		self.changerate.SetValue(str(result.data))
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
