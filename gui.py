#!/usr/bin/python
#!_*_coding:utf-8 _*_
import wx
import sys
from xlutils.copy import copy
from wx.lib.pubsub import Publisher
import threading
import os
import re
import urllib2
import xlrd
import xlwt
def getkoreanratechange(url='http://data.bank.hexun.com/other/cms/fxjhjson.ashx?callback=PereMoreData'):
	try:
		heads = {'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
			'Accept-Charset':'GB2312,utf-8;q=0.7,*;q=0.7',
			'Accept-Language':'zh-cn,zh;q=0.5',
			'Cache-Control':'max-age=0',
			'Connection':'keep-alive',
			'Keep-Alive':'115',
			'Referer':url,
			'User-Agent':'Mozilla/5.0 (X11; U; Linux x86_64; zh-CN; rv:1.9.2.14) Gecko/20110221 Ubuntu/10.10 (maverick) Firefox/3.6.14'}
		
		opener = urllib2.build_opener(urllib2.HTTPCookieProcessor())
		urllib2.install_opener(opener)
		req = urllib2.Request(url)
		opener.addheaders = heads.items()
		page = opener.open(req).read()
		pat=re.compile(r"refePrice:('[^']+)',code:'KRW")
		results=pat.findall(page)
		return float(results[0].strip('"\'')) if len(results) else 0
	except Exception,e:
		return 0
class myexception(Exception):pass
class recorder(object):
	def __init__(self,changerate=None):
		#if changerate is None:
		#	t=threading.Thread(target=self.getrate('http://data.bank.hexun.com/other/cms/fxjhjson.ashx?callback=PereMoreData'))
		#	t.setDaemon(True)
		#	t.start()
		#else:
		self.changerate=changerate
#		self.file='/home/dane/Desktop/recorder.xls'
#		self.file=os.path.join('result','default.xls')
		self.file=''
		self.style1=xlwt.easyxf('font: height 240,name SimSun, colour_index black, bold off, italic off; align:wrap on, vert centre, horiz centre;')
		self.style2=xlwt.easyxf('font: height 340, name Arial, colour_index blue, bold off, italic off; align:wrap on, vert centre, horiz centre;')
		self.styleboldred=xlwt.easyxf('font: color-index red,bold on;align:wrap on,vert centre, horiz centre;');
		self.background1=xlwt.easyxf('pattern: pattern solid,fore_colour red;align:wrap on, vert centre,horiz centre;')
		self.rxld=None
	def setchangerate(self,changerate):
		self.changerate=float(changerate)
	def setfilename(self,filename):
		self.file=filename
	def getrate(self,url):
		self.changerate=rate.getkoreanratechange(url)
	def writeexcel(self,**args):
		header=(u'用户名',u'地址',u'商品名',u'单价',u'数量',u'邮费',u'总数(韩元)',u'汇率',u'总数(人民币)')
		if os.path.isfile(self.file):
			self.rxld=xlrd.open_workbook(self.file,formatting_info=True)
			nsheets=self.rxld.nsheets
			flag=False
			whichone=nsheets
			for i in xrange(0,nsheets):
				sheetname=self.rxld.sheet_by_index(i).name
				if sheetname==args['name']:
					flag=True
					whichone=i
					break
			if not flag:
				headerstyle=self.styleboldred
				ws=copy(self.rxld)
				ws.add_sheet(args['name'])
				ws.save(self.file)
				wsheet=ws.get_sheet(nsheets)
				for i,e in enumerate(header):
					wsheet.write(0,i,e,self.background1 if i%2==0 else headerstyle)
					wsheet.col(i).width=10000 if e==u'地址' else 5000 if e!=u'数量' and e!=u'汇率' else 3000
					if e=='COUNTS':wsheet.col(i).width=2500
				wsheet.write(1,0,args['name'],self.style1)
				wsheet.write(1,1,args['address'],self.style1)
				wsheet.write(1,2,args['product'],self.style1)
				wsheet.write(1,3,args['price'],self.style1)
				wsheet.write(1,4,args['counts'],self.style1)
				wsheet.write(1,5,args['fee'],self.style1)
				wsheet.write(1,6,float(args['price'])*int(args['counts'])+float(args['fee']),self.style1)
				wsheet.write(1,7,'%.2f'%self.changerate,self.style1)
				wsheet.write(1,8,'%.2f'%((float(args['price'])*int(args['counts'])+float(args['fee']))/100.0*float(self.changerate)),self.style1)
				wsheet.write(2,5,args['fee'],self.style2)
				wsheet.write(2,6,float(args['price'])*int(args['counts'])+float(args['fee']),self.style2)
				wsheet.write(2,7,'%.2f'%self.changerate,self.style2)
				wsheet.write(2,8,'%.2f'%((float(args['price'])*int(args['counts'])+float(args['fee']))/100.0*float(self.changerate)),self.style2)
				ws.save(self.file)

			else:
				self.appendexcel(whichone,**args)
		else:	
			#header=(u'NAME',u'ADDRESS',u'PRODUCT',u'PRICE',u'COUNTS',u'FEE',u'TOTAL(KR)',u'RATE(KR/RMB)',u'TOTAL(RMB)')
			headerstyle=self.styleboldred
			w=xlwt.Workbook(encoding='utf-8')
			ws=w.add_sheet(args['name'],cell_overwrite_ok=True)
			for i,c in enumerate(header):
				ws.write(0,i,c,self.background1 if i%2==0 else headerstyle)
				ws.col(i).width=10000 if c==u'地址' else 5000  if c!=u'数量' and c!=u'汇率' else 3000
				if c=='COUNTS':ws.col(i).width=2500
			ws.write(1,0,args['name'],self.style1)
			ws.write(1,1,args['address'],self.style1)
			ws.write(1,2,args['product'],self.style1)
			ws.write(1,3,args['price'],self.style1)
			ws.write(1,4,args['counts'],self.style1)
			ws.write(1,5,args['fee'],self.style1)
			ws.write(1,6,float(args['price']*int(args['counts']))+float(args['fee']),self.style1)
			ws.write(1,7,'%.2f'%(self.changerate),self.style1)
			ws.write(1,8,'%.2f'%((float(args['price'])*int(args['counts'])+float(args['fee']))/100.0*float(self.changerate)),self.style1)
			ws.write(2,5,args['fee'],self.style2)
			ws.write(2,6,float(args['price'])*int(args['counts'])+float(args['fee']),self.style2)
			ws.write(2,7,'%.2f'%self.changerate,self.style2)
			ws.write(2,8,'%.2f'%((float(args['price'])*int(args['counts'])+float(args['fee']))/100.0*float(self.changerate)),self.style2)
			w.save(self.file)
			self.rxld=xlrd.open_workbook(self.file,formatting_info=True)
	def updateexcel(self,**args):
		try:
			sheetindex=args['sheetindex']
			row=args['row']
			wb=copy(self.rxld)
			sheet=wb.get_sheet(sheetindex)
			sheet.write(row,0,args['name'],self.style1)
			sheet.write(row,1,args['address'],self.style1)
			sheet.write(row,2,args['product'],self.style1)
			sheet.write(row,3,float(args['price']),self.style1)
			sheet.write(row,4,int(args['counts']),self.style1)
			sheet.write(row,5,float(args['fee']),self.style1)
			sheet.write(row,6,float(args['price'])*int(args['counts'])+float(args['fee']),self.style1)
			sheet.write(row,7,'%.2f'%args['rate'],self.style1)
			sheet.write(row,8,'%.2f'%((float(args['price'])*int(args['counts'])+float(args['fee']))/100.0*args['rate']),self.style1)
			rs=self.rxld.sheet_by_index(sheetindex)
			rows=rs.nrows
			while(rs.cell_value(rows-1,5)==''):
				rows-=1
			maxfee=float(args['fee'])
			total=0
			maxrate=float(args['rate'])
			for r in xrange(1,rows-1):
					fee=float(rs.cell_value(r,5))
					rate=float(rs.cell_value(r,7))
					if rate>maxrate:
						if r!=row:
							maxrate=rate
					if fee>maxfee:
						if r!=row:
							maxfee=fee
					
					tmp=(rs.cell_value(r,6) -fee)if r!=row else float(args['price'])*int(args['counts'])
					total+=tmp
			total+=maxfee
			sheet.write(rows-1,5,maxfee,self.style2)
			sheet.write(rows-1,6,total,self.style2)
#			sheet.write(rows-1,7,'%.2f'%args['rate'],self.style2)
			sheet.write(rows-1,7,'%.2f'%maxrate,self.style2)
			sheet.write(rows-1,8,'%.2f'%(total/100.0*maxrate),self.style2)
			wb.save(self.file)
			return (True,u'更新成功')
		except Exception,e:
			return (False,str(e))
#			raise e

	def appendexcel(self,whichone,**values):
		rsheet=self.rxld.sheet_by_index(whichone)
		rows=rsheet.nrows-1
		while(rsheet.cell_value(rows,5)==''):
			rows-=1
		wxls=copy(self.rxld)
		wsheet=wxls.get_sheet(whichone)
		wsheet.write(rows,0,values['name'],self.style1)
		wsheet.write(rows,1,values['address'],self.style1)
		wsheet.write(rows,2,values['product'],self.style1)
		wsheet.write(rows,3,values['price'],self.style1)
		wsheet.write(rows,4,values['counts'],self.style1)
		wsheet.write(rows,5,values['fee'],self.style1)
		wsheet.write(rows,6,float(values['price'])*int(values['counts'])+float(values['fee']),self.style1)
		wsheet.write(rows,7,'%.2f'%float(self.changerate),self.style1)
		wsheet.write(rows,8,'%.2f'%((float(values['price'])*int(values['counts'])+float(values['fee']))/100.0*float(self.changerate)),self.style1)
		total=0
		maxfee=float(values['fee'])
		maxrate=float(values['rate'])
		for i in xrange(1,rows):
			fee=float(rsheet.cell_value(i,5))
			rate=float(rsheet.cell_value(i,7))
			total+=float(rsheet.cell_value(i,6))-fee
			if fee>maxfee:
				maxfee=fee
			if rate>maxrate:
				maxrate=rate			
		wsheet.write(rows+1,5,maxfee,self.style2)
		wsheet.write(rows+1,6,total+maxfee+float(values['price'])*int(values['counts']),self.style2)
		wsheet.write(rows+1,7,'%.2f'%maxrate,self.style2)
		wsheet.write(rows+1,8,'%.2f'%(maxrate*(total+float(values['fee'])+float(values['price'])*int(values['counts']))/100.0),self.style2)
		wxls.save(self.file)
	
	def searchexcel(self,**args):
		self.rxld=xlrd.open_workbook(self.file,formatting_info=True)
		nsheets=self.rxld.nsheets
		flag=False
		matchlist={}
		if args.has_key('name'):
			for n in xrange(nsheets):
				sh=self.rxld.sheet_by_index(n)
				name=sh.name
				nrows=sh.nrows
				if name.find(args['name'])>=0:
					if not args.has_key('address'):
						for n in xrange(1,nrows-1):
							if sh.cell_value(n,0)=='':
								continue
							matchlist[name].append(n) if matchlist.has_key(name) else matchlist.setdefault(name,[n])
					else:
						for r in xrange(1,nrows-1):
							addr=sh.cell_value(r,1)
							if addr=='':
								continue
							if addr.find(args['address'])>=0:						
								matchlist[name].append(r) if matchlist.has_key(name) else matchlist.setdefault(name,[r])
								
			return (len(matchlist.keys())>0,u'没找到该用户:{0}'.format(args['name']) if not len(matchlist.keys())>0 else self.readexcel(matchlist))
		elif args.has_key('address'):
			for n in xrange(nsheets):
				sh=self.rxld.sheet_by_index(n)
				nrows=sh.nrows
				for r in xrange(1,nrows-1):
					addr=sh.cell_value(r,1)
					if addr=='':
						continue
					if addr.find(args['address'])>=0:
						matchlist[sh.name].append(r) if matchlist.has_key(sh.name) else matchlist.setdefault(sh.name,[r])
			return (len(matchlist.keys())>0,u'没找到该用户:{0}'.format(args['address']) if not len(matchlist.keys())>0 else self.readexcel(matchlist))
	def readexcel(self,matchlist):
		result={}
		for shname in matchlist.keys():
			sh=self.rxld.sheet_by_name(shname)
			result[shname]={}
			for row in matchlist[shname]:
				tmp=u'%s %s %d个'%(sh.cell_value(row,0),sh.cell_value(row,2),int(sh.cell_value(row,4)))
				result[shname]['value'].append(tmp) if result[shname].has_key('value') else result[shname].setdefault('value',[tmp])
				result[shname]['rows'].append(row) if result[shname].has_key('rows') else result[shname].setdefault('rows',[row])
		return result
	
	def getmoredetail(self,**args):
		if self.fileexist():
			result={}
			for n in xrange(self.rxld.nsheets):
				name=self.rxld.sheet_by_index(n).name
				if name==args['sheetname']:
					result['sheetindex']=n
					break
			sheet=self.rxld.sheet_by_name(args['sheetname'])
			row=args['row']
			result['name']=sheet.cell_value(row,0)
			result['address']=sheet.cell_value(row,1)
			result['product']=sheet.cell_value(row,2)
			result['price']=sheet.cell_value(row,3)
			result['counts']=sheet.cell_value(row,4)
			result['fee']=sheet.cell_value(row,5)
			result['totalkr']=sheet.cell_value(row,6)
			result['rate']=sheet.cell_value(row,7)
			result['totalrmb']=sheet.cell_value(row,8)		
			return result
		else:
			return None
	def fileexist(self):
		return os.path.exists(self.file)
	def deletesheet(self,**args):
		try:
			sheetname=args['sheetname']
			new_workbook=copy(self.rxld)
			for sheet in  new_workbook._Workbook__worksheets:
				if sheet.name==sheetname:
					new_workbook._Workbook__worksheets.remove(sheet)
			if len(new_workbook._Workbook__worksheets)==0:
				os.remove(self.file)
				return (True,u'这个文件已经被删除')
			else:
				new_workbook.save(self.file)
				return (True,u'已经成功删除该用户')
		except Exception,e:
			return (False,str(e))
	def gettotalsheets(self):
		if os.path.isfile(self.file):
			self.rxld= xlrd.open_workbook(self.file,formatting_info=True)
			sheets=self.rxld.nsheets
			result=[]
			for n in xrange(sheets):
				result.append(self.rxld.sheet_by_index(n).name)
			return result
		return None
	def getfilename(self):
		return self.file
	def deleteitem(self,**args):
		sheetname=args['sheetname']
		row=args['row']
		whichsheet=0
		for s in xrange(self.rxld.nsheets):
			
			if self.rxld.sheet_by_index(s).name==sheetname:
				whichsheet=s
				rsheet=self.rxld.sheet_by_index(s)
				break
		wxls=copy(self.rxld)
		sheet=wxls.get_sheet(whichsheet)
		nrows=rsheet.nrows
		while(rsheet.cell_value(nrows-1,5)==''):
			nrows-=1
		ncols=rsheet.ncols
		if nrows==3:
			return (True,'last item')
		else:
			try:
				for n in xrange(row,nrows):
					for c in xrange(ncols):
						if n==nrows-2 and c in [5,6,7,8]:
							maxfee=0
							maxrate=0
							totalkr=0
							totalrmb=0
							for a in xrange(1,nrows-1):
								if a!=row:
									if maxfee<float(rsheet.cell_value(a,5)):
										maxfee=float(rsheet.cell_value(a,5))
									if maxrate<float(rsheet.cell_value(a,7)):
										maxrate=float(rsheet.cell_value(a,7))
									totalkr+=float(rsheet.cell_value(a,6))-float(rsheet.cell_value(a,5))
#									totalrmb+=float(rsheet.cell_value(a,8))
#							for a in xrange(row+1,nrows-1):
#								if maxfee<float(rsheet.cell_value(a,5)):
#									maxfee=float(rsheet.cell_value(a,5))
#								totalkr+=float(rsheet.cell_value(a,6))-float(rsheet.cell_value(a,5))
#								totalrmb+=float(rsheet.cell_value(a,8))
							totalkr+=maxfee
							sheet.write(n,c,maxfee,self.style2) if c==5 else sheet.write(n,c,totalkr,self.style2) if c==6 else sheet.write(n,c,maxrate,self.style2) if c==7 else sheet.write(n,c,float('%.2f'%(totalkr/100.0*maxrate)),self.style2)
						elif n==nrows-1:
							sheet.write(n,c,'',self.style1)
						else:
							value=rsheet.cell_value(n+1,c)
							sheet.write(n,c,value,self.style1)
							
				wxls.save(self.file)
				return (True,u'已经成功删除该列')
			except Exception,e:
				return (False,str(e))
	def getcurrentusers(self,name):
		sheets=self.rxld.nsheets
		sheetnames=[]
		for i in xrange(sheets):
			rsheet=self.rxld.sheet_by_index(i)
			sname=rsheet.name
			nrows=rsheet.nrows
			while(rsheet.cell_value(nrows-1,5)==''):
				nrows-=1
			if sname!=name:
				totalkr=rsheet.cell_value(nrows-1,6)
				totalrmb=rsheet.cell_value(nrows-1,8)
				sheetnames.append(sname+'(%s,%s)'%(totalkr,totalrmb))
		return sheetnames
	def getfilename(self):
		return self.file
	def getrxld(self):
		return self.rxld
class mygui(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self,None,title=u'媳妇的记账器')
		self.SetSizeHintsSz((900,400),(900,400))
		self.panel=wx.Panel(self,-1,style=wx.SIMPLE_BORDER)
		self.panel.SetBackgroundColour(wx.Colour(230,255,255)) 
#		self.recorder=recorder.recorder()
		self.recorder=recorder()
		self.Bind(wx.EVT_CLOSE,self.closeaction)
		
		self.userlistselectionindex=0
		self.font = wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD)		
		

		self.currentusers=wx.StaticText(self.panel,label=u'当前记录的用户:')
		self.currentusers.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.currentusers.SetForegroundColour(wx.Colour(0,0,255))			
		self.currentusers.SetFont(self.font)

		self.searchheader=wx.StaticText(self.panel,label=u'搜素栏')
		self.searchheader.SetFont(self.font)
		self.searchheader.SetForegroundColour(wx.Colour(255,0,0))
		self.updateheader=wx.StaticText(self.panel,label=u'更新栏')		
		self.updateheader.SetFont(self.font)
		self.updateheader.SetForegroundColour(wx.Colour(255,0,0))
		self.leftheader=wx.StaticText(self.panel,label=u'填入用户购买的商品信息')
		self.leftheader.SetFont(self.font)
		self.leftheader.SetForegroundColour(wx.Colour(255,0,0))
	
		self.filename=wx.TextCtrl(self.panel,style=wx.TE_PROCESS_ENTER)
		self.filename.Bind(wx.EVT_TEXT_ENTER,self.enterfilenameaction)
		self.filename.Bind(wx.EVT_TEXT,self.filenameaction)
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
		self.lefthbox1.Add(wx.StaticText(self.panel,label=u'用户名:'),proportion=1,flag=wx.ALIGN_RIGHT,border=0)
		self.lefthbox1.Add(self.username,proportion=1,flag=wx.EXPAND,border=1)

		self.lefthbox2=wx.BoxSizer()
		self.lefthbox2.Add(wx.StaticText(self.panel,label=u'地址:'),proportion=1,flag=wx.EXPAND|wx.ALL,border=0)
		self.lefthbox2.Add(self.address,proportion=1,flag=wx.EXPAND,border=1)

#		self.lefthbox3=wx.BoxSizer()
		self.lefthbox1.Add(wx.StaticText(self.panel,label=u'商品名:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox1.Add(self.product,proportion=1,flag=wx.EXPAND,border=1)
  	
		self.lefthbox4=wx.BoxSizer()
		self.lefthbox4.Add(wx.StaticText(self.panel,label=u'单价:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox4.Add(self.price,proportion=1,flag=wx.EXPAND,border=1)

#		self.lefthbox5=wx.BoxSizer()
		self.lefthbox4.Add(wx.StaticText(self.panel,label=u'数量:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox4.Add(self.counts,proportion=1,flag=wx.EXPAND,border=1)

		self.lefthbox6=wx.BoxSizer()
		self.lefthbox6.Add(wx.StaticText(self.panel,label=u'邮费:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox6.Add(self.fee,proportion=1,flag=wx.EXPAND,border=1)

		self.lefthbox7=wx.BoxSizer()
		self.lefthbox7.Add(wx.StaticText(self.panel,label=u'总数(韩元):'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox7.Add(self.totalkr,proportion=1,flag=wx.EXPAND,border=1)

#		self.lefthbox8=wx.BoxSizer()
		self.lefthbox6.Add(wx.StaticText(self.panel,label=u'汇率:'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox6.Add(self.changerate,proportion=1,flag=wx.EXPAND,border=1)

#		self.lefthbox9=wx.BoxSizer()
		self.lefthbox7.Add(wx.StaticText(self.panel,label=u'总数(人民币):'),proportion=1,flag=wx.EXPAND,border=0)
		self.lefthbox7.Add(self.totalrmb,proportion=1,flag=wx.EXPAND,border=1)
  	
		self.lefthbox10=wx.BoxSizer()
		self.lefthbox10.Add(self.savebutton,proportion=1,flag=wx.EXPAND,border=1)
		self.lefthbox10.Add(self.clearbutton,proportion=1,flag=wx.EXPAND,border=1)
		
		self.lefthbox11=wx.BoxSizer()
		self.lefthbox11.Add(self.leftstatus,proportion=1,flag=wx.EXPAND|wx.ALIGN_LEFT|wx.ALIGN_RIGHT,border=1)

		self.lefthbox12=wx.BoxSizer()
		self.lefthbox12.Add(wx.StaticText(self.panel,label=u'另存为其他文件名:'),proportion=1)
		self.lefthbox12.Add(self.filename,proportion=1)
		
		self.lefthbox13=wx.BoxSizer()
		self.lefthbox13.Add(self.leftheader,proportion=1)
		
		self.leftvbox=wx.BoxSizer(orient=wx.VERTICAL)
		self.leftvbox.Add(self.lefthbox13)
		self.leftvbox.Add(self.lefthbox12)
		self.leftvbox.Add(self.lefthbox1)
		self.leftvbox.Add(self.lefthbox2)
#		self.leftvbox.Add(self.lefthbox3)
	   	self.leftvbox.Add(self.lefthbox4)
#		self.leftvbox.Add(self.lefthbox5)
		self.leftvbox.Add(self.lefthbox6)
		self.leftvbox.Add(self.lefthbox7)
#		self.leftvbox.Add(self.lefthbox8)
#		self.leftvbox.Add(self.lefthbox9)
		self.leftvbox.Add(self.lefthbox10)
		self.leftvbox.Add(self.lefthbox11)
		self.midhbox1=wx.BoxSizer()
		self.midhbox1.Add(self.searchheader,proportion=1,flag=wx.EXPAND,border=1)

		self.midhbox2=wx.BoxSizer()
		self.midhbox2.Add(wx.StaticText(self.panel,label=u'以用户名搜索:'),proportion=1,flag=wx.EXPAND,border=0)
		self.midhbox2.Add(self.searchusername,proportion=1,flag=wx.EXPAND,border=0)

		self.midhbox3=wx.BoxSizer()
		self.midhbox3.Add(wx.StaticText(self.panel,label=u'以地址搜索:'),proportion=1,flag=wx.EXPAND,border=0)
		self.midhbox3.Add(self.searchaddress,proportion=1,flag=wx.EXPAND,border=0)

		self.midhbox4=wx.BoxSizer()
		self.midhbox4.Add(self.userlist,proportion=1,flag=wx.EXPAND|wx.ALIGN_LEFT|wx.ALIGN_RIGHT,border=0)
  	
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
		self.righthbox1.Add(wx.StaticText(self.panel,label=u'用户名:'),proportion=1,flag=wx.EXPAND,border=1)
		self.righthbox1.Add(self.updateusername,proportion=1,flag=wx.EXPAND,border=1)

		self.righthbox2=wx.BoxSizer()
		self.righthbox2.Add(wx.StaticText(self.panel,label=u'地址:'),proportion=1,flag=wx.EXPAND,border=1)
		self.righthbox2.Add(self.updateaddress,proportion=1,flag=wx.EXPAND,border=1)

#		self.righthbox3=wx.BoxSizer()
		self.righthbox1.Add(wx.StaticText(self.panel,label=u'商品名:'),proportion=1,flag=wx.EXPAND,border=1)
		self.righthbox1.Add(self.updateproduct,proportion=1,flag=wx.EXPAND,border=1)
  	
		self.righthbox4=wx.BoxSizer()
		self.righthbox4.Add(wx.StaticText(self.panel,label=u'单价:'),proportion=1,flag=wx.EXPAND,border=1)
		self.righthbox4.Add(self.updateprice,proportion=1,flag=wx.EXPAND,border=1)

#		self.righthbox5=wx.BoxSizer()
		self.righthbox4.Add(wx.StaticText(self.panel,label=u'数量:'),proportion=1,flag=wx.EXPAND,border=1)
		self.righthbox4.Add(self.updatecounts,proportion=1,flag=wx.EXPAND,border=1)

		self.righthbox6=wx.BoxSizer()
		self.righthbox6.Add(wx.StaticText(self.panel,label=u'邮费:'),proportion=1,flag=wx.EXPAND,border=1)
		self.righthbox6.Add(self.updatefee,proportion=1,flag=wx.EXPAND,border=1)

		self.righthbox7=wx.BoxSizer()
		self.righthbox7.Add(wx.StaticText(self.panel,label=u'总数(韩元):'),proportion=1,flag=wx.EXPAND,border=1)
		self.righthbox7.Add(self.updatetotalkr,proportion=1,flag=wx.EXPAND,border=1)

#		self.righthbox8=wx.BoxSizer()
		self.righthbox6.Add(wx.StaticText(self.panel,label=u'汇率:'),proportion=1,flag=wx.EXPAND,border=1)
		self.righthbox6.Add(self.updatechangerate,proportion=1,flag=wx.EXPAND,border=1)

#		self.righthbox9=wx.BoxSizer()
		self.righthbox7.Add(wx.StaticText(self.panel,label=u'总数(人民币):'),proportion=1,flag=wx.EXPAND,border=5)
		self.righthbox7.Add(self.updatetotalrmb,proportion=1,flag=wx.EXPAND,border=1)
  	
		self.righthbox10=wx.BoxSizer()
		self.righthbox10.Add(self.updatebutton,proportion=1,flag=wx.EXPAND,border=1)
		self.righthbox10.Add(self.deletebutton,proportion=1,flag=wx.EXPAND,border=1)
	
		self.righthbox11=wx.BoxSizer()
		self.righthbox11.Add(self.updatestatus,proportion=1,flag=wx.EXPAND|wx.LEFT|wx.RIGHT,border=1)
		
		self.righthbox12=wx.BoxSizer()
		self.righthbox12.Add(self.updateheader)

		self.rightvbox=wx.BoxSizer(wx.VERTICAL)
		self.rightvbox.Add(self.righthbox12)		
		self.rightvbox.Add(self.righthbox1)
		self.rightvbox.Add(self.righthbox2)
#		self.rightvbox.Add(self.righthbox3)
		self.rightvbox.Add(self.righthbox4)
#		self.rightvbox.Add(self.righthbox5)
		self.rightvbox.Add(self.righthbox6)
		self.rightvbox.Add(self.righthbox7)
#		self.rightvbox.Add(self.righthbox8)
#		self.rightvbox.Add(self.righthbox9)
		self.rightvbox.Add(self.righthbox10)
  		self.rightvbox.Add(self.righthbox11)

		self.vbox=wx.BoxSizer()
		self.vbox.Add((20,0))
		self.vbox.Add(self.leftvbox,border=1)
		self.vbox.Add(self.midvbox,border=1)
		self.vbox.Add(self.rightvbox,border=1)
  		
		self.generalhbox=wx.BoxSizer(wx.VERTICAL)
		self.generalhbox.Add(self.vbox)
		self.bottomhbox=wx.BoxSizer()
		self.bottomhbox.Add(self.currentusers)
		self.generalhbox.Add(self.bottomhbox)
		self.panel.SetSizer(self.generalhbox)

		self.loadcurrentusers()
		Publisher.subscribe(self.getrateresult,'getrateresult')
		t=getratethread()
		t.start()
	def verifyfileexist(func):
		def wrapper(self,evt=None):
			if os.path.exists(self.recorder.getfilename()):
				func(self,evt)
			else:
				if self.filename.GetValue().strip()!='':
					if self.recorder.getfilename() =='':
						self.recorder.setfilename('%s.xls'%os.path.join('result',self.filename.GetValue().strip().split('.')[0]))
						func(self,evt)
					else:
						self.showmessage(u'这个文件已经被删除，请重新添加.')
						self.searchusername.SetValue('')
						self.searchaddress.SetValue('')
				else:
					self.showmessage(u'请填入要要搜索的文件名')
		return wrapper
	def verifyfileexist2(status=False):
		def _wrapper(func):
			def wrapper(self,evt=None):
				if os.path.exists(self.recorder.getfilename()):
					func(self,evt)
				else:
					if self.filename.GetValue().strip()!='':
						if self.recorder.getfilename() =='':
							self.recorder.setfilename('%s.xls'%os.path.join('result',self.filename.GetValue().strip().split('.')[0]))
							func(self,evt)
						else:
							if  status:
								self.recorder.setfilename('%s.xls'%os.path.join('result',self.filename.GetValue().strip().split('.')[0]))
							else:
								self.showmessage(u'这个文件已经被删除，请重新添加.')
								self.searchusername.SetValue('')
								self.searchaddress.SetValue('')
					else:
						self.showmessage(u'请填入要要搜索的文件名')
			return wrapper
		return _wrapper
	def filenameaction(self,evt):
		self.userlist.Clear()
		self.searchusername.SetValue('')
		self.searchaddress.SetValue('')
		self.clearupdateinfo()
		self.searchdeletebutton.Disable()
	def loadcurrentusers(self):
		usernames=self.recorder.gettotalsheets()
		if usernames is not None:
			tmp=''
			for key,name in enumerate(usernames):
				tmp+=name+self.gettotals(name)+('\n' if (key+1)%3==0 else '\t')
			self.currentusers.SetLabel('%s:%s'%(u'当前记录的用户',tmp))
	def gettotals(self,sheetname):
		rsheet=self.recorder.getrxld().sheet_by_name(sheetname)
		nrows=rsheet.nrows
		while(rsheet.cell_value(nrows-1,5)==''):
			nrows-=1
		totalkr=rsheet.cell_value(nrows-1,6)
		totalrmb=rsheet.cell_value(nrows-1,8)
		return '(%s,%s)'%(totalkr,totalrmb)
	def setfilename(self,evt):
		self.filenamevalue=self.filename.GetValue().strip()
		if self.filenamevalue!='':
			self.filenamevalue=os.path.join('result','%s.xls'%self.filenamevalue.split('.')[0])
			self.recorder.setfilename(self.filenamevalue)
	@verifyfileexist2(False)
	def searchdeleteaction(self,evt):
		sheetname,row=self.showitems[self.userlistselectionindex]
		dlg=wx.MessageDialog(self.panel,u'你确定要删除该项吗',u'警告',wx.YES_NO|wx.ICON_QUESTION)
		if dlg.ShowModal()==wx.ID_YES:			
			status,info=self.recorder.deleteitem(sheetname=sheetname,row=row)
			if status:
				if info=='last item':
					dlg=wx.MessageDialog(self.panel,u'这是该人最后一笔记录,你确定要删除<%s>吗?'%self.sheetname,u'警告',wx.YES_NO|wx.ICON_QUESTION)
					if dlg.ShowModal()==wx.ID_YES:
						sts,info=self.recorder.deletesheet(sheetname=sheetname)
						if sts:
							self.userlist.Clear()
							self._refreshuserlist()
							if self.userlist.IsEmpty():
								self.searchusername.SetValue('')
								self.searchaddress.SetValue('')
								self.searchdeletebutton.Disable()				
								self.searchbutton.Disable()
								self.clearupdateinfo()
							users=self.recorder.getcurrentusers(sheetname)
							tmp=u'当前记录的用户:'
							for i,user in enumerate(users):
								if (i+1)%3==0:
									tmp+=user+'\n'
								else:
									tmp+=user+'\t'
							self.currentusers.SetLabel(tmp)
					else:
						self.userlist.SetFocus()
					dlg.Destroy()
				else:
					self.loadcurrentusers()
		#			self.showitems=[n for i,n in enumerate(self.showitems) if i != self.userlistselectionindex ]
					self.clearupdateinfo()
					self.userlist.Clear()
					self._refreshuserlist()
					self.showmessage(info)
	
	def enterfilenameaction(self,evt):
		tmpname='%s.xls'%self.filename.GetValue().split('.')[0]
		if os.path.exists(os.path.join('result',tmpname)):
			self.recorder.setfilename(os.path.join('result',tmpname))
			self.loadcurrentusers()
		else:
			self.showmessage(u'<%s>这个文件不存在，请重新输入'%tmpname)
	def updatetextchangeaction(self,evt):
		self.updatestatus.SetLabel('')
	def updatetextchangeaction2(self,evt):
		self.updatetextchangeaction(self)
#		self.updateprice.SetValue('')
#		self.updatecounts.SetValue('')
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
		if len(self.showitems)>0:
			self.userlistselectionindex=evt.GetSelection() if evt!=None else self.userlistselectionindex
			self.updatebutton.Enable()
			self.deletebutton.Enable()
			self.searchdeletebutton.Enable(True)
			row=self.showitems[int(self.userlistselectionindex)][1]		
			self.sheetname=self.showitems[int(self.userlistselectionindex)][0]
			searchresult=self.recorder.getmoredetail(sheetname=self.sheetname,row=row)
			if searchresult!=None:
				self.sheetindex=searchresult['sheetindex']
				self.row=row
				self.updateusername.SetValue(searchresult['name'])
				self.updateaddress.SetValue(searchresult['address'])
				self.updateproduct.SetValue(searchresult['product'])
				self.updateprice.SetValue(str(searchresult['price']))
				self.updatecounts.SetValue(str(searchresult['counts']))
				self.updatefee.SetValue(str(searchresult['fee']))
				self.updatetotalkr.SetValue(str(searchresult['totalkr']))
				self.updatetotalrmb.SetValue(str(searchresult['totalrmb']))
				self.updatechangerate.SetValue(str(searchresult['rate']))
#			else:
#				self.showmessage(u'已经没有用户了，你需要再添加')				
		
	def searchvaluechangeaction(self,evt):
		self.userlist.Clear()
#		self.searchdeletebutton.Enable()
		self.searchbutton.Enable()
		self.clearupdateinfo()
	@verifyfileexist2(status=True)
	def searchaction(self,evt):
		
		self.userlist.Clear()
		self.searchusernamevalue=self.searchusername.GetValue().strip()
		self.searchaddressvalue=self.searchaddress.GetValue().strip()
		if self.searchusernamevalue=='' and self.searchaddressvalue=='':
			self.setfilename(self.filename.GetValue().strip())
			self.loadcurrentusers()
#			if self.recorder.getfilename()!='':
			if evt!=None:
				self.showmessage('请先填写搜索姓名或地址')
		else:
			if self.filename.GetValue().strip()!='':
				self.recorder.setfilename('%s.xls'%os.path.join('result',self.filename.GetValue().strip().split('.')[0]))
			if os.path.exists(self.recorder.getfilename()):
				status,self.result= self.recorder.searchexcel(name=u'%s'%self.searchusernamevalue) if self.searchusernamevalue!='' else self.recorder.searchexcel(address=u'%s'%self.searchaddressvalue)		
				if status:
					showlist=[]
					self.showitems=[]
					for key,value in self.result.items():
						showlist+=value['value']
						for x in value['rows']:
							self.showitems.append((key,x))
					self.searchdeletebutton.Enable()
					self.userlist.SetItems(showlist)
					self.userlist.SetFocus()
					self.userlist.SetSelection(self.userlistselectionindex)
					self.userlistselection()
					self.loadcurrentusers()
				else:
					self.searchdeletebutton.Disable()
					self.showmessage(u'没找到<%s>相应的匹配'%(self.searchusernamevalue if self.searchusernamevalue!='' else self.searchaddressvalue))
			else:
				self.showmessage(u'<%s.xls>这个文件不存在，请重新输入'%self.filename.GetValue().strip().split('.')[0])
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
			self.ratevalue=0.6 if self.changerate.GetValue()=='' else float(self.changerate.GetValue())
		except Exception,e:
			pass
		else:
			total=self.pricevalue*self.countsvalue+self.feevalue
			self.totalkr.SetValue(str(total))
			self.totalrmb.SetValue('%2.f'%(total/100.0*self.ratevalue))
	@verifyfileexist2(False)
	def deleteaction(self,evt):
		totalsheets=self.recorder.gettotalsheets()
		if len(totalsheets)==1 and totalsheets[0]==self.sheetname:
			dlg=wx.MessageDialog(self.panel,u'这是最后一个人,你确定要删除<%s>吗?'%self.sheetname,u'警告',wx.YES_NO|wx.ICON_QUESTION)
			if dlg.ShowModal()==wx.ID_YES:
				os.remove(self.recorder.getfilename())
				self.searchusername.SetValue('')
				self.searchaddress.SetValue('')
				self.searchbutton.Disable()
				self.searchdeletebutton.Disable()				
				self.clearupdateinfo()
				self.userlist.Clear()
				self.currentusers.SetLabel(u'当前记录的用户:')
				self.showmessage(u'已经成功删除此文件!')
			else:
				self.userlist.SetFocus()
				
		else:
			dlg=wx.MessageDialog(self.panel,u'你确定要删除<%s>吗?'%self.sheetname,u'警告',wx.YES_NO|wx.ICON_QUESTION)
			if dlg.ShowModal()==wx.ID_YES:
				status,info=self.recorder.deletesheet(sheetname=self.sheetname)
				if status:
					
					self.showitems=[(m[0],m[1]) for n,m in enumerate(self.showitems) if n!=self.userlistselectionindex]
					self.userlist.Clear()
					self._refreshuserlist()
					if self.userlist.IsEmpty():
						self.searchusername.SetValue('')
						self.searchaddress.SetValue('')
						self.searchbutton.Disable()
						self.searchdeletebutton.Disable()				
						self.clearupdateinfo()
					self.loadcurrentusers()
				self.updatestatus.SetLabel(info)
				self.showmessage(info)
			dlg.Destroy()
	def clearupdateinfo(self):
		self.updateusername.SetValue('')
		self.updateaddress.SetValue('')
		self.updateproduct.SetValue('')
		self.updateprice.SetValue('')
		self.updatecounts.SetValue('')
		self.updatefee.SetValue('')
		self.updatechangerate.SetValue('')
		self.updatetotalkr.SetValue('')
		self.updatetotalrmb.SetValue('')
		self.updatebutton.Disable()
		self.deletebutton.Disable()
	def saveaction(self,evt):
		try:
			self.filenamevalue=self.filename.GetValue().strip()
			tmp=self.recorder.getfilename()
			if self.filenamevalue!='':
				self.filenamevalue=os.path.join('result','%s.xls'%self.filenamevalue.split('.')[0])
				self.recorder.setfilename(self.filenamevalue)
			else:
				raise myexception('请填写需要保存的文件名!')
			self.usernamevalue=self.username.GetValue().strip()
			self.addressvalue=(' '.join(self.address.GetValue().strip().split('\n'))).strip()
			self.productvalue=self.product.GetValue().strip()
			self.pricevalue=float(self.price.GetValue().strip())
			self.countsvalue=int(float(self.counts.GetValue().strip()))
			self.feevalue=float(self.fee.GetValue().strip())
			self.ratevalue=0.6 if self.changerate.GetValue()=='' else float('%.2f'%float(self.changerate.GetValue().strip()))
			total=self.pricevalue*self.countsvalue+self.feevalue
			if self.totalkr.GetValue()=='':
				self.totalkr.SetValue(str(total))
			if self.totalrmb.GetValue()=='':
				self.totalrmb.SetValue('%.2f'%(total/100.0*self.ratevalue))
			self.recorder.setchangerate(self.ratevalue)
			self.recorder.writeexcel(name=self.usernamevalue,address=self.addressvalue,product=self.productvalue,price=self.pricevalue,counts=self.countsvalue,fee=self.feevalue,rate=self.ratevalue)
			if tmp==self.filenamevalue:
				self.searchaction(None)
				if self.userlist.IsEmpty():
					self.searchdeletebutton.Disable()
			else:
				self.userlist.Clear()
				self.searchusername.SetValue('')
				self.searchaddress.SetValue('')
				self.clearupdateinfo()
			self.loadcurrentusers()
			self.leftstatus.SetLabel('写入成功!')
			self.showmessage('写入成功!')
		except Exception,e:
#			raise e
			self.leftstatus.SetLabel(str(e))
			self.showmessage(str(e))
		self.leftstatus.SetLabel('')
			
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
		self.leftstatus.SetLabel('')
	def closeaction(self,evt):
		sys.exit(0)
	@verifyfileexist2(False)
	def updateaction(self,evt):
		try:
		
			name=self.updateusername.GetValue().strip()
			address=self.updateaddress.GetValue().strip()
			product=self.updateproduct.GetValue().strip()
			price=float(self.updateprice.GetValue().strip())
			counts=int(float(self.updatecounts.GetValue().strip()))
			fee=float(self.updatefee.GetValue().strip())
			rate=float(self.updatechangerate.GetValue().strip())
			self.updatetotalkr.SetValue('%.2f'%(price*counts+fee))
			self.updatetotalrmb.SetValue('%.2f'%((price*counts+fee)/100.0*rate))
			status,info=self.recorder.updateexcel(name=name,address=address,product=product,counts=counts,fee=fee,rate=rate,price=price,sheetindex=self.sheetindex,row=self.row)
			self.updatestatus.SetLabel(info)
			
			status,self.result= self.recorder.searchexcel(name=u'%s'%self.searchusernamevalue) if self.searchusernamevalue!='' else self.recorder.searchexcel(address=u'%s'%self.searchaddressvalue)
			self.userlist.Clear()
			self._refreshuserlist()
			self.loadcurrentusers()
			self.showmessage(info)
		except Exception,e:
			self.showmessage(str(e))
#			self.showmessage(u'请输入正确数字格式')
	@verifyfileexist2(False)
	def _refreshuserlist(self,evt=None):
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
	def getrateresult(self,result):
		self.changerate.SetValue('%.2f'%float(result.data))
		self.recorder.setchangerate(float(result.data))
class getratethread(threading.Thread):
	def run(self):
#		result=rate.getkoreanratechange()
		result=getkoreanratechange()
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
