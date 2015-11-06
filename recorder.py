#!/usr/bin/python
#!_*_ coding:utf-8 _*_
import xlrd
import xlwt
from xlutils.copy import copy
import getopt
import os
import rate
import threading
class recorder(object):
	def __init__(self,changerate=None):
		#if changerate is None:
		#	t=threading.Thread(target=self.getrate('http://data.bank.hexun.com/other/cms/fxjhjson.ashx?callback=PereMoreData'))
		#	t.setDaemon(True)
		#	t.start()
		#else:
		self.changerate=changerate
		self.file='/home/dane/Desktop/recorder.xls'
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
				return (True,u'这个文件 已经被删除')
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
			sname=self.rxld.sheet_by_index(i).name
			if sname!=name:
				sheetnames.append(sname)
		return sheetnames
if __name__=='__main__':
	r=recorder(.5)
#	r.writeexcel(name='daisongchen',address='weihai',product='product1',price=100.0,counts=10,fee=12.0)
#	r.writeexcel(name='jinxianzhu',address='qiqihaer',product='product2',price=110.0,counts=3,fee=22.0)
	r.searchexcel(name=u'dane')
#	print r.getmoredetail(sheetname='jinxianzhu',row=2)
#	r.updateexcel(sheetindex=1,row=1,name='jin',address='weihai',product='prod1',price=10,counts=20,fee=33)
#	r.deletesheet(sheetname=u'dane')
	r.deleteitem(sheetname='dai1',row=2)
