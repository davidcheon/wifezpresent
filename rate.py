#!/usr/bin/python
#!_*_ coding:utf-8 _*-
import re
import urllib2
def getkoreanratechange(url='http://data.bank.hexun.com/other/cms/fxjhjson.ashx?callback=PereMoreData'):
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
if __name__=='__main__':
	print getkoreanratechange('http://data.bank.hexun.com/other/cms/fxjhjson.ashx?callback=PereMoreData')
