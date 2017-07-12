# -*- coding: utf-8 -*-
import urllib.request
import urllib.parse
import json
import time
from xlwt import *

f=Workbook(encoding='utf-8')
table =f.add_sheet('海关失信企业信息')

num = 4360 # 由于直接把所有内容放在一页出了问题，所以修改为循环模式
pageSize = 20

def haiguan():
	labels = ['企业名称', '企业组织机构代码', '统一社会信用代码']
	for i, j in enumerate(labels):
			str_exp = 'table.write(0,i,j)'
			exec(str_exp)

	for curPage in range(1,int(num/pageSize)+1):
		data = {}
		data['ccppListQueryRequest.manaType'] = 'C'
		data['ccppListQueryRequest.casePage.curPage'] = curPage
		data['ccppListQueryRequest.casePage.pageSize'] = pageSize
		data = urllib.parse.urlencode(data).encode('utf-8')

		url = "http://credit.customs.gov.cn/ccppAjax/queryLostcreditList.action"
		req = urllib.request.Request(url, headers=headers, data=data) # 利用Request传入headers和data数据，以此方法来使用POST语句
		html = urllib.request.urlopen(req).read()
		html = html.decode('utf-8')
		target = json.loads(html) # 返回的结果是json数据
		result = target['responseResult']['responseData']['copInfoResultList'] # 失信名单所在list
		for i in range(0,len(result)):
			socialCreditCode = result[i]['socialCreditCode']
			saicSysNo = result[i]['saicSysNo']
			nameSaic = result[i]['nameSaic']
			table.write((curPage-1)*pageSize+i+1,0,nameSaic)
			table.write((curPage-1)*pageSize+i+1,1,socialCreditCode)
			table.write((curPage-1)*pageSize+i+1,2,saicSysNo)
		
		print("第%r页已经完成"%curPage)
		time.sleep(2)
	
headers = {}
headers["User-Agent"]="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36"
headers['Cookie'] = 'JSESSIONID=Jt7YZl2M0mW1cdWYj8y1wNZnr3vTSMF1hP87rjYrV1hQMkhVQnGV!-572474836'
headers['host'] = 'credit.customs.gov.cn'
headers['Referer'] = 'http://credit.customs.gov.cn/ccppCopAction/toLostcredit.action'

if __name__=="__main__":
	haiguan()

f.save("海关失信企业信息.xls")