#coding=utf-8
import re

import win32com.client as win32

def office(app):
	return win32.Dispatch('%s.Application' % app)

def get_outlook(res,name,foldername):
	#从收件箱中获取主题列表5已发送。6收件箱
	#app = 'Outlook'
	outlook = office('Outlook')
	#Senton 发送时间，Sender 发送人,Recipients[0],收件人列表

	ns = outlook.GetNamespace("MAPI")
	inbox = ns.Folders(name).Folders(foldername)
	count = inbox.Items.Count
	print(count)
	for i in range(0,count):
		tim = str(inbox.Items[i].Senton).split( )
		Sub.append([inbox.Items[i].Subject,tim[0],inbox.Items[i].Recipients])
	#提取出‘送验单的’
	for strin in Sub:
		relike = relikesearch(res,strin[0])
		if relike != None:
			Subject_like.append(strin)
			print(relike)
		else:
			pass

	#outlook.quit()

	for i in range(0,len(Subject_like)):
		print(Subject_like[i][1])
	return Subject_like
	#将这部分分隔然后放入Excel中
def excel(strings):
	#app = 'Excel'win32.Dispatch('%s.Application' % app)	
	x1 = office('Excel')
	sh = x1.Workbooks.Add().ActiveSheet

	x1.Visible = True

	sh.Cells(1,1).Value = '变更号' 
	sh.Cells(1,2).Value = '需求号'
	sh.Cells(1,3).Value = '需求名称'
	sh.Cells(1,4).Value = '送验时间'
	sh.Cells(1,5).Value = '验收人'

	for i in range(0,len(strings)):

		sub = Subject_like[i][0].split('_',3)
		Recip = Subject_like[i][2]
		#print(len(Recip))	
		st = str(Recip[0]).split('@')[0]	
		sh.Cells(i+2,1).Value = sub[1]
		sh.Cells(i+2,2).Value = sub[2]
		sh.Cells(i+2,3).Value = sub[3]
		sh.Cells(i+2,4).Value = Subject_like[i][1]
		for j in range(0,len(Recip)):
			st = st,str(Recip[j]).split('@')[0]
		sh.Cells(i+2,5).Value = st
def relikesearch(res,str):
	result = re.match(res,str)
	if result:
		return str
	else :
		pass	


if __name__ == '__main__':
	Subject_like=[]
	Sub = []
	res = '送验单'
	name = 'xyxjxc3@163.com'
	foldername = '收件箱'
	get_outlook(res,name,foldername)
	excel(Subject_like)

