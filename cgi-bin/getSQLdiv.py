#!/usr/bin/python
# -*- coding: UTF-8 -*-


import json
def makejson(data):
	jsondata=[]
	for row in data:
		j={}
		j['component_name']=row[3]
		j['component_type']=row[1]
		j['component_txm']=row[2]
		j['component_PartNo']=row[0]
		jsondata.append(j)
	print(json.dumps(jsondata)[1:-1])


def sub_query(conn,qc):#quick change over module
	cur=conn.cursor()
	qcid="203QCM"+qc[-2:]
	sql="""
	select b.GoodsCode,c.GoodsName,a.GoodsTxm, c.GoodsType as bbb
	from dbo.tab_GoodsInfo as b  
	LEFT JOIN dbo.tab_GoodsKcWz as a ON a.GoodsTxm=b.GoodsTxm 
	left join dbo.tab_GoodsCommon as c on c.GoodsBatch=b.GoodsCode 
	left join dbo.tab_GoodsType as d on d.GoodsType=c.GoodsType 
	where  RackID=?
	AND c.GoodsType<6
	AND a.StoreNum>0	
	"""
	rows=cur.execute(sql,qcid).fetchall()
	r=""
	for w in rows:
			print("<TR>")		
			for i in w:
				print("<TD>"+i+"</TD>")
			print("</TR>")




def infoget(TB_no,conn):	
	cur=conn.cursor()
	sql="""
	select b.GoodsCode,c.GoodsName,d.GoodsType,a.GoodsTxm,a.RackID,c.GoodsType as GoodsTypeCode,
	(select Lx from dbo.tab_RackInfo as e where e.RackId=a.RackID) as hjlx
	from dbo.tab_GoodsKcWz as a left join dbo.tab_GoodsInfo as b on b.GoodsTxm=a.GoodsTxm
	left join dbo.tab_GoodsCommon as c on c.GoodsBatch=b.GoodsCode left join dbo.tab_GoodsType as d on d.Id=c.GoodsType
	 where a.StoreNum>0 and (select Lx from dbo.tab_RackInfo as e where e.RackId=a.RackID)='试验台架'
	 and (select FatherId from dbo.tab_GoodsType as g where g.Id=c.GoodsType)=30
	 and RackId=?

	"""
	sql_dev="""
	select b.GoodsCode,c.GoodsName,a.GoodsTxm, c.GoodsType as bbb
	from dbo.tab_GoodsInfo as b  
	LEFT JOIN dbo.tab_GoodsKcWz as a ON a.GoodsTxm=b.GoodsTxm 
	left join dbo.tab_GoodsCommon as c on c.GoodsBatch=b.GoodsCode 
	left join dbo.tab_GoodsType as d on d.GoodsType=c.GoodsType 
	where  RackID=?
	AND (c.GoodsType<6 OR c.GoodsType=41)
	AND a.StoreNum>0

	"""
	rows=cur.execute(sql_dev,TB_no).fetchall()
	
	print("<div class='section ext'>")
	if (len(rows)>0):
		#makejson(rows)
		print("<table class='sin' border=1>")
		print("<tr>")
		for row in cur.description:
			print("<TD>"+row[0]+"</TD>")
		print("</TR>")
		for w in rows:
			
			#print("#".join(x for x in w)) #print(str(w).decode('GBK')+"<br>\n")

			if(int(w[3])==41):
				sub_query(conn,w[2])
			else:
				print("<TR>")
				for i in w:
					print("<TD>"+i+"</TD>")
				print("</TR>")
			#print("<br>\n")
		print("</table>\n</div>")
	else:
		print("There is nothing to display.<br><hr>")
print("Content-Type: text/html; charset='utf-8' \n")

print('')

#print('<h2>Hello Word </h2>')

import pyodbc
#import os
#env_dist=os.environ
#print(env_dist.get("QUERY_STRING"))

conn=pyodbc.connect("DSN=esd1;DATABASE=ESD_Store;")

infoget("H190100",conn)
infoget("H190200",conn)
infoget("H190300",conn)
infoget("H190400",conn)
infoget("H190500",conn)
infoget("H190600",conn)
infoget("H190700",conn)
infoget("H190800",conn)

