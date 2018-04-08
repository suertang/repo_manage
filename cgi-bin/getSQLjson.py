#!/usr/bin/python
# -*- coding: UTF-8 -*-


import json
import pyodbc

def sub_query(conn,qc,ret):#quick change over module
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
    for w in rows:
        dictadd(ret,w[3],w[1])




def infoget(TB_no,conn):
    ret={}	
    cur=conn.cursor()
    sql="""
	select b.GoodsCode,c.GoodsName,a.GoodsTxm, c.GoodsType as bbb
	from dbo.tab_GoodsInfo as b  
	LEFT JOIN dbo.tab_GoodsKcWz as a ON a.GoodsTxm=b.GoodsTxm 
	left join dbo.tab_GoodsCommon as c on c.GoodsBatch=b.GoodsCode 
	left join dbo.tab_GoodsType as d on d.GoodsType=c.GoodsType 
	where  RackID=?
	AND (c.GoodsType<6 OR c.GoodsType=41)
	AND a.StoreNum>0

	"""
    rows=cur.execute(sql,TB_no).fetchall()
    if (len(rows)>0):		
        for w in rows:
            if(int(w[3])==41):
                sub_query(conn,w[2],ret)
            else:
                dictadd(ret,w[3],w[1])
        return ret
    else:
        return ret


def dictadd(d,key,value):
    if(key not in d):
        d[key]=value

print("Content-Type: text/json; charset='utf-8' \n")
conn=pyodbc.connect("DSN=esd1;DATABASE=ESD_Store;")
var={}
for i in range(1,9):
    var["TB{0}".format(i)]=infoget("H190{0}00".format(i),conn)
print(json.dumps(var))
conn.close()

