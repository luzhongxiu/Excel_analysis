import pandas as pd
import numpy as np
import datetime

def excelwrite(df):
	year =datetime.datetime.now().year # 获取当前日期
	month =datetime.datetime.now().month
	day =datetime.datetime.now().day
	new_excel_name = str(year)+str(month)+str(day)
	basestation = "D:/pythoncode/20191226_投资交资分析/"+new_excel_name+".xls"
	df.to_excel(basestation)

def df_resolve(df):
	# 筛选非待摊费
	df1 = df[df['任务编号'].str.contains('\.01')==True] 
	print(df1)

	df2 = df1[df1['项目当前状态'].str.contains('已批准，允许事务处理')==True]
	print(df2)
	
	#筛选资本性支出缺口及比例
	rows = df2.shape[0]
	columns = df2.shape[1]

	list_rest=[]
	for i in range(rows):
		list_rest.append(df2.iloc[i]['立项批复总投资（元）']-df2.iloc[i]['截止到目前累计完成投资合计（元）'])
	
	list_rest_rate=[]
	for i in range(rows):
		list_rest_rate.append(list_rest[i]/df2.iloc[i]['立项批复总投资（元）'])

	
	# print(type(list_rest))
	# print(type(list_rest_rate))
	
	return df2
def ExcelRead():
	year =datetime.datetime.now().year # 获取当前日期
	month =datetime.datetime.now().month
	day =datetime.datetime.now().day
	date = str(year)+str(month)+str(day)
	ExcelName="全指标任务数据展示"+"_"+date+".xls"
	df=pd.read_excel(ExcelName)

	# df_rows = df.shape[0]
	# df_column = df.shape[1]
	# print(df_rows)
	# print(df_column)
	# column_Name = df.columns.values.tolist()
	column_new_name=['项目编码','任务编号','项目名称','工建中心三级部门','项目实际负责人','项目当前状态','立项批复总投资（元）','采购订单金额（元）','截止到目前累计完成投资合计（元）','截止上年底累计完成投资合计（元）','本年累计完成投资合计（元）','本月完成投资合计（元）','截止到目前累计付现金额（元）','截至上年底在建工程余额（元）']
	df_new= pd.DataFrame()
	for i in column_new_name:
		df_new[i]=df[i]

	df_new_resolve = df_resolve(df_new)
	return df_new_resolve
def main():
	df_new=ExcelRead()
	excelwrite(df_new)














if __name__ == '__main__':
	main()