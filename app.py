import streamlit as st
import pandas as pd
import os
import shutil
from stoc import stoc

from io import BytesIO


st.set_page_config(layout="wide")

# title
st.title('订单自动处理')
# file upload
uploaded_file = st.file_uploader("上传客户订单", type=["xlsx"])
# merge table upload
uploaded_id_table = st.file_uploader("上传物料对照", type=["xlsx"])

# parse order

def parse_order(df):
    df_orders = pd.DataFrame()
    # select order rows between
    # NO	产品编号	描述	规格	供方料号	数量	未税单价	未税金额	含税金额	交货日期
    # and
    # <NA>	<NA>	<NA>	合计	<NA>	<NA>	41814.05	<NA>
    start = None
    end = None
    for i in range(len(df)):
        if df.iloc[i, 0] == 'NO':
            start = i
        if df.iloc[i, 5] == '合计' or df.iloc[i, 0] == '[以下空白]':
            end = i
        if start and end:
            df_tmp = df.iloc[start:end, :]
            df_tmp.columns = df_tmp.iloc[0, :]
            df_tmp = df_tmp[1:]
            df_orders = pd.concat([df_orders, df_tmp])
            start = None
            end = None
    df_orders = df_orders.reset_index(drop=True)
    return df_orders


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.close()
    processed_data = output.getvalue()
    return processed_data


if uploaded_file and uploaded_id_table:
    toc = stoc()
    # read data
    df = pd.read_excel(uploaded_file)
    df_orders = parse_order(df)
    id_table = pd.read_excel(uploaded_id_table)
    id_table = id_table.iloc[:, :2]
    id_table.columns = ['公司物料', '产品编号']
    
    toc.h3('处理后订单数据')
    st.table(df_orders)
    
    toc.h3('物料对照表')
    st.table(id_table.head())
    
    df_orders.columns = ['NO', '产品编号', '描述', '规格', '供方料号', '数量', '未税单价', '未税金额', '含税金额','交货日期']
    df_merged = df_orders.merge(id_table, on='产品编号', how='left')
    
    toc.h3('合并后数据')
    st.table(df_merged)

    toc.h3('保存合并后数据')
    # save as excel
    df_xlsx = to_excel(df_merged)
    st.download_button(label='📥 下载合并后数据',
                                data=df_xlsx ,
                                file_name= '合并后数据.xlsx')
    
    # show items not found in id_table
    toc.h3('未找到的物料')
    st.table(df_merged[df_merged['公司物料'].isnull()]['产品编号'].unique())
    
    toc.toc()