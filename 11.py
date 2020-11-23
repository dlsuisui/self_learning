# -*- coding: utf-8 -*-
"""
Created on Tue Nov 17 13:38:49 2020

@author: admin
"""
import pandas as pd
import numpy as np
import xlsxwriter
path = 'E:\\firesoon\\02_DRG地区优化\\台州规则优化\\tz\\'
path_diag = path+'tz_rule_diag2.xlsx'
path_surg = path+'tz_rule_surg2.xlsx'
print(path_diag)
print(path_surg)

dat = pd.read_csv('step1_class_rule.csv')
dat.loc[:,'MDC'] = dat.ADRG_pre.str[0:1]
dat.loc[:,'icd_code5'] = dat.icd_code.str[0:5]

# 剔除先期,多因子及特殊ADRG
dat = dat[~(dat['MDC'].isin(['A', 'X', 'Y', 'Z', 'R']))]
dat = dat[~(dat['ADRG_pre'].isin(['CB6', 'CB7', 'FB29', 'FM29', 'FM21', 'FM19', 'FM31', 'FN49',
                                  'FN59', 'FN69', 'FN79', 'FQ08', 'GR1', 'HC1', 'HC2', 'HC3', 'HC4',
                                  'HC6', 'MC2', 'ND1', 'OB1', 'WR1', 'WR2', 'RR3']))]
# 剔除type_num1,1
dat = dat[~((dat['type_num_diag'] == 1) & (dat['type_num_surg'] == 1))].dropna()
print(dat.type.unique())

# 诊断
dat_diag = dat[dat['type'] == '诊断']
dat_diag = dat_diag[['ADRG_pre', 'ADRG_name','icd_code5', 'icd_code', 'icd_name']].drop_duplicates()
dat_diag = dat_diag.sort_values(by=['ADRG_pre', 'ADRG_name','icd_code5', 'icd_code', 'icd_name'])
dat_diag_count = dat.groupby(['ADRG_pre','icd_code5'], as_index=False)['icd_name'].count().rename(
    columns={'icd_name': 'count'})
dat_diag_counts = dat.groupby(['ADRG_pre'], as_index=False)['icd_name'].count().rename(columns={'icd_name': 'counts'})
dat_diag = pd.merge(dat_diag, dat_diag_count, on=['ADRG_pre','icd_code5'], how='left')
dat_diag = pd.merge(dat_diag, dat_diag_counts, on=['ADRG_pre'], how='left')
dat_diag['index0'] = dat_diag.index
dat_diag['RN'] = dat_diag[['ADRG_pre','icd_code5', 'index0']].groupby(['ADRG_pre','icd_code5'])['index0'].rank()
dat_diag.drop(columns=['index0'], inplace=True)
dat_diag['CN1'] = dat_diag['count']
dat_diag['CN2'] = dat_diag['counts']
dat_diag['percent'] = dat_diag['count'] / dat_diag['counts']



#手术操作
dat_surg = dat[dat['type'] == '手术操作']
dat_surg = dat_surg[['ADRG_pre', 'ADRG_name','icd_code5', 'icd_code', 'icd_name']].drop_duplicates()
dat_surg = dat_surg.sort_values(by=['ADRG_pre', 'ADRG_name','icd_code5', 'icd_code', 'icd_name'])
dat_surg_count = dat.groupby(['ADRG_pre','icd_code5'], as_index=False)['icd_name'].count().rename(
    columns={'icd_name': 'count'})
dat_surg_counts = dat.groupby(['ADRG_pre'], as_index=False)['icd_name'].count().rename(columns={'icd_name': 'counts'})
dat_surg = pd.merge(dat_surg, dat_surg_count, on=['ADRG_pre','icd_code5'], how='left')
dat_surg = pd.merge(dat_surg, dat_surg_counts, on=['ADRG_pre'], how='left')
dat_surg['index0'] = dat_surg.index
dat_surg['RN'] = dat_surg[['ADRG_pre','icd_code5', 'index0']].groupby(['ADRG_pre','icd_code5'])['index0'].rank()
dat_surg.drop(columns=['index0'], inplace=True)
dat_surg['CN1'] = dat_surg['count']
dat_surg['CN2'] = dat_surg['counts']

dat_surg['percent'] = dat_surg['count'] / dat_surg['counts']
dat_surg = dat_surg.head(100)

# dat_surg.to_csv(path+'dat_surg.csv',index=False)
# dat_diag.to_csv(path+'dat_diag.csv',index=False)


# line_cn=dat_surg.index.size
# cols=list(dat_surg.columns.values)
# merge_cols=['icd_code5','ADRG_pre','ADRG_name','count','counts']
# print(line_cn)
# wb2007 = xlsxwriter.Workbook(path_surg)
# worksheet2007 = wb2007.add_worksheet()
# format_top = wb2007.add_format({'border':1,'bold':True,'text_wrap':True})
# format_other = wb2007.add_format({'border':1,'valign':'vcenter'})
# for i,value in enumerate(cols):  #写表头
#     #print(value)
#     worksheet2007.write(0,i,value,format_top)
# 
# 
# for i in range(line_cn):
#     if dat_surg['CN'].tolist()[i]>1:
#         print('该行存在需要合并的单元格')
#         for j,col in enumerate(cols):
#             if col in merge_cols: #哪些列需要合并
#                 if dat_surg['RN'].to_list()[i]:#合并写第一个单元格，下一个第一个将不再写
#                     # print(dat_surg['CN'].to_list()[i])
#                     worksheet2007.merge_range(i+1,j,i+int(dat_surg['CN'].to_list()[i]),j,dat_surg[col].to_list()[i],format_other)##合并单元格
#                 else:
#                     pass
#             else:
#                 worksheet2007.write(i+1,j,dat_surg[col].to_list()[i],format_other)
#     else:
#         print('该行无需要合并的单元格')
#         for j,col in enumerate(cols):
#             #print(df.ix[i,col])
#             worksheet2007.write(i+1,j,dat_surg[col].to_list()[i],format_other)
# wb2007.close()
def transExcel(path,data,merge_cols_list):
    line_cn=data.index.size
    cols=list(data.columns.values)
    merge_cols=merge_cols_list
    wb2007 = xlsxwriter.Workbook(path)
    worksheet2007 = wb2007.add_worksheet()
    format_top = wb2007.add_format({'border':1,'bold':True,'text_wrap':True})
    format_other = wb2007.add_format({'border':1,'valign':'vcenter'})
    for i,value in enumerate(cols):  #写表头
        #print(value)
        worksheet2007.write(0,i,value,format_top)

    for i in range(line_cn):
        if data['CN2'].tolist()[i]==1:
            print('该行(ADRG_pre)不存在需要合并的单元格')
            for j,col in enumerate(cols):
                #print(df.ix[i,col])
                worksheet2007.write(i+1,j,data[col].to_list()[i],format_other)
        else:
            if data['CN1'].tolist()[i]==1:
                print('该行(ADRG_pre_icd_code5)不存在需要合并的单元格')
                for j, col in enumerate(cols):
                    # print(df.ix[i,col])
                    worksheet2007.write(i + 1, j, data[col].to_list()[i], format_other)
            else:
                for j,col in enumerate(cols):
                    if col in merge_cols: #哪些列需要合并
                        if data['RN'].to_list()[i]:#合并写第一个单元格，下一个第一个将不再写
                            # print(dat_diag['CN'].to_list()[i])
                            worksheet2007.merge_range(i+1,j,i+int(data['CN1'].to_list()[i]),j,data[col].to_list()[i],format_other)##合并单元格
                        else:
                            pass
                    else:
                        worksheet2007.write(i+1,j,data[col].to_list()[i],format_other)
    wb2007.close()



merge_cols_list=['icd_code5','ADRG_pre','ADRG_name','count','counts','percent']
#诊断输出
transExcel(path_diag,dat_diag,merge_cols_list)
#手术输出
transExcel(path_surg,dat_surg,merge_cols_list)



