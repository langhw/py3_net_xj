import os
import re
from pandas import DataFrame as df
from styleframe import StyleFrame, Styler


def get_info():
   pat_ip = re.compile(r'管理IP：(\d+\.\d+\.\d+\.\d+)\n', re.S)    #ip
   pat_version = re.compile(r'display version\n(.*?)\n')    #version
   info = {'ip':[], 'version':[]}
   for filename in os.listdir('D:\lhw\巡检结果\huawei'):
      # print(filename)
      with open(r'D:\lhw\巡检结果\huawei\\' + filename, 'r', encoding='utf-8') as f:
         file_r = f.read()
         ip = pat_ip.findall(file_r)[0]
         version = pat_version.findall(file_r)[0]
         # print(file_r)
         # print(ip)
         # print(version)
         info['ip'].append(ip)
         info['version'].append(version)
   return info

def Table_conversion(out_file):
   column = ['ip', 'version']
   info = get_info()
   data = df.from_dict(info, orient='index').T
   data.reset_index(inplace=True)
   data.rename(columns={'index': '序号'}, inplace=True)
   data.index = data.index + 1
   sf = StyleFrame(data)
   sf.apply_column_style(cols_to_style=column,
                         styler_obj=Styler(horizontal_alignment='left'),
                         style_header=False)
   writer = StyleFrame.ExcelWriter(out_file)
   sf.to_excel(
       excel_writer=writer,
       sheet_name=out_file,
       best_fit=column,
       columns_and_rows_to_freeze='B2',
       row_to_add_filters=0,
   )
   writer.save()
   writer.close()
if __name__ == '__main__':
   Table_conversion('ip.xlsx')