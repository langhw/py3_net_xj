import datetime
import os.path
import xlrd
import re
from netmiko import ConnectHandler as ch
from netmiko import NetMikoAuthenticationException
from netmiko import NetMikoTimeoutException

print("设备巡检中……")
switch_with_authentication_issue = []
switch_not_reachable = []
switch_OSError = []
switch_other_Error = []

xj_jg = "巡检结果//"#巡检后的记录全部在此目录中
xj_faill = "巡检结果//巡检失败记录//"#巡检失败的ip在此目录中
xj_sheet = "运维巡检表.xls"#巡检清单表，按照此表参数执行

if os.path.exists(xj_jg) == 0:
    os.mkdir(xj_jg)
if os.path.exists(xj_faill) == 0:
    os.mkdir(xj_faill)

workbook = xlrd.open_workbook(xj_sheet)#打开巡检清单表，读取表格中值赋给相应的变量，注意表格列数从0-8
sheet = workbook.sheet_by_index(0)
for index in range(1, sheet.nrows):
    hostname = sheet.row(index)[6].value
    ipaddr = sheet.row(index)[0].value
    username = sheet.row(index)[1].value
    password = sheet.row(index)[2].value
    enable_password = sheet.row(index)[3].value
    vendor = sheet.row(index)[4].value
    command = sheet.row(index)[7].value
    location = sheet.row(index)[8].value

    try:
        device = {
            'device_type': vendor,
            'ip': ipaddr,
            'username': username,
            'password': password,
            'secret': enable_password,
        }

        times = str(datetime.datetime.today().strftime('%Y-%m-%d-%H-%m'))
        time2 = str(datetime.datetime.today().strftime('%Y-%m-%d  %H:%m'))
        xj_dir = xj_jg + location + '//'#设备操作日志记录目录
        cmd = open(command, 'r')
        cmd.seek(0)

        conn = ch(**device)
        swname = conn.find_prompt().replace("#", "").replace("<", "").replace(">", "")#读取设备配置的名称
        xj_file = swname + '_' + ipaddr + '-' + times + '.txt'
        xj_log = xj_dir + xj_file#设备操作日志记录文件
        xj_log_info = "设备信息:" + hostname + "\n设备名称：" + swname + "\n管理IP：" + ipaddr + "\n巡检人：xxx" + "\n巡检时间：" + time2

        print("Successfully connected to " + swname + "\nIP:" + ipaddr)
        if re.search('cisco_xe', vendor.lower()):
            conn.send_command('terminal length 0')
            conn.enable()
        elif re.search('cisco_nxos', vendor.lower()):
            conn.send_command('terminal length 0')
            conn.enable()
        elif re.search('cisco_ios', vendor.lower()):
            conn.send_command('terminal length 0')
            conn.enable()
        elif re.search('hp', vendor.lower()):
            conn.send_command('screen-length disable')

        if os.path.exists(xj_dir) == 0:
            os.mkdir(xj_dir)

        with open(xj_log, 'a+', encoding='utf-8') as xj_logw:
            xj_logw.write(xj_log_info)
            for line in cmd.readlines():
                output = "\n********************************************\n" + line + "\n" + conn.send_command(
                    line.strip())
                xj_logw.write(str(output))

        cmd.close()

    except OSError:
        print("User OSError for " + hostname + ".")
        switch_OSError.append(ipaddr)
    except NetMikoAuthenticationException:
        print("User authentication failed for " + hostname + ".")
        switch_with_authentication_issue.append(ipaddr)
    except NetMikoTimeoutException:
        print(hostname + "is not reachable.")
        switch_not_reachable.append(ipaddr)
    except:
        print(hostname)
        switch_other_Error.append(ipaddr)
    finally:
        print(hostname + ipaddr + '巡检完成！')

print('\nUser authentication failed for below switches: ')
for i in switch_with_authentication_issue:
    print(i)
    with open(xj_faill + "/faillist_authentication.txt", 'a+') as faillist:
        faillist.write(i + '\n')

print('\nBelow switches are not reachable: ')
for i in switch_not_reachable:
    print(i)
    with open(xj_faill + "/faillist_not_reachable.txt", 'a+') as faillist:
        faillist.write(i + '\n')

print('\nBelow switches are OSError: ')
for i in switch_OSError:
    print(i)
    with open(xj_faill + "/faillist_OSError.txt", 'a+') as faillist:
        faillist.write(i + '\n')

print('\nBelow switches are Other_Error: ')
for i in switch_other_Error:
    print(i)
    with open(xj_faill + "/faillist_Other_Error.txt", 'a+') as faillist:
        faillist.write(i + '\n')

print("\n 巡检结束，请移至“巡检结果”查看。")
