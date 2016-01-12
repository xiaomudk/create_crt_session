#!/usr/bin/python
# coding=utf-8

import ConfigParser
import re
import xlrd
import os


if __name__ == '__main__':

    # 加载conf.ini里的配置
    config = ConfigParser.ConfigParser()
    config.read('conf.ini')

    xls_names = config.get('conf','xls_names')
    session_temp = config.get('conf','session_temp')
    ip_list = config.get('conf','ip_list')
    servername_list = config.get('conf','servername_list')
    user_list = config.get('conf','user_list')
    port_list = config.get('conf','port_list')

    session_name_format = config.get('conf','session_name_format')

    session_out_dir = config.get('conf','sesssion_out_dir')

    #得到程序运行目录
    root_dir = os.path.abspath(os.path.dirname(__file__))
    #得到session模板目录
    crt_session_temp_path = os.path.join(root_dir,session_temp)
    #生成的crt session目录
    crt_session_dir = os.path.join(root_dir,session_out_dir)

    xlrd.Book.encoding = "gbk"
    #遍历所有的xls文件
    for xls_name in xls_names.split(','):
        print "dealing with %s...." % xls_name
        myxls = xlrd.open_workbook(os.path.join(root_dir,xls_name))

        #遍历xls文件里的sheets
        for sheet_id in range(myxls.nsheets):
            sheet = myxls.sheet_by_index(sheet_id)
            sheetname =  sheet.name
            new_sheetname = sheetname.replace('->','/')
            print new_sheetname

            #根据sheet名称创建对应的目录
            sheet_session_dir = os.path.join(crt_session_dir.decode('gbk'),new_sheetname)
            if not os.path.exists(sheet_session_dir):
                os.makedirs(sheet_session_dir)

            ip_list_num = sheet.row_values(0).index(ip_list)
            servername_list_num = sheet.row_values(0).index(servername_list)
            user_list_num = sheet.row_values(0).index(user_list)
            port_list_num = sheet.row_values(0).index(port_list)

            conf_dict = {'ip':ip_list_num, \
                         'servername':servername_list_num, \
                         'user':user_list_num, \
                         'port':port_list_num,\
                         }

            for row in range(1,sheet.nrows):
                row_dict = {}
                for content,list_num in conf_dict.items():
                    if sheet.row_types(row)[list_num] == 1:
                        var = sheet.row_values(row)[list_num].strip()
                    elif sheet.row_types(row)[list_num] == 2:
                        var = str(int(sheet.row_values(row)[list_num]))
                    else:
                        var = ''
                    row_dict[content] = var

                #得到sesssion_name
                session_name = re.sub(r"\b%s\b" % ip_list,row_dict['ip'], session_name_format)
                session_name = re.sub(r"\b%s\b" % servername_list,row_dict['servername'], session_name)
                session_name = re.sub(r"\b%s\b" % user_list,row_dict['user'], session_name)
                session_name = re.sub(r"\b%s\b" % port_list,row_dict['port'], session_name)

                #session.ini的名字及路径
                ser_ini_name = "%s.ini" % session_name
                ser_ini_path = os.path.join(sheet_session_dir,ser_ini_name)


                # 读取模板，生成新的ini文件
                session_str = open(crt_session_temp_path.decode('gbk'),'r').read()
                if row_dict['ip']:
                    new_session_str = re.sub(r'''("Hostname"=).*''','\g<1>%s'% str(row_dict['ip']), session_str)
                if row_dict['user']:
                    new_session_str = re.sub(r'''("Username"=).*''','\g<1>%s' % str(row_dict['user']), new_session_str)
                if row_dict['port']:
                    #转换为8位的16进制
                    row_dict['port'] = hex(int(row_dict['port']))[2:].zfill(8)
                    new_session_str = re.sub(r'''(\[SSH2\] Port"=).*''','\g<1>%s' % str(row_dict['port']), new_session_str)

                open(ser_ini_path,'w').write(new_session_str)

    raw_input("Press Enter to continue...")
