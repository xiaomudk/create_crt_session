# create_crt.ini使用说明


## 1. 配置文件为conf.ini
```
[conf]
#当前目录下的xls名字, 多个xls以逗号分开,不支持中文文件名
xls_names = server_list.xls,junwang2.xls

#session ini模板名字
session_temp = crt_session_temp.ini

#ip,servername,用户名和端口对应的列名
ip_list = ip
servername_list = name
user_list = user
port_list = port

#session名字(也是ini的文件名,不需要加后缀.ini)
#文件名不能包含: \/：*？“<>|
session_name_format = ip(name)

#crt session输出目录
sesssion_out_dir = session
```

## 2.excel表格要求

1. 每个sheet的第一行必须为ip_list、servername_list、user_list、port_list对应的列名
2. 无用的sheet需要删掉！
3. 如果需要分多级目录，sheet名可以使用"->"分隔， 比如: sheet名为"test->123",表示test/123目录



