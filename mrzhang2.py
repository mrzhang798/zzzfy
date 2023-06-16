from paramiko import SSHClient
from paramiko import AutoAddPolicy
from time import sleep
import pandas as pd
from os import mkdir
from os import path

def getConfig(connect, command):
    connect.send(command+'\n')
    sleep(5)
    conf = connect.recv(65535).decode("utf8", "ignore")
    return conf

def getSheet(path):
    return pd.read_excel(path, sheet_name=None, converters={0: str}, names=None)




def save(path,data):
    with open(file=path, mode="a+", encoding='utf-8') as file:
        file.write(data)

def start(workBookName):
    sshClient = SSHClient()
    sshClient.set_missing_host_key_policy(AutoAddPolicy())

    if not path.exists('config'):
        mkdir('config')

    data = getSheet(workBookName)

    for sheet_name in data.keys():
        sheet_data = data[sheet_name].fillna(False)
        if not path.exists('./config/'+sheet_name):
            mkdir('./config/'+sheet_name)

        col_count = sheet_data.shape[0]

        for col in range(0, col_count):
            data = sheet_data.loc[col:col].to_dict()
            ip = data['ip'][col]
            uname = data['uname'][col]
            pwd = data['pwd'][col]
            del data['ip'], data['uname'], data['pwd']
            sshClient.connect(hostname=ip, username=uname, password=pwd)
            connect = sshClient.invoke_shell()
            for value in data.values():
                if value[col]:
                    print(ip+' 正在执行: '+str(value[col]))
                    conf = getConfig(connect, value[col])
                    filename = 'huawei'+ip
                    save('./config/'+sheet_name+'/'+filename+'.txt', conf)
            print()
def main():
    if not path.exists('1234.xlsx'):
        pd1 = pd.DataFrame(columns=['ip', 'uname', 'pwd'])
        pd1.to_excel('1234.xlsx', index=False)
        exit("已生成xlsx文件，填写xlsx后，再次运行本程序！")
    start('1234.xlsx')

if __name__=='__main__':
    main()