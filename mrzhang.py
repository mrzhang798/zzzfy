from paramiko import SSHClient
from paramiko import AutoAddPolicy
from time import sleep
from os import mkdir
from os import path
import pandas as pd
import yaml
import threading

def getYamlData(path):
    with open(path, 'r',encoding='utf-8')as file:
        data = file.read()
        result = yaml.load(data, Loader=yaml.FullLoader)
        return result

def getSheet(path):
    return pd.read_excel(path, sheet_name=None, converters={0: str}, names=['ip', 'uname', 'pwd'])

def runCommand(connect, command, timeout):
    connect.send(command + "\n")
    sleep(timeout)
    conf = connect.recv(65535).decode("utf8", "ignore")
    return conf

def save(path, data):
    with open(file=path, mode="w+", encoding='utf-8') as file:
        file.write(data)

def start(yamlPath: str, execlPath: str):
    """
    :param yamlPath : yaml配置文件路径
    :param execlPath : execl文件路径
    :return : None
    """
    sshClient = SSHClient()
    sshClient.set_missing_host_key_policy(AutoAddPolicy())
    timeout = 10
    saved = True

    if not path.exists('config'):
        mkdir('config')

    config = getYamlData(yamlPath)
    login_setting = getSheet(execlPath)

    if 'Sheet1' not in login_setting.keys():
        exit('将设备的配置存在在execl表中的 Sheet1 中')
    data = login_setting['Sheet1'].to_dict()
    index_list = list(login_setting['Sheet1']['ip'].to_dict().values())

    for key, values in config.items():
        if 'host' not in values.keys():
            print(key + ' 主机配置获取错误或不存在，已跳过')
            continue
        elif 'command' not in values.keys():
            print(key + ' 命令配置获取错误或不存在，已跳过')
            continue

        if not path.exists('./config/' + key):
            mkdir('./config/' + key)

        if 'timeout' in values.keys():
            timeout = values['timeout']
        if 'saved' in values.keys():
            saved = values['saved']


        for host in values['host']:
            if host in index_list:
                index = index_list.index(host)
                # print(data['ip'][index])
                # print(data['uname'][index])
                # print(data['pwd'][index])

                sshClient.connect(hostname=data['ip'][index], username=data['uname'][index],
                                  password=data['pwd'][index])
                connect = sshClient.invoke_shell()
                print("\033[0;32m[+]\033[0m"+'\t已连接至：'+ host)
                for command in values['command']:
                    print('正在执行: '+command+' 等待时间: '+str(timeout))
                    r = runCommand(connect=connect, command=command, timeout=timeout)
                    if saved:
                        with open('./config/' + key + '/' + host + '.txt', mode='a+', encoding='utf-8') as file:
                            print('./config/' + key + '/' + host + '.txt')
                            file.write("# command: "+command+"\n")
                            file.write(r+"\n")

            else:
                print("\033[0;31m[-]\033[0m"+'\t在execl中未找到ip为 '+host+' 的主机')



def main():
    if not path.exists('config.xlsx'):
        pd1 = pd.DataFrame(columns=['IP', '用户名', '密码'])
        pd1.to_excel('config.xlsx', index=False)
        exit("已生成xlsx文件，填写xlsx后，再次运行本程序！")

    if not path.exists('config.yaml'):
        with open(file='config.yaml', mode='w+', encoding='utf-8'):
            pass
        exit("已生成配置文件 config.yaml，参考example.yaml进行配置后，再次运行本程序！")

    start('config.yaml', 'config.xlsx')

if __name__ == '__main__':
    main()
