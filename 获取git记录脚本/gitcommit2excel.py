# -*- coding: utf-8 -*-
"""
Created on Tue Jul 24 10:01:04 2018

@author: luca.zhu
将git中的提交记录保存为releaseNote，放在git路径下的releaseNote文件夹中
1.将git配置到环境变量中
2.gitbash使用如下命令，适配中文文件名
git config --global core.quotepath false
"""
import subprocess,os,datetime,openpyxl

class Git2Excel(object):
    
    def __init__(self):
        '''
        初始化git路径,commitId,提交记录,编码格式
        '''
        self.gitPath = ""
        self.commitIdStart = ""
        self.commitIdEnd=""
        self.commitList = []
        self.revList=[]
        self.code ="utf-8"
        self.excelPath =''
        self.version='1.0.0'
    
    def __getCommitList(self):
        os.chdir(self.path)
        '''
        获取要显示的提交信息，现在显示近5周的提交记录，可以自行修改
        '''
        showcommandstr ='git log --since=5.weeks --pretty=format:"%H | %an | %cd | %s"'
        print(showcommandstr)
        showrev = subprocess.Popen(showcommandstr,shell=True,stdout=subprocess.PIPE)
        showrevlist = showrev.stdout.readlines()
        showrev.kill()
        for show in showrevlist:
            showstr = str(show, encoding = self.code);
            print(showstr)
            self.commitList.append(showstr)
     
    def __getResultList(self):
        '''
        获取要处理的提交信息
        '''
        os.chdir(self.path)
        #commmandstr ='git log  --pretty=format:"$%H | %an | %cd | %s |" --name-status '+self.commitIdStart+'...'
        commmandstr ='git log  --no-merges --pretty=format:"$%H | %an | %cd | %s |" --name-status '+self.commitIdStart+'...'+ self.commitIdEnd
        print(commmandstr)
        rev = subprocess.Popen(commmandstr,shell=True,stdout=subprocess.PIPE)
        self.revList = rev.stdout.readlines()
        rev.kill()
        
    def __makeExcelPath(self):
        '''
        生成releaseNote的文件夹以及文件路径
        '''
        nowTime=datetime.datetime.now().strftime('%Y%m%d,%H-%M-%S')#现在
        releasedir=self.path+'\\releaseNote'
        if(not os.path.exists(releasedir)):#文件夹不存在，创建文件夹
            os.makedirs(releasedir)
            print('releaseNote文件夹路径不存在，自动创建文件夹:\n'+releasedir)
        self.excelPath = releasedir+'\\releaseNote('+self.version+')['+nowTime+'].xlsx'
    
    def __setVersion(self):
        '''
        获取版本号
        '''
        self.version=input("请输入版本号:")
        
    def __makeExcel(self):
        '''
        将提交记录生成excel
        '''
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = 'releaseNoteSheet'
        #添加列名
        sheet.cell(row=1, column=1, value='序号')
        sheet.cell(row=1, column=2, value='git号')
        sheet.cell(row=1, column=3, value='责任人')
        sheet.cell(row=1, column=4, value='提交时间')
        sheet.cell(row=1, column=5, value='提交记录内容')
        sheet.cell(row=1, column=6, value='文件操作类型(增删改)')
        sheet.cell(row=1, column=7, value='文件路径')
        #循环获取每次提交的文件列表
        line =''
        for cmdline in self.revList:
            line =line+str(cmdline, encoding = self.code)
        commits = line.split("$")[1:]#根据输出切割,去掉最左边的空格
        commitlen = len(commits)
        print('共获取到'+str(commitlen)+'条提交记录...')
        i=2;#第一行为表头
        index=1;#序号从1开始
        for commit in commits:
            j=1;#列从1开始
            sheet.cell(row=i, column=j, value=index)
            index=index+1;
            j=j+1;
            messages = commit.split("|")#根据format输出切割
            for message in messages:
                #处理文件路径
                if(j==6):
                    resourcepath = message.split("\n")
                    for rsp in resourcepath:
                        if(rsp.strip()!=''):#注意空行问题
                            #换行
                            operate = rsp[0:1]
                            rpath= rsp[1:].lstrip()
                            #操作类型 A M D
                            sheet.cell(row=i, column=j, value=operate);
                            j=j+1;
                            #文件路径
                            sheet.cell(row=i, column=j, value=rpath);
                            j=j-1;
                            i=i+1;
                #提交时间处理
                elif(j==4):
                    # Mon Jul 23 09:36:01 2018 +0800,不要去掉前后的空格
                    date = datetime.datetime.strptime(message,' %a %b %d %H:%M:%S %Y %z ')
                    sheet.cell(row=i, column=j, value=date.strftime("%Y-%m-%d %H:%M:%S"))
                else:
                    sheet.cell(row=i, column=j, value=message)
                #换列
                j=j+1;
            i=i+1;#行
        wb.save(self.excelPath)
    
    def executeGit2Excel(self):
        print('欢迎使用releaseNote生成工具，请按照提示进行操作 @Author luca')
        print('************************************************************')
        self.path=input("请输入git项目文件夹路径:")
        self.__getCommitList()
        self.commitIdStart=input("请输入起始提交编号:")
        self.commitIdEnd=input("请输入结束提交编号:")
        self.__getResultList()
        self.__setVersion()
        self.__makeExcelPath()
        print('正在获取git提交数据...')
        self.__makeExcel()
        print('生成releaseNote结束，生成路径为'+self.excelPath)
        
if __name__ == "__main__":
    t = Git2Excel()
    t.executeGit2Excel()
    input('输入q关闭脚本:')

                