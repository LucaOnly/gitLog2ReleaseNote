本工具功能为将2个git提交间的提交记录生成releaseNote格式的excel,方便发布使用。

软件环境：python3.5

依赖包：openpyxl(生成excel使用此包) 使用pip安装

git:
将git配置到环境变量；
使用此命令兼容中文文件名：git config --global core.quotepath false

输入内容：
1.git项目文件夹路径 如：D:\git项目SourceTree\releaseManagement
2.起始id: 较早的commitid,生成时不包含此次提交记录
3.结束id:较晚的commitid，生成时包含此次提交记录
4.版本号：发布时的版本号

注意：
使用前请拉取最新代码，保证本地和远程git同步，如果使用中卡死，请重启此脚本。

