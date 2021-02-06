# lniHelper
execl转lni工具

适用于w3x2lni项目的物编功能，因为ini物编不如slk直观，所以用该工具预编译，将xlsx编译成ini

依赖环境jre8

使用命令：
$ java -jar LniHelper.jar "项目路径"

配合w2l命令：
test.bat:
taskkill /f /im war3.exe
java -jar LniHelper.jar "D:\WarWorks\war3project\BlueFantasy\" 
w2l obj "D:\WarWorks\war3project\BlueFantasy\BlueFantasy\.w3x"
ydweconfig.exe -launchwar3 -loadfile "D:\WarWorks\war3project\BlueFantasy\BlueFantasy.w3x"
