# lniHelper
execl转lni工具

适用于w3x2lni项目的物编功能，因为ini物编不如slk直观，所以用该工具预编译，将xlsx编译成ini

依赖环境jre8

使用命令：
$ java -jar LniHelper.jar "项目路径"

配合w2l命令：
test.bat:

===================================================================================================

@echo 结束war3进程  

taskkill /f /im war3.exe

@echo 调用lnihelper编译 

java -jar LniHelper.jar "D:\WarWorks\war3project\BlueFantasy\" 

@echo 调用w2l编译

w2l obj "D:\WarWorks\war3project\BlueFantasy\BlueFantasy\.w3x"

@echo 调用ydwe进入测试游戏  -windows窗口化 -fullscreen全屏

ydweconfig.exe -launchwar3 -loadfile "D:\WarWorks\war3project\BlueFantasy\BlueFantasy.w3x" -windows

===================================================================================================

excel.xlxs 和注意事项参考 example文件夹