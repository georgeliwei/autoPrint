# 说明
    一个简单的工具，利用mirosoft onedrive网盘进行手机与pc之间的文件同步，利用windows service周期性检测文件变化，  
    通过python win32api调用windows接口实现文件的后台打印。
# 使用的技术
    python win32api
    Windows service
    pyinstaller将python脚本转换为exe文件
    python wincom将word文件转换为pdf文件
    此外为了支持静默打印pdf文件，需要安装gs以及gsprint软件， 你可以在互联网上找到这两个软件。
    
#  遇到的问题
    1：安装windows service时提示异常：需要将python解释器以及python win32api的dll文件添加到windows的系统路径。
