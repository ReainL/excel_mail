# xlsemail

excel邮件发送统一模板。

# 部署通用邮件发送


 1 检查本机环境python3,pip,pandas
    如无则安装python3 https://www.python.org/downloads/windows/
    修改环境变量
        计算机 属性 高级系统设置 环境变量 选择“系统变量”窗口下面的"Path"
        python  C:\Python34;
    安装pip
        进入 C:\Python34\Script目录下执行pip.exe
    修改环境变量
        计算机 属性 高级系统设置 环境变量 选择“系统变量”窗口下面的"Path"
        pip     C:\Python34\Scripts;

    检查:->cmd->python
        ->cmd->pip
    安装pip install pandas


2 建立desc、src文件夹目录,desc目录存放发送结果,src存放需要群发的excel

3 执行xlsemail.py文件


