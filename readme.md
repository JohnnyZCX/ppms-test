# 省级系统自动化验证

## 一、环境搭建

### 1、安装谷歌浏览器

### 2、python安装

python版本==3.10

依赖安装

```text
打开Terminal终端窗口输入命令：pip3 install -r requirements.txt
```

## 二、程序说明

### 1、检查过程

1. 登录各个省级系统平台，检查每个页面中某个元素是否存在

2. 页面中存在指定元素则输出页面检查结果正常

3. 页面中找不到指定元素，则输出页面检查异常

4. 所有页面检查结果输出到excel文件中，一个sheet表代表一个省级系统

### 2、结果存储

程序会存储两份文件，xxx.log文件和省级功能巡检.xlsx文件。前者供脚本开发人员调试和查看报错用。后者供每日巡检结果汇报用。

注意在程序运行过程中不要直接打开省级功能巡检.xlsx文件，复制后打开或运行结束后打开，避免出现程序写入权限问题。

### 3、异常处理

1、浏览器开发工具断连。系统休眠或网络问题可能导致断连，程序会重启浏览器后，重新执行单个用户的检查工作，最多尝试五次，若失败则退出程序。

2、其它异常。若出现其它未预料到的异常，则重新执行单个用户的检查工作，最多尝试五次，若失败则退出程序。
