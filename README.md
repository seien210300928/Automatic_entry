## Language
[中文](#中文)
[English](#english)
[日本語](#日本語)

---
### 中文
# 自动输入

自动打开软件并模拟键盘输入

使用前请确保系统支持.vbs（Visual Basic Script）文件

```vba
' 声明变量
Dim WshShell 

' 创建一个 WScript.Shell 对象，用于执行系统命令和模拟键盘输入
Set WshShell=WScript.CreateObject("WScript.Shell") 

’运行指定程序
WshShell.Run "[Program]"

' 以下可自由发挥，这里以Windows命令提示符为例

```