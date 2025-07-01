' 此脚本用于实现通过命令行一键登录 SSH 服务器
' This script is used to achieve one-click SSH server login via the command line
' このスクリプトは、コマンドラインを通じて SSH サーバーにワンクリックでログインするために使用されます。
Dim WshShell 

' 创建一个 WScript.Shell 对象，用于执行系统命令和模拟键盘输入
' Create a WScript.Shell object to execute system commands and simulate keyboard input
' システムコマンドの実行とキーボード入力のシミュレーションに使用する WScript.Shell オブジェクトを作成します。
Set WshShell=WScript.CreateObject("WScript.Shell") 

' 运行命令提示符窗口
' Run the command prompt window
' コマンドプロンプトウィンドウを起動します。
WshShell.Run "cmd.exe"

' 暂停脚本执行 50 毫秒，等待命令提示符窗口打开
' Pause the script execution for 50 milliseconds to wait for the command prompt window to open
' コマンドプロンプトウィンドウが開くのを待つために、スクリプトの実行を 50 ミリ秒間一時停止します。
WScript.Sleep 50

' 向命令提示符窗口发送 SSH 登录命令
' Send the SSH login command to the command prompt window
' コマンドプロンプトウィンドウに SSH ログインコマンドを送信します。
WshShell.SendKeys "ssh [user]@[hostname]"

' 模拟按下回车键，执行 SSH 登录命令
' Simulate pressing the Enter key to execute the SSH login command
' Enter キーを押す動作をシミュレートして、SSH ログインコマンドを実行します。
WshShell.SendKeys "{ENTER}"

' 暂停脚本执行 150 毫秒，等待 SSH 登录提示
' Pause the script execution for 150 milliseconds to wait for the SSH login prompt
' SSH ログインプロンプトが表示されるのを待つために、スクリプトの実行を 150 ミリ秒間一時停止します。
WScript.Sleep 150

' 向命令提示符窗口发送 SSH 登录密码
' Send the SSH login password to the command prompt window
' コマンドプロンプトウィンドウに SSH ログインパスワードを送信します。
WshShell.SendKeys "[password]"

' 模拟按下回车键，完成 SSH 登录
' Simulate pressing the Enter key to complete the SSH login
' Enter キーを押す動作をシミュレートして、SSH ログインを完了します。
WshShell.SendKeys "{ENTER}"