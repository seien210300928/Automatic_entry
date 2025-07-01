

## Language
[中文](#中文)
[English](#english)
[日本語](#日本語)

---
### 中文
# 自动输入

自动打开软件并模拟键盘输入

使用前请确保系统支持.vbs（Visual Basic Script）脚本文件

<code class="language-vba" style="font-family: monospace"><pre>
' 声明变量
Dim WshShell 

' 创建一个 WScript.Shell 对象，用于执行系统命令和模拟键盘输入
Set WshShell=WScript.CreateObject("WScript.Shell") 

' 运行指定程序
WshShell.Run "[Program]"

' 等待指定时间,以毫秒为单位（一般来说是必须的，因为程序需要启动时间）
WScript.Sleep [Time.ms]

' 键入指定内容
WshShell.SendKeys "[Elements]"

' 模拟回车键
WshShell.SendKeys "{ENTER}"

' 等待一会儿继续输入
WScript.Sleep [Time.ms]
WshShell.SendKeys "[Elements]"
WshShell.SendKeys "{ENTER}"
......
</pre></code>
各种非语句模拟举例  
1. 单独使用按键{[Button] [times]}
<code class="language-vba" style="font-family: monospace"><pre>
WshShell.SendKeys "{ENTER}"                 ' 回车键
WshShell.SendKeys "{TAB}"                   ' 制表键
WshShell.SendKeys "{BACKSPACE}"             ' 退格键
WshShell.SendKeys "{DELETE}"                ' 删除键
WshShell.SendKeys "{SPACE}"                 ' 空格键
WshShell.SendKeys "{ESC}"                   ' ESC键
WshShell.SendKeys "{UP}"                    ' 上箭头
WshShell.SendKeys "{DOWN}"                  ' 下箭头
WshShell.SendKeys "{LEFT}"                  ' 左箭头
WshShell.SendKeys "{RIGHT}"                 ' 右箭头
WshShell.SendKeys "{F1}"                    ' F1键（F1-F12同理）
WshShell.SendKeys "{CAPSLOCK}"              ' 切换 Caps Lock 状态
WshShell.SendKeys "{NUMLOCK}"               ' 切换Num Lock状态
WshShell.SendKeys "{SPACE}"                 ' 空格键
</pre></code>

2. 使用组合键`[Modifiers][Modifiers]...{[Button]}... `
如果是字母按键可将`{[Button]}`直接替换为`[Alphabet]`

    以下三个按键不能直接使用`{[Button]}`模拟，这里先展示其使用`[Modifiers]`模拟的情况
    * `+`：`Shift`键
    * `^`：`Ctrl`键
    * `%`：`Alt`键

* 一次性组合键`[Modifiers]...{[Button]}... `
<code class="language-vba" style="font-family: monospace"><pre>
WshShell.SendKeys "^C"                      ' Ctrl+C（复制）
WshShell.SendKeys "^V"                      ' Ctrl+V（粘贴）
WshShell.SendKeys "%{F4}"                   ' Alt+F4（关闭窗口）
WshShell.SendKeys "+{TAB}"                  ' Shift+Tab（反向制表）
WshShell.SendKeys "^+{ESC}"                 ' Ctrl+Shift+ESC（打开任务管理器）
WshShell.SendKeys "+2"                      ' 输入 "@"（因为按 Shift+2 是 @）
</pre></code>


* 多修饰符+多个普通键`[Modifiers]...({[Button]}...) `
<pre>
<code class="language-vba" style="font-family: monospace">
WshShell.SendKeys "+(abc)"                       ' 输入无大写锁定状态下的"ABC"
</code>
</pre>

3. 模拟特殊字符
* 某些字符因为有关键字的作用所以不能直接输入，所以只能使用`{[Button]}`模拟
<code class="language-vba" style="font-family: monospace"><pre>
WshShell.SendKeys "{{}"                    ' 输入左花括号 "{"
WshShell.SendKeys "{}}"                    ' 输入右花括号 "}"
WshShell.SendKeys "{+}"                    ' 输入加号 "+"
WshShell.SendKeys "{^}"                    ' 输入脱字符 "^"
WshShell.SendKeys "{%}"                    ' 输入百分号 "%"
WshShell.SendKeys "{~}"                    ' 输入波浪号 "~"
WshShell.SendKeys "\("                     ' 输入 "("（部分场景需配合转义）
WshShell.SendKeys "\)"                     ' 输入 ")"（部分场景需配合转义）
</pre></code>




