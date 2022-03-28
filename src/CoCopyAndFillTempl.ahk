#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Author: Coande
; Date: 2022-03-27

;Gui, GuiName:New, , 填充模板式复制 ;创建一个新的名为 GuiName 的窗口
Gui +AlwaysOnTop ; 固定在最上面

Gui, Add, Text, w400, 选项： ;选项的标签
Gui, Add, CheckBox, wp h20 visToTrimValueId, 去掉复制内容的前后空格换行 ; 去掉空格选项
GuiControl,, isToTrimValueId , %True% ; 默认勾选
Gui, Add, Text, w100, 占位符 ; 占位符标签
Gui, Add, Edit, vplaceholderValueId w100, {} ; 占位符的值

Gui, Add, Text, w400, 模板： ;模板标签
Gui, Add, Edit, vtemplValueId r8 wp, @ExcelProperty("{}")`nString {}`;`n`n ; 模板默认值

Gui, Add, Text, w400, 处理中： ;处理中标签
Gui, Add, Edit, vcurrentValueId r8 wp ReadOnly ; 模板默认值

Gui, Add, Text, voutputLabelId wp, 结果： ; 输出结果标签
Gui, Add, Edit, voutputValueId r24 wp, ctrl + c 复制文本看效果 ; 输出结果内容

Gui, Show ; 显示窗口


; 查找模板中的变量个数
GetTemplVarCount(str, startingPos:=1, initCount:=0) {
    global placeholderValueId
    GuiControlGet, placeholderValueVal,, placeholderValueId ; 获取占位符到 placeholderValueVal 变量
    pos := InStr(str, placeholderValueVal, , startingPos)
    if (pos) {
        initCount := initCount + 1
        return GetTemplVarCount(str, pos + 1, initCount)
    } else {
        return initCount
    }
}

; 填充模板
FillTempl() {
    static currentTempl
    static isFirst := True
    global templValueId ; 声明使用的是外部的全局变量
    global outputValueId ; 声明使用的是外部的全局变量
    global currentValueId ; 声明使用的是外部的全局变量

    if (!currentTempl) { ; 赋初始值
        GuiControlGet, currentTempl,, templValueId ; 获取模板内容到 currentTempl 变量
    }

    ; 获取剪切板中的内容
    copyText := Clipboard

    ; 获取是否需要去掉前后空格的选项
    GuiControlGet, isToTrimValueVal,, isToTrimValueId
    if (isToTrimValueVal) {
        copyText := Trim(copyText, " `t`n`r")
    }
    
    ; 替换操作
    GuiControlGet, placeholderValueVal,, placeholderValueId ; 获取占位符到 placeholderValueVal 变量
    replacedStr := StrReplace(currentTempl, placeholderValueVal, copyText, replaceCount, 1) ; 替换1个，并把替换数量输出到 replaceCount 变量
    currentTempl := replacedStr

    ; 获取剩余变量个数
    remainVarCount := GetTemplVarCount(currentTempl)

    if (!remainVarCount) { ; 全部替换完后保存到输出
        ; 更新处理中的值为空
        GuiControl,, currentValueId ,  ; 设置输出结果

        GuiControlGet, outputValueVal,, outputValueId ; 获取已存在结果内容到 outputValueVal 变量
        if (isFirst) {
            isFirst := False
            GuiControl,, outputValueId , %currentTempl% ; 设置输出结果
        } else {
            GuiControl,, outputValueId , %outputValueVal%%currentTempl% ; 设置输出结果
        }
        ; 设置回初始值，依次循环
        GuiControlGet, currentTempl,, templValueId ; 获取模板内容到 currentTempl 变量
    } else {
        ; 更新处理中的值
        GuiControl,, currentValueId , %currentTempl% 
    }
}

; 粘贴板变更时回调函数
ClipChanged(Type) {
    if (Type = 1) { ; 内容为文本
        ;GuiControlGet, outputValueVal,GuiName:, outputValueId ; 获取模板内容
        ;MsgBox % matchCount . "测试"
        FillTempl()

        ToolTip 已复制 ; 显示复制反馈
        SetTimer, RemoveToolTip, -600
    } else { ; 内容为非文本
        ToolTip 复制非文本，已忽略
        SetTimer, RemoveToolTip, -600
    }
}

OnClipboardChange("ClipChanged")

; 这是一个标签（label），用于关闭 Tooltip
RemoveToolTip:
ToolTip ; 关闭提示
return