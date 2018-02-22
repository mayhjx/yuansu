menu,tray, nostandard
menu,tray, add, 如何使用, menuhandler1
menu,tray, add, 重新加载, menuhandler2
menu,tray, add, 退出, menuhandler3
return

menuhandler1:
msgbox, F1: 复查快捷键`nF2: 重新加载脚本`nF3: 退出
return

menuhandler2:
reload
return

menuhandler3:
exitapp
return


#ifwinExist, 复查 3.0 
{
F1::
MouseGetPos, xpos, ypos                                        ; 获取鼠标的位置
MouseClickDrag, Left , , ,622, 0, 4 ,R
MouseMove, xpos , ypos, 0
Clipboard =                                                              ; 清空剪切板
SendInput ^c
ClipWait                                                                    ; 等待剪切板出现文本  
SetControlDelay -1                                                   ; 避免在点击时按住鼠标，减少用户对鼠标的干扰   
ControlClick, , 复查 3.0,,,,NA x150 y100               ; NA 避免激活目标窗口
Return

*F2:: Reload
*F3:: ExitApp
}

