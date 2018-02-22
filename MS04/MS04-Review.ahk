
#ifwinExist, 复查 3.0 ; 上下文相关 
{
    F1:: 
    MouseGetPos, xpos, ypos
    MouseClickDrag, L, , , 840, ypos 
    clipboard =
    sendinput ^c
    ClipWait 
    SetControlDelay -1 
    ControlClick, , 复查 3.0,,,,NA x120 y120 
    MouseMove, xpos , ypos, 0
    Return
}

