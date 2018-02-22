
#ifwinExist, 复查 3.0
{
    F1:: 
    MouseGetPos, xpos, ypos 
    MouseClickDrag, L, , ,900, ypos
    Clipboard = 
    Sendinput ^c
    ClipWait   
    SetControlDelay -1   
    ControlClick, , 复查 3.0,,,,NA x120 y120 
    MouseMove, xpos , ypos, 0
    Return
}

