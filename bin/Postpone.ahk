#Persistent

Loop
{
    if WinExist("Windows Update", "&Postpone")
    {
        WinActivate, Windows Update, &Postpone
        Send, {Enter}
    }
    Sleep, 10000
}