This is my portfolio site.
www.marchenko-profile.com


Set WshShell = WScript.CreateObject("WScript.Shell")
interval = 120000 ' 2 минуты

Do
    WshShell.SendKeys "{NUMLOCK}"
    WScript.Sleep interval
Loop
