' 校园网助手 启动器（无控制台窗口）
Set sh = CreateObject("WScript.Shell")
script = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & _
         Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\")) & "campus_net.ps1"""
sh.Run script, 0, False
