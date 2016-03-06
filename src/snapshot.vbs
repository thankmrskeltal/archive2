Dim FriendlyDate
Dim FriendlyTime
FriendlyDate = Replace(Date, "/", "-")
FriendlyTime = Mid(Replace(Time, ":", "-"), 1, 5)

Dim WShell
Set WShell = CreateObject("Wscript.Shell")

WShell.Run ("cmd /c move " + Wscript.Arguments(0) + " .\Backups\" + FriendlyDate + "_" + FriendlyTime + "_" + Wscript.Arguments(0))

Set WShell = Nothing