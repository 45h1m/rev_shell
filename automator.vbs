Set WshShell = WScript.CreateObject("WScript.Shell")

WScript.Sleep 1000

WshShell.Run "chrome.exe", 9

WScript.Sleep 3000

WshShell.SendKeys "https://eol2022.org/CitizenFeedback?id=FB_79528#/"

WScript.Sleep 200

WshShell.SendKeys "{ENTER}"

WScript.Sleep 8000

Do While True

a = InputBox("0[Quit]     1[Continue]","Enter Value",1,,500)

If a < 1 Then
    Exit Do
End If

WScript.Sleep 1000

WshShell.SendKeys "%{TAB}"

WScript.Sleep 2000

WshShell.SendKeys "^(c)"

WScript.Sleep 200
      
WshShell.SendKeys "{DOWN}"

WScript.Sleep 200

WshShell.SendKeys "%{TAB}"

WScript.Sleep 200

x=0
Do While x<12
    WshShell.SendKeys "{TAB}"
    WScript.Sleep 150
    
    x=x+1
Loop

WshShell.SendKeys "^V{ENTER}"

WScript.Sleep 200

WshShell.SendKeys "{TAB}"

WScript.Sleep 200

WshShell.SendKeys "^V{ENTER}"

WScript.Sleep 200

WshShell.SendKeys "{TAB}"

r = int(rnd*10) + 1
x = 0
Do While x<r
    WshShell.SendKeys "{DOWN}"
    WScript.Sleep 150
    
    x=x+1
Loop

WshShell.SendKeys "{TAB}"

r = int(rnd*2) + 1
x = 0
Do While x<r
    WshShell.SendKeys "{DOWN}"
    WScript.Sleep 200
    
    x=x+1
Loop

WshShell.SendKeys "{TAB}"

r = int(rnd*78) + 17

WScript.Sleep 200

WshShell.SendKeys r

WScript.Sleep 200

WshShell.SendKeys "{TAB}"

r = int(rnd*6) + 1
x = 0
Do While x<r
    WshShell.SendKeys "{DOWN}"
    WScript.Sleep 50
    
    x=x+1
Loop

WshShell.SendKeys "{TAB}"

x = 0
Do While x<33
    WshShell.SendKeys "{DOWN}"
    WScript.Sleep 50
    
    x=x+1
Loop

WshShell.SendKeys "{ENTER}"

WScript.Sleep 200

WshShell.SendKeys "{ENTER}"

WScript.Sleep 1000

WshShell.SendKeys "{TAB}"

WScript.Sleep 500

WshShell.SendKeys "{DOWN}"

WScript.Sleep 500

WshShell.SendKeys "{DOWN}"

WScript.Sleep 200

WshShell.SendKeys "{TAB}"

WScript.Sleep 200

WshShell.SendKeys "{TAB}"

WScript.Sleep 200

WshShell.SendKeys "{TAB}"

WScript.Sleep 200

r = int(rnd*6) + 1
x = 0
Do While x<r
    WshShell.SendKeys "{RIGHT}"
    WScript.Sleep 20
    
    x=x+1
Loop

WScript.Sleep 200

WshShell.SendKeys "{TAB}"

WScript.Sleep 200

WshShell.SendKeys "{ENTER}"

WScript.Sleep 2000

x = 0
Do While x<22
    WshShell.SendKeys "{TAB}"
    WScript.Sleep 50
    
    r = int(rnd*4) + 1
    y=0
    Do While y<r
        WshShell.SendKeys "{RIGHT}"
        WScript.Sleep 50
        y=y+1
    Loop

    x=x+1
Loop

WshShell.SendKeys "{TAB}"

WScript.Sleep 5000

WshShell.SendKeys "{ENTER}"

WScript.Sleep 4000

' WshShell.SendKeys "{TAB}"

' WScript.Sleep 50

' WshShell.SendKeys "{TAB}"

' WScript.Sleep 50

' WshShell.SendKeys "{TAB}"

' WScript.Sleep 50

' WshShell.SendKeys "{ENTER}"

' WScript.Sleep 4000

' WshShell.SendKeys "^R{ENTER}"

WScript.Sleep 3000

Loop

WScript.Quit