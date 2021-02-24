Set objShell = WScript.CreateObject("WScript.Shell" )             
Do 
  ' Set the focus to the browser whose name starts with "Codeguru"
  objShell.AppActivate "Chrome"

  ' Send the F5 key for a refresh
  objShell.SendKeys "{F5}"

  iRetVal = objShell.Popup("Stop refreshing script?", 3, "", 4)
  If (iRetVal = 6) Then
     Exit Do
  End If

  'Wait for 5 seconds
  Wscript.Sleep 5000 
Loop
