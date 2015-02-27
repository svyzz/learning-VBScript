Option Explicit

Dim objResult

Set objShell = WScript.CreateObject("WScript.Shell")    
infiniteLoop = 0

'Initialize an infinite loop that toggles Number lock on and off (together) every 10 seconds!
'Hopefully this keeps your machine open and prevents automated screen-locking kicking in.
Do While infiniteLoop = 0
  objResult = objShell.sendkeys("{NUMLOCK}{NUMLOCK}")
  Wscript.Sleep (10000)
Loop
