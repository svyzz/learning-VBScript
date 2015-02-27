Option Explicit

''''''''''''''''''''''''''''''''''''''''''
' Author: A very bored Arun Venkatram who attempts
' to learn VBScript by writing useless games!

' Purpose: Learn VBScript and also manage a useful
' distraction in the process?

' Notes to people reading the source code, DO NOT CHEAT!
''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''
' Variables and Constants
''''''''''''''''''''''''''''''''''''''''''
Public player, startTime, endTime, elapsed, allItems, allPlayers, fastestTime, fastestPlayer
Public objFSO, logFile, objFile
Public fastestClick: fastestClick = False
Public newPlayer: newPlayer = False

Public d: Set d = CreateObject("Scripting.Dictionary")

Const TEMP_FOLDER = 2
Set objFSO=CreateObject("Scripting.FileSystemObject")
Public tempFolder: tempFolder = objFSO.GetSpecialFolder(TEMP_FOLDER)


''''''''''''''''''''''''''''''''''''''''''
' Get the user to play and record their time
''''''''''''''''''''''''''''''''''''''''''
' Let the game begin!
player = InputBox("Enter Your Name!" & vbNewLine & vbNewLine & "Note that user names are CASE-SENSITIVE.", "Fastest Click Ever.")
If player = "" Then
	MsgBox "No Name, No Game!"
	Wscript.Quit
End If

' Calculate time spent to click on buttons
MsgBox "Click to BEGIN.", vbExclamation
startTime = Timer

MsgBox "Click to END!", vbExclamation
elapsed = Timer - startTime


''''''''''''''''''''''''''''''''''''''''''
' Get a read/write handle on the log file in question
''''''''''''''''''''''''''''''''''''''''''
logFile = tempFolder & "\FastestClick.log"

If objFSO.FileExists(logFile) Then
	' Read from file and populate dictionary
	readFromLogFile()
	If Not d.exists(player) Then
		newPlayer = True
		d.Add player, 42
	End If
	
	SortDictionary d, 2
	allItems = d.Items
	fastestTime = allItems(0)
	allPlayers = d.Keys
	fastestPlayer = allPlayers(0)
	
	' This works only if midnight does not occur between the two clicks.
	' How about that for a frickin bug?
	If elapsed < cDbl(fastestTime) Then
		MsgBox "Fastest time EVER! Congratulations!" & vbNewLine & vbNewLine & "You clocked an impressive " & elapsed & " seconds, beating " & fastestPlayer & " who held the previous record for " & fastestTime & " seconds", vbInformation
		d.Item(player) = elapsed
	ElseIf elapsed < cDbl(d.Item(player)) and Not newPlayer Then
		MsgBox "Personal best!" & vbNewLine & vbNewLine & "You just set a new personal record of " & elapsed & " seconds!" & vbNewLine & vbNewLine & "The all time RECORD is " & fastestTime & " seconds, set by " & fastestPlayer, vbInformation
		d.Item(player) = elapsed
	ElseIf newPlayer Then
		MsgBox "Good Job! You clocked " & elapsed & " seconds" & vbNewLine & vbNewLine & "The all time RECORD is " & fastestTime & " seconds, set by " & fastestPlayer, vbInformation
		d.Item(player) = elapsed
	Else
		MsgBox "Good Job! You clocked " & elapsed & " seconds" & vbNewLine & vbNewLine & "Your personal best is " & d.Item(player) & " seconds" & vbNewLine & vbNewLine & "The all time RECORD is " & fastestTime & " seconds, set by " & fastestPlayer, vbInformation
	End If
	writeToLogfile()
Else
	' Write the dictionary to log file if one does not already exist for this session
	' This block automatically implies that the user is the first to play the game!
	MsgBox "Fastest time EVER! Congratulations!", vbInformation
	d.Add player, elapsed
	writeToLogFile()
End If


''''''''''''''''''''''''''''''''''''''''''
' Subroutines
''''''''''''''''''''''''''''''''''''''''''

Sub readFromLogfile()
	' TO-DO: Check if file is empty!
	Dim line, userResults
	Dim row: row = 0
	Set objFile = objFSO.OpenTextFile(logFile)
	
	Do Until objFile.AtEndOfStream
		line = objFile.Readline
		userResults = Split(line, ",")
		d.Add userResults(0), cDbl(userResults(1))
		row = row + 1
	Loop
	
	objFile.Close
End Sub


Sub writeToLogFile()
	Dim objFile, k
	Set objFile = objFSO.CreateTextFile(logFile, True)

	For Each k in d.Keys
		objFile.Write k
		objFile.Write ","
		objFile.Write d(k) & vbCrLf
	Next
	
	objFile.Close
End Sub


Function SortDictionary(objDict, intSort)
   ' This defines the value of intSort. 1 sorts by key and 2 sorts by value
   Const dictKey  = 1
   Const dictItem = 2
 
   Dim strDict()
   Dim objKey
   Dim strKey,strItem
   Dim X,Y,Z
 
   ' Get the dictionary count
   Z = objDict.Count
 
   ' We need more than one item to warrant sorting
   If Z > 1 Then
     ' Create an array to store dictionary information
     ReDim strDict(Z,2)
     X = 0
     ' Populate the string array
     For Each objKey In objDict
         strDict(X,dictKey)  = CStr(objKey)
         strDict(X,dictItem) = CStr(objDict(objKey))
         X = X + 1
     Next
 
     ' Perform a a shell sort of the string array
     For X = 0 To (Z - 2)
       For Y = X To (Z - 1)
         If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
             strKey  = strDict(X,dictKey)
             strItem = strDict(X,dictItem)
             strDict(X,dictKey)  = strDict(Y,dictKey)
             strDict(X,dictItem) = strDict(Y,dictItem)
             strDict(Y,dictKey)  = strKey
             strDict(Y,dictItem) = strItem
         End If
       Next
     Next
 
     ' Erase the contents of the dictionary object
     objDict.RemoveAll
 
     ' Repopulate the dictionary with the sorted information
     For X = 0 To (Z - 1)
       objDict.Add strDict(X,dictKey), strDict(X,dictItem)
     Next
 
   End If
 End Function
