Option Explicit

Dim objHTTP, rex, score, matches, ptrn

'Change the teams here based on the scores for the match you want
Dim team1: team1 = "Australia"
Dim team2: team2 = "India"

'Instantiate an object for sending HTTP requests and request the live-scores RSS feed from CricInfo
Set objHTTP = CreateObject("MSXML2.XMLHTTP.3.0")
objHTTP.open "GET", "http://static.espncricinfo.com/rss/livescores.xml", False
objHTTP.send

If objHTTP.Status = 200 Then
	Set rex = New RegExp
	'Instantiating the pattern variant here. Perhaps a future iteration could ASK users for team1 and team2?
	ptrn = "(<title>)" & "(" & team1 & ")" & "(.*)" & "(" & team2 & ")"& "(.*)(</title>)"
	With rex
		.Pattern    = ptrn
		.IgnoreCase = True
		.Global     = False
	End With

	Set matches = rex.Execute(objHTTP.responseText)
	If matches.Count = 1 Then
		'This was the best way to include details about which team's presently batting and scores from the second innnigs
		score = matches.Item(0).Submatches(1) & matches.Item(0).Submatches(2) & matches.Item(0).Submatches(3) & matches.Item(0).Submatches(4)
		MsgBox score
	End If
Else
	MsgBox "You MAY have connectivity issues. Please check your internet connection and try again.", VBExclamation
End If
