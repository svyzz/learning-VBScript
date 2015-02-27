Option Explicit

'@Documentation: Checks to see if an array is empty! Good learning exercise and shows exactly why people moved on from using VBScript
Public Function isArrayEmpty(ByVal x)
   If IsArray(x) Then
		If Len(Join(x, "")) = 0 Then
			isArrayEmpty = True
		Else
			isArrayEmpty = False
		End If
	Else
		isArrayEmpty = "Error!"
   End If
End Function

'@Documentation: Generates a random numer! Notice the use of the keyword Randomize to generate a random seed?
Public Function generateRandomNumber(ByVal min, ByVal max)
	If (IsNumeric(min) and IsNumeric(max)) and (min < max) Then
		Randomize
		generateRandomNumber = Int((max - min + 1) * Rnd + min)
	Else
		generateRandomNumber = "Error!"
	End If
End Function

'@Documentation: Generates 42 impressions of the given character. I use them in my comments! Why 42 you ask? :)
Public Function fortyTwo(ByVal character)
	Dim i, characterString
	For i = 0 to 41
		characterString = characterString & character
	Next
	fortyTwo = characterString
End Function

'@Documentation: Checks if a given year is a leap year or not. January has 31 days and February has 29 days in a leap year.
'That makes a grand total of 61 days including the 1st of March on a leap year. Small, beautiful code!
Public Function isLeapYear(ByVal year)	
	If DatePart("y", Cdate("March 1, " & year)) > 60 Then
		isLeapYear = True
	Else
		isLeapYear = False
	End If
End Function

'@Documentation: Disco lights on your keyboard! Such awesomeness, so majestic, much fancy, very bored.
Public Sub disco()
	Set WshShell = WScript.CreateObject("WScript.Shell")
	do
		WshShell.SendKeys "{NUMLOCK}"
		WScript.Sleep 200
		WshShell.SendKeys "{CAPSLOCK}"
		WScript.Sleep 200
		WshShell.SendKeys "{SCROLLLOCK}"
		WScript.Sleep 200
	loop
End Sub
