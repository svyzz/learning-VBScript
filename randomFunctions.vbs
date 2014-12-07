Option Explicit

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


Public Function generateRandomNumber(ByVal min, ByVal max)
	If (IsNumeric(min) and IsNumeric(max)) and (min < max) Then
		Randomize
		generateRandomNumber = Int((max - min + 1) * Rnd + min)
	Else
		generateRandomNumber = "Error!"
	End If
End Function


Public Function fortyTwo(ByVal character)
	Dim i, characterString
	For i = 0 to 41
		characterString = characterString & character
	Next
	fortyTwo = characterString
End Function


Public Function isLeapYear(ByVal year)	
	If DatePart("y", Cdate("March 1, " & year)) > 60 Then
		isLeapYear = True
	Else
		isLeapYear = False
	End If
End Function
