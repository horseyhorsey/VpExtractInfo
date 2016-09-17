'Save a VP collection to JSON
'Saves name, timerinterval, x , y. Uses: exporting lamps
'tablename, name = name of collection, collection = collection from VP
Sub PrintCollectionFull(ByVal tableName, ByVal name, ByVal collection)
	dim str
	Dim Tab
	Dim FSO, oFile
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set oFile = FSO.OpenTextFile(tableName + "-" + name + ".json", 2, True)

	Tab = "    "
	str = "{" & vbCrLf & Chr(34) & name & Chr(34) & ": ["
	
	dim item
	For each item in collection:
		str = str & vbCrLf & "{" & vbCrLf & Tab & Chr(34) & "Name" & Chr(34) & ": " & Chr(34) & item.Name & Chr(34) & ","	& vbCrLf		
		str = str & Tab & Chr(34) & "Interval" & Chr(34) & ": " & item.timerinterval & "," & vbCrLf
		str = str & Tab & Chr(34) & "X" & Chr(34) & ": " & item.x & "," & vbCrLf
		str = str & Tab & Chr(34) & "Y" & Chr(34) & ": " & item.y & vbCrLf
		str = str & "}," & vbCrLf
	next
	
	str = str & "]" & vbCrLf & "}"	

	oFile.Write(str & vbCrLf)

	Set oFile = Nothing
	Set FSO = Nothing
End Sub	

Sub PrintCollectionJSON(ByVal tableName, ByVal name, ByVal collection)
	dim str
	Dim Tab
	Dim FSO, oFile
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set oFile = FSO.OpenTextFile(tableName + "-" + name + ".json", 2, True)

	Tab = "    "
	str = "{" & vbCrLf & Chr(34) & name & Chr(34) & ": ["
	
	dim item
	For each item in collection:
		str = str & vbCrLf & "{" & vbCrLf & Tab & Chr(34) & "Name" & Chr(34) & ": " & Chr(34) & item.Name & Chr(34) & ","	& vbCrLf		
		str = str & Tab & Chr(34) & "Interval" & Chr(34) & ": " & item.timerinterval & "," & vbCrLf
		str = str & "}," & vbCrLf
	next
	
	str = str & "]" & vbCrLf & "}"	

	oFile.Write(str & vbCrLf)

	Set oFile = Nothing
	Set FSO = Nothing
End Sub	

'Save a VP collection to JSON
'Saves just the name
Sub PrintCollectionSafe(ByVal tableName, ByVal name, ByVal collection)
	dim str
	Dim Tab
	Dim FSO, oFile
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set oFile = FSO.OpenTextFile(tableName + "-" + name + ".json", 2, True)

	Tab = "    "
	str = "{" & vbCrLf & Chr(34) & name & Chr(34) & ": ["
	
	dim item
	For each item in collection:
		str = str & vbCrLf & "{" & vbCrLf & Tab & Chr(34) & "Name" & Chr(34) & ": " & Chr(34) & item.Name & Chr(34) & ","	& vbCrLf
		str = str & "}," & vbCrLf
	next
	
	str = str & "]" & vbCrLf & "}"	

	oFile.Write(str & vbCrLf)

	Set oFile = Nothing
	Set FSO = Nothing
End Sub

Sub PrintCollectionLedShowJSON(ByVal tableName, ByVal collection)
	dim str
	Dim Tab
	Dim FSO, oFile
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set oFile = FSO.OpenTextFile(tableName + "-LedShow.json", 2, True)

	Tab = "    "
	str = "{" & vbCrLf & Chr(34) & "PlayfieldImage" & Chr(34) & ": " & Chr(34) & "pf.jpg" & Chr(34) & "," & vbCrLf 
	str = str & Chr(34) & "PlayfieldToLedsScale" & Chr(34) & ": " & 0.25 & "," & vbCrLf
	str = str & Chr(34) & "Leds" & Chr(34) & ": [" & vbCrLf
	
	dim item
	For each item in collection:
		str = str & "{" & vbCrLf
		str = str & Tab & Chr(34) & "Id" & Chr(34) & ": " & item.timerinterval & "," & vbCrLf
		str = str & Tab & Chr(34) & "Name" & Chr(34) & ": " & Chr(34) & item.Name & Chr(34) & "," & vbCrLf		
		str = str & Tab & Chr(34) & "IsSingleColor" & Chr(34) & ": " & "true" & "," & vbCrLf
		str = str & Tab & Chr(34) & "SingleColor" & Chr(34) & ": " & Chr(34) & "#FFADFF2F" & Chr(34) & "," & vbCrLf		
		str = str & Tab & Chr(34) & "LocationX" & Chr(34) & ": " & 17.702 & "," & vbCrLf
		str = str & Tab & Chr(34) & "LocationY" & Chr(34) & ": " & 19.702 & "," & vbCrLf
		str = str & Tab & Chr(34) & "Angle" & Chr(34) & ": " & 0.0 & "," & vbCrLf
		str = str & Tab & Chr(34) & "Scale" & Chr(34) & ": " & 1.0 & "," & vbCrLf
		str = str & Tab & Chr(34) & "Shape" & Chr(34) & ": " & 1 & "," & vbCrLf
		str = str & "}," & vbCrLf
	next
	
	str = str & "]," & vbCrLf & "}"	

	oFile.Write(str & vbCrLf)

	Set oFile = Nothing
	Set FSO = Nothing
End Sub	

'prType = type of PR ..eg PRLamps , PRCoils, PRSwitches
Function  PrintCollectionYaml(ByVal tableName, ByVal prType, ByVal collection)

	numPrefix = "number: "
	If StrComp(prType,"PRSwitches",vbTextCompare) = 0 Then
		numPrefix = numPrefix & "S"
	ElseIf StrComp(prType,"PRlamps",vbTextCompare) = 0 Then
		numPrefix = numPrefix & "L"
	ElseIf StrComp(prType,"PRCoils",vbTextCompare) = 0 Then
		numPrefix = numPrefix & "C"	
	End If
	
	if StrComp(numPrefix,"number: ",vbTextCompare) =  0 Then	
		msgbox "PrintCollectionYaml failed to export" & vbCrLf & "prType: " & pyType & " Doesn't match PRLamps, PRSwitches, PRCoils"
	Else
			dim str
		Dim Tab
		Dim FSO, oFile
		Set FSO = CreateObject("Scripting.FileSystemObject")
		Set oFile = FSO.OpenTextFile(tableName & prType & ".yaml", 2, True)

		Tab = "    "
		str = prType & ":" & vbCrLf

		dim item
		For each item in collection:
			str = str & Tab & item.Name & ":" & vbCrLf & Tab & Tab & numPrefix & item.timerinterval	& vbCrLf
			
			If StrComp(prType, "PRlamps", vbTextCompare) = 1 Then
				str = str & Tab & Tab & ballsearch & vbCrLf
			End If		
		next	

		oFile.Write(str & vbCrLf)

		Set oFile = Nothing
		Set FSO = Nothing
		
	End iF

End Function 	
