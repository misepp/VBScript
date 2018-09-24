Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set args = WScript.Arguments

Dim strLine
Dim newLine
Dim quotePos
Dim hPos
Dim hasLink

If args.Count = 0 Then
	WScript.echo "*** Give a source .txt file name without the suffix and try again ***"
	WScript.Quit
End If

Set objFile = objFSO.OpenTextFile(args.Item(0) + ".txt", ForReading)
Set objOutFile = objFSO.CreateTextFile("RESULT.txt",True)  

Do Until objFile.AtEndOfStream
	Call loopFile()
Loop

Function loopFile()
	strLine = objFile.ReadLine	
	hasLink = InStr(1, strLine, "<a href=", vbTextCompare)
	
	if hasLink > 0 Then
		quotePos = InStr(1, strLine, "'", vbTextCompare)
	End If
		
	hPos = InStr(1, strLine, "h3", vbTextCompare)
	If hPos > 0 Then
		newLine = Replace(strLine, "h3", "h2", 1)
		If hasLink > 0 Then
			newLine = Replace(newLine, "'", """", 1, 2)
			objOutFile.WriteLine(newLine)
			Exit Function
		End If
		objOutFile.WriteLine(newLine)
	ElseIf hasLink Then
		If quotePos > 0 Then
			newLine = Replace(strLine, "'", """", 1, 2)
			objOutFile.WriteLine(newLine)
		End If
	Else
		objOutFile.WriteLine(strLine)
	End If
End Function

objFile.Close
objOutFile.Close
