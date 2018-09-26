Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set args = WScript.Arguments

Dim strLine
Dim newLine
Dim quotePos
Dim hasLink
header1 = "Bugs"
header2 = "Project Specific"
header3 = "Generic"

If args.Count = 0 Then
	WScript.echo "*** Give a source text file name without any suffix and try again ***"
	WScript.Quit
End If

Set objFile = objFSO.OpenTextFile(args.Item(0), ForReading)
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
		 
	If InStr(1, strLine, "h2", vbTextCompare) Then
		newLine = Replace(strLine, "h2", "h3", 1)
		If InStr(1, newLine, "Bug", vbTextCompare) Then
			newLine = Replace(newLine, "Bug", header1, 1)
		End If
		If InStr(1, newLine, "Epic", vbTextCompare) Then
			newLine = Replace(newLine, "Epic", header2, 1)
		End If
		If InStr(1, newLine, "Story", vbTextCompare) Then
			newLine = Replace(newLine, "Story", header3, 1)
		End If
		If hasLink > 0 Then
			newLine = Replace(newLine, "'", """", 1, 2)
			objOutFile.WriteLine(newLine)
			Exit Function
		End If
		objOutFile.WriteLine(newLine)
	ElseIf hasLink > 0 Then
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
