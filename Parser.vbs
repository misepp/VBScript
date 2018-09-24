Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set args = WScript.Arguments

Dim strLine
Dim newLine
Dim pos

If args.Count = 0 Then
	WScript.echo "*** Give a source .txt file name without the suffix and try again ***"
	WScript.Quit
End If

Set objFile = objFSO.OpenTextFile(args.Item(0) + ".txt", ForReading)
Set objOutFile = objFSO.CreateTextFile("RESULT.txt",True)  

Do Until objFile.AtEndOfStream
	strLine = objFile.ReadLine
	if InStr(1, strLine, "<a href=", vbTextCompare) Then
		pos = InStr(1, strLine, "'", vbTextCompare)
		If pos > 0 Then
			newLine = Replace(strLine, "'", """", 1, 2)
			objOutFile.WriteLine(newLine)
		End If
	Else
		objOutFile.WriteLine(strLine)
	End If
Loop

objFile.Close
objOutFile.Close
