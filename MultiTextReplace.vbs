Const ForReading = 1
Const ForWriting = 2
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFSO.OpenTextFile("C:\text.txt", ForReading)

strText = objFile.ReadAll
objFile.Close
strNewText = Replace(strText, "hola", "Hoy es - primero de " &Date)
strNewText1 = Replace(strNewText, "no", "casi")
strNewText2 = Replace(strNewText1, "12", " por nada")

Set objFile = objFSO.OpenTextFile("C:\text.txt", ForWriting)
objFile.WriteLine strNewText2 

objFile.Close
