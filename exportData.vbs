Dim targetFile, srcFile, originFile, objFs, objCsvSrcFile, strSrc, objExcel, workbookSrc, workbookOrigin, firstStrFlag, isFirstStr, index 

'operation power
Const ForReading   = 1
Const ForAppending = 8

'is first line
firstStrFlag = true
isFirstStr = true
'telephone
index = 30

If WScript.Arguments.length < 2 Then
    WScript.Echo("请依次添加分行,总行文件")
    WScript.Quit()
End If

srcFile = WScript.Arguments.Item(0)
originFile = WScript.Arguments.Item(1)
'open file system
Set objFS  = WScript.CreateObject("Scripting.FileSystemObject")

If Not objFS.FileExists(srcFile) Then
    WScript.Echo("分行文件已存在")
    WScript.Quit()
End If
If Not objFS.FileExists(originFile) Then
    WScript.Echo("总行文件已存在")
    WScript.Quit()
End If

'excel save to csv
Set objExcel     = WScript.CreateObject("Excel.Application")
Set workbookSrc     = objExcel.Application.Workbooks.Open(srcFile, Null, True)
Set workbookOrigin     = objExcel.Application.Workbooks.Open(originFile, Null, True)
objExcel.Visible = false
srcFile = left(srcFile, instrrev(srcFile, "."))+"csv"
originFile = left(originFile , instrrev(originFile , "."))+"csv"
If objFS.FileExists(srcFile) Then
    Call objFS.DeleteFile(srcFile)
End If
If objFS.FileExists(originFile) Then
    Call objFS.DeleteFile(originFile)
End If
Call workbookSrc.SaveAs(srcFile, 6)
Call workbookOrigin.SaveAs(originFile , 6)
objExcel.Application.DisplayAlerts = False
Call objExcel.Quit()

'output file url
targetFile = left(srcFile, instrrev(srcFile, "\"))+"output.csv"
If objFS.FileExists(targetFile) Then
    Call objFS.DeleteFile(targetFile)
    objFs.CreateTextFile(targetFile)
End If

'reading file
Set objCsvSrcFile= objFs.OpenTextFile(srcFile, ForReading)
Do Until objCsvSrcFile.AtEndOfStream
strSrc = objCsvSrcFile.ReadLine
	If firstStrFlag=true And strSrc="卡号" Then
		index = 7
	End If
	readOriginFile strSrc
firstStrFlag = false
Loop
objCsvSrcFile.Close
Call objFS.DeleteFile(srcFile)
Call objFS.DeleteFile(originFile)
MsgBox "操作成功"

Function readOriginFile(str)
Dim strOrigin, s, objCsvOriginFile, fw
Set objCsvOriginFile = objFs.OpenTextFile(originFile, ForReading)
Do Until objCsvOriginFile.AtEndOfStream
strOrigin = objCsvOriginFile.ReadLine
s = split(strOrigin, ",")
If (isFirstStr=true) Or (inStr(s(index), str)>0)  Then 
	isFirstStr = false
	Set fw = objFs.openTextFile(targetFile, ForAppending, true)
	fw.WriteLine(strOrigin)
	fw.Close
End If
Loop
objCsvOriginFile.Close
End Function


