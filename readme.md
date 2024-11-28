```vbScript
' 脚本：将当前目录中的 .doc 和 .docx 文件用 Word 打开并用 WordPad 保存为 .rtf 文件
' 转换结果保存到当前目录的 RTF 文件夹中

Dim objArgs, objFSO, objWord, objShell, sourceFolder, rtfFolder, file, tempFile, destFile, wordPadPath

' 获取传入的文件夹路径
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    MsgBox "请通过右键菜单发送文件夹到此脚本。", vbExclamation, "错误"
    WScript.Quit
End If

sourceFolder = objArgs(0)

' 检查目标路径是否为文件夹
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FolderExists(sourceFolder) Then
    MsgBox "指定的路径不是文件夹：" & vbCrLf & sourceFolder, vbExclamation, "错误"
    WScript.Quit
End If

' 创建 RTF 文件夹
rtfFolder = objFSO.BuildPath(sourceFolder, "RTF")
If Not objFSO.FolderExists(rtfFolder) Then
    objFSO.CreateFolder rtfFolder
End If

' 获取 WordPad 的路径
Set objShell = CreateObject("WScript.Shell")
wordPadPath = objShell.ExpandEnvironmentStrings("%ProgramFiles%\Windows NT\Accessories\wordpad.exe")

If Not objFSO.FileExists(wordPadPath) Then
    MsgBox "未找到 WordPad 程序，请检查路径是否正确。", vbExclamation, "错误"
    WScript.Quit
End If

' 启动 Word 应用程序
Set objWord = CreateObject("Word.Application")
objWord.Visible = False

' 遍历 .doc 和 .docx 文件
For Each file In objFSO.GetFolder(sourceFolder).Files
    If LCase(objFSO.GetExtensionName(file.Name)) = "doc" Or LCase(objFSO.GetExtensionName(file.Name)) = "docx" Then
        tempFile = objFSO.BuildPath(sourceFolder, objFSO.GetBaseName(file.Name) & "_temp.rtf") ' 临时文件路径
        destFile = objFSO.BuildPath(rtfFolder, objFSO.GetBaseName(file.Name) & ".rtf") ' 最终文件路径

        ' 用 Word 打开文件并保存为临时 RTF 文件
        On Error Resume Next
        Set objDoc = objWord.Documents.Open(file.Path)
        If Err.Number = 0 Then
            objDoc.SaveAs2 tempFile, 6 ' 6 = wdFormatRTF
            objDoc.Close False
        Else
            MsgBox "无法打开文件：" & vbCrLf & file.Path, vbExclamation, "错误"
            On Error GoTo 0
            Continue For
        End If
        On Error GoTo 0

        ' 用 WordPad 打开临时 RTF 文件并保存到最终文件
        If objFSO.FileExists(tempFile) Then
            objShell.Run Chr(34) & wordPadPath & Chr(34) & " /p " & Chr(34) & tempFile & Chr(34), 0, True
            WScript.Sleep 1000 ' 等待 WordPad 打开
            objShell.SendKeys "^s" ' Ctrl+S 保存文件
            WScript.Sleep 500
            objShell.SendKeys "%{F4}" ' Alt+F4 关闭 WordPad
            WScript.Sleep 500

            ' 将临时文件移动到最终文件夹
            objFSO.MoveFile tempFile, destFile
        End If
    End If
Next

' 清理
objWord.Quit
Set objWord = Nothing
Set objFSO = Nothing
Set objShell = Nothing

MsgBox "转换完成！RTF 文件已保存在：" & vbCrLf & rtfFolder, vbInformation, "完成"

```

```vbscript
' *********************************************************************
' This VBS script converts all .doc and .docx files in the current
' directory to .rtf files using WordPad.
' The converted files are saved in a subfolder named "RTF".
' To use, place the script in the "SendTo" folder for the right-click menu.
' *********************************************************************

Option Explicit

Dim objFSO, objShell, objArgs, currentFolder, rtfFolder, file, fileExtension
Dim wordPadPath, rtfFilePath, command

' Initialize FileSystemObject and Shell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Get the arguments passed to the script (the folder or file that was sent to the script)
Set objArgs = WScript.Arguments

' Ensure that a folder was sent to the script
If objArgs.Count = 0 Then
    MsgBox "Please use the 'SendTo' context menu to run this script.", vbExclamation, "No Input"
    WScript.Quit
End If

' Get the first argument (the folder or file sent to the script)
currentFolder = objArgs(0)

' Ensure the argument is a folder
If Not objFSO.FolderExists(currentFolder) Then
    MsgBox "The selected item is not a folder.", vbExclamation, "Invalid Input"
    WScript.Quit
End If

' Create the RTF folder inside the current directory
rtfFolder = objFSO.BuildPath(currentFolder, "RTF")
If Not objFSO.FolderExists(rtfFolder) Then
    objFSO.CreateFolder(rtfFolder)
End If

' Define the path to WordPad (usually located at this path)
wordPadPath = objShell.ExpandEnvironmentStrings("%ProgramFiles%\Windows NT\Accessories\wordpad.exe")

' Check if WordPad exists
If Not objFSO.FileExists(wordPadPath) Then
    MsgBox "WordPad is not found on this system.", vbCritical, "Error"
    WScript.Quit
End If

' Process each file in the folder
For Each file In objFSO.GetFolder(currentFolder).Files
    ' Get the file extension
    fileExtension = LCase(objFSO.GetExtensionName(file.Name))
    
    ' Check if the file is a .doc or .docx file
    If fileExtension = "doc" Or fileExtension = "docx" Then
        ' Define the output RTF file path
        rtfFilePath = objFSO.BuildPath(rtfFolder, objFSO.GetBaseName(file.Name) & ".rtf")
        
        ' Construct the command to convert the file using WordPad
        ' /p ensures that WordPad opens the file and /s saves it in RTF format
        command = """" & wordPadPath & """ """ & file.Path & """ /p /s /o:""" & rtfFilePath & """"
        
        ' Run the command
        objShell.Run command, 0, True
    End If
Next

' Notify the user of completion
MsgBox "Conversion completed. RTF files are saved in the 'RTF' folder.", vbInformation, "Done"

' Clean up
Set objFSO = Nothing
Set objShell = Nothing
Set objArgs = Nothing
```
