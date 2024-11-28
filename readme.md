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
'==========================================================================
' This script converts all .doc and .docx files in the current directory 
' into .rtf files using Microsoft Word and saves them in a subdirectory 
' named "RTF".
' Place this script in the "SendTo" folder for quick access from the 
' right-click context menu.
'==========================================================================

Dim objArgs, objFSO, objWord, inputFolder, outputFolder, inputFile, outputFile, files
Dim folderPath, fileExtension

' Get the command-line arguments (the selected file/folder sent to the script)
Set objArgs = WScript.Arguments

' Ensure a file/folder was passed in
If objArgs.Count = 0 Then
    MsgBox "Please select a file or folder and use 'Send To' to run this script.", vbExclamation, "Error"
    WScript.Quit
End If

' Get the path of the selected file/folder
folderPath = objArgs(0)

' Create a FileSystemObject to handle file and folder operations
Set objFSO = CreateObject("Scripting.FileSystemObject")

' If the selected item is a file, get its parent directory
If objFSO.FileExists(folderPath) Then
    folderPath = objFSO.GetParentFolderName(folderPath)
End If

' Ensure the folder exists
If Not objFSO.FolderExists(folderPath) Then
    MsgBox "The selected folder does not exist.", vbExclamation, "Error"
    WScript.Quit
End If

' Create the output folder named "RTF" in the current directory
outputFolder = objFSO.BuildPath(folderPath, "RTF")
If Not objFSO.FolderExists(outputFolder) Then
    objFSO.CreateFolder(outputFolder)
End If

' Initialize Microsoft Word application
On Error Resume Next
Set objWord = CreateObject("Word.Application")
If Err.Number <> 0 Then
    MsgBox "Microsoft Word is not installed on this system.", vbExclamation, "Error"
    WScript.Quit
End If
On Error GoTo 0

' Process all .doc and .docx files in the selected directory
Set files = objFSO.GetFolder(folderPath).Files

For Each inputFile In files
    fileExtension = LCase(objFSO.GetExtensionName(inputFile.Name))
    
    ' Check if the file is .doc or .docx
    If fileExtension = "doc" Or fileExtension = "docx" Then
        ' Construct the output file path
        outputFile = objFSO.BuildPath(outputFolder, objFSO.GetBaseName(inputFile.Name) & ".rtf")
        
        ' Open the file in Word
        On Error Resume Next
        Set doc = objWord.Documents.Open(inputFile.Path, False, True) ' Open in read-only mode
        If Err.Number <> 0 Then
            MsgBox "Failed to open file: " & inputFile.Path, vbExclamation, "Error"
            Err.Clear
            On Error GoTo 0
            Continue For
        End If
        On Error GoTo 0

        ' Save the document as RTF
        On Error Resume Next
        doc.SaveAs2 outputFile, 6 ' 6 corresponds to the RTF format
        If Err.Number <> 0 Then
            MsgBox "Failed to save file: " & outputFile, vbExclamation, "Error"
            Err.Clear
        End If
        doc.Close False ' Close the document without saving changes
        On Error GoTo 0
    End If
Next

' Quit Word application
objWord.Quit
Set objWord = Nothing

' Notify the user
MsgBox "Conversion completed. RTF files are saved in: " & outputFolder, vbInformation, "Done"

' Clean up
Set objFSO = Nothing
Set objArgs = Nothing
```

```vbscript

'==========================================================================
' This script converts all .doc and .docx files in the current directory 
' into .rtf files using Microsoft Word and saves them in a subdirectory 
' named "RTF".
' Place this script in the "SendTo" folder for quick access from the 
' right-click context menu.
'==========================================================================

Dim objArgs, objFSO, objWord, inputFolder, outputFolder, inputFile, outputFile, files
Dim folderPath, fileExtension

' Get the command-line arguments (the selected file/folder sent to the script)
Set objArgs = WScript.Arguments

' Ensure a file/folder was passed in
If objArgs.Count = 0 Then
    MsgBox "Please select a file or folder and use 'Send To' to run this script.", vbExclamation, "Error"
    WScript.Quit
End If

' Get the path of the selected file/folder
folderPath = objArgs(0)

' Create a FileSystemObject to handle file and folder operations
Set objFSO = CreateObject("Scripting.FileSystemObject")

' If the selected item is a file, get its parent directory
If objFSO.FileExists(folderPath) Then
    folderPath = objFSO.GetParentFolderName(folderPath)
End If

' Ensure the folder exists
If Not objFSO.FolderExists(folderPath) Then
    MsgBox "The selected folder does not exist.", vbExclamation, "Error"
    WScript.Quit
End If

' Create the output folder named "RTF" in the current directory
outputFolder = objFSO.BuildPath(folderPath, "RTF")
If Not objFSO.FolderExists(outputFolder) Then
    objFSO.CreateFolder(outputFolder)
End If

' Initialize Microsoft Word application
On Error Resume Next
Set objWord = CreateObject("Word.Application")
If Err.Number <> 0 Then
    MsgBox "Microsoft Word is not installed on this system.", vbExclamation, "Error"
    WScript.Quit
End If
On Error GoTo 0

' Process all .doc and .docx files in the selected directory
Set files = objFSO.GetFolder(folderPath).Files

For Each inputFile In files
    fileExtension = LCase(objFSO.GetExtensionName(inputFile.Name))
    
    ' Check if the file is .doc or .docx
    If fileExtension = "doc" Or fileExtension = "docx" Then
        ' Construct the output file path
        outputFile = objFSO.BuildPath(outputFolder, objFSO.GetBaseName(inputFile.Name) & ".rtf")
        
        ' Open the file in Word
        On Error Resume Next
        Set doc = objWord.Documents.Open(inputFile.Path, False, True) ' Open in read-only mode
        If Err.Number <> 0 Then
            ' Failed to open the file, skip to the next file
            Err.Clear
            On Error GoTo 0
            ContinueLoop:
        Else
            ' Save the document as RTF
            doc.SaveAs2 outputFile, 6 ' 6 corresponds to the RTF format
            If Err.Number <> 0 Then
                ' Failed to save the file, skip to the next file
                Err.Clear
            End If
            doc.Close False ' Close the document without saving changes
        End If
        On Error GoTo 0
    End If
Next

' Quit Word application
objWord.Quit
Set objWord = Nothing

' Notify the user
MsgBox "Conversion completed. RTF files are saved in: " & outputFolder, vbInformation, "Done"

' Clean up
Set objFSO = Nothing
Set objArgs = Nothing
```
