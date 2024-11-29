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
' Convert Word Documents to RTF 
' Purpose: Open .doc and .docx files in Word, copy content, save as RTF using WordPad
' Language: VBScript
' Compatible with: Windows

Option Explicit

' Declare variables
Dim fso, wordApp, wshShell
Dim sourceFolder, rtfFolder
Dim file, sourceFile, rtfFile
Dim wordDoc

' Create File System Object and Windows Shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

' Get the current directory
sourceFolder = fso.GetParentFolderName(WScript.ScriptFullName)

' Create RTF subdirectory if it doesn't exist
rtfFolder = fso.BuildPath(sourceFolder, "RTF")
If Not fso.FolderExists(rtfFolder) Then
    fso.CreateFolder(rtfFolder)
End If

' Create Word Application object
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False

' Iterate through .doc and .docx files
For Each file In fso.GetFolder(sourceFolder).Files
    sourceFile = file.Path
    
    ' Check if file is a Word document
    If LCase(fso.GetExtensionName(sourceFile)) = "doc" Or _
       LCase(fso.GetExtensionName(sourceFile)) = "docx" Then
        
        ' Construct RTF filename
        rtfFile = fso.BuildPath(rtfFolder, fso.GetBaseName(sourceFile) & ".rtf")
        
        ' Open document in Word
        Set wordDoc = wordApp.Documents.Open(sourceFile)
        
        ' Copy all content to clipboard
        wordDoc.Content.Copy
        
        ' Close Word document
        wordDoc.Close False
        
        ' Open WordPad and wait a moment
        wshShell.Run "write.exe", 1, False
        WScript.Sleep 1000
        
        ' Paste content into WordPad
        wshShell.SendKeys "^v"
        WScript.Sleep 500
        
        ' Save file
        wshShell.SendKeys "^s"
        WScript.Sleep 500
        
        ' Type the RTF filename
        wshShell.SendKeys rtfFile
        WScript.Sleep 500
        
        ' Confirm save
        wshShell.SendKeys "{ENTER}"
        WScript.Sleep 500
        
        ' Close WordPad
        wshShell.SendKeys "%{F4}"
    End If
Next

' Clean up
wordApp.Quit
Set wordDoc = Nothing
Set wordApp = Nothing
Set fso = Nothing
Set wshShell = Nothing

' Show completion message
MsgBox "Conversion complete. RTF files saved in " & rtfFolder, 64, "Conversion Successful"
```

```vbscript
' 脚本：将当前目录中的 .doc 和 .docx 文件转换为 .rtf 文件
' 使用方法：将脚本放入 SendTo 文件夹，通过右键菜单调用

Dim objArgs, objWord, objFSO, sourceFolder, rtfFolder, file, destFile

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

' 启动 Word 应用程序
Set objWord = CreateObject("Word.Application")
objWord.Visible = False

' 转换 .doc 和 .docx 文件
For Each file In objFSO.GetFolder(sourceFolder).Files
    If LCase(objFSO.GetExtensionName(file.Name)) = "doc" Or LCase(objFSO.GetExtensionName(file.Name)) = "docx" Then
        destFile = objFSO.BuildPath(rtfFolder, objFSO.GetBaseName(file.Name) & ".rtf")
        On Error Resume Next
        ' 打开 Word 文件
        Set objDoc = objWord.Documents.Open(file.Path)
        If Err.Number = 0 Then
            ' 保存为 RTF 格式
            objDoc.SaveAs2 destFile, 6 ' 6 = wdFormatRTF
            objDoc.Close False
        Else
            MsgBox "无法打开文件：" & vbCrLf & file.Path, vbExclamation, "错误"
        End If
        On Error GoTo 0
    End If
Next

' 清理
objWord.Quit
Set objWord = Nothing
Set objFSO = Nothing

MsgBox "转换完成！RTF 文件已保存在：" & vbCrLf & rtfFolder, vbInformation, "完成"
```

```vbscript
' 设置包含.doc和.docx文件的目录路径
strSourceDir = WScript.Arguments(0)

' 创建Word和WordPad对象
Set objWord = CreateObject("Word.Application")
Set objWordPad = CreateObject("WordPad.Document")

' 设置Word和WordPad为可见
objWord.Visible = True
objWordPad.Visible = True

' 获取所有.doc和.docx文件
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strSourceDir)
Set colFiles = objFolder.Files

' 创建RTF目录
strRtfDir = strSourceDir & "\RTF"
If Not objFSO.FolderExists(strRtfDir) Then
    objFSO.CreateFolder(strRtfDir)
End If

' 遍历文件
For Each objFile in colFiles
    strFileName = objFile.Name
    ' 检查文件扩展名
    If LCase(objFSO.GetExtensionName(strFileName)) = "doc" Or LCase(objFSO.GetExtensionName(strFileName)) = "docx" Then
        ' 打开Word文档
        Set objDoc = objWord.Documents.Open(objFile.Path)
        
        ' 复制内容
        objDoc.Content.Copy
        
        ' 在WordPad中粘贴内容
        objWordPad.Range.Paste
        
        ' 设置RTF文件名和路径
        strRtfFileName = strRtfDir & "\" & objFSO.GetBaseName(strFileName) & ".rtf"
        
        ' 保存为RTF文件
        objWordPad.SaveAs strRtfFileName, 2 ' 2代表RTF格式
        
        ' 清除WordPad内容
        objWordPad.Range.Delete
        
        ' 关闭Word文档
        objDoc.Close
    End If
Next

' 关闭WordPad和Word
objWordPad.Close
objWord.Quit

' 清理对象
Set objWordPad = Nothing
Set objDoc = Nothing
Set objWord = Nothing
Set objFSO = Nothing
Set objFolder = Nothing
Set colFiles = Nothing

```
```vbscript
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' 获取传递的文件夹路径
If WScript.Arguments.Count = 0 Then
    WScript.Echo "请通过右键菜单的“发送到”选项运行此脚本。"
    WScript.Quit
End If

folderPath = WScript.Arguments(0)

' 检查传递的路径是否为文件夹
If Not objFSO.FolderExists(folderPath) Then
    WScript.Echo "传递的路径无效，请确认选择的是文件夹。"
    WScript.Quit
End If

' 创建 RTF 目录
Set folder = objFSO.GetFolder(folderPath)
rtfFolderPath = objFSO.BuildPath(folderPath, "RTF")
If Not objFSO.FolderExists(rtfFolderPath) Then
    objFSO.CreateFolder rtfFolderPath
End If

' 初始化 Word 应用程序
On Error Resume Next
Set wordApp = CreateObject("Word.Application")
If Err.Number <> 0 Then
    WScript.Echo "无法启动 Word 应用程序，请确保已安装 Microsoft Word。"
    WScript.Quit
End If
On Error GoTo 0

wordApp.Visible = False ' 后台运行 Word

' 遍历文件夹中的 .doc 和 .docx 文件
For Each file In folder.Files
    If LCase(objFSO.GetExtensionName(file.Name)) = "doc" Or LCase(objFSO.GetExtensionName(file.Name)) = "docx" Then
        docFilePath = file.Path
        rtfFileName = objFSO.GetBaseName(file.Name) & ".rtf"
        rtfFilePath = objFSO.BuildPath(rtfFolderPath, rtfFileName)

        ' 打开 Word 文档
        On Error Resume Next
        Set doc = wordApp.Documents.Open(docFilePath, False, True) ' 打开为只读模式
        If Err.Number <> 0 Then
            WScript.Echo "无法打开文件：" & docFilePath
            Err.Clear
            On Error GoTo 0
            Continue For
        End If
        On Error GoTo 0

        ' 复制内容到剪贴板
        doc.Content.Copy
        doc.Close False

        ' 打开 WordPad 并粘贴内容
        On Error Resume Next
        objShell.Run "write.exe", 1, False ' 启动 WordPad
        WScript.Sleep 1000 ' 等待 WordPad 启动

        objShell.SendKeys "^v" ' 粘贴内容
        WScript.Sleep 500

        ' 保存为 RTF 文件
        objShell.SendKeys "^s" ' 保存文件
        WScript.Sleep 500
        objShell.SendKeys rtfFilePath & "{ENTER}" ' 输入文件路径并保存
        WScript.Sleep 1000

        objShell.SendKeys "%{F4}" ' 关闭 WordPad
        WScript.Sleep 500
        On Error GoTo 0

        WScript.Echo "已转换文件：" & rtfFileName
    End If
Next

' 关闭 Word 应用程序
wordApp.Quit
Set wordApp = Nothing

WScript.Echo "转换完成！RTF 文件已保存到 " & rtfFolderPath

```
