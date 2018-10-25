Attribute VB_Name = "執行腳本"

Sub RunScript()

    Dim jar As String
    ActiveWorkbook.Save
    Application.Wait Now() + TimeValue("00:00:02") '暫緩2秒
    
    
    CheckAPPandDeviceResult = CheckAPPandDevice()
    CheckValueResult = CheckValue()
    CheckCommandResult = CheckCommand()
    
    If CheckAPPandDeviceResult = True And CheckValueResult = True And CheckCommandResult = True Then

        jar = "java -jar " & Sheets("APP&Device").Cells(2, "G")
        r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & jar, 1)  '啟動cmd，指定路徑cd

        '& ">>log.txt"

    End If

End Sub


