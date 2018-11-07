Attribute VB_Name = "執行腳本"

Sub RunScript()

    Dim jar As String
    ActiveWorkbook.Save
    Application.Wait Now() + TimeValue("00:00:02") '暫緩2秒
    
    
    CheckAPPandDeviceResult = CheckAPPandDevice()
    CheckValueResults = CheckValueResult()
    CheckCommandResult = CheckCommand()
    CheckExpectResult2 = CheckExpectResult_Ver2()
    
    If CheckAPPandDeviceResult = True And CheckValueResults = True And CheckCommandResult = True And CheckExpectResult2 = True Then

        jar = "java -jar " & Sheets("APP&Device").Cells(2, "G")
        r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & jar, 1)  '啟動cmd，指定路徑cd

        '& ">>log.txt"

    End If

End Sub


