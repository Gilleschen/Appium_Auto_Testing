Attribute VB_Name = "����}��"

Sub RunScript()

    Dim jar As String
    ActiveWorkbook.Save
    Application.Wait Now() + TimeValue("00:00:02") '�Ƚw2��
    
    
    CheckAPPandDeviceResult = CheckAPPandDevice()
    CheckValueResults = CheckValueResult()
    CheckCommandResult = CheckCommand_Ver2()
    'CheckExpectResult2 = CheckExpectResult_Ver2()
    JarPath = checkJarPath()
    
    If CheckAPPandDeviceResult = True And CheckValueResults = True And CheckCommandResult = True And JarPath = True Then 'And CheckExpectResult2 = True

        jar = "java -jar " & "C:\TUTK_QA_TestTool\TestTool\Appium_Android.jar"
        r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & jar, 1)  '�Ұ�cmd�A���w���|cd

        '& ">>log.txt"

    End If

End Sub


Function checkJarPath()

    If Dir(CStr("C:\TUTK_QA_TestTool\TestTool\Appium_Android.jar")) = "" Then

        x = MsgBox("�䤣�� C:\TUTK_QA_TestTool\TestTool\Appium_Android.jar" & vbNewLine & "Appium_Android.jar����mTestTool��Ƨ�", 0 + 16, "Error")
        checkJarPath = False
        Exit Function
    Else
        
        checkJarPath = True

    End If

End Function
