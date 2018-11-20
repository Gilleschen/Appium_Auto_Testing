Attribute VB_Name = "CreateGitVersion"
Sub createGitVersion()
    
    x = MsgBox("確認建立Git Version ?", 4 + 32, "Message")
    
    If x = 6 Then
    
        x = MsgBox("再次確認建立Git Version ?", 4 + 32, "Message")
    
        If x = 6 Then
        
            Application.ScreenUpdating = False
            Call deleteSheets
            Call hideCommandCodeSheet
            Call removeiOSnote
            Call defaultInfo
            Call copyDemoTestScript
            Call clearAPPandDevice_data
            Sheets("APP&Device").Select
            Application.ScreenUpdating = True
            x = MsgBox("Done.", 0 + 64, "Message")
            
        End If
        
    End If
End Sub


Sub deleteSheets()
    i = 0
    Do
    
        If ThisWorkbook.Sheets(i + 1).Visible = False Then ThisWorkbook.Sheets(i + 1).Visible = True
            
        If ThisWorkbook.Sheets(i + 1).Name <> "APP&Device" And ThisWorkbook.Sheets(i + 1).Name <> "APP&Device_Data" And _
        ThisWorkbook.Sheets(i + 1).Name <> "說明" And ThisWorkbook.Sheets(i + 1).Name <> "CommandCode" Then
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets(i + 1).Select
            ActiveWindow.SelectedSheets.delete
            Application.DisplayAlerts = True
            i = i - 1
        End If
        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count

End Sub

Sub hideCommandCodeSheet()
    ' Hide CommandCode sheet
    If Sheets("CommandCode").Visible = True Then Sheets("CommandCode").Visible = False
End Sub

Sub removeiOSnote()
    ' Delete 說明 sheets中ByXpath_Swipe_FindText_Click_iOS
    Sheets("說明").Select
    i = 1
    Do
        If Sheets("說明").Cells(i, "A") = "ByXpath_Swipe_FindText_Click_iOS" Then
        
            Rows(i & ":" & i).Select
            Selection.delete Shift:=xlUp
            Exit Do
        End If
    
        i = i + 1
    Loop Until Sheets("說明").Cells(i, "A") = ""
End Sub

Sub clearAPPandDevice_data()
    ' Clear APP&Device_Data data
    Sheets("APP&Device_Data").Select
    i = 2
    Do
        Rows(i & ":" & i).Select
        Selection.delete Shift:=xlUp

    Loop Until Sheets("APP&Device_Data").Cells(i, "A") = "" And Sheets("APP&Device_Data").Cells(i, "B") = "" And Sheets("APP&Device_Data").Cells(i, "A") = "" And Sheets("APP&Device_Data").Cells(i, "D") = ""
    
    ' Input packagename and activity
    i = 2
    Do
        Sheets("APP&Device_Data").Cells(i, "A") = Workbooks("TestScript_git.xlsm").Sheets("APP&Device_Data").Cells(i, "A")
        Sheets("APP&Device_Data").Cells(i, "B") = Workbooks("TestScript_git.xlsm").Sheets("APP&Device_Data").Cells(i, "B")
        i = i + 1
    Loop Until Workbooks("TestScript_git.xlsm").Sheets("APP&Device_Data").Cells(i, "A") = ""
    
     ' Input UDID and OS version
    i = 2
    Do
        Sheets("APP&Device_Data").Cells(i, "C") = Workbooks("TestScript_git.xlsm").Sheets("APP&Device_Data").Cells(i, "C")
        Sheets("APP&Device_Data").Cells(i, "D") = Workbooks("TestScript_git.xlsm").Sheets("APP&Device_Data").Cells(i, "D")
        i = i + 1
    Loop Until Workbooks("TestScript_git.xlsm").Sheets("APP&Device_Data").Cells(i, "A") = ""
    
    ' Close TestScript_git.xlsm
    Application.DisplayAlerts = False
    Workbooks("TestScript_git.xlsm").Activate
    Workbooks("TestScript_git.xlsm").Close
    Application.DisplayAlerts = True
End Sub

Sub defaultInfo()
    Sheets("APP&Device").Select
    Sheets("APP&Device").Cells(2, "C") = ""
    Sheets("APP&Device").Cells(2, "D") = ""
    Sheets("APP&Device").Cells(2, "E") = ""
    Sheets("APP&Device").Cells(2, "F") = ""
    Sheets("APP&Device").Cells(2, "G") = "C:\Users\Desktop\Appium_Android.jar"
End Sub

Sub copyDemoTestScript()
    ' Launch TestScript_git.xlsm
    Workbooks.Open Filename:="C:\Users\jhih_chen\Desktop\TestScript_git.xlsm"
    Workbooks("TestScript_git.xlsm").Activate
    Workbooks("TestScript_git.xlsm").Sheets(Array("Example_TestScript", "Example2_TestScript", "ExpectResult")).Select
    ' Copy Sheets to TestScript.xlsm
    Workbooks("TestScript_git.xlsm").Sheets(Array("Example_TestScript", "Example2_TestScript", "ExpectResult")).Copy Before:=Workbooks("TestScript.xlsm").Sheets("說明")
    Windows("TestScript.xlsm").Activate

End Sub


