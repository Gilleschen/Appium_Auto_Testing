VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestInformation 
   Caption         =   "APP & Device"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11505
   OleObjectBlob   =   "TestInformation.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "TestInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancelCcaseList_Click()
    
    For i = 0 To CaseList.ListCount - 1
    
        If CaseList.selected(i) = True Then
            CaseList.selected(i) = False
        End If
        
    Next i
    
End Sub

Private Sub CommandButton1_Click()
    Application.ScreenUpdating = False
    If PackageBox.Text = "" Then
        
        x = MsgBox("請選擇Package Name", 0 + 16, "Message")
        
    ElseIf DeviceBox.Text = "" Then
    
        x = MsgBox("請選擇Device Udid", 0 + 16, "Message")
        
    ElseIf ScriptBox.Text = "" Then
    
        x = MsgBox("請選擇Script", 0 + 16, "Message")
        
'    ElseIf checkCaseList = False Then
'
'        x = MsgBox("請選擇Case", 0 + 16, "Message")
        
    ElseIf checkJarPath = False Then
    
        'Error
    Else
        Sheets("APP&Device").Cells(2, "A") = PackageBox.Text
        Sheets("APP&Device").Cells(2, "B") = Right(ActivityLabel.Caption, Len(ActivityLabel.Caption) - Len("Activity: "))
        Sheets("APP&Device").Cells(2, "C") = DeviceBox.Text
        Sheets("APP&Device").Cells(2, "D") = Right(OSLabel.Caption, Len(OSLabel.Caption) - Len("OS Version: "))
        Sheets("APP&Device").Cells(2, "E") = ScriptBox.Text
        
        Sheets("APP&Device").Cells(2, "F").clear
        Count = 0
        For i = 0 To CaseList.ListCount - 1
            
            If CaseList.selected(i) = True Then
                Count = Count + 1
                
                If Count = 1 Then
                    Sheets("APP&Device").Cells(2, "F").Value = CaseList.List(i)
                Else
                    Sheets("APP&Device").Cells(2, "F").Value = Sheets("APP&Device").Cells(2, "F").Value & "," & CaseList.List(i)
                End If
        
            End If
        
        Next
        
        Sheets("APP&Device").Cells(2, "G") = JarPath.Text
        
        If resetTrue.Value = True Then
            Sheets("APP&Device").Cells(2, "H") = "TRUE"
        ElseIf resetFalse.Value = True Then
            Sheets("APP&Device").Cells(2, "H") = "FALSE"
        End If
        
        
        If UITrue.Value = True Then
            Sheets("APP&Device").Cells(2, "I") = "TRUE"
        ElseIf UIFalse.Value = True Then
            Sheets("APP&Device").Cells(2, "I") = "FALSE"
        End If
        
        x = MsgBox("Done.", 0 + 64, "Message")
    End If
    Application.ScreenUpdating = True
End Sub

Function checkJarPath()

    If JarPath.Text = "" Then
    
        x = MsgBox("請填入Jar Path", 0 + 16, "Message")
        checkJarPath = False
        Exit Function
        
    ElseIf Dir(JarPath.Text) = "" Then
        
        x = MsgBox("找不到" & JarPath.Text, 0 + 16, "Message")
        checkJarPath = False
        Exit Function
        
    Else
    
        checkJarPath = True
        
    End If
    
End Function

Function checkCaseList()
    
    For i = 0 To CaseList.ListCount - 1

        If CaseList.selected(i) = True Then
        
            checkCaseList = True
            Exit Function
        End If
    Next i
    
    checkCaseList = False
End Function


Private Sub CommandButton2_Click()
    Unload TestInformation
End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub DeviceBox_Change()
    OSLabel.Caption = "OS Version: " & Sheets("APP&Device_Data").Cells(DeviceBox.ListIndex + 2, "D")
    
    x = Left(Sheets("APP&Device_Data").Cells(DeviceBox.ListIndex + 2, "D"), 1)
    If x >= 7 Then
        UITrue.Value = True
    Else
        UIFalse.Value = True
    End If
    
End Sub


Private Sub PackageBox_Change()
    ActivityLabel.Caption = "Activity: " & Sheets("APP&Device_Data").Cells(PackageBox.ListIndex + 2, "B")
End Sub

Private Sub ScriptBox_Change()
    CaseList.clear
    i = 1
    Do
        If Sheets(ScriptBox.Text).Cells(i, "A") = "CaseName" Then
            CaseList.AddItem Sheets(ScriptBox.Text).Cells(i, "B")
        End If
        i = i + 1
    Loop Until Sheets(ScriptBox.Text).Cells(i, "A") = ""
    
End Sub

Private Sub UserForm_Activate()
            
    resetFalse.Value = True
    UIFalse.Value = True
    JarPath.Text = "C:\Users\Desktop\Appium_Android.jar"
            
    i = 2
    Do
        
        PackageBox.AddItem Sheets("APP&Device_Data").Cells(i, "a")
        
        i = i + 1
    Loop Until Sheets("APP&Device_Data").Cells(i, "A") = ""
    
    j = 2
    Do
        DeviceBox.AddItem Sheets("APP&Device_Data").Cells(j, "C")
        j = j + 1
    Loop Until Sheets("APP&Device_Data").Cells(j, "C") = ""
    
    
    m = 0
    Do
        
        If ThisWorkbook.Sheets(m + 1).Visible = True And Right(ThisWorkbook.Sheets(m + 1).Name, 11) = "_TestScript" Then
            
            ScriptBox.AddItem ThisWorkbook.Sheets(m + 1).Name
        
        End If
        m = m + 1
    Loop Until m = ThisWorkbook.Sheets.Count
    
End Sub

