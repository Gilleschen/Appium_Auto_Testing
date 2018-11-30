VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "¤u§@ªí1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub importScriptAndCase_Click()

    Sheets("APP&Device").Cells(2, "E").clear
    For i = 0 To ListBox1.ListCount - 1
        
        If ListBox1.selected(i) = True Then
            
            Sheets("APP&Device").Cells(2, "E").Value = ListBox1.List(i)
        
        End If
    
    Next
    
    Sheets("APP&Device").Cells(2, "F").clear
    Count = 0
    For i = 0 To ListBox2.ListCount - 1
        
        If ListBox2.selected(i) = True Then
            Count = Count + 1
            
            If Count = 1 Then
                Sheets("APP&Device").Cells(2, "F").Value = ListBox2.List(i)
            Else
                Sheets("APP&Device").Cells(2, "F").Value = Sheets("APP&Device").Cells(2, "F").Value & "," & ListBox2.List(i)
            End If
    
        End If
    
    Next
End Sub

Private Sub ListBox1_Click()
    Application.ScreenUpdating = False
    ListBox3.clear
    
    For i = 0 To ListBox1.ListCount - 1
        
        If ListBox1.selected(i) = True Then
            j = 1
            Do
                If Sheets(ListBox1.List(i)).Cells(j, "A") = "CaseName" Then
                
                  ListBox3.AddItem (Sheets(ListBox1.List(i)).Cells(j, "B"))
                
                End If
                
            
            j = j + 1
            Loop Until Sheets(ListBox1.List(i)).Cells(j, "A") = ""
        
        End If
    
    Next
    
    Sheets("APP&Device").Cells(2, "E").clear
    Sheets("APP&Device").Cells(2, "F").clear
    For i = 0 To ListBox1.ListCount - 1
        
        If ListBox1.selected(i) = True Then
            
            Sheets("APP&Device").Cells(2, "E").Value = ListBox1.List(i)
        
        End If
    
    Next
    
    
End Sub

Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Application.ScreenUpdating = False
    Sheets("APP&Device").Cells(2, "F").clear
    Count = 0
    For i = 0 To ListBox3.ListCount - 1
        
        If ListBox3.selected(i) = True Then
            Count = Count + 1
            
            If Count = 1 Then
                Sheets("APP&Device").Cells(2, "F").Value = ListBox3.List(i)
            Else
                Sheets("APP&Device").Cells(2, "F").Value = Sheets("APP&Device").Cells(2, "F").Value & "," & ListBox3.List(i)
            End If
    
        End If
    
    Next
End Sub

Private Sub DeviceList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Clear UDID and OS version
    Application.ScreenUpdating = False
    
    i = 2
    Do
        Sheets("APP&Device").Cells(i, "C").Select
        Selection.delete Shift:=xlUp
        Sheets("APP&Device").Cells(i, "D").Select
        Selection.delete Shift:=xlUp
        'i = i + 1
    Loop Until Sheets("APP&Device").Cells(i, "C") = ""
    
    
    For i = 0 To DeviceList.ListCount - 1
    
        If DeviceList.selected(i) = True Then
        
            Count = Count + 1
            
            If Count = 1 Then
                Sheets("APP&Device").Cells(Count + 1, "C").Value = DeviceList.List(i)
                Sheets("APP&Device").Cells(Count + 1, "D").Value = Sheets("APP&Device_Data").Cells(i + 2, "D")
            Else
                Sheets("APP&Device").Cells(Count + 1, "C").Value = DeviceList.List(i)
                Sheets("APP&Device").Cells(Count + 1, "D").Value = Sheets("APP&Device_Data").Cells(i + 2, "D")
            End If
        
        End If
    
    Next i

End Sub
Private Sub loadingTestCase_Click()
   ListBox3.clear
    
    For i = 0 To ListBox1.ListCount - 1
        
        If ListBox1.selected(i) = True Then
            j = 1
            Do
                If Sheets(ListBox1.List(i)).Cells(j, "A") = "CaseName" Then
                
                  ListBox3.AddItem (Sheets(ListBox1.List(i)).Cells(j, "B"))
                
                End If
                
            
            j = j + 1
            Loop Until Sheets(ListBox1.List(i)).Cells(j, "A") = ""
        
        End If
    
    Next
End Sub

Private Sub loadingTestScript_Click()

    ListBox1.clear
    i = 0
    Do
        If ThisWorkbook.Sheets(i + 1).Visible = True And Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
            ListBox1.AddItem (ThisWorkbook.Sheets(i + 1).Name)
        End If
    
    i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
End Sub
Sub loadingTestScript()

    ListBox1.clear
    i = 0
    Do
        If ThisWorkbook.Sheets(i + 1).Visible = True And Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
            ListBox1.AddItem (ThisWorkbook.Sheets(i + 1).Name)
        End If
    
    i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
End Sub

Sub loadingDeivceUdid()
    DeviceList.clear
    i = 2
    Do
        DeviceList.AddItem (Sheets("APP&Device_Data").Cells(i, "C"))
        
    
        i = i + 1
    Loop Until Sheets("APP&Device_Data").Cells(i, "C") = ""
    
    
End Sub


Private Sub Package_Change()
    Application.ScreenUpdating = False
    Sheets("APP&Device").Cells(2, "A") = Package.Text
    Sheets("APP&Device").Cells(2, "B") = Sheets("APP&Device_Data").Cells(Package.ListIndex + 2, "B")
End Sub

Private Sub ResetAPP_Change()
    Sheets("APP&Device").Cells(2, "G") = ResetAPP.Text
End Sub
Sub loadingPackage()
    
    Dim PackageCurrent_index As Integer
    
    PackageCurrent_index = Package.ListIndex
    
    If PackageCurrent_index > -1 Then
    
    
        Package.clear
        i = 2
        Do
    
            Package.AddItem (Sheets("APP&Device_Data").Cells(i, "A"))
    
        i = i + 1
        Loop Until Sheets("APP&Device_Data").Cells(i, "A") = ""
    
        Package.Text = Sheets("APP&Device_Data").Cells(PackageCurrent_index + 2, "A")
    Else
        Package.clear
        i = 2
        Do
    
            Package.AddItem (Sheets("APP&Device_Data").Cells(i, "A"))
    
        i = i + 1
        Loop Until Sheets("APP&Device_Data").Cells(i, "A") = ""
    
    End If


End Sub
Sub loadingDeviceUDID()

    Dim DeviceUDIDCurrent_index As Integer
    DeviceUDIDCurrent_index = DeviceUDID.ListIndex
    DeviceUDID.clear
    i = 2
    Do

        DeviceUDID.AddItem (Sheets("APP&Device_Data").Cells(i, "C"))

    i = i + 1
    Loop Until Sheets("APP&Device_Data").Cells(i, "C") = ""
    DeviceUDID.Text = Sheets("APP&Device_Data").Cells(DeviceUDIDCurrent_index + 2, "C")
    
End Sub
Sub loadingResetAPP()

     If Sheets("APP&Device").Cells(2, "G") = "" Then
        ResetAPP.clear
        ResetAPP.AddItem ("TRUE")
        ResetAPP.AddItem ("FALSE")
        ResetAPP.Text = Sheets("APP&Device").Cells(2, "G")
    Else
        tmp = Sheets("APP&Device").Cells(2, "G")
        ResetAPP.clear
        ResetAPP.AddItem ("TRUE")
        ResetAPP.AddItem ("FALSE")
        ResetAPP.Text = tmp
    End If

End Sub


Private Sub Worksheet_Activate()
    'Sheets("APP&Device").Cells(2, "B").Clear
    'Sheets("APP&Device").Cells(2, "D").Clear
    Application.ScreenUpdating = False
    Call loadingPackage
    Call loadingDeivceUdid
    Call loadingTestScript
    Call loadingResetAPP

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    If Sheets("APP&Device").Cells(2, "F") = "" Then
        
        For w = 1 To 10
            Sheets("APP&Device").Cells(2, "F").Select
            Selection.delete Shift:=xlUp
        Next w
        
    End If
    Call reSetResetAPP
    Application.EnableEvents = True
    
End Sub

Sub reSetResetAPP()
    ResetAPP.Text = Sheets("APP&Device").Cells(2, "G")
End Sub

