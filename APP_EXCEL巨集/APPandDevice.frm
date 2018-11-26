VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} APPandDevice 
   Caption         =   "Edit Udid and Package"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6675
   OleObjectBlob   =   "APPandDevice.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "APPandDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub Add_Click()
    Dim isSelected As Boolean
    Application.ScreenUpdating = False
    isSelected = False
    If Device.Value = True Then
        
        If TextBox1.Text = "" Then
    
            x = MsgBox("請輸入UDID", 0 + 16, "Error")
    
        ElseIf TextBox2.Text = "" Then
    
            x = MsgBox("請輸入OS Version", 0 + 16, "Error")
    
        ElseIf TextBox1.Text <> "" And TextBox2.Text <> "" Then
            
            For i = 0 To dataList.ListCount - 1
                
                If dataList.selected(i) = True Then isSelected = True: Exit For
                
            Next i
            
            If isSelected = True Then
                
                If checkDuplicate(i + 2, TextBox1.Text, TextBox2.Text, 3) = False Then
                
                    Sheets("APP&Device_Data").Cells(i + 2, "C") = TextBox1.Text
                    Sheets("APP&Device_Data").Cells(i + 2, "D") = TextBox2.Text
                    dataList.AddItem Sheets("APP&Device_Data").Cells(i + 2, "C") & " / " & Sheets("APP&Device_Data").Cells(i + 2, "D"), i
                    dataList.RemoveItem (i + 1)
                    x = MsgBox("Done.", 0 + 64, "Message")
                    
                End If
                
            ElseIf isSelected = False Then
                
                lastRow = Sheets("APP&Device_Data").Cells(1, "C").End(xlDown).row
                If checkDuplicate(lastRow, TextBox1.Text, TextBox2.Text, 3) = False Then
                    Sheets("APP&Device_Data").Cells(lastRow + 1, "C") = TextBox1.Text
                    Sheets("APP&Device_Data").Cells(lastRow + 1, "D") = TextBox2.Text
                    x = MsgBox("Done.", 0 + 64, "Message")
                    dataList.AddItem (Sheets("APP&Device_Data").Cells(lastRow + 1, "C") & " / " & Sheets("APP&Device_Data").Cells(lastRow + 1, "D"))
                End If
            End If
            
        End If
        
        
    ElseIf Package.Value = True Then
        
        If TextBox1.Text = "" Then
    
            x = MsgBox("請輸入PackageName", 0 + 16, "Error")
    
        ElseIf TextBox2.Text = "" Then
    
            x = MsgBox("請輸入Activity", 0 + 16, "Error")
    
        ElseIf TextBox1.Text <> "" And TextBox2.Text <> "" Then
            
            For i = 0 To dataList.ListCount - 1
                
                If dataList.selected(i) = True Then isSelected = True: Exit For
                
            Next i
            
            If isSelected = True Then
            
                If checkDuplicate(i + 2, TextBox1.Text, TextBox2.Text, 1) = False Then
                
                    Sheets("APP&Device_Data").Cells(i + 2, "A") = TextBox1.Text
                    Sheets("APP&Device_Data").Cells(i + 2, "B") = TextBox2.Text
                    dataList.AddItem Sheets("APP&Device_Data").Cells(i + 2, "A") & " / " & Sheets("APP&Device_Data").Cells(i + 2, "B"), i
                    dataList.RemoveItem (i + 1)
                    x = MsgBox("Done.", 0 + 64, "Message")
                
                End If
                
            ElseIf isSelected = False Then
                
                lastRow = Sheets("APP&Device_Data").Cells(1, "A").End(xlDown).row
                If checkDuplicate(lastRow, TextBox1.Text, TextBox2.Text, 1) = False Then
                    Sheets("APP&Device_Data").Cells(lastRow + 1, "A") = TextBox1.Text
                    Sheets("APP&Device_Data").Cells(lastRow + 1, "B") = TextBox2.Text
                    x = MsgBox("Done.", 0 + 64, "Message")
                    dataList.AddItem (Sheets("APP&Device_Data").Cells(lastRow + 1, "A") & " / " & Sheets("APP&Device_Data").Cells(lastRow + 1, "B"))
                    
                End If
                       
            End If
    
        End If
    
    Else
    
        x = MsgBox("請選擇項目", 0 + 16, "Error")
    
    End If

    Application.ScreenUpdating = True
End Sub
Function checkDuplicate(lastRow, data1 As String, data2 As String, col As Integer)
    
    For i = 2 To lastRow
        
        If data1 = Sheets("APP&Device_Data").Cells(i, col) Then
            
            If data2 = Sheets("APP&Device_Data").Cells(i, col + 1) Then
                checkDuplicate = True
                Exit Function
            Else
                checkDuplicate = False
                'Exit Function
            End If
            
        End If

    Next i

End Function



Private Sub Cancel_Click()
    For i = 0 To dataList.ListCount - 1
        
        If dataList.selected(i) = True Then
            dataList.selected(i) = False
        End If
    Next i
    
    TextBox1.Text = ""
    TextBox2.Text = ""
    
End Sub

Private Sub dataList_Change()

    If dataList.ListCount > 0 Then
    
        For i = 0 To dataList.ListCount - 1
            
            If dataList.selected(i) = True Then
                
                If Device.Value = True Then
                    TextBox1.Text = Sheets("APP&Device_Data").Cells(i + 2, "C")
                    TextBox2.Text = Sheets("APP&Device_Data").Cells(i + 2, "D")
                ElseIf Package.Value = True Then
                    TextBox1.Text = Sheets("APP&Device_Data").Cells(i + 2, "A")
                    TextBox2.Text = Sheets("APP&Device_Data").Cells(i + 2, "B")
                End If
                
            End If
            
        Next i
        
    End If

End Sub

Private Sub delete_Click()
    Application.ScreenUpdating = False
    Dim Delete As Boolean
    Delete = False
    
    x = MsgBox("確定移除?", 1 + 32, "Message")
    
    If x = 1 Then
    
        If Device.Value = True Then
            i = 0
            Do
            
                If dataList.selected(i) = True Then
                    Delete = True
                    dataList.RemoveItem (i)
                    Sheets("APP&Device_Data").Cells(i + 2, "C").Delete Shift:=xlUp
                    Sheets("APP&Device_Data").Cells(i + 2, "D").Delete Shift:=xlUp
                    i = i - 1
                End If
            
                i = i + 1
            Loop Until i = dataList.ListCount
            
        ElseIf Package.Value = True Then
            i = 0
            Do
                
                If dataList.selected(i) = True Then
                    Delete = True
                    dataList.RemoveItem (i)
                    Sheets("APP&Device_Data").Cells(i + 2, "A").Delete Shift:=xlUp
                    Sheets("APP&Device_Data").Cells(i + 2, "B").Delete Shift:=xlUp
                    i = i - 1
                End If
            
                i = i + 1
            Loop Until i = dataList.ListCount
        
        End If
    
    End If
    
    If Delete = True Then
        x = MsgBox("Done.", 0 + 64, "Message")
        TextBox1.Text = ""
        TextBox2.Text = ""
    End If
    
    Application.ScreenUpdating = True
End Sub

Private Sub Device_Click()
    If Device.Value = True Then
        text1.Caption = "UDID:"
        text2.Caption = "OS Version:"
        text3.Caption = "Udid/OS version:"
        text1.Visible = True
        text2.Visible = True
        text3.Visible = True
        TextBox1.Text = ""
        TextBox2.Text = ""
        
        dataList.clear
        
        i = 2
        Do
           dataList.AddItem (Sheets("APP&Device_Data").Cells(i, "C") & " / " & Sheets("APP&Device_Data").Cells(i, "D"))
        
        i = i + 1
        Loop Until Sheets("APP&Device_Data").Cells(i, "C") = ""
        
    End If
    
End Sub


Private Sub Package_Click()
    If Package.Value = True Then
        text1.Caption = "Package Name:"
        text2.Caption = "APP Activity:"
        text3.Caption = "PackageName/Activity:"
        text1.Visible = True
        text2.Visible = True
        text3.Visible = True
        TextBox1.Text = ""
        TextBox2.Text = ""
        
        dataList.clear
            i = 2
        Do
           dataList.AddItem (Sheets("APP&Device_Data").Cells(i, "A") & " / " & Sheets("APP&Device_Data").Cells(i, "B"))
    
        i = i + 1
        Loop Until Sheets("APP&Device_Data").Cells(i, "A") = ""
        
    End If
    
End Sub


Private Sub UserForm_Activate()

    dataList.clear
    
    text1.Visible = False
    text2.Visible = False
    text3.Visible = False
    
'    i = 2
'    Do
'       dataList.AddItem (Sheets("APP&Device_Data").Cells(i, "C") & " / " & Sheets("APP&Device_Data").Cells(i, "D"))
'
'    i = i + 1
'    Loop Until Sheets("APP&Device_Data").Cells(i, "C") = ""
    
'    i = 2
'    Do
'       APPList.AddItem (Sheets("APP&Device_Data").Cells(i, "A") & " / " & Sheets("APP&Device_Data").Cells(i, "B"))
'
'    i = i + 1
'    Loop Until Sheets("APP&Device_Data").Cells(i, "A") = ""
    
    
End Sub


