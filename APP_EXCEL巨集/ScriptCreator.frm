VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScriptCreator 
   Caption         =   "TestScript Creator"
   ClientHeight    =   9465.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9645.001
   OleObjectBlob   =   "ScriptCreator.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "ScriptCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
    Dim selected As Boolean
    selected = False
    casenamestate = False
    QuitAPP = False
    
    For i = 0 To CommandList.ListCount - 1
    
        If CommandList.selected(i) = True Then
                
                For k = 0 To StepList.ListCount - 1
                    
                    If CommandList.List(i) = StepList.List(k) Then
                    
                        If StepList.List(k) = "CaseName" Then
                            casenamestate = True
                            x = MsgBox("CaseName已存在", 0 + 64, "Message")
                            Exit For
                        ElseIf StepList.List(k) = "QuitAPP" Then
                            QuitAPP = True
                            x = MsgBox("QuitAPP已存在", 0 + 64, "Message")
                            Exit For
                        End If
                        
                    End If

                Next k
            
                For j = 0 To StepList.ListCount - 1
                
                    If StepList.selected(j) = True Then
                        selected = True
                        Exit For
                    End If
    
                Next j
                
                If selected = True Then
                    If StepList.List(j) <> "CaseName" Then
                        StepList.AddItem CommandList.List(i), j
                        StepList.selected(j + 1) = True
                        'Exit For
                    End If
                Else
                    StepList.AddItem CommandList.List(i), StepList.ListCount - 1
                    'Exit For
                End If
        End If
 
    Next i
End Sub

Private Sub APP_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "B"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "B") = ""
End Sub

Private Sub cancelSelect_Click()
    
    For j = 0 To StepList.ListCount - 1
    
        StepList.selected(j) = False
    
    Next j
    
End Sub

Private Sub clear_Click()
    StepList.clear
    StepList.AddItem ("CaseName")
    StepList.AddItem ("QuitAPP")
End Sub

Private Sub ClearElement_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "E"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "E") = ""
End Sub

Private Sub Click_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "C"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "C") = ""
End Sub

Private Sub CommandBox_Change()
    i = 2
    CommandList.clear
     
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(i, CommandBox.ListIndex + 2))
    i = i + 1
    Loop Until Sheets("CommandCode").Cells(i, CommandBox.ListIndex + 2) = ""
End Sub



Private Sub CommandList_Change()
    For i = 0 To CommandList.ListCount - 1
    
        If CommandList.selected(i) = True Then
        
        j = 2
        Do
            
            If CommandList.List(i) = Sheets("說明").Cells(j, "A") Then
                    
                    x = Mid(Sheets("說明").Cells(j, "A").NoteText, 12, Len(Sheets("說明").Cells(j, "A").NoteText) - 12 + 1)
                    Exit Do
            End If
            
            j = j + 1
        Loop Until Sheets("說明").Cells(j, "A") = ""
            
            Command.Caption = "Command:" + CommandList.List(i) + vbNewLine + x
            Exit For
        
        End If
        
    Next i
End Sub

Private Sub CreateCase_Click()
    Application.ScreenUpdating = False
    Dim exist As Boolean
    exist = False
    
    If scriptname.Text <> "" And CaseName.Text <> "" And StepList.ListCount > 2 Then
    
        If Right(scriptname.Text, 11) = "_TestScript" Then
            i = 0
            Do
               
                If ThisWorkbook.Sheets(i + 1).Name = scriptname.Text Then
                
                    exist = True
                    Exit Do
            
                End If
            i = i + 1
            Loop Until i = ThisWorkbook.Sheets.Count
            
            If exist = False Then
                
                Sheets.add After:=Sheets(Sheets.Count - 1)
                Sheets(Sheets.Count - 1).Name = scriptname.Text
            
            End If
            '起始列
            start_row = Sheets(scriptname.Text).Cells(Sheets(scriptname.Text).Rows.Count, 1).End(xlUp).row
            original_start_row = start_row
            '填入Step
            If start_row = 1 Then
                Sheets(scriptname.Text).Cells(start_row, "B") = CaseName.Text
                For i = 0 To StepList.ListCount - 1
                    Sheets(scriptname.Text).Cells(start_row, 1) = StepList.List(i)
                    start_row = start_row + 1
                Next i
            Else
                Sheets(scriptname.Text).Cells(start_row + 1, "B") = CaseName.Text
                For i = 0 To StepList.ListCount - 1
                    Sheets(scriptname.Text).Cells(start_row + 1, 1) = StepList.List(i)
                    start_row = start_row + 1
                Next i
            End If
            
            Call ImportData(original_start_row)
            x = MsgBox("Done.", 0 + 64, "Message")
        Else
        
            x = MsgBox("Script名稱必須以_TestScript結尾", 0 + 16, "Error")
        
        End If
    ElseIf scriptname.Text = "" And CaseName.Text = "" Then
        
        x = MsgBox("請輸入Script名稱及Case名稱", 0 + 16, "Error")
        
    ElseIf scriptname.Text = "" Then
        
         x = MsgBox("請輸入Script名稱", 0 + 16, "Error")
        
    ElseIf CaseName.Text = "" Then
        
         x = MsgBox("請輸入Case名稱", 0 + 16, "Error")
         
    ElseIf StepList.ListCount = 2 Then
    
        x = MsgBox("請加入指令", 0 + 16, "Error")
    
    End If
    
End Sub

Private Sub delete_Click()
    For i = 0 To StepList.ListCount - 1
    
        If StepList.selected(i) = True Then
            If StepList.List(i) <> "CaseName" And StepList.List(i) <> "QuitAPP" Then
                StepList.RemoveItem (i)
                StepList.selected(i) = False
            End If
            
        End If
    
    Next i
End Sub

Private Sub HideKeyboard_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "K"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "K") = ""
End Sub

Private Sub down_Click()
    For i = StepList.ListCount - 1 To 0 Step -1
    
        If StepList.ListIndex <> StepList.ListCount - 1 And StepList.selected(i) = True And StepList.ListIndex <> StepList.ListCount - 2 And StepList.List(i) <> "CaseName" Then
        
            temp = StepList.List(i)
            StepList.RemoveItem (i)
            StepList.AddItem temp, i + 1
            StepList.selected(i + 1) = True
            Exit For
            
        End If
        
    Next i
End Sub

Private Sub Invisibility_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "H"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "H") = ""
End Sub

Private Sub Others_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "K"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "K") = ""
End Sub

Private Sub SendKey_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "D"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "D") = ""
End Sub

Private Sub StepList_Change()
    For i = 0 To StepList.ListCount - 1
    
        If StepList.selected(i) = True Then
            
            StepCommand.Caption = "Command:" + StepList.List(i)
            Exit For
        End If
    Next i
End Sub


Private Sub Swipe_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "G"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "G") = ""
End Sub

Private Sub System_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "J"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "J") = ""
End Sub

Private Sub up_Click()
    For i = 0 To StepList.ListCount - 1
    
        If StepList.ListIndex > 0 And StepList.selected(i) = True And StepList.ListIndex <> 1 And StepList.List(i) <> "QuitAPP" Then
        
            temp = StepList.List(i)
            StepList.RemoveItem (i)
            StepList.AddItem temp, i - 1
            StepList.selected(i - 1) = True
            Exit For
            
        End If
        
    Next i
End Sub

Private Sub UserForm_Activate()
    StepList.clear
    StepList.AddItem ("CaseName")
    StepList.AddItem ("QuitAPP")
End Sub

Private Sub Verify_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "I"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "I") = ""
End Sub

Private Sub Wait_Click()
    CommandList.clear
    j = 2
    Do
        CommandList.AddItem (Sheets("CommandCode").Cells(j, "F"))

    j = j + 1
    Loop Until Sheets("CommandCode").Cells(j, "F") = ""
End Sub

Sub ImportData(startj)
    x = scriptname.Text
    Sheets(scriptname.Text).Select
    j = startj + 1
    'j = Sheets(scriptname.Text).Cells(Sheets(scriptname.Text).Rows.Count, 1).End(xlUp).row
    Do
        i = 3
        Do
            If Sheets(scriptname.Text).Cells(j, "A") = Sheets("說明").Cells(i, "A") Then
                
                k = 2
                Do Until Sheets("說明").Cells(i, k) = ""
                    
                    Sheets(scriptname.Text).Cells(j, k).Select
                    Call line
                    k = k + 1
                Loop
                
            End If
            i = i + 1
        Loop Until Sheets("說明").Cells(i, "A") = ""
    
        j = j + 1
    Loop Until Sheets(scriptname.Text).Cells(j, "A") = ""

End Sub


Sub line()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDashDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDashDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDashDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
