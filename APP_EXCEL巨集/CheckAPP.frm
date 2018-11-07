VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CheckAPP 
   Caption         =   "指令檢查"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   OleObjectBlob   =   "CheckAPP.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "CheckAPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
    Dim ScriptCheck As Boolean
    
    For i = 0 To TotalScriptList.ListCount - 1
    ScriptCheck = True
        If TotalScriptList.selected(i) = True Then
        
            For j = 0 To CheckScriptList.ListCount - 1
                
                If TotalScriptList.List(i) = CheckScriptList.List(j) Then
                    
                    ScriptCheck = False
                    Exit For
                
                End If
                
            Next j
            
            If ScriptCheck = True Then
                
                  CheckScriptList.AddItem TotalScriptList.List(i)
            
            End If
        
        End If

    Next i
End Sub

Private Sub cancelSelect_Click()
    For i = 0 To CheckScriptList.ListCount - 1
    
        If CheckScriptList.selected(i) = True Then
            CheckScriptList.selected(i) = False
        End If

    Next i
    
    For i = 0 To TotalScriptList.ListCount - 1
    
        If TotalScriptList.selected(i) = True Then
            TotalScriptList.selected(i) = False
        End If

    Next i
End Sub





Private Sub clear_Click()
    If CheckScriptList.ListCount > 0 Then
    
        CheckScriptList.clear
        
    End If
    
End Sub

Private Sub CommandButton1_Click()
    s = xxx("abc")
End Sub

Private Sub CreateCase_Click()
    If (CheckCommand.Value = True Or CheckValue.Value = True Or CheckExpectResult.Value = True) And CheckScriptList.ListCount > 0 Then
         
        For i = 0 To CheckScriptList.ListCount - 1
         
            If CheckCommand.Value = True Then
            
                x = CheckCommand_2(CheckScriptList.List(i))
                
            End If
            
            If CheckValue.Value = True Then
            
                x = CheckValueResult_2(CheckScriptList.List(i))
                
            End If
            
            If CheckExpectResult.Value = True Then
                
                x = CheckExpectResult_Ver2_2(CheckScriptList.List(i))
                
            End If
            
            CheckAPP.Caption = "指令檢查 " & ((i + 1) / CheckScriptList.ListCount) * 100 & "%"
            
        Next i
        CheckAPP.Hide
        Unload CheckAPP
    ElseIf (CheckCommand.Value = False And CheckValue.Value = False And CheckExpectResult.Value = False And CheckScriptList.ListCount = 0) Then
    
        x = MsgBox("請勾選檢查項目及TestScript", 0 + 16, "Error")
        
    ElseIf (CheckCommand.Value = False And CheckValue.Value = False And CheckExpectResult.Value = False) Then
    
        x = MsgBox("請勾選檢查項目", 0 + 16, "Error")
    
    ElseIf CheckScriptList.ListCount = 0 Then
    
        x = MsgBox("請選擇TestScript", 0 + 16, "Error")
    
    End If
    
End Sub

Private Sub delete_Click()
    For i = 0 To CheckScriptList.ListCount - 1
    
        If CheckScriptList.selected(i) = True Then
            CheckScriptList.RemoveItem (i)
        End If

    Next i
End Sub

Private Sub UserForm_Activate()
    i = 0
    Do
        If ThisWorkbook.Sheets(i + 1).Visible = True And Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
            TotalScriptList.AddItem (ThisWorkbook.Sheets(i + 1).Name)
        End If
    
    i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
End Sub

Private Sub UserForm_Click()

End Sub
