VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExpectResultCreator 
   Caption         =   "ExpectResult Creator"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5820
   OleObjectBlob   =   "ExpectResultCreator.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "ExpectResultCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub addCase_Click()
    Dim isCaseDuplicate As Boolean
    isCaseDuplicate = False
    Do
        newCaseName = InputBox("請輸入Case Name", "新增Case")
        
        If StrPtr(newCaseName) = 0 Then
            
            clickCancel = 1
    
         Else
         
            If newCaseName <> "" Then
                
                '判斷是否存在相同Case
                For i = 0 To CaseName.ListCount - 1
                
                    If CaseName.List(i) = newCaseName Then
                        
                        isCaseDuplicate = True
                        Exit For
                    
                    End If
                
                Next i
'                i = 2
'                Do
'                    If Sheets("ExpectResult").Cells(i, "A") = newCaseName Then
'                        isCaseDuplicate = True
'                        Exit Do
'                    End If
'                    i = i + 1
'                Loop Until Sheets("ExpectResult").Cells(i, "A") = ""
                
                If isCaseDuplicate = True Then
                    x = MsgBox("Case已存在!", 0 + 64, "Message")
                    CaseName.Text = newCaseName
                Else
                    CaseName.AddItem (newCaseName)
                    CaseName.Text = newCaseName
                End If
        
            Else
    
                Z = MsgBox("請填入Case Name", 0 + 48, "Message")
            End If
        
        End If
        
        If clickCancel = 1 Then Exit Do
    Loop Until newCaseName <> ""
End Sub

Private Sub addString_Click()
    Dim StringSelected As Boolean
    Dim duplicate As Boolean
    newString = ExpectString.Text
    
    If newString <> "" Then
        
        '確認Sting是否重複
        For i = 0 To StringList.ListCount - 1
            
            If StringList.List(i) = newString Then
                
                duplicate = True
                Exit For
                
            Else
                duplicate = False
            End If
    
        Next i
        
        
        If duplicate = False Then
            For i = 0 To StringList.ListCount - 1
            
                If StringList.selected(i) = True Then
                    StringSelected = True
                    StringList.AddItem newString, i
                    StringList.RemoveItem (i + 1)
                    StringList.selected(i) = True
                    x = MsgBox("Edit Successfully.", 0 + 64, "Message")
                    Exit For
                    
                Else
                    StringSelected = False
                
                End If
            Next i

            If StringSelected = False And newString <> "" Then
                StringList.AddItem newString
                x = MsgBox("Add Successfully.", 0 + 64, "Message")
            End If

        End If
        
        
    End If
    
    
End Sub

Private Sub cancelSelect_Click()
    ExpectString.Text = ""
    
    For i = 0 To StringList.ListCount - 1
        
        If StringList.selected(i) = True Then StringList.selected(i) = False: Exit For
    
    Next i
End Sub

Private Sub CaseName_Change()
    StringList.clear
    i = 2
    Do
        If CaseName.Text = Sheets("ExpectResult").Cells(i, "A") Then
        
            j = 2
            Do
            
                StringList.AddItem (Sheets("ExpectResult").Cells(i, j))
                
                j = j + 1
            Loop Until Sheets("ExpectResult").Cells(i, j) = ""
            Exit Do
        End If
        
        i = i + 1
    Loop Until Sheets("ExpectResult").Cells(i, "A") = ""
End Sub


Private Sub Create_Click()
    Dim newCase As Boolean
    Dim isNewString As Boolean
    newCase = True
    isNewString = True
    Application.ScreenUpdating = False
    If CaseName.Text = "" Then
        
        x = MsgBox("請選擇Case Name", 0 + 48, "Message")
    
    ElseIf StringList.ListCount = 0 Then
    
        x = MsgBox("請加入比對字串", 0 + 48, "Message")
    Else
        ' 寫入ExpectResult sheet
        ' 判斷是否為新Case
        i = 2
        Do
            If Sheets("ExpectResult").Cells(i, "A") = CaseName.Text Then
                
                newCase = False
                Exit Do
                
            End If
            i = i + 1
        Loop Until Sheets("ExpectResult").Cells(i, "A") = ""
        
        If newCase = False Then
            Sheets("ExpectResult").Select
            ' 刪除舊有字串
            j = 2
            Do
                Sheets("ExpectResult").Cells(i, j).Select
                Selection.delete Shift:=xlToLeft
            Loop Until Sheets("ExpectResult").Cells(i, j) = ""
            
            For y = 0 To StringList.ListCount - 1
                
                Sheets("ExpectResult").Cells(i, j) = StringList.List(y)
                j = j + 1
            Next y
        Else
            j = 2
            Sheets("ExpectResult").Cells(i, "A") = CaseName.Text
            For y = 0 To StringList.ListCount - 1
                Sheets("ExpectResult").Cells(i, j) = StringList.List(y)
                j = j + 1
            Next y
        
        End If
        Sheets("ExpectResult").Cells(i, "A").Select
        x = MsgBox("Done.", 0 + 64, "Message")
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub delete_Click()
    For i = 0 To StringList.ListCount - 1
    
        If StringList.selected(i) = True Then
            
            x = MsgBox("確定移除" & vbNewLine & "「" & StringList.List(i) & "」", 32 + 1, "詢問")
            
            If x = 1 Then
                StringList.RemoveItem (i)
                StringList.selected(i) = False
                ExpectString.Text = ""
            End If
        
        End If
    Next i
End Sub



Private Sub StringList_Click()
    For i = 0 To StringList.ListCount - 1

        If StringList.selected(i) = True Then

            ExpectString.Text = StringList.List(i)
            Exit For
        End If
    Next i
End Sub



Private Sub UserForm_Activate()
    i = 2
    Do
        CaseName.AddItem (Sheets("ExpectResult").Cells(i, "A"))
        i = i + 1
    Loop Until Sheets("ExpectResult").Cells(i, "A") = ""
    
    If Sheets("ExpectResult").Visible = False Then Sheets("ExpectResult").Visible = True
    
End Sub

