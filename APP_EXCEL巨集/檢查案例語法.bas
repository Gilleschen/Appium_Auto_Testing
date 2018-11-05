Attribute VB_Name = "檢查案例語法"
Function CheckCommand_2(TestScriptName As String)
    Application.ScreenUpdating = False
    Dim sheetname As String
    Dim CaseName, LaunchAPP, Byid_Result, ByXpath_Result, QuitAPP As Integer
    
    
    CaseName = 0: LaunchAPP = 0: Byid_Result = 0: ByXpath_Result = 0: QuitAPP = 0
    'sheetname = ThisWorkbook.Sheets(i + 1).Name
    sheetname = TestScriptName
    j = 1
    Do
    
        Select Case Sheets(TestScriptName).Cells(j, "A")
        
        Case "CaseName"
        
            CaseName = CaseName + 1
        
        Case "LaunchAPP"
        
            LaunchAPP = LaunchAPP + 1
        
        Case "QuitAPP"
        
            QuitAPP = QuitAPP + 1
        
        'Case "Byid_Result"

            'Byid_Result = Byid_Result + 1
            
        'Case "ByXpath_Result"
            
            'ByXpath_Result = ByXpath_Result + 1

        End Select
        
    j = j + 1
    Loop Until Sheets(TestScriptName).Cells(j, "A") = ""
    
    If LaunchAPP <> CaseName Then
        x = MsgBox(TestScriptName & "中缺少LaunchAPP或CaseName", 0 + 16, "Error")
        CheckCommand_2 = False
        Exit Function
    Else
        CheckCommand_2 = True
    End If

    If QuitAPP <> CaseName Then
        x = MsgBox(TestScriptName & "中缺少QuitAPP或CaseName", 0 + 16, "Error")
        CheckCommand_2 = False
        Exit Function
    Else
        CheckCommand_2 = True
    End If

    Call Classification_TestCase_2(TestScriptName)
    Sheets("APP&Device").Select
End Function



Function CheckCommand()
    Application.ScreenUpdating = False
    Dim sheetname As String
    Dim CaseName, LaunchAPP, Byid_Result, ByXpath_Result, QuitAPP As Integer
    
    i = 0
    Do
        
        If Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" And ThisWorkbook.Sheets(i + 1).Visible = True Then
            'If ThisWorkbook.Sheets(i + 1).Visible = True Then
                CaseName = 0: LaunchAPP = 0: Byid_Result = 0: ByXpath_Result = 0: QuitAPP = 0
                sheetname = ThisWorkbook.Sheets(i + 1).Name
                j = 1
                Do
                
                    Select Case Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A")
                    
                    Case "CaseName"
                    
                        CaseName = CaseName + 1
                    
                    Case "LaunchAPP"
                    
                        LaunchAPP = LaunchAPP + 1
                    
                    Case "QuitAPP"
                    
                        QuitAPP = QuitAPP + 1
                    
                    'Case "Byid_Result"
    
                        'Byid_Result = Byid_Result + 1
                        
                    'Case "ByXpath_Result"
                        
                        'ByXpath_Result = ByXpath_Result + 1
    
                    End Select
                    
                j = j + 1
                Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = ""
                
                If LaunchAPP <> CaseName Then
                    x = MsgBox(sheetname & "中缺少LaunchAPP或CaseName", 0 + 16, "Error")
                    CheckCommand = False
                    Exit Function
                Else
                    CheckCommand = True
                End If
 
                If QuitAPP <> CaseName Then
                    x = MsgBox(sheetname & "中缺少QuitAPP或CaseName", 0 + 16, "Error")
                    CheckCommand = False
                    Exit Function
                Else
                    CheckCommand = True
                End If

                'If Byid_Result <> CaseName Or ByXpath_Result <> CaseName Then x = MsgBox(sheetname & "中缺少Byid_Result或CaseName", 0 + 16, "Error")
              
            'End If
        
        End If

        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    Call Classification_TestCase
    Sheets("APP&Device").Select
End Function



Sub Classification_TestCase_2(TestScriptName As String)
    Dim row As String
    Dim color As Integer
    color = 1
    
    Application.ScreenUpdating = False
   
    start_count = 1
    Count = 1
    
    
    
        sheetname = TestScriptName 'ThisWorkbook.Sheets(i + 1).Name
        Sheets(sheetname).Select
        j = 1
        
        Do
           
            Do
            
                Count = Count + 1
        
            Loop Until Sheets(sheetname).Cells(Count, "A") = "QuitAPP"
            
            row = start_count & ":" & Count
            start_count = Count + 1
            
            color = color * (-1)
            
            Rows(row).Select
            
            If color < 0 Then
            
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.8
                .PatternTintAndShade = 0
            End With
            Else
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.8
                .PatternTintAndShade = 0
            End With
            End If
        j = start_count
        Loop Until Sheets(sheetname).Cells(j, "A") = ""
        
End Sub

Sub Classification_TestCase()
    Dim row As String
    Dim color As Integer
    color = 1
    
    Application.ScreenUpdating = False
   
    i = 0
    Do
        start_count = 1
        Count = 1
        If ThisWorkbook.Sheets(i + 1).Visible = True And Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
        
        
            sheetname = ThisWorkbook.Sheets(i + 1).Name
            Sheets(sheetname).Select
            j = 1
            
            Do
               
                Do
                
                    Count = Count + 1
            
                Loop Until Sheets(sheetname).Cells(Count, "A") = "QuitAPP"
                
                row = start_count & ":" & Count
                start_count = Count + 1
                
                color = color * (-1)
                
                Rows(row).Select
                
                If color < 0 Then
                
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.8
                    .PatternTintAndShade = 0
                End With
                Else
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.8
                    .PatternTintAndShade = 0
                End With
                End If
            j = start_count
            Loop Until Sheets(sheetname).Cells(j, "A") = ""
            
        End If
    i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
End Sub

