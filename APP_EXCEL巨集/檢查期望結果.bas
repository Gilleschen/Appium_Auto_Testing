Attribute VB_Name = "檢查期望結果"
Function CheckExpectResult_2(TestScriptName As String)
    Application.ScreenUpdating = False
    Dim result As String
    Dim x As Integer
    
    j = 1: x = 0
    Do
        If Sheets(TestScriptName).Cells(j, "A") = "CaseName" Then
            
            x = x + 1
        
        End If
    
        
        j = j + 1
    Loop Until Sheets(TestScriptName).Cells(j, "A") = ""
    
    ReDim casename(x - 1)
    
    j = 1: x = 0
    
    Do
        If Sheets(TestScriptName).Cells(j, "A") = "CaseName" Then
            
            casename(x) = Sheets(TestScriptName).Cells(j, "B")
            x = x + 1
            
        End If
    
        j = j + 1
    Loop Until Sheets(TestScriptName).Cells(j, "A") = ""



    k = 0
    Do
        j = 2
        Do
            
            If casename(k) = Sheets("ExpectResult").Cells(j, "A") Then
                result = "Pass"
                Exit Do
            End If
            
            j = j + 1
        Loop Until Sheets("ExpectResult").Cells(j, "A") = ""
        
        If result <> "Pass" Then
            x = MsgBox(casename(k) + "的期望結果為未寫入ExpectResult", 0 + 16, "Error")
            CheckExpectResult_2 = False
            Exit Function
        End If
        
        result = ""
        
        k = k + 1
    Loop Until k = UBound(casename) - LBound(casename) + 1
    
    CheckExpectResult_2 = True
        
    
End Function



Sub CheckExpectResult()
    Application.ScreenUpdating = False
    Dim result As String
    Dim x As Integer
    i = 0
    Do
        If Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" And ThisWorkbook.Sheets(i + 1).Visible = True Then
        
            
            j = 1: x = 0
            Do
                If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = "CaseName" Then
                    
                    x = x + 1
                
                End If
            
                
                j = j + 1
            Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = ""
            
            ReDim casename(x - 1)
            
            j = 1: x = 0
            
            Do
                If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = "CaseName" Then
                    
                    casename(x) = Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B")
                    x = x + 1
                    
                End If
            
                j = j + 1
            Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = ""
        
        
        
            k = 0
            Do
                j = 2
                Do
                    
                    If casename(k) = Sheets("ExpectResult").Cells(j, "A") Then
                        result = "Pass"
                        Exit Do
                    End If
                    
                    j = j + 1
                Loop Until Sheets("ExpectResult").Cells(j, "A") = ""
                
                If result <> "Pass" Then x = MsgBox(casename(k) + "的期望結果為未寫入ExpectResult工作表", 0 + 16, "Error")
                
                result = ""
                
                k = k + 1
            Loop Until k = UBound(casename) - LBound(casename) + 1
        
        End If

        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
End Sub


Function CheckExpectResult_Ver2_2(TestScriptName As String)
    Application.ScreenUpdating = False
    Dim tempCaseName As String
    
    k = 1
    Do
        j = k
        Do
            If Sheets(TestScriptName).Cells(j, "A") = "CaseName" Then
            
                tmpCaseName = Sheets(TestScriptName).Cells(j, "B")
                CaseRow = j
            
            End If
            
            If Sheets(TestScriptName).Cells(j, "A") = "Byid_VerifyText" Or Sheets(TestScriptName).Cells(j, "A") = "ByXpath_VerifyText" Then
                
                Call CompareExpectResult_2(tmpCaseName, CaseRow, TestScriptName)
            
            End If
            
            j = j + 1
        Loop Until Sheets(TestScriptName).Cells(j, "A") = "QuitAPP"
    
        k = j + 1
    Loop Until Sheets(TestScriptName).Cells(k, "A") = ""
    
    x = CheckExpectResultisEmpty()
        
End Function


Function CheckExpectResult_Ver2()
    Application.ScreenUpdating = False
    Dim tempCaseName As String
    i = 0
    Do
        If Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" And ThisWorkbook.Sheets(i + 1).Visible = True Then
        
            k = 1
            Do
                j = k
                Do
                    If ThisWorkbook.Sheets(i + 1).Cells(j, "A") = "CaseName" Then
                    
                        tmpCaseName = ThisWorkbook.Sheets(i + 1).Cells(j, "B")
                        CaseRow = j
                    
                    End If
                    
                    If ThisWorkbook.Sheets(i + 1).Cells(j, "A") = "Byid_VerifyText" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A") = "ByXpath_VerifyText" Then
                        
                        CheckExpectResult_Ver2 = CompareExpectResult(tmpCaseName, CaseRow, i)
                    
                    End If
                    
                    j = j + 1
                Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = "QuitAPP"
            
                k = j + 1
            Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(k, "A") = ""
    
        End If
    
        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
    
    x = CheckExpectResultisEmpty()
End Function

Sub CompareExpectResult_2(casename, CaseRow, scriptname)
    Dim result As Boolean
    j = 2
    Do
        If casename = Sheets("ExpectResult").Cells(j, "A") Then
        
            result = True
            Exit Do
        End If
    
        j = j + 1
    Loop Until Sheets("ExpectResult").Cells(j, "A") = ""
    
    If result = False Then
        x = MsgBox(casename + "的期望結果為未寫入ExpectResult", 0 + 16, "Error")
        Sheets(scriptname).Cells(CaseRow, "B").Font.color = RGB(255, 0, 0)
        
    Else
        
        Sheets(scriptname).Cells(CaseRow, "B").Font.color = RGB(0, 0, 0)
    End If
    
End Sub
Function CompareExpectResult(casename, CaseRow, i)
    Dim result As Boolean
    j = 2
    Do
        If casename = Sheets("ExpectResult").Cells(j, "A") Then
            
            result = True
            Exit Do
        End If
    
        j = j + 1
    Loop Until Sheets("ExpectResult").Cells(j, "A") = ""
    
    If result = False Then
        x = MsgBox(casename + "的期望結果為未寫入ExpectResult", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(CaseRow, "B").Font.color = RGB(255, 0, 0)
        CompareExpectResult = False
        
    Else
        
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(CaseRow, "B").Font.color = RGB(0, 0, 0)
        CompareExpectResult = True
    End If
    
End Function


Function CheckExpectResultisEmpty()

    
    Sheets("ExpectResult").Select
    i = 2
    Do
        j = 1: x = 0
        Do While Sheets("ExpectResult").Cells(i, j + 1) <> ""
            x = x + 1
            j = j + 1
        Loop
        
        If x > 0 Then
            Sheets("ExpectResult").Cells(i, "A").Select
            With Selection.Font
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
            End With
            CheckExpectResultisEmpty = True
        Else
            x = MsgBox(Sheets("ExpectResult").Cells(i, "A") & "缺少ExpectResult", 0 + 16, "Error")
            Sheets("ExpectResult").Cells(i, "A").Select
            With Selection.Font
                .color = -16776961
                .TintAndShade = 0
            End With
            CheckExpectResultisEmpty = False
            Exit Function
            
        End If

        i = i + 1
    Loop Until Sheets("ExpectResult").Cells(i, "A") = ""
    


End Function
