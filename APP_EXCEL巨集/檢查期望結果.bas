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
    
    ReDim CaseName(x - 1)
    
    j = 1: x = 0
    
    Do
        If Sheets(TestScriptName).Cells(j, "A") = "CaseName" Then
            
            CaseName(x) = Sheets(TestScriptName).Cells(j, "B")
            x = x + 1
            
        End If
    
        j = j + 1
    Loop Until Sheets(TestScriptName).Cells(j, "A") = ""



    k = 0
    Do
        j = 2
        Do
            
            If CaseName(k) = Sheets("ExpectResult").Cells(j, "A") Then
                result = "Pass"
                Exit Do
            End If
            
            j = j + 1
        Loop Until Sheets("ExpectResult").Cells(j, "A") = ""
        
        If result <> "Pass" Then
            x = MsgBox(CaseName(k) + "的期望結果為未寫入ExpectResult", 0 + 16, "Error")
            CheckExpectResult_2 = False
            Exit Function
        End If
        
        result = ""
        
        k = k + 1
    Loop Until k = UBound(CaseName) - LBound(CaseName) + 1
    
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
            
            ReDim CaseName(x - 1)
            
            j = 1: x = 0
            
            Do
                If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = "CaseName" Then
                    
                    CaseName(x) = Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B")
                    x = x + 1
                    
                End If
            
                j = j + 1
            Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = ""
        
        
        
            k = 0
            Do
                j = 2
                Do
                    
                    If CaseName(k) = Sheets("ExpectResult").Cells(j, "A") Then
                        result = "Pass"
                        Exit Do
                    End If
                    
                    j = j + 1
                Loop Until Sheets("ExpectResult").Cells(j, "A") = ""
                
                If result <> "Pass" Then x = MsgBox(CaseName(k) + "的期望結果為未寫入ExpectResult", 0 + 16, "Error")
                
                result = ""
                
                k = k + 1
            Loop Until k = UBound(CaseName) - LBound(CaseName) + 1
        
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
End Function

Sub CompareExpectResult_2(CaseName, CaseRow, ScriptName)
    Dim result As Boolean
    j = 2
    Do
        If CaseName = Sheets("ExpectResult").Cells(j, "A") Then
        
            result = True
            Exit Do
        End If
    
        j = j + 1
    Loop Until Sheets("ExpectResult").Cells(j, "A") = ""
    
    If result = False Then
        x = MsgBox(CaseName + "的期望結果為未寫入ExpectResult", 0 + 16, "Error")
        Sheets(ScriptName).Cells(CaseRow, "B").Font.color = RGB(255, 0, 0)
        
    Else
        
        Sheets(ScriptName).Cells(CaseRow, "B").Font.color = RGB(0, 0, 0)
    End If
    
End Sub
Function CompareExpectResult(CaseName, CaseRow, i)
    Dim result As Boolean
    j = 2
    Do
        If CaseName = Sheets("ExpectResult").Cells(j, "A") Then
        
            result = True
            Exit Do
        End If
    
        j = j + 1
    Loop Until Sheets("ExpectResult").Cells(j, "A") = ""
    
    If result = False Then
        x = MsgBox(CaseName + "的期望結果為未寫入ExpectResult", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(CaseRow, "B").Font.color = RGB(255, 0, 0)
        CompareExpectResult = False
        
    Else
        
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(CaseRow, "B").Font.color = RGB(0, 0, 0)
        CompareExpectResult = True
    End If
    
End Function
