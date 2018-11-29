Attribute VB_Name = "檢查資訊"

Function CheckAPPandDevice()
    Dim sheetname As String
    Dim scriptnumber, result As Integer
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    i = 1
    Do
        If Sheets("APP&Device").Cells(1, i) <> "CaseName" Then
            If Sheets("APP&Device").Cells(2, i) = "" Then
            
                x = MsgBox("請填入" & Sheets("APP&Device").Cells(1, i), 0 + 16, "Error")
                Sheets("APP&Device").Cells(2, i).Interior.color = RGB(255, 0, 0)
                CheckAPPandDevice = False
                Exit Function
            Else
            
                Sheets("APP&Device").Cells(2, i).Interior.Pattern = xlNone
                CheckAPPandDevice = True
                
            End If
        End If
        i = i + 1
    Loop Until Sheets("APP&Device").Cells(1, i) = ""
    
    i = 2
    Do
        If Sheets("APP&Device").Cells(i, "C") = "" Then
            Sheets("APP&Device").Cells(i, "C").Interior.color = RGB(255, 0, 0)
            x = MsgBox("請填入OS " & Sheets("APP&Device").Cells(i, "D").Value & " UDID", 0 + 16, "Error")
            CheckAPPandDevice = False
            Exit Function
            Exit Do
            
        ElseIf Sheets("APP&Device").Cells(i, "D") = "" Then
            Sheets("APP&Device").Cells(i, "D").Interior.color = RGB(255, 0, 0)
            x = MsgBox("請填入" & Sheets("APP&Device").Cells(i, "C") & " OS Version", 0 + 16, "Error")
            CheckAPPandDevice = False
            Exit Function
            Exit Do
        Else
            
            Sheets("APP&Device").Cells(i, "C").Interior.Pattern = xlNone
            Sheets("APP&Device").Cells(i, "D").Interior.Pattern = xlNone
            CheckAPPandDevice = True
            
        End If
        
        i = i + 1
    Loop Until Sheets("APP&Device").Cells(i, "C") = "" And Sheets("APP&Device").Cells(i, "D") = ""
    
    j = 2: scriptnumber = 0
    Do
        scriptnumber = scriptnumber + 1
    j = j + 1
    Loop Until Sheets("APP&Device").Cells(j, "E") = ""
    
    ReDim scriptarray(scriptnumber - 1) As String
    
    
    j = 2: x = 0
    Do
        scriptarray(x) = Sheets("APP&Device").Cells(j, "E")
    j = j + 1: x = x + 1
    Loop Until Sheets("APP&Device").Cells(j, "E") = ""
    
    
    i = 0
    Do
        j = 0: result = 0
        Do
            sheetname = ThisWorkbook.Sheets(j + 1).Name
            If scriptarray(i) <> sheetname Then result = result + 1
    
            j = j + 1
        Loop Until j = ThisWorkbook.Sheets.Count
        If result = ThisWorkbook.Sheets.Count Then
            y = MsgBox("找不到" & scriptarray(i) & "工作表", 0 + 16, "Error")
            CheckAPPandDevice = False
            Exit Function
        End If

        i = i + 1
    Loop Until i = UBound(scriptarray) - LBound(scriptarray) + 1
    
    i = 2
    Do
    
        If Right(Sheets("APP&Device").Cells(i, "E"), 11) <> "_TestScript" Then
            
            y = MsgBox("ScriptName欄位請填入以_TestScript為結尾之工作表(大小寫有分)", 0 + 16, "Error")
            Sheets("APP&Device").Cells(i, "E").Font.color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
        Else
            Sheets("APP&Device").Cells(i, "E").Font.color = RGB(0, 0, 0)
            CheckAPPandDevice = True
        End If
    
    i = i + 1
    Loop Until Sheets("APP&Device").Cells(i, "E") = ""
    
'    If Sheets("APP&Device").Cells(2, "G") = "" Then
'
'        x = MsgBox("請填入Jar檔路徑" & vbNewLine & "例如：C:\Users\Desktop\檔名.jar", 0 + 16, "Error")
'        CheckAPPandDevice = False
'        Exit Function
'    ElseIf Dir(CStr(Sheets("APP&Device").Cells(2, "G"))) = "" Then
'
'        x = MsgBox("找不到" & Sheets("APP&Device").Cells(2, "G"), 0 + 16, "Error")
'        CheckAPPandDevice = False
'        Exit Function
'
'    End If
    
    '確認CaseName欄位
    Sheets("APP&Device").Select
    TotalCaseName = Sheets("APP&Device").Cells(2, "F").Text
    For w = 1 To 10
        Sheets("APP&Device").Cells(2, "F").Select
        Selection.delete Shift:=xlUp
    Next w
    Sheets("APP&Device").Cells(2, "F") = TotalCaseName
    i = 2
    Do
        
        If Sheets("APP&Device").Cells(i, "F") <> "" Then
        
            strArray = Split(Sheets("APP&Device").Cells(i, "F"), ",")
            
            For intCount = LBound(strArray) To UBound(strArray)
                
                j = 1
                
                Do
                    
                    If strArray(intCount) = Sheets(Sheets("APP&Device").Cells(i, "E").Text).Cells(j, "B") Then
                        strResult = True
                        CheckAPPandDevice = True
                        Sheets("APP&Device").Cells(i, "F").Font.color = RGB(0, 0, 0)
                        Exit Do
                    Else
                        strResult = False
                    End If
                    
                    If strResult = False And Sheets(Sheets("APP&Device").Cells(i, "E").Text).Cells(j + 1, "A") = "" Then
                        y = MsgBox(Sheets("APP&Device").Cells(i, "E") & "工作表中，找不到" & strArray(intCount) & "案例", 0 + 16, "Error")
                        CheckAPPandDevice = False
                        Sheets("APP&Device").Cells(i, "F").Font.color = RGB(255, 0, 0)
                        Application.EnableEvents = True
                        Exit Function
                    End If
                    
                j = j + 1
                Loop Until Sheets(Sheets("APP&Device").Cells(i, "E").Text).Cells(j, "A") = ""
                
                
            Next intCount
            
        End If
       
    i = i + 1
    Loop Until Sheets("APP&Device").Cells(i, "E") = ""

    
    '確認ReSet APP欄位
    Sheets("APP&Device").Cells(2, "G").NumberFormatLocal = "G/通用格式"
    If Sheets("APP&Device").Cells(2, "G") = "False" Or Sheets("APP&Device").Cells(2, "G") = "FALSE" Or Sheets("APP&Device").Cells(2, "G") = "false" Then
        
        Sheets("APP&Device").Cells(2, "G") = "False"
        Sheets("APP&Device").Cells(2, "G").NumberFormatLocal = "G/通用格式"
        Sheets("APP&Device").Cells(2, "G").Font.color = RGB(0, 0, 0)
        CheckAPPandDevice = True
        
    ElseIf Sheets("APP&Device").Cells(2, "G") = "True" Or Sheets("APP&Device").Cells(2, "G") = "TRUE" Or Sheets("APP&Device").Cells(2, "G") = "true" Then
    
        Sheets("APP&Device").Cells(2, "G") = "True"
        Sheets("APP&Device").Cells(2, "G").NumberFormatLocal = "G/通用格式"
        Sheets("APP&Device").Cells(2, "G").Font.color = RGB(0, 0, 0)
        CheckAPPandDevice = True
    Else
        y = MsgBox("ResetAPP欄位請輸入大寫TRUE或FALSE", 0 + 16, "Error")
        Sheets("APP&Device").Cells(2, "G").Font.color = RGB(255, 0, 0)
        CheckAPPandDevice = False
        Exit Function
        
    End If
    
    
    '確認UIAutomator 2欄位
    Sheets("APP&Device").Cells(2, "H").NumberFormatLocal = "G/通用格式"
    If Sheets("APP&Device").Cells(2, "H") = "False" Or Sheets("APP&Device").Cells(2, "H") = "FALSE" Or Sheets("APP&Device").Cells(2, "H") = "false" Then
        
        Sheets("APP&Device").Cells(2, "H") = "False"
        Sheets("APP&Device").Cells(2, "H").NumberFormatLocal = "G/通用格式"
        Sheets("APP&Device").Cells(2, "H").Font.color = RGB(0, 0, 0)
        CheckAPPandDevice = True
        
    ElseIf Sheets("APP&Device").Cells(2, "H") = "True" Or Sheets("APP&Device").Cells(2, "H") = "TRUE" Or Sheets("APP&Device").Cells(2, "H") = "true" Then
    
        Sheets("APP&Device").Cells(2, "H") = "True"
        Sheets("APP&Device").Cells(2, "H").NumberFormatLocal = "G/通用格式"
        Sheets("APP&Device").Cells(2, "H").Font.color = RGB(0, 0, 0)
        CheckAPPandDevice = True
    Else
        y = MsgBox("UIAutomator 2欄位請輸入大寫TRUE或FALSE", 0 + 16, "Error")
        Sheets("APP&Device").Cells(2, "H").Font.color = RGB(255, 0, 0)
        CheckAPPandDevice = False
        Exit Function
        
    End If
    
    Application.EnableEvents = True
End Function
