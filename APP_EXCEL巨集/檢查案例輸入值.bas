Attribute VB_Name = "檢查案例輸入值"
Function CheckValue()
    Application.ScreenUpdating = False
    Dim sheetname As String
    Dim xpath, id As String
    xpath = "xpath": id = "id"
    i = 0
    Do
        If ThisWorkbook.Sheets(i + 1).Visible = True And Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
            'If Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
        
                sheetname = ThisWorkbook.Sheets(i + 1).Name
                j = 1
                Do
                
                    Select Case Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A")
                    
                    Case "CaseName"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行缺少CaseName", 0 + 16, "Error")
                           Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                           CheckValue = False
                           Exit Function
                           
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                         CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                    
                    Case "Byid_Click"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    
                    Case "ByXpath_Click"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "Byid_LongPress"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_LongPress"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function

                    Case "Byid_SendKey"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkSendKey(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "D")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_SendKey"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkSendKey(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "D")
                        If CheckValue = False Then Exit Function
                        
                    Case "Byid_Clear"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_Clear"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "Byid_invisibility"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_invisibility"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                              
                    Case "Byid_Swipe"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkSwipe(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "D")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_Swipe"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkSwipe(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkSwipeXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "D")
                        If CheckValue = False Then Exit Function
                    
                    Case "HideKeyboard"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入HideKeyboard", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0): CheckValue = False: Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                    Case "LaunchAPP"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入LaunchAPP", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0): CheckValue = False: Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                    
                    Case "QuitAPP"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入QuitAPP", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0): CheckValue = False: Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                    
                    Case "Byid_VerifyText"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        'If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
                        CheckValue = checkResult(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_VerifyText"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        'If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
                        CheckValue = checkResult(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
    
                    Case "Byid_VerifyRadioButton"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        'If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
                        CheckValue = checkResult(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkVerifyRadioButton(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "D")
                        If CheckValue = False Then Exit Function
   
                    Case "ByXpath_VerifyRadioButton"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        'If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
                        CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkVerifyRadioButton(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "D")
                        If CheckValue = False Then Exit Function
    
                    Case "ResetAPP"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "僅能填入ResetAPP", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0): CheckValue = False: Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                    
                    Case "Byid_Wait"
                    
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkWait(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                    
                    Case "ByXpath_Wait"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkWait(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                    
                    Case "Sleep"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行缺少秒數", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        ElseIf IsNumeric(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B")) = False Then
                            x = MsgBox(sheetname & "中，第" & j & "行請輸入數值", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        
                        Else
                            If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Value) <> "String" Then
                               Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B")
                            End If
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "ScreenShot"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入ScreenShot", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                    Case "Orientation"
                    
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
                            x = MsgBox(sheetname & "中，第" & j & "請填入Landscape或Portrait", 0 + 16, "Error")
                             Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                             CheckValue = False
                             Exit Function
                        ElseIf Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "Landscape" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "Portrait" Then
                        
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入Landscape或Portrait (大小寫有分)", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "Back"
                    
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入Back", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                    
                    Case "Home"
                    
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入Home", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                    Case "Power"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入Power", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                    Case "Recent"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入Recent", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                    Case "Customized"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入Customized", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                    Case "WiFi"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行請填入On或Off (大小寫有分)", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        ElseIf Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "Off" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "On" Then
                            x = MsgBox(sheetname & "中，第" & j & "行僅能填入On或Off (大小寫有分)", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                    
                    Case "Swipe"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkSwipeData(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "G")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_Swipe_Vertical"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkSwipeVertical_HorHorizontal(sheetname, i, j, "Vertical")
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "E")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_Swipe_Horizontal"
                        
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkSwipeVertical_HorHorizontal(sheetname, i, j, "Horizontal")
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "E")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_Swipe_FindText_Click_Android"
                    
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = ByXpath_Swipe_FindText_Click_Android(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "G")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_Swipe_FindText_Click_iOS"
                    
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(0, 0, 0)
                        CheckValue = ByXpath_Swipe_FindText_Click_iOS(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "E")
                        If CheckValue = False Then Exit Function
                    
                    Case Else
                        
                        x = MsgBox(sheetname & "中，第" & j & "行語法有誤，" & "無" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Value & " 語法", 0 + 16, "Error")
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.color = RGB(255, 0, 0)
                        CheckValue = False
                        Exit Function
                    End Select
                    
                    
                j = j + 1
                Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = ""
            
           ' End If
    
            
        End If
        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
    CheckValue2 = Delete_All_Blank_Cells
End Function

Function checkVerifyRadioButton(sheetname, i, j)
    
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "True" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "False" Then

        x = MsgBox(sheetname & "中，第" & j & "列第C欄請填入TRUE/FALSE", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
        checkVerifyRadioButton = False
        Exit Function
    ElseIf TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Value) = "Boolean" Then

        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
        checkVerifyRadioButton = True
    Else
        checkVerifyRadioButton = True
    End If

End Function



Function checkExcessData(sheetname, i, j, col) '檢查所有指令最後一欄是否為空白或無資料

    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, col) <> "" Then
                                
        x = MsgBox(sheetname & "中，第" & j & "列第" & col & "欄請保持空白", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, col).Interior.color = RGB(255, 0, 0)
        checkExcessData = False
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, col).Interior.Pattern = xlNone
        checkExcessData = True
    End If

End Function




Function ByXpath_Swipe_FindText_Click_iOS(sheetname, i, j)

    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
        x = MsgBox(sheetname & "中，第" & j & "行缺少xpath", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_iOS = False
        Exit Function
    ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) <> "//*[@" And Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 6) <> "(//*[@" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_iOS = False
        Exit Function
    ElseIf Right(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 1) <> "]" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_iOS = False
        Exit Function
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Font.color = RGB(0, 0, 0)
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
        ByXpath_Swipe_FindText_Click_iOS = True
    End If
    
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
        
        x = MsgBox(sheetname & "中，第" & j & "列第C欄缺少UP/DOWN", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_iOS = False
        Exit Function
        
    ElseIf Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "LEFT" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "RIGHT" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "UP" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "DOWN" Then
        
        x = MsgBox(sheetname & "中，第" & j & "列第C欄ByXpath_Swipe_FindText_Click_iOS方法只包含UP/DOWN/LEFT/RIGHT", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_iOS = False
        Exit Function
    Else
        
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Font.color = RGB(0, 0, 0)
        ByXpath_Swipe_FindText_Click_iOS = True
    End If

    If Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D"), 11) <> "//*[@text='" Then
        x = MsgBox(sheetname & "中，第" & j & "行僅能輸入//*[@text='String']格式", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_iOS = False
        Exit Function
        
    ElseIf Right(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D"), 2) <> "']" Then
        x = MsgBox(sheetname & "中，第" & j & "行僅能輸入//*[@text='String']格式", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_iOS = False
        Exit Function
        
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Font.color = RGB(0, 0, 0)
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.Pattern = xlNone
        ByXpath_Swipe_FindText_Click_iOS = True
    End If

End Function


Function ByXpath_Swipe_FindText_Click_Android(sheetname, i, j)
    
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
        x = MsgBox(sheetname & "中，第" & j & "行缺少xpath", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
        
    ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) <> "//*[@" And Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 6) <> "(//*[@" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
        
    ElseIf Right(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 1) <> "]" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
        
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Font.color = RGB(0, 0, 0)
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
        ByXpath_Swipe_FindText_Click_Android = True
    End If
    
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D") = "" Then
        x = MsgBox(sheetname & "中，第" & j & "行缺少xpath", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
    
    ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D"), 5) <> "//*[@" And Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D"), 6) <> "(//*[@" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
    ElseIf Right(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D"), 1) <> "]" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
        
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Font.color = RGB(0, 0, 0)
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.Pattern = xlNone
        ByXpath_Swipe_FindText_Click_Android = True
    End If
    
    If Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F"), 11) <> "//*[@text='" Then
        x = MsgBox(sheetname & "中，第" & j & "行僅能輸入//*[@text='String']格式", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
        
    ElseIf Right(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F"), 2) <> "']" Then
        x = MsgBox(sheetname & "中，第" & j & "行僅能輸入//*[@text='String']格式", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F").Interior.Pattern = xlNone
        ByXpath_Swipe_FindText_Click_Android = True
    End If
    
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
        
        x = MsgBox(sheetname & "中，第" & j & "列第C欄缺少UP/DOWN/LEFT/RIGHT", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
        
    ElseIf Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "UP" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "DOWN" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "LEFT" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "RIGHT" Then
        
        x = MsgBox(sheetname & "中，第" & j & "列第C欄ByXpath_Swipe_FindText_Click_Android方法只包含UP/DOWN/LEFT/RIGHT", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
        
    Else
        
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
        ByXpath_Swipe_FindText_Click_Android = True
    End If
    
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E") = "" Then
        
        x = MsgBox(sheetname & "中，第" & j & "列第E欄缺少字串", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E").Interior.color = RGB(255, 0, 0)
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
        
    ElseIf TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E").Value) <> "String" Then
    
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E")
        ByXpath_Swipe_FindText_Click_Android = False
        Exit Function
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E").Interior.Pattern = xlNone
        ByXpath_Swipe_FindText_Click_Android = True
    End If
    
End Function



Function checkXpath(sheetname, i, j)
    
    If Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) <> "//*[@" And Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 6) <> "(//*[@" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        checkXpath = False
        Exit Function
    ElseIf Right(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 1) <> "]" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        checkXpath = False
        Exit Function
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E").Interior.Pattern = xlNone
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Font.color = RGB(0, 0, 0)
        checkXpath = True
    End If
    
End Function

Function checkSwipeXpath(sheetname, i, j)
    If Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) <> "//*[@" Or Right(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 1) <> "]" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        checkSwipeXpath = False
        Exit Function
    ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C"), 5) <> "//*[@" Or Right(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C"), 1) <> "]" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
        checkSwipeXpath = False
        Exit Function
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Font.color = RGB(0, 0, 0)
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Font.color = RGB(0, 0, 0)
        checkSwipeXpath = True
    End If
End Function

Function checkClick(sheetname, i, j, status)
    
    If status = "xpath" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkClick = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkClick = True
        End If
        
    Else
    
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkClick = False
            Exit Function
        ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) = "//*[@" Then
            x = MsgBox(sheetname & "中，第" & j & "使用Byid定位，卻輸入Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkClick = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkClick = True
        End If
        
    End If

End Function

Function checkSwipe(sheetname, i, j, status)
    
    If status = "xpath" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkSwipe = False
            Exit Function
        ElseIf Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
            checkSwipe = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
            checkSwipe = True
        End If
        
    Else
    
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkSwipe = False
            Exit Function
            
        ElseIf Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
            checkSwipe = False
            Exit Function
            
        ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) = "//*[@" Then
            x = MsgBox(sheetname & "中，第" & j & "使用Byid_Swipe，卻輸入Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkSwipe = False
            Exit Function
            
        ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C"), 5) = "//*[@" Then
            x = MsgBox(sheetname & "中，第" & j & "使用Byid_Swipe，卻輸入Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
            checkSwipe = False
            Exit Function
            
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
            checkSwipe = True
        End If
        
    End If

End Function
Function checkSwipeVertical_HorHorizontal(sheetname, i, j, state)

    If state = "Vertical" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
        
            x = MsgBox(sheetname & "中，第" & j & "缺少UP/DOWN", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
            checkSwipeVertical_HorHorizontal = False
            Exit Function
            
        ElseIf Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "UP" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "DOWN" Then
            
            x = MsgBox(sheetname & "中，第" & j & "Vertical方法只包含UP與DOWN", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
            checkSwipeVertical_HorHorizontal = False
            Exit Function
            
        Else
            
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
            checkSwipeVertical_HorHorizontal = True
        End If
        
    
    Else
         If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
        
            x = MsgBox(sheetname & "中，第" & j & "缺少RIGHT/LEFT", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
            checkSwipeVertical_HorHorizontal = False
            Exit Function
            
        ElseIf Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "RIGHT" And Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") <> "LEFT" Then
            
            x = MsgBox(sheetname & "中，第" & j & "Vertical方法只包含RIGHT與LEFT", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
            checkSwipeVertical_HorHorizontal = False
            Exit Function
            
        Else
            
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
            checkSwipeVertical_HorHorizontal = True
        End If
    
    End If
    
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D") = "" Then
        x = MsgBox(sheetname & "中，第" & j & "行缺少滑動次數", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.color = RGB(255, 0, 0)
        checkSwipeVertical_HorHorizontal = False
        Exit Function
        
    ElseIf IsNumeric(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D")) = False Then
        x = MsgBox(sheetname & "中，第" & j & "行請輸入數值", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.color = RGB(255, 0, 0)
        checkSwipeVertical_HorHorizontal = False
        Exit Function
        
    Else
        If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Value) <> "String" Then
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D")
        End If
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.Pattern = xlNone
        checkSwipeVertical_HorHorizontal = True
    End If
    
    
End Function

Function checkSwipeData(sheetname, i, j)
    
    '起始X座標
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
        x = MsgBox(sheetname & "中，第" & j & "行缺少起始X座標", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        checkSwipeData = False
        Exit Function
        
    ElseIf IsNumeric(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B")) = False Then
        x = MsgBox(sheetname & "中，第" & j & "行請輸入數值", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
        checkSwipeData = False
        Exit Function
        
    Else
        If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Value) <> "String" Then
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B")
        End If
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
        checkSwipeData = True
    End If
    
    '起始Y座標
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
        x = MsgBox(sheetname & "中，第" & j & "行缺少起始Y座標", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
        checkSwipeData = False
        Exit Function
        
    ElseIf IsNumeric(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C")) = False Then
        x = MsgBox(sheetname & "中，第" & j & "行請輸入數值", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
        checkSwipeData = False
        Exit Function
        
    Else
        If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Value) <> "String" Then
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C")
        End If
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
        checkSwipeData = True
    End If
    
    '結束X座標
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D") = "" Then
        x = MsgBox(sheetname & "中，第" & j & "行缺少結束X座標", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.color = RGB(255, 0, 0)
        checkSwipeData = False
        Exit Function
        
    ElseIf IsNumeric(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D")) = False Then
        x = MsgBox(sheetname & "中，第" & j & "行請輸入數值", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.color = RGB(255, 0, 0)
        checkSwipeData = False
        Exit Function
        
    Else
        If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Value) <> "String" Then
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D")
        End If
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "D").Interior.Pattern = xlNone
        checkSwipeData = True
    End If
    
    '結束Y座標
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E") = "" Then
        x = MsgBox(sheetname & "中，第" & j & "行缺少結束Y座標", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E").Interior.color = RGB(255, 0, 0)
        checkSwipeData = False
        Exit Function
        
    ElseIf IsNumeric(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E")) = False Then
        x = MsgBox(sheetname & "中，第" & j & "行請輸入數值", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E").Interior.color = RGB(255, 0, 0)
        checkSwipeData = False
        Exit Function
        
    Else
        If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E").Value) <> "String" Then
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E")
        End If
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E").Interior.Pattern = xlNone
        checkSwipeData = True
    End If
    
    '滑動次數
    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F") = "" Then
        x = MsgBox(sheetname & "中，第" & j & "行缺少滑動次數", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F").Interior.color = RGB(255, 0, 0)
        checkSwipeData = False
        Exit Function
        
    ElseIf IsNumeric(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F")) = False Then
        x = MsgBox(sheetname & "中，第" & j & "行請輸入數值", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F").Interior.color = RGB(255, 0, 0)
        checkSwipeData = False
        Exit Function
        
    Else
        If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F").Value) <> "String" Then
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F")
        End If
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "F").Interior.Pattern = xlNone
        checkSwipeData = True
    End If

End Function
Function checkWait(sheetname, i, j, status)
    
    If status = "xpath" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkWait = False
            Exit Function
            
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkWait = True
        End If
        
    Else
    
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkWait = False
            Exit Function
            
        ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) = "//*[@" Then
            x = MsgBox(sheetname & "中，第" & j & "使用Byid_Wait，卻輸入Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkWait = False
            Exit Function
            
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkWait = True
        End If

        
    End If

End Function

Function checkResult(sheetname, i, j, status)
    If status = "xpath" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkResult = False
            Exit Function
            
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkResult = True
        End If
    
    Else
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkResult = False
            Exit Function
            
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkResult = True
        End If
        
        If Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) = "//*[@" Then
            x = MsgBox(sheetname & "中，第" & j & "使用Byid_VerifyText，卻輸入Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkResult = False
            Exit Function
            
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkResult = True
        End If
    
    End If
End Function

Function checkSendKey(sheetname, i, j, status)

    If status = "xpath" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkSendKey = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkSendKey = True
        End If
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少輸入值", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
            checkSendKey = False
            Exit Function
            
        Else
            
            If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Value) <> "String" Then
                Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C")
            End If
        
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
            checkSendKey = True
        End If
    Else
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.color = RGB(255, 0, 0)
            checkSendKey = False
            Exit Function
            
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkSendKey = True
        End If
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少輸入值", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.color = RGB(255, 0, 0)
            checkSendKey = False
            Exit Function
            
        Else
            If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Value) <> "String" Then
                Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C")
            End If
        
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
            checkSendKey = True
        End If
                
    End If
    
End Function
Function Clear_Hidekeyboard_LaunchAPP_QuitAPP()
    Application.ScreenUpdating = False
    Dim sheetname As String
    
    i = 0
    Do
        
        If Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
        
            If ThisWorkbook.Sheets(i + 1).Visible = True Then
                        
                'sheetname = ThisWorkbook.Sheets(i + 1).Name
                'Sheets(sheetname).Select
                ThisWorkbook.Sheets(i + 1).Select
                j = 1
                Do
                    If ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "ScreenShot" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "ResetAPP" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "Power" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "Home" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "Back" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "QuitAPP" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "LaunchAPP" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "HideKeyboard" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "Recent" Then
                        For k = 1 To 5
                            ThisWorkbook.Sheets(i + 1).Cells(j, "B").Select
                            Selection.delete Shift:=xlToLeft
                        Next k
                    End If
                    
                
                    j = j + 1
                Loop Until ThisWorkbook.Sheets(i + 1).Cells(j, "A") = ""
    
            End If
        End If

        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
    Sheets("APP&Device").Select
End Function

Function Delete_All_Blank_Cells()
    Application.ScreenUpdating = False
    Dim sheetname As String
    
    i = 0
    Do
        
        If Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
        
            If ThisWorkbook.Sheets(i + 1).Visible = True Then
                 
                ThisWorkbook.Sheets(i + 1).Select
                j = 1
                Do
                    k = 1
                    Do While ThisWorkbook.Sheets(i + 1).Cells(j, k) <> ""
                        k = k + 1
                    Loop
                       
                    For w = 1 To 10
                        ThisWorkbook.Sheets(i + 1).Cells(j, k).Select
                        Selection.delete Shift:=xlToLeft
                    Next w
      
                j = j + 1
                Loop Until ThisWorkbook.Sheets(i + 1).Cells(j, "A") = ""
        
            End If
        End If

        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    Sheets("APP&Device").Select
End Function

