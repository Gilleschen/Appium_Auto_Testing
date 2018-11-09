Attribute VB_Name = "¸}¥»½s¿è"
Sub test()

    Dim CaseName As String
    Dim arrayNumber As Integer
    Dim CaseRow As Integer
    Sheets("T1_TestScript").Select
    CaseName = "N3"
    arrayNumber = 0
    maxCol = 0
    
    i = 1
    Do
        
        If Sheets("T1_TestScript").Cells(i, "B") = CaseName Then
            CaseRow = i + 1
            j = i + 1 'row
            Do
                k = 1
                Do
                    x = x + 1
                    k = k + 1 ' col
                Loop Until Sheets("T1_TestScript").Cells(j, k) = ""
                
                If k > maxCol Then maxCol = k
            
                j = j + 1
            Loop Until Sheets("T1_TestScript").Cells(j, "A") = "QuitAPP"
            
            Exit Do
        
        End If
        
        i = i + 1
    Loop Until Sheets("T1_TestScript").Cells(i, "A") = ""
    
    'ReDim OriginalCase(j - i - 2, maxCol - 1 - 1)
    ReDim OriginalCase(x - 1)
    Indexi = 0
    
    Do
        k = 1: Indexj = 0
        Do
            
            OriginalCase(Indexi) = Sheets("T1_TestScript").Cells(CaseRow, k)
            
            k = k + 1: Indexi = Indexi + 1
        Loop Until Sheets("T1_TestScript").Cells(CaseRow, k) = ""
        
        
        CaseRow = CaseRow + 1
    Loop Until Sheets("T1_TestScript").Cells(CaseRow, "A") = "QuitAPP"
    
    Erase OriginalCase 'Clear matrix momery
    
    
End Sub


Sub creatTempSheets()
    Application.ScreenUpdating = False
    i = 0
    Do
    
        If ThisWorkbook.Sheets(i + 1).Name = "EditCase" Then
    
            exist = True
            Exit Do
    
        End If
        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
    If exist = False Then
    
        Sheets.add After:=Sheets(Sheets.Count - 1)
        Sheets(Sheets.Count - 1).Name = "EditCase"
    
    End If
    Sheets("EditCase").Select
    Cells.Select
    Selection.ClearContents

End Sub

Sub CopyCasetoTempSheet()
    
End Sub

