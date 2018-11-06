Attribute VB_Name = "«e¸m§@·~"
Sub loadingDeviceInfo()
Application.ScreenUpdating = False
    Dim n As Integer
    n = 0
    i = 0
    Do
        
        If ThisWorkbook.Sheets(i + 1).Visible = True Then
            
            n = n + 1
        End If
    
        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
    
    If n > 1 Then
        i = 0
        Do
            If ThisWorkbook.Sheets(i + 1).Visible = True And ThisWorkbook.Sheets(i + 1).Name <> "APP&Device" Then
                ThisWorkbook.Sheets(i + 1).Select
                Exit Do
            End If
        
            i = i + 1
        Loop Until i = ThisWorkbook.Sheets.Count
        Sheets("APP&Device").Select
        
    End If

End Sub
