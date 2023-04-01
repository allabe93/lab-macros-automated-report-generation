Attribute VB_Name = "Módulo1"
Sub lab_macros_automated_report_generation()
Attribute lab_macros_automated_report_generation.VB_ProcData.VB_Invoke_Func = " \n14"
'
' lab_macros_automated_report_generation Macro
'

'
    'Removing the headers
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    
    'Removing hyphens from the first column and aligning to the left
    Range("A:A").Replace What:="-", Replacement:=""
    Range("A:A").HorizontalAlignment = xlLeft
    
    'Replacing B column´s content with the constant value "EE_DDUCT"
    Range("B1:B4").Value = "EE_DDUCT"
    
    'Passing the first three letters from AG column to C column and changing font in destination
    Dim i As Long
    For i = 1 To 4
        Cells(i, 3).Value = Left(Cells(i, 33).Value, 3)
    Next i
    
    Range("C1:C4").Font.Name = "Calibri"
    
    'Passing the numerical value from the column AK to the column D followed by six vertical bars
    For i = 1 To 4
        Cells(i, 4).Value = Cells(i, 37).Value & "||||||"
    Next i
    
    'Deleting the remaining unnecessary data
    Range("E1:AR4").ClearContents
    
End Sub
