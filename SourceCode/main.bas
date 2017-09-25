Attribute VB_Name = "main"
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''Desarrollado:       Juan Antonio Barragán  ''''
''''Email:              jabarragann@unal.edu.co''''
'''''''''''''''''''''''''''''''''''''''''''''''''''


Function searchPattern(testString As String, pattern As String) As Boolean
    
    Dim RegX As Object
    Set RegX = CreateObject("VBScript.RegExp")
    
    With RegX
        .Global = False
        .IgnoreCase = False
        .pattern = pattern
    End With
    
    
    searchPattern = RegX.Test(testString)
End Function

Sub search()
    
    Dim c1 As Range
    Dim testString As String, pattern As String
    
    Set c1 = Range("A7:A7")
    pattern = Range("B3:B3")
    
    Do Until c1.Value = ""
        testString = c1.Value

        If searchPattern(testString, pattern) Then
            c1.Offset(0, 1).Value = "si"
            With c1.Offset(0, 1).Interior
                .Color = 5287936
            End With
        Else
            c1.Offset(0, 1).Value = "no"
            With c1.Offset(0, 1).Interior
                .Color = 255
            End With
        End If
        
        Set c1 = c1.Offset(1, 0)
    Loop
    
End Sub
Sub delete()
    
    With Range("B7:B18")
        .Clear
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
End Sub
