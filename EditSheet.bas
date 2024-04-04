Attribute VB_Name = "EditSheet"
Public Sub EditSheet()
    
    Dim AvrCredDes As Double
    Dim CMFactor As Double
    Dim CostperCred As Double
    Dim Rev As Double
    Dim Customers As Integer
    
    Customers = Worksheets("Asumptions").Range("L9").Value
    Debug.Print (Customers)
    
    
    Rev = (AvrCredDes * CostperCred) + (AvrCredDes * CMFactor * (CostperCred / 3))
    Debug.Print (Rev)
    
    
    For i = 2 To (Customers + 1)
        
        For j = 17 To 64
        
                If Not IsEmpty(Worksheets("Assumption Projection").Cells(i, j).Value) Then
                    AvrCredDes = Worksheets("Assumption Projection").Cells(i, 2).Value
                    'Debug.Print (AvrCredDes)
                    CMFactor = Worksheets("Assumption Projection").Cells(i, 3).Value
                    'Debug.Print (CMFactor)
                    CostperCred = Worksheets("Assumption Projection").Cells(i, 4).Value
                    'Debug.Print (CostperCred)

                    Worksheets("Assumption Projection").Cells(i, j).Value = (AvrCredDes * CostperCred) + (AvrCredDes * CMFactor * (CostperCred / 3))
                    'MsgBox "The cell is not empty!"
                Else
                    '
                End If
            
        Next
    Next
End Sub

