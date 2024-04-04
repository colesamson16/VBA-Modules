Attribute VB_Name = "ClearSheet"
Public Sub ClearSheet()
    
    Dim Customers As Integer

    Customers = Worksheets("Asumptions").Range("L9").Value
    Debug.Print (Customers)
    
    Worksheets("Assumption Projection").Range("Q2:BL" & Customers + 1).ClearContents

End Sub

