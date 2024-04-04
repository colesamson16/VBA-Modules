Attribute VB_Name = "ClearEditorSheet"
Public Sub ClearEditorSheet()
    
    Dim Customers As Integer

        
    Customers = Worksheets("Asumptions").Range("L9").Value
    Debug.Print (Customers)
    

    
    Worksheets("Assumption Projection Editor").Range("Q2:BL" & Customers + 1).ClearContents

End Sub


