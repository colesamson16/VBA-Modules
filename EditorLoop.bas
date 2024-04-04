Attribute VB_Name = "EditorLoop"
Public Sub EditorLoop()

        Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim AdoptRate As Integer
    Dim AvrCredDes As Double
    Dim CMFactor As Double
    Dim CostperCred As Double
    Dim Rev As Double
    Dim StartDate As Date
    Dim Customers As Integer
    Dim ShName As Worksheet
    
    Dim ActiveCell As CellFormat
    

    
    AvrCredDes = Worksheets("Asumptions").Range("L4").Value
    Debug.Print (AvrCredDes)
    
    CMFactor = Worksheets("Asumptions").Range("L5").Value
    Debug.Print (CMFactor)
    
    CostperCred = Worksheets("Asumptions").Range("L6").Value
    Debug.Print (CostperCred)
    
    Rev = (AvrCredDes * CostperCred) + (AvrCredDes * CMFactor * (CostperCred / 3))
    Debug.Print (Rev)
    
    AdoptRate = Worksheets("Asumptions").Range("L7").Value
    Debug.Print (AdoptRate)
    
    Customers = Worksheets("Asumptions").Range("L9").Value
    Debug.Print (Customers)
    
    ShName = Worksheets("Asumptions")
    
    k = ActiveCell.Column
    Debug.Print (k)
    'StartDate = 17
    Do While k < 64
    
        For i = 2 To (Customers + 1)
        
            For j = 64 To k Step -1

                ActiveCell = Cells(i, j)
                Worksheets("Assumption Projection Editor").Cells(i, j) = Rev
                ActiveCell.FormulaR1C1 = "=((ShName & " '!$L$4')*(ShName & '!$L$6'))+((ShName & '!$L$4')*(ShName & '!$L$5')*(ShName & '!$L$6'/3)"
                 
            Next
            
            k = k + AdoptRate
            
        Next
        
    Loop

End Sub
