Attribute VB_Name = "CreateLoop"
Public Sub CreateLoop()

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim AdoptRate As Integer
    Dim AvrCredDes As Double
    Dim CMFactor As Double
    Dim CostperCred As Double
    Dim Rev As Double
    Dim StartDate As Date
    'Dim rng As Range
    
    'Set rng = Worksheets("Sheet6").Range("Q1:BL1")
    Dim DirArray As Variant
    DirArray = Worksheets("Sheet6").Range("Q1:BL10").Value
    
    'Debug.Print (Cells(1, 17))
    
    AvrCredDes = Worksheets("Asumptions").Range("K30").Value
    Debug.Print (AvrCredDes)
    
    CMFactor = Worksheets("Asumptions").Range("K34").Value
    Debug.Print (CMFactor)
    
    CostperCred = Worksheets("Asumptions").Range("K31").Value
    Debug.Print (CostperCred)
    
    Rev = (AvrCredDes * CostperCred) + (AvrCredDes * CMFactor * (CostperCred / 3))
    Debug.Print (Rev)
    
    AdoptRate = Worksheets("Asumptions").Range("K28").Value
    Debug.Print (AdoptRate)
    
    ' Loop Through Sheet and fill in Revenue
    
    For i = 2 To 10
        
        For j = 47 To i Step -1

                For k = AdoptRate To 1 Step -1
                    
                    'Debug.Print (k)
       
                    Worksheets("Sheet6").Cells(i, (AdoptRate * j) + k).Value = Rev
                    
                Next
        Next
    
    Next

End Sub
