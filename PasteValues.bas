Attribute VB_Name = "Paste_Values"
Sub Paste_Values()
    
    Worksheets("Assumption Projection").Range("A1:BL50").copy
    Worksheets("Assumption Projection Editor").Range("A1:BL50").PasteSpecial Paste:=xlPasteValues

End Sub
