Attribute VB_Name = "Module1"
Public Function SUMByLabel(Categories As Range, Label As String, SumRange As Range)
    
    ' Define n�mero de elementos na categoria
    Dim CategoriesLength As Integer
    CategoriesLength = Categories.Count
    
    For i = 1 To CategoriesLength
        If Categories.Cells(i).Value = Label Then
        ' Se valor do r�tulo da categoria � igual a valor pr�-especificado para categoria
            ' Adiciona valor ao quadrado da c�lula a contagem da soma de quadrados
            SUMByLabel = SUMByLabel + SumRange.Cells(i).Value
        End If
    Next
    
End Function

Public Function SUMSQByLabel(Categories As Range, Label As String, SumRange As Range)
    
    ' Define n�mero de elementos na categoria
    Dim CategoriesLength As Integer
    CategoriesLength = Categories.Count
    
    For i = 1 To CategoriesLength
        If Categories.Cells(i).Value = Label Then
        ' Se valor do r�tulo da categoria � igual a valor pr�-especificado para categoria
            ' Adiciona valor ao quadrado da c�lula a contagem da soma de quadrados
            SUMSQByLabel = SUMSQByLabel + SumRange.Cells(i).Value ^ 2
        End If
    Next
    
End Function

Public Function VARByLabel(Categories As Range, Label As String, SumRange As Range)
    
    ' Define n�mero de elementos na categoria
    Dim CategoriesLength As Integer
    CategoriesLength = Categories.Count
    
    ' Inicia m�dia em 0
    Dim MEANByLabel As Double
    MEANByLabel = 0
    
    ' Inicia contagem de valores por categoria em 0
    Dim LabelCount As Integer
    LabelCount = 0
    
    For i = 1 To CategoriesLength
        If Categories.Cells(i).Value = Label Then
        ' Se valor do r�tulo da categoria � igual a valor pr�-especificado para categoria
            ' Adiciona valor da c�lula a contagem da m�dia
            MEANByLabel = MEANByLabel + SumRange.Cells(i).Value
            ' Adiciona valor a contagem de elementos do respectivo r�tulo (equivale a n nas f�rmula de m�dia e vari�ncia)
            LabelCount = LabelCount + 1
        End If
    Next
    
    MEANByLabel = MEANByLabel / LabelCount ' Divide soma dos valores por n�mero de elementos, calculando a m�dia
    
     For i = 1 To CategoriesLength
        If Categories.Cells(i).Value = Label Then
            ' Soma desvios em rela��o a m�dia ao quadrado ao valor da vari�ncia
            VARByLabel = VARByLabel + (SumRange.Cells(i).Value - MEANByLabel) ^ 2
        End If
    Next
    
    VARByLabel = VARByLabel / (LabelCount - 1) ' Divide a soma dos valores dos desvios ao quadrado pelo n�mero de elementos - 1, calcuando a vari�ncia
    
End Function

Public Function MEANByLabel(Categories As Range, Label As String, SumRange As Range)

' Define n�mero de elementos na categoria
Dim CategoriesLength As Integer
CategoriesLength = Categories.Count

' Inicia contagem de valores por categoria em 0
Dim LabelCount As Integer
LabelCount = 0

For i = 1 To CategoriesLength
    If Categories.Cells(i).Value = Label Then
    ' Se valor do r�tulo da categoria � igual a valor pr�-especificado para categoria
        ' Adiciona valor da c�lula a contagem da m�dia
        MEANByLabel = MEANByLabel + SumRange.Cells(i).Value
        ' Adiciona valor a contagem de elementos do respectivo r�tulo (equivale a n nas f�rmula de m�dia e vari�ncia)
        LabelCount = LabelCount + 1
    End If
Next
    
MEANByLabel = MEANByLabel / LabelCount ' Divide soma dos valores por n�mero de elementos, calculando a m�dia
    

End Function

Public Function COUNTByLabel(Categories As Range, Label As String)
    
' Define n�mero de elementos na categoria
Dim CategoriesLength As Integer
CategoriesLength = Categories.Count

For i = 1 To CategoriesLength
    If Categories.Cells(i).Value = Label Then
    ' Se valor do r�tulo da categoria � igual a valor pr�-especificado para categoria
        ' Adiciona valor a contagem de elementos do respectivo r�tulo (equivale a n nas f�rmula de m�dia e vari�ncia)
        COUNTByLabel = COUNTByLabel + 1
    End If
Next
    
End Function
