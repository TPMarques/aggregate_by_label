Attribute VB_Name = "Module1"
Public Function SUMByLabel(Categories As Range, Label As String, SumRange As Range)
    
    ' Define número de elementos na categoria
    Dim CategoriesLength As Integer
    CategoriesLength = Categories.Count
    
    For i = 1 To CategoriesLength
        If Categories.Cells(i).Value = Label Then
        ' Se valor do rótulo da categoria é igual a valor pré-especificado para categoria
            ' Adiciona valor ao quadrado da célula a contagem da soma de quadrados
            SUMByLabel = SUMByLabel + SumRange.Cells(i).Value
        End If
    Next
    
End Function

Public Function SUMSQByLabel(Categories As Range, Label As String, SumRange As Range)
    
    ' Define número de elementos na categoria
    Dim CategoriesLength As Integer
    CategoriesLength = Categories.Count
    
    For i = 1 To CategoriesLength
        If Categories.Cells(i).Value = Label Then
        ' Se valor do rótulo da categoria é igual a valor pré-especificado para categoria
            ' Adiciona valor ao quadrado da célula a contagem da soma de quadrados
            SUMSQByLabel = SUMSQByLabel + SumRange.Cells(i).Value ^ 2
        End If
    Next
    
End Function

Public Function VARByLabel(Categories As Range, Label As String, SumRange As Range)
    
    ' Define número de elementos na categoria
    Dim CategoriesLength As Integer
    CategoriesLength = Categories.Count
    
    ' Inicia média em 0
    Dim MEANByLabel As Double
    MEANByLabel = 0
    
    ' Inicia contagem de valores por categoria em 0
    Dim LabelCount As Integer
    LabelCount = 0
    
    For i = 1 To CategoriesLength
        If Categories.Cells(i).Value = Label Then
        ' Se valor do rótulo da categoria é igual a valor pré-especificado para categoria
            ' Adiciona valor da célula a contagem da média
            MEANByLabel = MEANByLabel + SumRange.Cells(i).Value
            ' Adiciona valor a contagem de elementos do respectivo rótulo (equivale a n nas fórmula de média e variância)
            LabelCount = LabelCount + 1
        End If
    Next
    
    MEANByLabel = MEANByLabel / LabelCount ' Divide soma dos valores por número de elementos, calculando a média
    
     For i = 1 To CategoriesLength
        If Categories.Cells(i).Value = Label Then
            ' Soma desvios em relação a média ao quadrado ao valor da variância
            VARByLabel = VARByLabel + (SumRange.Cells(i).Value - MEANByLabel) ^ 2
        End If
    Next
    
    VARByLabel = VARByLabel / (LabelCount - 1) ' Divide a soma dos valores dos desvios ao quadrado pelo número de elementos - 1, calcuando a variância
    
End Function

Public Function MEANByLabel(Categories As Range, Label As String, SumRange As Range)

' Define número de elementos na categoria
Dim CategoriesLength As Integer
CategoriesLength = Categories.Count

' Inicia contagem de valores por categoria em 0
Dim LabelCount As Integer
LabelCount = 0

For i = 1 To CategoriesLength
    If Categories.Cells(i).Value = Label Then
    ' Se valor do rótulo da categoria é igual a valor pré-especificado para categoria
        ' Adiciona valor da célula a contagem da média
        MEANByLabel = MEANByLabel + SumRange.Cells(i).Value
        ' Adiciona valor a contagem de elementos do respectivo rótulo (equivale a n nas fórmula de média e variância)
        LabelCount = LabelCount + 1
    End If
Next
    
MEANByLabel = MEANByLabel / LabelCount ' Divide soma dos valores por número de elementos, calculando a média
    

End Function

Public Function COUNTByLabel(Categories As Range, Label As String)
    
' Define número de elementos na categoria
Dim CategoriesLength As Integer
CategoriesLength = Categories.Count

For i = 1 To CategoriesLength
    If Categories.Cells(i).Value = Label Then
    ' Se valor do rótulo da categoria é igual a valor pré-especificado para categoria
        ' Adiciona valor a contagem de elementos do respectivo rótulo (equivale a n nas fórmula de média e variância)
        COUNTByLabel = COUNTByLabel + 1
    End If
Next
    
End Function
