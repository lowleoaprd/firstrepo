Attribute VB_Name = "copar"
Sub nomatchyet()
 
 Dim X, Z, i As Long, j As Long, n As Long, Y(), dic As Object, k
 
With Worksheets("Static")
    X = .Range("L2").CurrentRegion.Value
 End With
 
  Set dic = CreateObject("scripting.dictionary")
 dic.comparemode = 1
 With dic
 
    For i = 3 To UBound(X, 1)
      If Len(X(i, 1)) Then
        .Item(X(i, 1)) = Empty
      End If
    Next
 End With
 
 With Worksheets("Table_Source")
    X = .Range("L2").CurrentRegion.Value
 End With
 
 ReDim Y(1 To UBound(X), 1 To UBound(X, 2) + 2)
 With dic
    For i = 3 To UBound(X)
       If Len(X(i, 1)) Then
         If Not .exists(X(i, 1)) Then
            k = k + 1
            For j = 1 To UBound(X, 2)
                Y(k, j) = X(i, j)
            Next
         End If
       End If
    Next
 End With

 With Worksheets("Sheet3")
    .Range("a1").End(xlUp).Offset(1).Resize(k, UBound(Y, 2)) = Y
 End With
End Sub

