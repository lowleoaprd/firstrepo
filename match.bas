Attribute VB_Name = "Module1"
Option Explicit

Sub CompareIt()
    Dim ar As Variant
    Dim arr As Variant
    Dim Var As Variant
    Dim v()
    Dim i As Long
    Dim n As Long
    Dim j As Long
    Dim str As String
    
    ar = Sheet1.Cells(10, 1).CurrentRegion.Value
    With CreateObject("Scripting.Dictionary")
        .CompareMode = 1
        ReDim v(1 To UBound(ar, 2))
        For i = 2 To UBound(ar, 1)
            For n = 1 To UBound(ar, 2)
                str = str & Chr(2) & ar(i, n)
                v(n) = ar(i, n)
            Next
            .Item(str) = v: str = ""
        Next
        ar = Sheet2.Cells(10, 1).CurrentRegion.Resize(, UBound(v)).Value
        For i = 2 To UBound(ar, 1)
            For n = 1 To UBound(ar, 2)
                str = str & Chr(2) & ar(i, n)
                v(n) = ar(i, n)
            Next
            If .exists(str) Then
                .Item(str) = Empty
            Else
                .Item(str) = v
            End If
            str = ""
        Next
        For Each arr In .keys
            If IsEmpty(.Item(arr)) Then .Remove arr
        Next
        Var = .items: j = .Count
    End With
    With Sheet3.Range("a1").Resize(, UBound(ar, 2))
        .CurrentRegion.ClearContents
        .Value = ar
        If j > 0 Then
            .Offset(1).Resize(j).Value = Application.Transpose(Application.Transpose(Var))
        End If
    End With
End Sub



