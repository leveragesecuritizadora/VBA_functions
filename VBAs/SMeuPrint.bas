Attribute VB_Name = "SMeuPrint"

Sub MeuPrint(ParamArray args() As Variant)
    Dim i As Long
    Dim resultado As String

    For i = LBound(args) To UBound(args)
        If Not IsEmpty(args(i)) And Not IsNull(args(i)) Then
            resultado = resultado & CStr(args(i))
        End If
    Next i

   Debug.Print resultado
End Sub

