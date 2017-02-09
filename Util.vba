'Module Name : Util
Public result() As Variant
Public lastResult() As Variant
Function CombRow(iSht As Worksheet, rng As Range, N As Single)
    
    tmpRng = Application.Transpose(rng.value)
    
    Dim i As Single
    
    ReDim result(N - 1, 0)
    Call Recursive(tmpRng, N, 1, 0)
    
    ReDim Preserve result(UBound(result, 1), UBound(result, 2) - 1)
    
    ReDim lastResult(UBound(result, 2))
    
    For i = 0 To UBound(result, 2)
        lastResult(i) = result(0, i) & "_" & result(1, i)
        'Debug.Print lastResult(i)
    Next i
    
    CombRow = lastResult
    
End Function

Function Recursive(r As Variant, c As Single, d As Single, e As Single)
    
    Dim f As Single
    
    For f = d To UBound(r, 1)
        result(e, UBound(result, 2)) = r(f, 1)
        
        If e = (c - 1) Then
            ReDim Preserve result(UBound(result, 1), UBound(result, 2) + 1)
            
            For g = 0 To UBound(result, 1)
                result(g, UBound(result, 2)) = result(g, UBound(result, 2) - 1)
            Next g
        Else
            Call Recursive(r, c, f + 1, e + 1)
        End If
    Next f

End Function

Public Function Contains(key As Variant, col As Collection) As Boolean
Dim obj As Variant
On Error GoTo err
    Contains = True
    obj = col(key)
    Exit Function
err:
    Contains = False
    err.Clear
End Function

Public Function IsArrayAllocated(arr As Variant) As Boolean
Dim N As Long
On Error Resume Next

If IsArray(arr) = False Then
    IsArrayAllocated = False
    Exit Function
End If

N = UBound(arr, 1)
If (err.Number = 0) Then
    If LBound(arr) <= UBound(arr) Then
        IsArrayAllocated = True
    Else
        IsArrayAllocated = False
    End If
Else
    IsArrayAllocated = False
End If

End Function
