Function Flatten(arr As Variant) As Variant
    Dim Rows As Long, Cols As Long, Length As Long
    Dim i As Long, r As Long, c As Long
    Dim Results() As Variant
    
    ' 行列の要素数を取得
    Rows = UBound(arr, 1) - LBound(arr, 1) + 1
    Cols = UBound(arr, 2) - LBound(arr, 2) + 1
    
    ' 2次元配列として取得された1次元配列のため、行列の掛け算で1次元配列の要素数を取得
    Length = Rows * Cols
    ' 取得した要素数に合わせて配列を作成
    ReDim Results(Length)
    
    For i = 1 To Length
        ' 行が1なら、列を伸ばす
        If Rows = 1 Then
            r = 1: c = i
        ' 列が1なら、行を伸ばす
        Else
            r = i: c = 1
        End If
        ' 行列で値を取得し、配列に格納
        Results(i - 1) = arr(r, c)
    Next
    
    Flatten = Results
End Function

Function getDimension(arr As Variant) As Long
    Dim Dimension As Long
    
    Dimension = 1
    On Error Resume Next
    Do While UBound(arr, Dimension) >= LBound(arr, Dimension)
        Dimension = Dimension + 1
    Loop
    
    getDimension = Dimension - 1
End Function

Function arrLength(arr As Variant) As Long
    If getDimension(arr) = 1 Then
        arrLength = UBound(arr) - LBound(arr) + 1
    ElseIf getDimension(arr) = 2 Then
        arrLength = UBound(arr, 2) - LBound(arr) + 1
    Else
        arrLength = 0
    End If
End Function

Function Extend(arr1 As Variant, arr2 As Variant) As Variant
    Dim Length As Long
    Dim Results()
    
    ' 配列1と配列2の要素数の合計を取得
    Length = arrLength(arr1) + arrLength(arr2)
    ' 取得した要素数に合わせて合成配列を作成
    If Length = 0 Then
        ' どちらも空配列なら空配列を返す
        ReDim Results(o To -1)
    Else
        ReDim Results(Length - 1)
    End If
    
    ' 配列1の内容を合成配列に代入
    If arrLength(arr1) > 0 Then
        For i = 0 To arrLength(arr1) - 1
            Results(i) = arr1(i)
        Next
    End If
    
    ' 配列2の内容を合成配列に代入
    If arrLength(arr2) > 0 Then
        For i = arrLength(arr1) To Length - 1
            Results(i) = arr2(i - arrLength(arr1))
        Next
    End If
    
    Extend = Results
End Function

Function Append(arr1 As Variant, arr2 As Variant) As Variant

End Function
