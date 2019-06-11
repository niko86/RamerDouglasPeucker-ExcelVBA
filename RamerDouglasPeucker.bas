Sub CallDouglasPeucker()

Dim pointList As Variant
Dim rowCount As Integer
Dim result As Variant
Dim epsilon As Double
    
    Call ClearOutput
    
    pointList = Sheets("Sheet1").Range("Table1")
    
    rowCount = UBound(pointList)
    
    epsilon = Sheets("Sheet1").Range("epsilon")
    
    result = DouglasPeucker(pointList, epsilon, rowCount)
    
    WriteResult (result)

End Sub

Function DouglasPeucker(pointList As Variant, epsilon As Double, rowCount As Integer) As Variant

    Dim dMax As Double
    Dim Index As Integer
    Dim d As Double
    Dim arrResults1 As Variant
    Dim arrResults2 As Variant
    Dim resultList As Variant
    Dim recResults1 As Variant
    Dim recResults2 As Variant
    
    ' Find the point with the maximum distance
    dMax = 0
    Index = 0
    
    For i = 2 To (rowCount)
        d = Abs((pointList(rowCount, 1) - pointList(1, 1)) * (pointList(1, 2) - pointList(i, 2)) - (pointList(1, 1) - pointList(i, 1)) * (pointList(rowCount, 2) - pointList(1, 2))) / Sqr((pointList(rowCount, 1) - pointList(1, 1)) ^ 2 + (pointList(rowCount, 2) - pointList(1, 2)) ^ 2)
        If d > dMax Then
            Index = i
            dMax = d
        End If
    Next i
    
    'Testing if can stop cut going to index 0
    If Index > 1 Then
        arrResults1 = Cut_Array(pointList, 1, Index)
        arrResults2 = Cut_Array(pointList, Index, rowCount)
        
        ' If max distance is greater than epsilon, recursively simplify
        If (dMax > epsilon) Then
            ' Recursive call
            recResults1 = DouglasPeucker(arrResults1, epsilon, UBound(arrResults1))
            recResults2 = DouglasPeucker(arrResults2, epsilon, UBound(arrResults2))
    
            ' Build the result list
            resultList = Join_Array(recResults1, recResults2)
        Else
            ReDim resultList(1 To 2, 1 To 2)
            
            resultList(1, 1) = pointList(1, 1)
            resultList(1, 2) = pointList(1, 2)
            resultList(2, 1) = pointList(rowCount, 1)
            resultList(2, 2) = pointList(rowCount, 2)
        End If
    Else
        resultList = pointList
    End If

    ' Return the result
    DouglasPeucker = resultList

End Function

Function Cut_Array(arr As Variant, arrStart As Integer, arrEnd As Integer) As Variant
    
    Dim resultList As Variant
    
    ReDim resultList(1 To (arrEnd - arrStart) + 1, 1 To 2)
    
    For i = arrStart To arrEnd
        For j = 1 To 2
            resultList((i - arrStart) + 1, j) = arr(i, j)
        Next j
    Next i
    
    Cut_Array = resultList
        
End Function

Function Join_Array(arr1 As Variant, arr2 As Variant) As Variant

    Dim resultList As Variant
    Dim arr1Length As Integer
    Dim arr2Length As Integer
    
    arr1Length = UBound(arr1) - 1
    arr2Length = UBound(arr2)
    
    newArrLength = arr1Length + arr2Length
    
    ReDim resultList(1 To newArrLength, 1 To 2)
    
    For i = 1 To arr1Length
        For j = 1 To 2
          resultList(i, j) = arr1(i, j)
        Next j
    Next i
    
    For i = 1 To arr2Length
        For j = 1 To 2
          resultList(i + arr1Length, j) = arr2(i, j)
        Next j
    Next i
    
    Join_Array = resultList
    
End Function

Sub WriteResult(result As Variant)
        
    Dim Table2 As ListObject
    
    Set Table2 = Sheets("Sheet2").ListObjects("Table2")

    'Copy information loop
    For i = 1 To UBound(result)
        For j = 1 To UBound(result, 2)
            Table2.Range.Cells(i + 1, j).Value = result(i, j)
        Next j
    Next i
    
End Sub

Sub ClearOutput()

    Dim OutputTable As ListObject
    Dim StartRow As Integer
    
    Set OutputTable = Sheets("Sheet2").ListObjects("Table2")
    
    StartRow = OutputTable.Range.Cells(1, 1).Row + 1
    If Not OutputTable.InsertRowRange Is Nothing Then
        'Pass
    Else
        Sheets("Sheet2").Rows(StartRow & ":" & (OutputTable.DataBodyRange.Rows.Count + StartRow)).Delete
    End If
    
End Sub

