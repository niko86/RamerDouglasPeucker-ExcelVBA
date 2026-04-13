Attribute VB_Name = "RamerDouglasPeucker"
Option Explicit

' ===========================================================================
' Ramer-Douglas-Peucker polyline simplification for Excel VBA
'
' Two entry points:
'   CallDouglasPeucker      — run with a fixed Epsilon value (cell E1, Sheet1)
'   CallDouglasPeuckerGoalSeek — binary-search for Epsilon to hit a target
'                                  point count (cell E2, Sheet1)
'
' Toolbar (Add-ins tab) is created/removed via CreateRDPToolbar /
' DeleteRDPToolbar, which are called from ThisWorkbook Workbook_Open and
' Workbook_BeforeClose.
' ===========================================================================


' ---------------------------------------------------------------------------
' Entry point: standard run with fixed epsilon
' ---------------------------------------------------------------------------
Sub CallDouglasPeucker()

    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "RDP: Reading input..."

    Dim pointList As Variant
    Dim rowCount As Long
    Dim result As Variant
    Dim epsilon As Double

    Call ClearOutput

    pointList = Sheets("Sheet1").Range("Table1")
    rowCount  = UBound(pointList, 1)
    epsilon   = CDbl(Sheets("Sheet1").Range("epsilon").Value)

    If IsEmpty(pointList) Or rowCount < 2 Then
        MsgBox "Table1 must contain at least 2 data rows.", vbExclamation
        GoTo CleanUp
    End If
    If epsilon <= 0 Then
        MsgBox "The named range 'epsilon' must be a positive number.", vbExclamation
        GoTo CleanUp
    End If

    Application.StatusBar = "RDP: Running simplification (epsilon = " & epsilon & ")..."

    If rowCount > 5000 Then
        result = DouglasPeuckerIterative(pointList, epsilon, rowCount)
    Else
        result = DouglasPeucker(pointList, epsilon, rowCount)
    End If

    Application.StatusBar = "RDP: Writing results..."
    WriteResult result

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    End If

End Sub


' ---------------------------------------------------------------------------
' Entry point: goal-seek — binary-search for epsilon to reach a target count
' ---------------------------------------------------------------------------
Sub CallDouglasPeuckerGoalSeek()

    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "RDP: Reading input..."

    Dim pointList As Variant
    Dim rowCount As Long
    Dim result As Variant
    Dim targetCount As Long

    Call ClearOutput

    pointList   = Sheets("Sheet1").Range("Table1")
    rowCount    = UBound(pointList, 1)
    targetCount = CLng(Sheets("Sheet1").Range("rdpTargetCount").Value)

    If IsEmpty(pointList) Or rowCount < 2 Then
        MsgBox "Table1 must contain at least 2 data rows.", vbExclamation
        GoTo CleanUp
    End If
    If targetCount < 2 Then
        MsgBox "Target count must be at least 2 (start and end points are always kept).", vbExclamation
        GoTo CleanUp
    End If
    If targetCount >= rowCount Then
        MsgBox "Target count (" & targetCount & ") must be less than the number of input points (" & rowCount & ").", vbExclamation
        GoTo CleanUp
    End If

    Application.StatusBar = "RDP: Estimating epsilon for " & targetCount & " points..."

    Dim bestEpsilon As Double
    bestEpsilon = EstimateEpsilonForTargetCount(pointList, rowCount, targetCount)

    ' Write found epsilon back to the sheet so the user can see and reuse it
    Sheets("Sheet1").Range("epsilon").Value = bestEpsilon

    Application.StatusBar = "RDP: Running with epsilon = " & Format(bestEpsilon, "0.0000") & "..."
    result = DouglasPeuckerIterative(pointList, bestEpsilon, rowCount)

    Application.StatusBar = "RDP: Writing results..."
    WriteResult result

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    End If

End Sub


' ---------------------------------------------------------------------------
' Recursive RDP — used when rowCount <= 5000
' Optimisations vs original:
'   - dx/dy/lineLenSq hoisted outside the inner loop (one computation per call)
'   - squared cross-product used throughout — zero Sqr() calls in the hot path
'   - all index/loop variables are Long (not Integer, which overflows at 32 767)
' ---------------------------------------------------------------------------
Function DouglasPeucker(pointList As Variant, epsilon As Double, rowCount As Long) As Variant

    Dim dMaxSq As Double
    Dim Index As Long
    Dim crossSq As Double
    Dim arrResults1 As Variant
    Dim arrResults2 As Variant
    Dim resultList As Variant
    Dim recResults1 As Variant
    Dim recResults2 As Variant
    Dim i As Long

    ' Cache line-vector components once — invariant across the inner loop
    Dim dx As Double, dy As Double, lineLenSq As Double
    dx        = pointList(rowCount, 1) - pointList(1, 1)
    dy        = pointList(rowCount, 2) - pointList(1, 2)
    lineLenSq = dx * dx + dy * dy

    ' Degenerate segment: both endpoints are identical
    If lineLenSq = 0 Then
        DouglasPeucker = pointList
        Exit Function
    End If

    ' Find the point with the maximum squared perpendicular distance
    dMaxSq = 0
    Index  = 0

    For i = 2 To rowCount
        ' Squared cross-product — equivalent to d^2 * lineLenSq, no Sqr needed
        crossSq = (dx * (pointList(1, 2) - pointList(i, 2)) _
                 - dy * (pointList(1, 1) - pointList(i, 1))) ^ 2
        If crossSq > dMaxSq Then
            Index  = i
            dMaxSq = crossSq
        End If
    Next i

    If Index > 1 Then
        arrResults1 = Cut_Array(pointList, 1, Index)
        arrResults2 = Cut_Array(pointList, Index, rowCount)

        ' d > epsilon  ⟺  crossSq > epsilon^2 * lineLenSq  (no Sqr)
        If dMaxSq > (epsilon * epsilon * lineLenSq) Then
            recResults1 = DouglasPeucker(arrResults1, epsilon, UBound(arrResults1, 1))
            recResults2 = DouglasPeucker(arrResults2, epsilon, UBound(arrResults2, 1))
            resultList  = Join_Array(recResults1, recResults2)
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

    DouglasPeucker = resultList

End Function


' ---------------------------------------------------------------------------
' Iterative RDP — explicit stack, safe for any dataset size
' Used automatically for rowCount > 5000, and always used by goal-seek.
' Produces identical output to the recursive version.
' ---------------------------------------------------------------------------
Function DouglasPeuckerIterative(pointList As Variant, epsilon As Double, rowCount As Long) As Variant

    ' keepFlag marks which input points survive into the output
    Dim keepFlag() As Boolean
    ReDim keepFlag(1 To rowCount)
    keepFlag(1)        = True
    keepFlag(rowCount) = True

    ' Explicit stack: two parallel arrays of (segStart, segEnd) pairs
    Dim stkStart() As Long
    Dim stkEnd()   As Long
    Dim stkTop     As Long
    ReDim stkStart(1 To rowCount)
    ReDim stkEnd(1 To rowCount)
    stkTop         = 1
    stkStart(1)    = 1
    stkEnd(1)      = rowCount

    Dim dx As Double, dy As Double, lineLenSq As Double
    Dim crossSq As Double, dMaxSq As Double
    Dim maxIdx As Long, i As Long
    Dim segStart As Long, segEnd As Long
    Dim epsilonSq As Double
    epsilonSq = epsilon * epsilon

    Do While stkTop > 0
        segStart = stkStart(stkTop)
        segEnd   = stkEnd(stkTop)
        stkTop   = stkTop - 1

        dx        = pointList(segEnd, 1) - pointList(segStart, 1)
        dy        = pointList(segEnd, 2) - pointList(segStart, 2)
        lineLenSq = dx * dx + dy * dy

        dMaxSq = 0
        maxIdx = segStart

        If lineLenSq > 0 Then
            For i = segStart + 1 To segEnd - 1
                crossSq = (dx * (pointList(segStart, 2) - pointList(i, 2)) _
                          - dy * (pointList(segStart, 1) - pointList(i, 1))) ^ 2
                If crossSq > dMaxSq Then
                    dMaxSq = crossSq
                    maxIdx = i
                End If
            Next i
        End If

        If dMaxSq > epsilonSq * lineLenSq Then
            ' Push both sub-segments onto the stack
            stkTop           = stkTop + 1
            stkStart(stkTop) = segStart
            stkEnd(stkTop)   = maxIdx
            stkTop           = stkTop + 1
            stkStart(stkTop) = maxIdx
            stkEnd(stkTop)   = segEnd
        Else
            keepFlag(segStart) = True
            keepFlag(segEnd)   = True
        End If
    Loop

    ' Single pass: collect kept points into a compact output array
    Dim keptCount As Long
    keptCount = 0
    For i = 1 To rowCount
        If keepFlag(i) Then keptCount = keptCount + 1
    Next i

    Dim resultList As Variant
    ReDim resultList(1 To keptCount, 1 To 2)
    Dim r As Long
    r = 0
    For i = 1 To rowCount
        If keepFlag(i) Then
            r = r + 1
            resultList(r, 1) = pointList(i, 1)
            resultList(r, 2) = pointList(i, 2)
        End If
    Next i

    DouglasPeuckerIterative = resultList

End Function


' ---------------------------------------------------------------------------
' Goal-seek helpers
' ---------------------------------------------------------------------------

' Run iterative RDP and return the output point count only (no sheet writes).
' Called inside the binary search loop — must be fast.
Function CountResultPoints(pointList As Variant, epsilon As Double, rowCount As Long) As Long
    Dim result As Variant
    result = DouglasPeuckerIterative(pointList, epsilon, rowCount)
    CountResultPoints = UBound(result, 1)
End Function

' Find the epsilon that produces output closest to targetCount points.
'
' Why no closed-form formula exists:
'   epsilon → output count is a monotonically non-increasing step function
'   with N-2 steps. Each step corresponds to the perpendicular distance at
'   which one interior point transitions from kept to removed. Those distances
'   are computed hierarchically (each recursion level uses a different chord),
'   so there is no algebraic mapping without running the algorithm.
'
' What IS analytically computable: the search bounds.
'   lo = 0        → keeps all N points
'   hi = max perpendicular distance from the baseline chord (one O(N) pass)
'       At epsilon >= hi the result is always exactly 2 points.
'   Binary search over [lo, hi] converges in <= 50 iterations.
Function EstimateEpsilonForTargetCount(pointList As Variant, _
                                       rowCount   As Long, _
                                       targetCount As Long) As Double
    Const MAX_ITER  As Integer = 50
    Const TOLERANCE As Double  = 0.000000001   ' 1e-9

    ' Compute the tight upper bound in one O(N) pass
    Dim dx As Double, dy As Double, lineLenSq As Double
    dx        = pointList(rowCount, 1) - pointList(1, 1)
    dy        = pointList(rowCount, 2) - pointList(1, 2)
    lineLenSq = dx * dx + dy * dy

    If lineLenSq = 0 Then
        EstimateEpsilonForTargetCount = 0
        Exit Function
    End If

    Dim i As Long, crossSq As Double, maxCrossSq As Double
    maxCrossSq = 0
    For i = 2 To rowCount - 1
        crossSq = (dx * (pointList(1, 2) - pointList(i, 2)) _
                 - dy * (pointList(1, 1) - pointList(i, 1))) ^ 2
        If crossSq > maxCrossSq Then maxCrossSq = crossSq
    Next i

    Dim lo As Double, hi As Double
    lo = 0
    hi = Sqr(maxCrossSq / lineLenSq) * 1.01   ' 1% above true max distance

    ' Binary search: track the best candidate seen so far
    Dim mid As Double, cnt As Long
    Dim bestEpsilon As Double, bestDiff As Long
    Dim iter As Integer
    bestDiff    = rowCount   ' worst possible difference
    bestEpsilon = lo

    For iter = 1 To MAX_ITER
        If (hi - lo) < TOLERANCE Then Exit For
        mid = (lo + hi) / 2
        cnt = CountResultPoints(pointList, mid, rowCount)
        If Abs(cnt - targetCount) < bestDiff Then
            bestDiff    = Abs(cnt - targetCount)
            bestEpsilon = mid
        End If
        If cnt > targetCount Then
            lo = mid           ' too many points: increase epsilon
        ElseIf cnt < targetCount Then
            hi = mid           ' too few points: decrease epsilon
        Else
            Exit For           ' exact match
        End If
    Next iter

    EstimateEpsilonForTargetCount = bestEpsilon

End Function


' ---------------------------------------------------------------------------
' Array helpers (used only by the recursive path)
' ---------------------------------------------------------------------------
Function Cut_Array(arr As Variant, arrStart As Long, arrEnd As Long) As Variant

    Dim resultList As Variant
    Dim i As Long, j As Long

    ReDim resultList(1 To (arrEnd - arrStart) + 1, 1 To 2)

    For i = arrStart To arrEnd
        For j = 1 To 2
            resultList((i - arrStart) + 1, j) = arr(i, j)
        Next j
    Next i

    Cut_Array = resultList

End Function

Function Join_Array(arr1 As Variant, arr2 As Variant) As Variant

    Dim resultList  As Variant
    Dim arr1Length  As Long
    Dim arr2Length  As Long
    Dim newArrLength As Long
    Dim i As Long, j As Long

    ' Exclude last point of arr1 — it is the first point of arr2 (shared vertex)
    arr1Length   = UBound(arr1, 1) - 1
    arr2Length   = UBound(arr2, 1)
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


' ---------------------------------------------------------------------------
' Output helpers
' ---------------------------------------------------------------------------
Sub WriteResult(result As Variant)

    On Error GoTo WriteErr

    Dim Table2 As ListObject
    Set Table2 = Sheets("Sheet2").ListObjects("Table2")

    Dim nRows As Long
    nRows = UBound(result, 1)

    ' Single bulk array-to-range assignment: one VBA-Excel boundary crossing
    ' instead of the original 2×N cell-by-cell loop.
    Table2.HeaderRowRange.Offset(1, 0).Resize(nRows, 2).Value = result

    Exit Sub
WriteErr:
    MsgBox "WriteResult failed: " & Err.Description, vbCritical

End Sub

Sub ClearOutput()

    On Error GoTo ClearErr

    Dim OutputTable As ListObject
    Set OutputTable = Sheets("Sheet2").ListObjects("Table2")

    ' DataBodyRange is Nothing when the table has no data rows (header only)
    If Not OutputTable.DataBodyRange Is Nothing Then
        OutputTable.DataBodyRange.Delete
    End If

    Exit Sub
ClearErr:
    MsgBox "ClearOutput failed: " & Err.Description, vbCritical

End Sub


' ===========================================================================
' Toolbar management
' Called from ThisWorkbook:
'   Private Sub Workbook_Open()         → CreateRDPToolbar
'   Private Sub Workbook_BeforeClose()  → DeleteRDPToolbar
'
' The toolbar appears under the Add-ins tab in Excel 2007+.
' No external tools or ribbon XML required.
' ===========================================================================
Private Const TOOLBAR_NAME As String = "RDP Simplify"

Sub CreateRDPToolbar()

    DeleteRDPToolbar   ' remove any stale copy from a previous session

    Dim bar As CommandBar
    Set bar = Application.CommandBars.Add( _
        Name:=TOOLBAR_NAME, _
        Position:=msoBarTop, _
        Temporary:=True)    ' Temporary removes the bar automatically if Excel closes unexpectedly

    ' --- Run Simplification ---
    Dim btnRun As CommandBarButton
    Set btnRun             = bar.Controls.Add(Type:=msoControlButton)
    btnRun.Caption         = "Run Simplification"
    btnRun.OnAction        = "CallDouglasPeucker"
    btnRun.Style           = msoButtonIconAndCaption
    btnRun.FaceId          = 186
    btnRun.TooltipText     = "Simplify Table1 using the Epsilon value in Sheet1 cell E1"

    ' --- Goal Seek ---
    Dim btnGoal As CommandBarButton
    Set btnGoal            = bar.Controls.Add(Type:=msoControlButton)
    btnGoal.Caption        = "Goal Seek"
    btnGoal.OnAction       = "CallDouglasPeuckerGoalSeek"
    btnGoal.Style          = msoButtonIconAndCaption
    btnGoal.FaceId         = 585
    btnGoal.TooltipText    = "Find Epsilon to reach the target point count in Sheet1 cell E2"

    ' --- Clear Output ---
    Dim btnClear As CommandBarButton
    Set btnClear           = bar.Controls.Add(Type:=msoControlButton)
    btnClear.Caption       = "Clear Output"
    btnClear.OnAction      = "ClearOutput"
    btnClear.Style         = msoButtonIconAndCaption
    btnClear.FaceId        = 1786
    btnClear.TooltipText   = "Clear all rows from Table2 on Sheet2"

    bar.Visible = True

End Sub

Sub DeleteRDPToolbar()
    On Error Resume Next
    Application.CommandBars(TOOLBAR_NAME).Delete
    On Error GoTo 0
End Sub
