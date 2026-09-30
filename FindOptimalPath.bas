Attribute VB_Name = "modFindOptimalPath"
Option Explicit

'===============================================================================
' FindOptimalPath
'
' Finds an optimal routing from the flight/train data on the sheet named
' SHEET_NAME, maximizing the MINIMUM slack (actual time spent in a city minus
' the required minimum) across the "optimized" connecting cities (MPU, CUN,
' GIG, FCO, AMM, DEL, and PEK on routes 9-12), subject to the total trip time (first departure -> last
' arrival) being UNDER the stated budget. On each run, a prompt asks which of
' twelve supported routings to solve:
'   Route 1: MPU -> CUZ -> CUN -> GIG -> FCO -> AMM -> DEL -> PEK                 (7 legs)
'   Route 2: MPU -> CUZ -> GIG -> CUN -> FCO -> AMM -> DEL -> PEK                 (7 legs)
'   Route 3: GIG -> CUZ -> MPU -> CUZ -> CUN -> FCO -> AMM -> DEL -> PEK          (8 legs)
'   Route 4: CUN -> CUZ -> MPU -> CUZ -> GIG -> FCO -> AMM -> DEL -> PEK          (8 legs)
'   Route 5: PEK -> DEL -> AMM -> FCO -> GIG -> CUN -> CUZ -> MPU                 (7 legs)
'   Route 6: PEK -> DEL -> AMM -> FCO -> CUN -> GIG -> CUZ -> MPU                 (7 legs)
'   Route 7: PEK -> DEL -> AMM -> FCO -> GIG -> CUZ -> MPU -> CUZ -> CUN          (8 legs)
'   Route 8: PEK -> DEL -> AMM -> FCO -> CUN -> CUZ -> MPU -> CUZ -> GIG          (8 legs)
'   Route 9: AMM -> DEL -> PEK -> FCO -> GIG -> CUZ -> MPU -> CUZ -> CUN          (8 legs)
'  Route 10: AMM -> DEL -> PEK -> FCO -> CUN -> CUZ -> MPU -> CUZ -> GIG          (8 legs)
'  Route 11: GIG -> CUZ -> MPU -> CUZ -> CUN -> FCO -> PEK -> DEL -> AMM          (8 legs)
'  Route 12: CUN -> CUZ -> MPU -> CUZ -> GIG -> FCO -> PEK -> DEL -> AMM          (8 legs)
' Routes 5-8 are routes 1-4 run backwards (PEK -> ... -> MPU/CUN/GIG instead of
' the other way around). Routes 9-10 start at AMM and fly out to PEK before
' doubling back west to FCO; routes 11-12 are routes 9-10 run backwards,
' ending at AMM. On routes 9-12 PEK is a connection, with a 3-hour minimum
' that counts toward the objective like any other optimized city.
' On routes 1-8 PEK is only ever the first or last city, never a connection,
' so its minimum never applies there. Routes 3, 4,
' 7, 8, 9, 10, 11 & 12 detour out to MPU and back through
' CUZ, so CUZ is visited (and evaluated as a connection) twice. Route length is
' NOT a fixed constant -- N_LEGS is a module-level variable set by InitRoute
' per route choice, and LegOrigin/LegDest are resized (ReDim) to match, since
' VBA won't allow a fixed array bound to vary. Every other function already
' keyed off N_LEGS, LegOrigin(0), and LegDest(N_LEGS-1) as runtime values, so
' no other logic needed to change to support the longer OR the reversed
' routes -- notably TZCorrectionMinutes (below) is a plain UTCOffset(dest) -
' UTCOffset(origin) subtraction, which comes out negative for routes 5-8
' (since PEK is far ahead of MPU/CUN/GIG) and so correctly ADDS time back
' rather than subtracting it, with no special-casing required.
'
' Because routes 3-12 all start somewhere other than MPU, that starting city
' (GIG, CUN, PEK, or AMM) is only ever LegOrigin(0) -- it never appears as a
' LegDest (a connection), so its minimum-stay requirement is never evaluated
' for that route, exactly as intended ("no minimum needed for the city you
' start in", which includes PEK on routes 5-8 and AMM on routes 9-10). No special-casing needed:
' RequiredMinutes/IsOptimizedCity are only ever queried at LegDest(i), so
' this falls out of the existing design automatically.
'
' CUZ is excluded from the objective (no minimum required there, and its
' connection doesn't count toward the max-min slack), per spec. Every other
' connecting city visited -- MPU, CUN, GIG, FCO, AMM, DEL, PEK -- has a required
' minimum and counts toward the objective.
'
' ALGORITHM
'   The one truly free choice in the whole problem is which flight you start
'   on -- every leg after that is a forced, greedy-optimal choice for whatever
'   slack floor is being targeted. For a candidate floor S, requiring every
'   optimized connection to have actual_layover >= required_minimum + S,
'   Propagate walks the route choosing, at each step, the candidate flight
'   with the EARLIEST ARRIVAL among those whose departure satisfies the
'   floor. This greedy choice is optimal: an earlier arrival can only relax
'   every downstream constraint, never tighten it (standard "earliest
'   arrival" argument for time-expanded networks), so it minimizes total
'   finish time for a fixed start and floor, and feasibility at a given
'   (start, S) is monotonically non-increasing in S.
'
'   BestForStart binary-searches S for ONE start to find that start's own
'   best achievable floor. ComputeAllCandidates runs BestForStart for EVERY
'   possible starting flight and sorts the results by minimum slack
'   descending, then by TOTAL slack (sum across every optimized connection)
'   descending as the tie-break. Rank 1 of that sort is the optimum: its
'   minimum slack equals the true global maximum (proof: the path achieving
'   the true optimum has ALL slacks >= that optimum by definition, so it
'   would appear among the candidates at exactly that value; nothing higher
'   is achievable by any start, or feasibility would have found it). Ranks 2+
'   are reported in the output's Top-10 table, each the best COMPLETE
'   itinerary for a different starting flight -- not an arbitrary partial
'   trade-off -- so a route with slightly lower worst-case slack but much
'   more slack everywhere else is visible for you to judge on its own merits,
'   rather than the macro guessing at an exchange rate between the two.
'
' REQUIRED SHEET LAYOUT
'   A sheet named SHEET_NAME (below) containing, in row 1, headers named
'   exactly: Origin, Destination, Departure Date, Arrival Date, Departure Time,
'   Arrival Time. Column order doesn't matter; the macro finds columns by
'   header name. A "Layover Duration" column is optional -- if present, its
'   text (the mid-flight stopover at a connecting airport, e.g. "1 hr 35 min
'   LIM") is carried through to the output's "Flight Layover" column purely
'   for reference. That is NOT the same thing as the city-to-city gap this
'   macro actually optimizes -- see the "Actual Layover" / "Slack" columns in
'   the output's connections table for that.
'
' OUTPUT
'   A sheet named "Optimal Path - Route N" (N = 1-12, per your choice at the
'   prompt) is (re)created each run with:
'     - a summary banner (feasible/infeasible, total time, budget)
'     - the chosen flights/trains, with a "Source Row" column pointing back to
'       the exact row on SHEET_NAME so you can audit against your raw data
'     - a per-city connection table showing arrival, departure, actual
'       layover, required minimum, and slack, with the binding (tightest)
'       connection highlighted gold and the rest shaded green
'     - a trip timeline (linear, color-coded flight/train vs. layover view)
'     - a Top-10 Alternative Routes table ranking every feasible starting
'       flight by minimum slack then total slack, so you can compare the
'       optimum against near-optimal alternatives that trade a little
'       worst-case slack for a lot more slack elsewhere
'   Your source sheet is never modified.
'
' HOW TO USE
'   Alt+F11 -> Insert -> Module -> paste this file's contents (or File ->
'   Import File... and pick this .bas directly) -> F5 or Run > FindOptimalPath.
'   Adjust SHEET_NAME / minimums / budget constants below if yours differ.
'===============================================================================

' --------------------------- CONFIGURATION ------------------------------------
Private Const SHEET_NAME As String = "Flight Schedule"   ' <-- change if needed

Private Const BUDGET_DAYS As Long = 5
Private Const BUDGET_HOURS As Long = 17
Private Const BUDGET_MINUTES As Long = 28

Private Const MAX_S_CAP As Long = 43200   ' safety cap on the slack search, minutes (30 days)
Private Const HUGE_BUDGET As Long = 2000000000   ' used only to find the true best-case total when infeasible

' Route length varies by route choice (7 legs for routes 1-2 & 5-6, 8 for
' routes 3-4 & 7-12), so this is a runtime variable, not a Const -- InitRoute sets it and
' ReDims LegOrigin/LegDest to match before filling them in.
Private N_LEGS As Long
Private LegOrigin() As String
Private LegDest() As String

' Source rows whose Arrival Date/Time came out BEFORE their departure (once
' time zones are accounted for) and were pushed forward a day -- see
' BuildLeg. Reset each run and reported in the final message box.
Private FixedArrivalRows As String
' -------------------------------------------------------------------------------


Public Sub FindOptimalPath()
    Dim ws As Worksheet
    Dim data As Variant
    Dim lastRow As Long, lastCol As Long, c As Long
    Dim colOrigin As Long, colDest As Long, colDepDate As Long, colArrDate As Long, colDepTime As Long, colArrTime As Long, colLayover As Long
    Dim LegData() As Variant
    Dim i As Long
    Dim budget As Long
    Dim finalChosen() As Long, finalTotal As Double, Sstar As Long
    Dim tmpChosen() As Long, tmpTotal As Double
    Dim candMinSlack() As Double, candTotalSlack() As Double, candTotalTime() As Double
    Dim candChosen() As Variant, candCount As Long
    Dim routeChoice As Long
    Dim resp As String
    Dim routeLabel As String, outSheetName As String

    Dim promptMsg As String
    promptMsg = "Which route do you want to solve? Enter 1-12:" & vbCrLf & vbCrLf & _
                "1)  MPU -> CUZ -> CUN -> GIG -> FCO -> AMM -> DEL -> PEK" & vbCrLf & _
                "2)  MPU -> CUZ -> GIG -> CUN -> FCO -> AMM -> DEL -> PEK" & vbCrLf & _
                "3)  GIG -> CUZ -> MPU -> CUZ -> CUN -> FCO -> AMM -> DEL -> PEK" & vbCrLf & _
                "4)  CUN -> CUZ -> MPU -> CUZ -> GIG -> FCO -> AMM -> DEL -> PEK" & vbCrLf & _
                "5)  PEK -> DEL -> AMM -> FCO -> GIG -> CUN -> CUZ -> MPU" & vbCrLf & _
                "6)  PEK -> DEL -> AMM -> FCO -> CUN -> GIG -> CUZ -> MPU" & vbCrLf & _
                "7)  PEK -> DEL -> AMM -> FCO -> GIG -> CUZ -> MPU -> CUZ -> CUN" & vbCrLf & _
                "8)  PEK -> DEL -> AMM -> FCO -> CUN -> CUZ -> MPU -> CUZ -> GIG" & vbCrLf & _
                "9)  AMM -> DEL -> PEK -> FCO -> GIG -> CUZ -> MPU -> CUZ -> CUN" & vbCrLf & _
                "10) AMM -> DEL -> PEK -> FCO -> CUN -> CUZ -> MPU -> CUZ -> GIG" & vbCrLf & _
                "11) GIG -> CUZ -> MPU -> CUZ -> CUN -> FCO -> PEK -> DEL -> AMM" & vbCrLf & _
                "12) CUN -> CUZ -> MPU -> CUZ -> GIG -> FCO -> PEK -> DEL -> AMM" & vbCrLf & vbCrLf & _
                "(Leave blank or Cancel to stop without running.)"
    Do
        resp = InputBox(promptMsg, "Choose Route")
        If resp = "" Then Exit Sub   ' Cancel, or blank + OK
        If resp = "1" Or resp = "2" Or resp = "3" Or resp = "4" Or _
           resp = "5" Or resp = "6" Or resp = "7" Or resp = "8" Or _
           resp = "9" Or resp = "10" Or resp = "11" Or resp = "12" Then Exit Do
        MsgBox "Please enter a number from 1 to 12.", vbExclamation
    Loop
    routeChoice = CLng(resp)

    Call InitRoute(routeChoice)
    FixedArrivalRows = ""
    ReDim LegData(0 To N_LEGS - 1)
    routeLabel = "Route " & routeChoice & ": " & LegOrigin(0)
    For i = 0 To N_LEGS - 1
        routeLabel = routeLabel & " -> " & LegDest(i)
    Next i
    outSheetName = "Optimal Path - Route " & routeChoice

    budget = BudgetMinutes()

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Could not find a sheet named '" & SHEET_NAME & "'." & vbCrLf & _
               "Update the SHEET_NAME constant at the top of the module.", vbCritical
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then
        MsgBox "'" & SHEET_NAME & "' has no data rows below the header.", vbCritical
        Exit Sub
    End If
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value

    colOrigin = 0: colDest = 0: colDepDate = 0: colArrDate = 0: colDepTime = 0: colArrTime = 0: colLayover = 0
    For c = 1 To UBound(data, 2)
        Select Case Trim(CStr(data(1, c)))
            Case "Origin": colOrigin = c
            Case "Destination": colDest = c
            Case "Departure Date": colDepDate = c
            Case "Arrival Date": colArrDate = c
            Case "Departure Time": colDepTime = c
            Case "Arrival Time": colArrTime = c
            Case "Layover Duration": colLayover = c   ' optional -- mid-flight stopover text, not required
        End Select
    Next c

    If colOrigin = 0 Or colDest = 0 Or colDepDate = 0 Or colArrDate = 0 Or colDepTime = 0 Or colArrTime = 0 Then
        MsgBox "Could not find all required headers (Origin, Destination, Departure Date, " & _
               "Arrival Date, Departure Time, Arrival Time) in row 1 of '" & SHEET_NAME & "'.", vbCritical
        Exit Sub
    End If

    For i = 0 To N_LEGS - 1
        LegData(i) = BuildLeg(data, colOrigin, colDest, colDepDate, colArrDate, colDepTime, colArrTime, colLayover, _
                               UCase(LegOrigin(i)), UCase(LegDest(i)))
        If IsEmpty(LegData(i)) Then
            MsgBox "No rows found for " & LegOrigin(i) & " -> " & LegDest(i) & " on '" & SHEET_NAME & "'. Cannot continue.", vbCritical
            Exit Sub
        End If
    Next i

    ' --- Search every possible starting flight, once, for a consistent answer ---
    Call ComputeAllCandidates(LegData, budget, candMinSlack, candTotalSlack, candTotalTime, candChosen, candCount)

    If candCount = 0 Then
        ' Find the true best-case total, ignoring the budget. If even that
        ' fails, no starting flight can be chained all the way to the end
        ' (e.g. the schedule runs out before the last leg), so there is no
        ' itinerary to show at all.
        If Not FeasibleAtS(LegData, 0, HUGE_BUDGET, tmpChosen, tmpTotal) Then
            ' Remove any results sheet from an earlier run so a stale itinerary isn't left looking current.
            Application.DisplayAlerts = False
            On Error Resume Next
            ThisWorkbook.Sheets(outSheetName).Delete
            On Error GoTo 0
            Application.DisplayAlerts = True
            MsgBox routeLabel & vbCrLf & vbCrLf & _
                   "No complete itinerary exists for this route: no starting flight can be connected " & _
                   "all the way to " & LegDest(N_LEGS - 1) & " while meeting every minimum stay, " & _
                   "no matter how long the trip takes. The schedule on '" & SHEET_NAME & "' likely " & _
                   "runs out of flights before the final leg." & vbCrLf & vbCrLf & _
                   FixedArrivalsNote(), vbExclamation
            Exit Sub
        End If
        Call WriteResults(LegData, tmpChosen, 0, tmpTotal, budget, False, routeLabel, outSheetName)
        MsgBox routeLabel & vbCrLf & vbCrLf & _
               "No itinerary fits within the " & FormatDuration(CDbl(budget)) & " budget." & vbCrLf & _
               "Tightest possible total, time-zone adjusted (zero extra slack beyond the minimums): " & FormatDuration(tmpTotal) & vbCrLf & _
               "That is over budget by " & FormatDuration(tmpTotal - budget) & "." & vbCrLf & vbCrLf & _
               FixedArrivalsNote() & _
               "See the '" & outSheetName & "' sheet for the closest itinerary found.", vbExclamation
        Exit Sub
    End If

    ' Rank 1 of the sorted candidates (max min-slack, tie-broken by max total
    ' slack -- see ComputeAllCandidates) IS the optimal path.
    Sstar = CLng(Round(candMinSlack(1), 0))
    finalTotal = candTotalTime(1)
    finalChosen = candChosen(1)

    Call WriteResults(LegData, finalChosen, Sstar, finalTotal, budget, True, routeLabel, outSheetName)
    Call WriteAlternativesTable(LegData, outSheetName, candMinSlack, candTotalSlack, candTotalTime, candChosen, candCount)
    MsgBox routeLabel & vbCrLf & vbCrLf & _
           "Optimal path found." & vbCrLf & _
           "Guaranteed minimum layover slack: " & Sstar & " min (" & FormatDuration(CDbl(Sstar)) & ")." & vbCrLf & _
           "Total trip time, time-zone adjusted: " & FormatDuration(finalTotal) & "  (budget " & FormatDuration(CDbl(budget)) & ")." & vbCrLf & vbCrLf & _
           "See the '" & outSheetName & "' sheet for full details, including a Top 10 Alternative Routes table " & _
           "further down so you can compare the trade-off between worst-case slack and total slack." & _
           IIf(FixedArrivalRows = "", "", vbCrLf & vbCrLf & FixedArrivalsNote()), vbInformation
End Sub


' routeChoice 1: MPU -> CUZ -> CUN -> GIG -> FCO -> AMM -> DEL -> PEK                 (7 legs)
' routeChoice 2: MPU -> CUZ -> GIG -> CUN -> FCO -> AMM -> DEL -> PEK                 (7 legs)
' routeChoice 3: GIG -> CUZ -> MPU -> CUZ -> CUN -> FCO -> AMM -> DEL -> PEK          (8 legs)
' routeChoice 4: CUN -> CUZ -> MPU -> CUZ -> GIG -> FCO -> AMM -> DEL -> PEK          (8 legs)
' routeChoice 5: PEK -> DEL -> AMM -> FCO -> GIG -> CUN -> CUZ -> MPU                 (7 legs)
' routeChoice 6: PEK -> DEL -> AMM -> FCO -> CUN -> GIG -> CUZ -> MPU                 (7 legs)
' routeChoice 7: PEK -> DEL -> AMM -> FCO -> GIG -> CUZ -> MPU -> CUZ -> CUN          (8 legs)
' routeChoice 8: PEK -> DEL -> AMM -> FCO -> CUN -> CUZ -> MPU -> CUZ -> GIG          (8 legs)
' routeChoice 9: AMM -> DEL -> PEK -> FCO -> GIG -> CUZ -> MPU -> CUZ -> CUN          (8 legs)
' routeChoice 10: AMM -> DEL -> PEK -> FCO -> CUN -> CUZ -> MPU -> CUZ -> GIG         (8 legs)
' routeChoice 11: GIG -> CUZ -> MPU -> CUZ -> CUN -> FCO -> PEK -> DEL -> AMM         (8 legs)
' routeChoice 12: CUN -> CUZ -> MPU -> CUZ -> GIG -> FCO -> PEK -> DEL -> AMM         (8 legs)
' Routes 5-8 are routes 1-4 traversed in reverse. Routes 9-10 are routes 7-8
' started from AMM instead, with PEK visited as a mid-route connection.
' Routes 11-12 are routes 9-10 traversed in reverse. RequiredMinutes/IsOptimizedCity
' key off city name only, and TZCorrectionMinutes/every other function keys off
' N_LEGS, LegOrigin(0) and LegDest(N_LEGS-1) as runtime values, so no other
' change is needed to support any of these orderings, the longer 8-leg routes,
' or the reversed direction.
Private Sub InitRoute(routeChoice As Long)
    Select Case routeChoice
        Case 1, 2, 5, 6
            N_LEGS = 7
        Case 3, 4, 7, 8, 9, 10, 11, 12
            N_LEGS = 8
    End Select
    ReDim LegOrigin(0 To N_LEGS - 1)
    ReDim LegDest(0 To N_LEGS - 1)

    Select Case routeChoice
        Case 1
            LegOrigin(0) = "MPU": LegDest(0) = "CUZ"
            LegOrigin(1) = "CUZ": LegDest(1) = "CUN"
            LegOrigin(2) = "CUN": LegDest(2) = "GIG"
            LegOrigin(3) = "GIG": LegDest(3) = "FCO"
            LegOrigin(4) = "FCO": LegDest(4) = "AMM"
            LegOrigin(5) = "AMM": LegDest(5) = "DEL"
            LegOrigin(6) = "DEL": LegDest(6) = "PEK"
        Case 2
            LegOrigin(0) = "MPU": LegDest(0) = "CUZ"
            LegOrigin(1) = "CUZ": LegDest(1) = "GIG"
            LegOrigin(2) = "GIG": LegDest(2) = "CUN"
            LegOrigin(3) = "CUN": LegDest(3) = "FCO"
            LegOrigin(4) = "FCO": LegDest(4) = "AMM"
            LegOrigin(5) = "AMM": LegDest(5) = "DEL"
            LegOrigin(6) = "DEL": LegDest(6) = "PEK"
        Case 3
            LegOrigin(0) = "GIG": LegDest(0) = "CUZ"
            LegOrigin(1) = "CUZ": LegDest(1) = "MPU"
            LegOrigin(2) = "MPU": LegDest(2) = "CUZ"
            LegOrigin(3) = "CUZ": LegDest(3) = "CUN"
            LegOrigin(4) = "CUN": LegDest(4) = "FCO"
            LegOrigin(5) = "FCO": LegDest(5) = "AMM"
            LegOrigin(6) = "AMM": LegDest(6) = "DEL"
            LegOrigin(7) = "DEL": LegDest(7) = "PEK"
        Case 4
            LegOrigin(0) = "CUN": LegDest(0) = "CUZ"
            LegOrigin(1) = "CUZ": LegDest(1) = "MPU"
            LegOrigin(2) = "MPU": LegDest(2) = "CUZ"
            LegOrigin(3) = "CUZ": LegDest(3) = "GIG"
            LegOrigin(4) = "GIG": LegDest(4) = "FCO"
            LegOrigin(5) = "FCO": LegDest(5) = "AMM"
            LegOrigin(6) = "AMM": LegDest(6) = "DEL"
            LegOrigin(7) = "DEL": LegDest(7) = "PEK"
        Case 5
            LegOrigin(0) = "PEK": LegDest(0) = "DEL"
            LegOrigin(1) = "DEL": LegDest(1) = "AMM"
            LegOrigin(2) = "AMM": LegDest(2) = "FCO"
            LegOrigin(3) = "FCO": LegDest(3) = "GIG"
            LegOrigin(4) = "GIG": LegDest(4) = "CUN"
            LegOrigin(5) = "CUN": LegDest(5) = "CUZ"
            LegOrigin(6) = "CUZ": LegDest(6) = "MPU"
        Case 6
            LegOrigin(0) = "PEK": LegDest(0) = "DEL"
            LegOrigin(1) = "DEL": LegDest(1) = "AMM"
            LegOrigin(2) = "AMM": LegDest(2) = "FCO"
            LegOrigin(3) = "FCO": LegDest(3) = "CUN"
            LegOrigin(4) = "CUN": LegDest(4) = "GIG"
            LegOrigin(5) = "GIG": LegDest(5) = "CUZ"
            LegOrigin(6) = "CUZ": LegDest(6) = "MPU"
        Case 7
            LegOrigin(0) = "PEK": LegDest(0) = "DEL"
            LegOrigin(1) = "DEL": LegDest(1) = "AMM"
            LegOrigin(2) = "AMM": LegDest(2) = "FCO"
            LegOrigin(3) = "FCO": LegDest(3) = "GIG"
            LegOrigin(4) = "GIG": LegDest(4) = "CUZ"
            LegOrigin(5) = "CUZ": LegDest(5) = "MPU"
            LegOrigin(6) = "MPU": LegDest(6) = "CUZ"
            LegOrigin(7) = "CUZ": LegDest(7) = "CUN"
        Case 8
            LegOrigin(0) = "PEK": LegDest(0) = "DEL"
            LegOrigin(1) = "DEL": LegDest(1) = "AMM"
            LegOrigin(2) = "AMM": LegDest(2) = "FCO"
            LegOrigin(3) = "FCO": LegDest(3) = "CUN"
            LegOrigin(4) = "CUN": LegDest(4) = "CUZ"
            LegOrigin(5) = "CUZ": LegDest(5) = "MPU"
            LegOrigin(6) = "MPU": LegDest(6) = "CUZ"
            LegOrigin(7) = "CUZ": LegDest(7) = "GIG"
        Case 9
            LegOrigin(0) = "AMM": LegDest(0) = "DEL"
            LegOrigin(1) = "DEL": LegDest(1) = "PEK"
            LegOrigin(2) = "PEK": LegDest(2) = "FCO"
            LegOrigin(3) = "FCO": LegDest(3) = "GIG"
            LegOrigin(4) = "GIG": LegDest(4) = "CUZ"
            LegOrigin(5) = "CUZ": LegDest(5) = "MPU"
            LegOrigin(6) = "MPU": LegDest(6) = "CUZ"
            LegOrigin(7) = "CUZ": LegDest(7) = "CUN"
        Case 10
            LegOrigin(0) = "AMM": LegDest(0) = "DEL"
            LegOrigin(1) = "DEL": LegDest(1) = "PEK"
            LegOrigin(2) = "PEK": LegDest(2) = "FCO"
            LegOrigin(3) = "FCO": LegDest(3) = "CUN"
            LegOrigin(4) = "CUN": LegDest(4) = "CUZ"
            LegOrigin(5) = "CUZ": LegDest(5) = "MPU"
            LegOrigin(6) = "MPU": LegDest(6) = "CUZ"
            LegOrigin(7) = "CUZ": LegDest(7) = "GIG"
        Case 11
            LegOrigin(0) = "GIG": LegDest(0) = "CUZ"
            LegOrigin(1) = "CUZ": LegDest(1) = "MPU"
            LegOrigin(2) = "MPU": LegDest(2) = "CUZ"
            LegOrigin(3) = "CUZ": LegDest(3) = "CUN"
            LegOrigin(4) = "CUN": LegDest(4) = "FCO"
            LegOrigin(5) = "FCO": LegDest(5) = "PEK"
            LegOrigin(6) = "PEK": LegDest(6) = "DEL"
            LegOrigin(7) = "DEL": LegDest(7) = "AMM"
        Case 12
            LegOrigin(0) = "CUN": LegDest(0) = "CUZ"
            LegOrigin(1) = "CUZ": LegDest(1) = "MPU"
            LegOrigin(2) = "MPU": LegDest(2) = "CUZ"
            LegOrigin(3) = "CUZ": LegDest(3) = "GIG"
            LegOrigin(4) = "GIG": LegDest(4) = "FCO"
            LegOrigin(5) = "FCO": LegDest(5) = "PEK"
            LegOrigin(6) = "PEK": LegDest(6) = "DEL"
            LegOrigin(7) = "DEL": LegDest(7) = "AMM"
    End Select
End Sub


' --------------------------- TIME ZONE DATA -----------------------------------
' Flight/train schedules record LOCAL time at each airport. An arrival and the
' next departure AT THE SAME CITY can be subtracted directly -- the offset is
' identical on both sides and cancels out -- so the connection/layover figures
' elsewhere in this module need no adjustment.
'
' But the OVERALL trip length (first departure -> final arrival) and any single
' leg's flight duration cross time zones, so naive subtraction of two local
' timestamps is off by the difference in UTC offset between the two endpoints.
' This table supplies that offset (minutes from UTC) for every airport in the
' route, valid for November 2026 -- none of these observe DST, so a fixed
' offset is safe for the whole trip regardless of the exact date:
'   MPU/CUZ (Peru)               UTC-5     Peru has not observed DST since 1994
'   CUN (Quintana Roo, Mexico)   UTC-5     "Zona Sureste", fixed since 2015
'   GIG (Rio de Janeiro, Brazil) UTC-3     Brazil abolished DST in 2019
'   FCO (Rome, Italy)            UTC+1     CET; EU summer time ends late Oct
'   AMM (Amman, Jordan)          UTC+3     Jordan went to permanent DST-time Oct 2022
'   DEL (New Delhi, India)       UTC+5:30  India has never observed DST
'   PEK (Beijing, China)         UTC+8     one national time zone, no DST
' If the route ever changes to include a different airport, add it here.
Private Function UTCOffsetMinutes(city As String) As Long
    Select Case UCase(city)
        Case "MPU", "CUZ", "CUN": UTCOffsetMinutes = -5 * 60
        Case "GIG": UTCOffsetMinutes = -3 * 60
        Case "FCO": UTCOffsetMinutes = 1 * 60
        Case "AMM": UTCOffsetMinutes = 3 * 60
        Case "DEL": UTCOffsetMinutes = 5 * 60 + 30
        Case "PEK": UTCOffsetMinutes = 8 * 60
        Case Else: UTCOffsetMinutes = 0   ' unrecognized city -- update the table above
    End Select
End Function


' Net correction to subtract from a naive (local-clock) total-trip-time figure
' to get the TRUE elapsed time: the final destination sits this many minutes
' further "ahead" of UTC than the origin, so local clocks read further forward
' by that much over the course of the trip, on top of time actually elapsed.
Private Function TZCorrectionMinutes() As Long
    TZCorrectionMinutes = UTCOffsetMinutes(LegDest(N_LEGS - 1)) - UTCOffsetMinutes(LegOrigin(0))
End Function
' -------------------------------------------------------------------------------


Private Function RequiredMinutes(city As String) As Long
    Select Case UCase(city)
        Case "MPU": RequiredMinutes = 2 * 60
        Case "CUZ": RequiredMinutes = 60
        Case "CUN": RequiredMinutes = 8 * 60 + 5
        Case "GIG": RequiredMinutes = 3 * 60 + 1
        Case "FCO": RequiredMinutes = 3 * 60 + 0
        Case "AMM": RequiredMinutes = 7 * 60 + 38
        Case "DEL": RequiredMinutes = 9 * 60 + 0
        Case "PEK": RequiredMinutes = 3 * 60 + 0   ' only a connection on routes 9-12
        Case Else: RequiredMinutes = 0   ' any other city: no minimum
    End Select
End Function


' A city here means "this connection counts toward the max-min slack
' objective." CUZ is excluded even though it now has a real 60-minute
' minimum (see RequiredMinutes above): that minimum is still enforced as a
' hard floor on every CUZ connection (Propagate always applies
' RequiredMinutes regardless of this function), it just isn't stretched by
' the slack search the way CUN/GIG/FCO/AMM/DEL/MPU/PEK are. CUZ appears twice in
' routes 3-4 and 7-12 -- both occurrences are handled independently and
' correctly, since this is keyed by city name only. Note that a route's OWN
' starting city (e.g. GIG in route 3, PEK in route 5, AMM in route 9) never gets evaluated
' here in the first place, since RequiredMinutes/IsOptimizedCity are only
' ever queried at LegDest(i) -- a route's starting city is LegOrigin(0) and
' is never also a LegDest unless the route revisits it later (which none of
' these do for their own start city).
Private Function IsOptimizedCity(city As String) As Boolean
    Select Case UCase(city)
        Case "MPU", "CUN", "GIG", "FCO", "AMM", "DEL", "PEK": IsOptimizedCity = True
        Case Else: IsOptimizedCity = False   ' CUZ excluded from the max-min objective
    End Select
End Function


Private Function BudgetMinutes() As Long
    BudgetMinutes = BUDGET_DAYS * 1440 + BUDGET_HOURS * 60 + BUDGET_MINUTES
End Function


Private Function FormatDuration(totalMin As Double) As String
    Dim t As Long, d As Long, h As Long, m As Long
    t = CLng(Round(totalMin, 0))
    d = t \ 1440
    h = (t Mod 1440) \ 60
    m = t Mod 60
    FormatDuration = d & "d " & h & "h " & m & "m"
End Function


' Combines a date cell and a time cell into one Date/time value. Deliberately
' avoids TimeValue(), which requires a String or Date and throws "Type
' mismatch" if the time is stored as a raw Double (day-fraction) instead --
' which is how this workbook's Departure/Arrival Time columns are stored.
' A VBA Date is internally a Double (integer part = day, fractional part =
' time-of-day), so this works uniformly whether the cell holds a true
' Date/Time or a plain numeric fraction.
Private Function TryCombine(dateVal As Variant, timeVal As Variant, ByRef outDT As Date) As Boolean
    On Error GoTo Fail
    Dim dPart As Double, tRaw As Double, tPart As Double

    dPart = Int(CDbl(dateVal))
    tRaw = CDbl(timeVal)
    tPart = tRaw - Int(tRaw)

    outDT = CDate(dPart + tPart)
    TryCombine = True
    Exit Function
Fail:
    TryCombine = False
End Function


' Returns Empty if no rows match; otherwise a Variant holding a
' (1 To n, 1 To 4) array: col1=DepDT, col2=ArrDT, col3=source row number,
' col4=Flight Layover text (mid-flight stopover, e.g. "1 hr 35 min LIM" --
' blank if colLayover=0 (column not present on the sheet) or the cell is
' empty, e.g. a nonstop flight or a train leg), sorted ascending by DepDT.
Private Function BuildLeg(data As Variant, colOrigin As Long, colDest As Long, _
                           colDepDate As Long, colArrDate As Long, colDepTime As Long, colArrTime As Long, _
                           colLayover As Long, originCode As String, destCode As String) As Variant
    Dim i As Long, n As Long, cnt As Long
    Dim depDT As Date, arrDT As Date
    Dim isMatch As Boolean
    Dim layoverVal As String
    Dim tzGap As Long
    Dim result() As Variant, final() As Variant

    n = UBound(data, 1)
    cnt = 0
    tzGap = UTCOffsetMinutes(destCode) - UTCOffsetMinutes(originCode)

    For i = 2 To n   ' row 1 = headers
        isMatch = False
        If Not IsEmpty(data(i, colOrigin)) And Not IsEmpty(data(i, colDest)) Then
            If UCase(Trim(CStr(data(i, colOrigin)))) = originCode And UCase(Trim(CStr(data(i, colDest)))) = destCode Then
                isMatch = True
            End If
        End If

        If isMatch Then
            If TryCombine(data(i, colDepDate), data(i, colDepTime), depDT) And _
               TryCombine(data(i, colArrDate), data(i, colArrTime), arrDT) Then
                ' An arrival that lands at or before its own departure (in
                ' true elapsed time, i.e. after removing the time-zone gap)
                ' is an overnight trip whose Arrival Date was entered as the
                ' departure day. Taken literally it looks like the fastest
                ' option of all, so the earliest-arrival search picks it and
                ' the itinerary appears to run backwards in time. Push it
                ' forward a day at a time until it's after the departure.
                If (arrDT - depDT) * 1440# - tzGap <= 0 Then
                    Do While (arrDT - depDT) * 1440# - tzGap <= 0
                        arrDT = arrDT + 1
                    Loop
                    If FixedArrivalRows <> "" Then FixedArrivalRows = FixedArrivalRows & ", "
                    FixedArrivalRows = FixedArrivalRows & i
                End If
                cnt = cnt + 1
                ReDim Preserve result(1 To 4, 1 To cnt)   ' last dimension only -- VBA's ReDim Preserve rule
                result(1, cnt) = depDT
                result(2, cnt) = arrDT
                result(3, cnt) = i

                layoverVal = ""
                If colLayover > 0 Then
                    If Not IsEmpty(data(i, colLayover)) Then layoverVal = Trim(CStr(data(i, colLayover)))
                End If
                result(4, cnt) = layoverVal
            End If
        End If
    Next i

    If cnt = 0 Then Exit Function   ' returns Empty

    ReDim final(1 To cnt, 1 To 4)
    For i = 1 To cnt
        final(i, 1) = result(1, i)
        final(i, 2) = result(2, i)
        final(i, 3) = result(3, i)
        final(i, 4) = result(4, i)
    Next i

    Call SortLegByDep(final)
    BuildLeg = final
End Function


' Message-box text listing the source rows BuildLeg corrected, or "" if none.
Private Function FixedArrivalsNote() As String
    If FixedArrivalRows = "" Then Exit Function
    FixedArrivalsNote = "Note: these '" & SHEET_NAME & "' rows had an arrival before their departure " & _
                        "and were treated as arriving the next day: row(s) " & FixedArrivalRows & "." & vbCrLf & _
                        "Consider correcting their Arrival Date on the sheet." & vbCrLf & vbCrLf
End Function


' Simple, obviously-correct insertion sort by column 1 (DepDT) ascending.
' n is at most a few hundred rows here, so O(n^2) is instant.
Private Sub SortLegByDep(arr As Variant)
    Dim n As Long, i As Long, j As Long
    Dim keyDep As Date, keyArr As Date, keyRow As Long, keyLayover As String

    n = UBound(arr, 1)
    For i = 2 To n
        keyDep = arr(i, 1): keyArr = arr(i, 2): keyRow = arr(i, 3): keyLayover = arr(i, 4)
        j = i - 1
        Do While j >= 1
            If arr(j, 1) <= keyDep Then Exit Do
            arr(j + 1, 1) = arr(j, 1)
            arr(j + 1, 2) = arr(j, 2)
            arr(j + 1, 3) = arr(j, 3)
            arr(j + 1, 4) = arr(j, 4)
            j = j - 1
        Loop
        arr(j + 1, 1) = keyDep
        arr(j + 1, 2) = keyArr
        arr(j + 1, 3) = keyRow
        arr(j + 1, 4) = keyLayover
    Next i
End Sub


' Among rows in legArr with DepDT >= minDep, returns the row index (1..n)
' with the earliest ArrDT, or 0 if none qualify.
Private Function EarliestAfterIdx(legArr As Variant, minDep As Date) As Long
    Dim n As Long, i As Long
    Dim bestIdx As Long, bestArr As Date

    bestIdx = 0
    n = UBound(legArr, 1)
    For i = 1 To n
        If legArr(i, 1) >= minDep Then
            If bestIdx = 0 Then
                bestIdx = i
                bestArr = legArr(i, 2)
            ElseIf legArr(i, 2) < bestArr Then
                bestIdx = i
                bestArr = legArr(i, 2)
            End If
        End If
    Next i
    EarliestAfterIdx = bestIdx
End Function


' Walks all N_LEGS legs starting from LegData(0)'s row `startIdx`, requiring
' each optimized connection's layover to be >= required + S (CUZ requires only
' a non-negative connection). Returns False if any leg has no qualifying flight.
Private Function Propagate(LegData() As Variant, startIdx As Long, S As Long, _
                            ByRef chosenIdx() As Long, ByRef totalMinutes As Double) As Boolean
    Dim i As Long, idx As Long, req As Long
    Dim arrDT As Date, minDep As Date
    Dim curLeg As Variant, nextLeg As Variant

    ReDim chosenIdx(0 To N_LEGS - 1)
    chosenIdx(0) = startIdx
    curLeg = LegData(0)
    arrDT = curLeg(startIdx, 2)

    For i = 1 To N_LEGS - 1
        req = RequiredMinutes(LegDest(i - 1))
        If IsOptimizedCity(LegDest(i - 1)) Then req = req + S
        minDep = DateAdd("n", req, arrDT)   ' "n" = minutes in DateAdd

        nextLeg = LegData(i)
        idx = EarliestAfterIdx(nextLeg, minDep)
        If idx = 0 Then
            Propagate = False
            Exit Function
        End If
        chosenIdx(i) = idx
        arrDT = nextLeg(idx, 2)
    Next i

    curLeg = LegData(0)
    ' Raw local-clock difference minus the origin/destination time-zone gap
    ' gives the TRUE elapsed time -- see TZCorrectionMinutes above. Since the
    ' correction is a constant (fixed start/end cities), minimizing this
    ' still minimizes the raw difference too, so the earlier-arrival greedy
    ' logic above remains valid unchanged.
    totalMinutes = ((arrDT - curLeg(startIdx, 1)) * 1440#) - TZCorrectionMinutes()
    Propagate = True
End Function


' True if ANY MPU->CUZ start yields a full chain with total < budget under
' floor S. Also returns (via ByRef) the min-total such path found.
Private Function FeasibleAtS(LegData() As Variant, S As Long, budget As Long, _
                              ByRef bestChosen() As Long, ByRef bestTotal As Double) As Boolean
    Dim leg0 As Variant
    Dim n0 As Long, s0 As Long
    Dim chosenIdx() As Long
    Dim total As Double
    Dim found As Boolean

    leg0 = LegData(0)
    n0 = UBound(leg0, 1)

    found = False
    bestTotal = -1
    For s0 = 1 To n0
        If Propagate(LegData, s0, S, chosenIdx, total) Then
            If total < budget Then
                If Not found Then
                    found = True
                    bestTotal = total
                    bestChosen = chosenIdx
                ElseIf total < bestTotal Then
                    bestTotal = total
                    bestChosen = chosenIdx
                End If
            End If
        End If
    Next s0
    FeasibleAtS = found
End Function


' Runs a per-start binary search -- identical logic to the global one in
' FindOptimalPath, just restricted to ONE starting flight -- to find the best
' achievable itinerary FOR THAT START: the maximum floor S such that
' Propagate(startIdx, S) is feasible and under budget. By the same argument
' that makes the global search correct (see the header comment block), that
' itinerary's true minimum slack among optimized connections equals exactly
' that S. This also sums every optimized connection's slack into a TOTAL,
' which the global search never needed but comparing alternative routes does
' -- two routes can share the same worst-case slack while differing a lot on
' how generous the OTHER connections are.
Private Function BestForStart(LegData() As Variant, startIdx As Long, budget As Long, _
                               ByRef outChosen() As Long, ByRef outMinSlack As Double, _
                               ByRef outTotalSlack As Double, ByRef outTotalTime As Double) As Boolean
    Dim lo As Long, hi As Long, mid As Long
    Dim chosenIdx() As Long, total As Double
    Dim bestChosen() As Long, bestTotal As Double

    If Not Propagate(LegData, startIdx, 0, chosenIdx, total) Then
        BestForStart = False
        Exit Function
    End If
    If Not (total < budget) Then
        BestForStart = False
        Exit Function
    End If
    bestChosen = chosenIdx
    bestTotal = total

    lo = 0
    hi = 1
    Do While hi <= MAX_S_CAP
        If Propagate(LegData, startIdx, hi, chosenIdx, total) Then
            If total < budget Then
                lo = hi
                bestChosen = chosenIdx
                bestTotal = total
                hi = hi * 2
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    If hi > MAX_S_CAP Then hi = MAX_S_CAP + 1

    Do While hi - lo > 1
        mid = (lo + hi) \ 2
        If Propagate(LegData, startIdx, mid, chosenIdx, total) Then
            If total < budget Then
                lo = mid
                bestChosen = chosenIdx
                bestTotal = total
            Else
                hi = mid
            End If
        Else
            hi = mid
        End If
    Loop

    outChosen = bestChosen
    outTotalTime = bestTotal

    Dim i As Long, req As Long, thisSlack As Double
    Dim leg As Variant, nextLeg As Variant
    Dim first As Boolean
    first = True
    outTotalSlack = 0
    outMinSlack = 0
    For i = 0 To N_LEGS - 2
        If IsOptimizedCity(LegDest(i)) Then
            req = RequiredMinutes(LegDest(i))
            leg = LegData(i)
            nextLeg = LegData(i + 1)
            thisSlack = (nextLeg(outChosen(i + 1), 1) - leg(outChosen(i), 2)) * 1440# - req
            outTotalSlack = outTotalSlack + thisSlack
            If first Then
                outMinSlack = thisSlack
                first = False
            ElseIf thisSlack < outMinSlack Then
                outMinSlack = thisSlack
            End If
        End If
    Next i

    BestForStart = True
End Function


' ===================== ALL-CANDIDATES SEARCH =====================
' The one truly free choice in this whole problem is which flight you start
' on -- every leg after that is then the forced, greedy-optimal choice for
' whatever slack floor is being targeted (see Propagate). So rather than a
' single global search that reports just one "the" answer, this runs
' BestForStart (a per-start binary search) for EVERY possible starting
' flight, keeping every one that's feasible and under budget, and sorts the
' results by minimum slack descending, then TOTAL slack descending as the
' tie-break. Rank 1 of that sort is used as THE optimal path (fed to
' WriteResults) and the same sorted list is reused for the Top-10 table below
' -- one computation, one consistent answer, instead of two separate searches
' that could disagree on how to break a tie.
'
' This tie-break is the "smarter" answer to a real question: among routes
' that share the same worst-case connection, prefer the one that's also most
' generous everywhere else, rather than an arbitrary pick (e.g. shortest
' total trip time) that ignores how much slack the rest of the trip has.
Private Sub ComputeAllCandidates(LegData() As Variant, budget As Long, _
                                  ByRef candMinSlack() As Double, ByRef candTotalSlack() As Double, _
                                  ByRef candTotalTime() As Double, ByRef candChosen() As Variant, ByRef cnt As Long)
    Dim leg0 As Variant
    Dim n0 As Long, s0 As Long
    Dim chosenIdx() As Long
    Dim minSlack As Double, totalSlack As Double, totalTime As Double

    Application.StatusBar = "Searching every possible starting flight..."

    leg0 = LegData(0)
    n0 = UBound(leg0, 1)
    cnt = 0
    ReDim candMinSlack(1 To n0)
    ReDim candTotalSlack(1 To n0)
    ReDim candTotalTime(1 To n0)
    ReDim candChosen(1 To n0)

    For s0 = 1 To n0
        If BestForStart(LegData, s0, budget, chosenIdx, minSlack, totalSlack, totalTime) Then
            cnt = cnt + 1
            candMinSlack(cnt) = minSlack
            candTotalSlack(cnt) = totalSlack
            candTotalTime(cnt) = totalTime
            candChosen(cnt) = chosenIdx
        End If
    Next s0

    Application.StatusBar = False
    If cnt = 0 Then Exit Sub

    ' Insertion sort, minSlack desc then totalSlack desc -- same simple,
    ' obviously-correct approach as SortLegByDep; cnt is at most a few hundred.
    Dim i As Long, j As Long
    Dim kMin As Double, kTot As Double, kTime As Double, kChosen As Variant
    For i = 2 To cnt
        kMin = candMinSlack(i): kTot = candTotalSlack(i): kTime = candTotalTime(i): kChosen = candChosen(i)
        j = i - 1
        Do While j >= 1
            If candMinSlack(j) > kMin Or (candMinSlack(j) = kMin And candTotalSlack(j) >= kTot) Then Exit Do
            candMinSlack(j + 1) = candMinSlack(j)
            candTotalSlack(j + 1) = candTotalSlack(j)
            candTotalTime(j + 1) = candTotalTime(j)
            candChosen(j + 1) = candChosen(j)
            j = j - 1
        Loop
        candMinSlack(j + 1) = kMin
        candTotalSlack(j + 1) = kTot
        candTotalTime(j + 1) = kTime
        candChosen(j + 1) = kChosen
    Next i
End Sub


' Writes the Top-10 (or fewer, if fewer exist) alternative-routes table to
' the bottom of the already-written output sheet, from the pre-sorted
' candidate list ComputeAllCandidates produced. Rank 1 here is exactly the
' same itinerary already detailed in full above -- same array, same sort --
' so there is no possibility of the two disagreeing.
Private Sub WriteAlternativesTable(LegData() As Variant, outSheetName As String, _
                                    candMinSlack() As Double, candTotalSlack() As Double, _
                                    candTotalTime() As Double, candChosen() As Variant, cnt As Long)
    Dim topN As Long
    topN = cnt
    If topN > 10 Then topN = 10

    Application.ScreenUpdating = False
    Dim outWs As Worksheet
    Set outWs = ThisWorkbook.Sheets(outSheetName)
    Dim startRow As Long, hdrRow As Long, r As Long, i As Long
    startRow = outWs.Cells(outWs.Rows.Count, 1).End(xlUp).Row + 2

    With outWs.Range("A" & startRow & ":G" & startRow)
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .Value = "TOP " & topN & " ALTERNATIVE ROUTES  (out of " & cnt & " feasible starting flights)"
    End With

    hdrRow = startRow + 1
    outWs.Cells(hdrRow, 1).Resize(1, 7).Value = Array("Rank", "Min Slack (min)", "Min Slack", _
        "Total Slack (min)", "Total Slack", "Total Trip Time", "First Departure")
    With outWs.Cells(hdrRow, 1).Resize(1, 7)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
    End With

    Dim leg As Variant, firstDep As Date, thisChosen As Variant, firstDepIdx As Long
    leg = LegData(0)
    For i = 1 To topN
        r = hdrRow + i
        thisChosen = candChosen(i)
        firstDepIdx = thisChosen(0)
        firstDep = leg(firstDepIdx, 1)

        outWs.Cells(r, 1).Value = i
        outWs.Cells(r, 2).Value = CLng(Round(candMinSlack(i), 0))
        outWs.Cells(r, 3).Value = FormatDuration(candMinSlack(i))
        outWs.Cells(r, 4).Value = CLng(Round(candTotalSlack(i), 0))
        outWs.Cells(r, 5).Value = FormatDuration(candTotalSlack(i))
        outWs.Cells(r, 6).Value = FormatDuration(candTotalTime(i))
        outWs.Cells(r, 7).Value = Format(firstDep, "mmm d, yyyy h:mm AM/PM")

        If i = 1 Then
            outWs.Cells(r, 1).Resize(1, 7).Interior.Color = RGB(255, 193, 7)     ' gold = the optimal path detailed in full above
        ElseIf (i Mod 2) = 0 Then
            outWs.Cells(r, 1).Resize(1, 7).Interior.Color = RGB(222, 235, 247)
        Else
            outWs.Cells(r, 1).Resize(1, 7).Interior.Color = RGB(242, 242, 242)
        End If
        outWs.Cells(r, 1).Resize(1, 7).HorizontalAlignment = xlCenter
    Next i
    outWs.Cells(hdrRow, 1).Resize(topN + 1, 7).Borders.LineStyle = xlContinuous

    Dim noteRow As Long
    noteRow = hdrRow + topN + 2
    outWs.Range("A" & noteRow & ":J" & noteRow).Merge
    outWs.Range("A" & noteRow).Value = _
        "Rank 1 (gold) is exactly the optimal path detailed in full above. Every other rank is the best COMPLETE " & _
        "itinerary achievable starting from a different flight -- not an arbitrary partial trade-off -- so compare " & _
        "Min Slack against Total Slack to judge whether a lower rank's larger total is worth its smaller worst case."
    outWs.Range("A" & noteRow).Font.Italic = True
    outWs.Range("A" & noteRow & ":J" & noteRow).WrapText = True

    outWs.Columns("A:G").AutoFit
    Application.ScreenUpdating = True
End Sub


Private Sub WriteResults(LegData() As Variant, chosenIdx() As Long, _
                          Sstar As Long, totalMinutes As Double, budget As Long, isFeasible As Boolean, _
                          routeLabel As String, outSheetName As String)
    Dim outWs As Worksheet
    Dim i As Long, r As Long
    Dim depDT As Date, arrDT As Date
    Dim srcRow As Long
    Dim cityAfter As String
    Dim req As Long
    Dim leg As Variant, nextLeg As Variant
    Dim flightLayover As String
    Dim hdrRow As Long, connHdr As Long
    Dim minSlackAmongOptimized As Double, thisSlack As Double
    Dim minSlackRow As Long
    Dim arrRow As Long, depRow As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(outSheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set outWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    outWs.Name = outSheetName

    ' ===================== SUMMARY BANNER =====================
    With outWs.Range("A1:J1")
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        If isFeasible Then
            .Value = routeLabel & "  |  OPTIMAL PATH FOUND  |  Guaranteed slack: " & Sstar & " min (" & FormatDuration(CDbl(Sstar)) & ")"
            .Interior.Color = RGB(46, 125, 50)
        Else
            .Value = routeLabel & "  |  NO PATH FITS THE BUDGET  |  Showing tightest possible itinerary"
            .Interior.Color = RGB(198, 40, 40)
        End If
    End With

    outWs.Range("A2").Value = "Total trip time (time-zone adjusted):"
    outWs.Range("B2").Value = FormatDuration(totalMinutes)
    outWs.Range("D2").Value = "Budget:"
    outWs.Range("E2").Value = FormatDuration(CDbl(budget))
    outWs.Range("G2").Value = "Under budget?"
    outWs.Range("H2").Value = IIf(totalMinutes < budget, "YES", "NO")
    outWs.Range("H2").Font.Bold = True
    outWs.Range("H2").Font.Color = IIf(totalMinutes < budget, RGB(46, 125, 50), RGB(198, 40, 40))
    outWs.Range("A2:H2").Font.Bold = True

    Dim naiveTotal As Double, tzGap As Long
    naiveTotal = totalMinutes + TZCorrectionMinutes()
    tzGap = TZCorrectionMinutes()
    With outWs.Range("A3:J3")
        .Merge
        .Font.Italic = True
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
        .Value = "Note: " & LegDest(N_LEGS - 1) & " is " & (Abs(tzGap) \ 60) & "h " & (Abs(tzGap) Mod 60) & "m " & _
                 IIf(tzGap >= 0, "ahead of", "behind") & " " & LegOrigin(0) & "'s time zone (both fixed, no DST). " & _
                 "The total above already accounts for this. The raw difference between the timestamps as they " & _
                 "literally appear on the departure/arrival cells would read " & FormatDuration(naiveTotal) & "."
    End With

    ' ===================== FLIGHT TABLE =====================
    hdrRow = 4
    outWs.Cells(hdrRow, 1).Resize(1, 10).Value = Array("Leg", "Origin", "Destination", "Departure Date", _
        "Departure Time", "Arrival Date", "Arrival Time", "Leg Duration", "Source Row (" & SHEET_NAME & ")", "Flight Layover")
    With outWs.Cells(hdrRow, 1).Resize(1, 10)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With

    For i = 0 To N_LEGS - 1
        r = hdrRow + 1 + i
        leg = LegData(i)
        depDT = leg(chosenIdx(i), 1)
        arrDT = leg(chosenIdx(i), 2)
        srcRow = leg(chosenIdx(i), 3)
        flightLayover = leg(chosenIdx(i), 4)

        outWs.Cells(r, 1).Value = i + 1
        outWs.Cells(r, 2).Value = LegOrigin(i)
        outWs.Cells(r, 3).Value = LegDest(i)
        outWs.Cells(r, 4).Value = DateValue(depDT)
        outWs.Cells(r, 4).NumberFormat = "mmm d, yyyy (ddd)"
        outWs.Cells(r, 5).Value = TimeValue(depDT)
        outWs.Cells(r, 5).NumberFormat = "h:mm AM/PM"
        outWs.Cells(r, 6).Value = DateValue(arrDT)
        outWs.Cells(r, 6).NumberFormat = "mmm d, yyyy (ddd)"
        outWs.Cells(r, 7).Value = TimeValue(arrDT)
        outWs.Cells(r, 7).NumberFormat = "h:mm AM/PM"
        ' Leg Duration = (local arrival - local departure), corrected for the
        ' UTC-offset gap between this leg's two airports (see UTCOffsetMinutes
        ' above) so it reflects true flight time, not a raw clock difference.
        outWs.Cells(r, 8).Formula = "=(F" & r & "+G" & r & ")-(D" & r & "+E" & r & ")-(" & _
            (UTCOffsetMinutes(LegDest(i)) - UTCOffsetMinutes(LegOrigin(i))) & "/1440)"
        outWs.Cells(r, 8).NumberFormat = "[h]:mm"
        outWs.Cells(r, 9).Value = srcRow
        ' Mid-flight stopover text pulled straight from the source sheet's
        ' "Layover Duration" column, if present -- informational only, NOT
        ' the same thing as the "Actual Layover" in the connections table
        ' below (which is the real city-to-city gap this macro optimizes).
        ' Blank for legs with no stopover (nonstop flights, train legs) or
        ' when the source sheet has no such column at all.
        If flightLayover <> "" Then outWs.Cells(r, 10).Value = flightLayover

        If (i Mod 2) = 0 Then
            outWs.Cells(r, 1).Resize(1, 10).Interior.Color = RGB(222, 235, 247)
        Else
            outWs.Cells(r, 1).Resize(1, 10).Interior.Color = RGB(242, 242, 242)
        End If
    Next i
    outWs.Cells(hdrRow, 1).Resize(N_LEGS + 1, 10).Borders.LineStyle = xlContinuous

    ' ===================== CONNECTIONS / AUDIT TABLE =====================
    connHdr = hdrRow + N_LEGS + 3
    outWs.Cells(connHdr, 1).Resize(1, 10).Value = Array("Connecting City", "Arrival (prior leg)", _
        "Departure (next leg)", "Actual Layover (min)", "Actual Layover", "Minimum Required (min)", _
        "Minimum Required", "Slack (min)", "Slack = Layover - Min", "Counts Toward Objective?")
    With outWs.Cells(connHdr, 1).Resize(1, 10)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With

    minSlackAmongOptimized = -1
    minSlackRow = 0

    For i = 0 To N_LEGS - 2   ' N_LEGS-1 connecting cities (CUZ may appear twice on routes 3-4 & 7-12)
        r = connHdr + 1 + i
        cityAfter = LegDest(i)
        req = RequiredMinutes(cityAfter)
        arrRow = hdrRow + 1 + i
        depRow = hdrRow + 1 + i + 1

        outWs.Cells(r, 1).Value = cityAfter
        outWs.Cells(r, 2).Formula = "=F" & arrRow & "+G" & arrRow
        outWs.Cells(r, 2).NumberFormat = "mmm d h:mm AM/PM"
        outWs.Cells(r, 3).Formula = "=D" & depRow & "+E" & depRow
        outWs.Cells(r, 3).NumberFormat = "mmm d h:mm AM/PM"
        outWs.Cells(r, 4).Formula = "=ROUND((C" & r & "-B" & r & ")*1440,0)"
        outWs.Cells(r, 5).Formula = "=INT(D" & r & "/60)&""h ""&MOD(D" & r & ",60)&""m"""
        outWs.Cells(r, 6).Value = req
        outWs.Cells(r, 7).Formula = "=INT(F" & r & "/60)&""h ""&MOD(F" & r & ",60)&""m"""
        outWs.Cells(r, 8).Formula = "=D" & r & "-F" & r
        outWs.Cells(r, 9).Formula = "=IF(H" & r & ">=0,INT(H" & r & "/60)&""h ""&MOD(H" & r & ",60)&""m"",""-""&INT(ABS(H" & r & ")/60)&""h ""&MOD(ABS(H" & r & "),60)&""m"")"
        outWs.Cells(r, 10).Value = IIf(IsOptimizedCity(cityAfter), "YES", "no (CUZ)")

        If IsOptimizedCity(cityAfter) Then
            leg = LegData(i)
            nextLeg = LegData(i + 1)
            thisSlack = (nextLeg(chosenIdx(i + 1), 1) - leg(chosenIdx(i), 2)) * 1440# - req
            If minSlackRow = 0 Then
                minSlackAmongOptimized = thisSlack
                minSlackRow = r
            ElseIf thisSlack < minSlackAmongOptimized Then
                minSlackAmongOptimized = thisSlack
                minSlackRow = r
            End If
        End If
    Next i
    outWs.Cells(connHdr, 1).Resize(N_LEGS, 10).Borders.LineStyle = xlContinuous

    For i = 0 To N_LEGS - 2
        r = connHdr + 1 + i
        cityAfter = LegDest(i)
        If IsOptimizedCity(cityAfter) Then
            If r = minSlackRow Then
                outWs.Cells(r, 1).Resize(1, 10).Interior.Color = RGB(255, 193, 7)     ' gold = the binding constraint
            Else
                outWs.Cells(r, 1).Resize(1, 10).Interior.Color = RGB(198, 239, 206)   ' green = has surplus slack
            End If
        Else
            outWs.Cells(r, 1).Resize(1, 10).Interior.Color = RGB(230, 230, 230)       ' gray = not part of objective (CUZ)
        End If
    Next i

    ' ===================== TRIP TIMELINE (linear, step-by-step) =====================
    ' A colorful, top-to-bottom retelling of the same itinerary above: one row
    ' per FLIGHT/TRAIN segment (you're moving) alternating with one row per
    ' LAYOVER segment (you're waiting in a city). Every value here is a
    ' formula referencing the flight table or connections table already built
    ' above, so this is purely a more readable *view* of that same data, not
    ' a second calculation that could drift out of sync with it.
    Dim tlHdr As Long, tlRow As Long, stepNum As Long, mode As String
    Dim flightRow As Long, connRow As Long

    tlHdr = connHdr + N_LEGS + 2
    With outWs.Range("A" & tlHdr & ":J" & tlHdr)
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .Value = "TRIP TIMELINE  --  " & routeLabel
    End With

    tlRow = tlHdr + 1
    stepNum = 1
    For i = 0 To N_LEGS - 1
        flightRow = hdrRow + 1 + i
        mode = IIf(UCase(LegOrigin(i)) = "MPU" Or UCase(LegDest(i)) = "MPU", "TRAIN", "FLIGHT")

        outWs.Cells(tlRow, 1).Value = stepNum
        outWs.Cells(tlRow, 2).Value = mode
        outWs.Range("C" & tlRow & ":D" & tlRow).Merge
        outWs.Cells(tlRow, 3).Formula = "=B" & flightRow & "&""  ->  ""&C" & flightRow
        outWs.Range("E" & tlRow & ":F" & tlRow).Merge
        outWs.Cells(tlRow, 5).Formula = "=TEXT(D" & flightRow & ",""mmm d"")&"" ""&TEXT(E" & flightRow & ",""h:mm AM/PM"")"
        outWs.Range("G" & tlRow & ":H" & tlRow).Merge
        outWs.Cells(tlRow, 7).Formula = "=TEXT(F" & flightRow & ",""mmm d"")&"" ""&TEXT(G" & flightRow & ",""h:mm AM/PM"")"
        outWs.Range("I" & tlRow & ":J" & tlRow).Merge
        outWs.Cells(tlRow, 9).Formula = "=H" & flightRow
        outWs.Cells(tlRow, 9).NumberFormat = "[h]:mm"" hrs"""

        outWs.Range("A" & tlRow & ":J" & tlRow).Interior.Color = RGB(31, 119, 180)     ' blue = in motion
        outWs.Range("A" & tlRow & ":J" & tlRow).Font.Color = RGB(255, 255, 255)
        outWs.Range("A" & tlRow & ":J" & tlRow).Font.Bold = True
        outWs.Range("A" & tlRow & ":J" & tlRow).HorizontalAlignment = xlCenter
        tlRow = tlRow + 1
        stepNum = stepNum + 1

        If i < N_LEGS - 1 Then
            connRow = connHdr + 1 + i
            outWs.Cells(tlRow, 1).Value = stepNum
            outWs.Cells(tlRow, 2).Value = "LAYOVER"
            outWs.Range("C" & tlRow & ":D" & tlRow).Merge
            outWs.Cells(tlRow, 3).Formula = "=""Time in ""&A" & connRow
            outWs.Range("E" & tlRow & ":F" & tlRow).Merge
            outWs.Cells(tlRow, 5).Formula = "=TEXT(B" & connRow & ",""mmm d h:mm AM/PM"")"
            outWs.Range("G" & tlRow & ":H" & tlRow).Merge
            outWs.Cells(tlRow, 7).Formula = "=TEXT(C" & connRow & ",""mmm d h:mm AM/PM"")"
            outWs.Range("I" & tlRow & ":J" & tlRow).Merge
            outWs.Cells(tlRow, 9).Formula = "=E" & connRow

            If IsOptimizedCity(LegDest(i)) Then
                If connRow = minSlackRow Then
                    outWs.Range("A" & tlRow & ":J" & tlRow).Interior.Color = RGB(255, 193, 7)    ' gold = the binding connection
                Else
                    outWs.Range("A" & tlRow & ":J" & tlRow).Interior.Color = RGB(198, 239, 206)  ' green = has surplus slack
                End If
            Else
                outWs.Range("A" & tlRow & ":J" & tlRow).Interior.Color = RGB(230, 230, 230)      ' gray = CUZ, no minimum
            End If
            outWs.Range("A" & tlRow & ":J" & tlRow).Font.Bold = True
            outWs.Range("A" & tlRow & ":J" & tlRow).HorizontalAlignment = xlCenter
            tlRow = tlRow + 1
            stepNum = stepNum + 1
        End If
    Next i
    outWs.Range("A" & tlHdr & ":J" & (tlRow - 1)).Borders.LineStyle = xlContinuous

    ' legend
    tlRow = tlRow + 1
    outWs.Cells(tlRow, 1).Value = "Key:"
    outWs.Cells(tlRow, 1).Font.Bold = True
    outWs.Range("B" & tlRow & ":C" & tlRow).Merge
    outWs.Range("B" & tlRow & ":C" & tlRow).Value = "In motion (flight/train)"
    outWs.Range("B" & tlRow & ":C" & tlRow).Interior.Color = RGB(31, 119, 180)
    outWs.Range("B" & tlRow & ":C" & tlRow).Font.Color = RGB(255, 255, 255)
    outWs.Range("D" & tlRow & ":E" & tlRow).Merge
    outWs.Range("D" & tlRow & ":E" & tlRow).Value = "Tightest layover (binding)"
    outWs.Range("D" & tlRow & ":E" & tlRow).Interior.Color = RGB(255, 193, 7)
    outWs.Range("F" & tlRow & ":G" & tlRow).Merge
    outWs.Range("F" & tlRow & ":G" & tlRow).Value = "Layover with surplus slack"
    outWs.Range("F" & tlRow & ":G" & tlRow).Interior.Color = RGB(198, 239, 206)
    outWs.Range("H" & tlRow & ":J" & tlRow).Merge
    outWs.Range("H" & tlRow & ":J" & tlRow).Value = "CUZ layover (no minimum, not optimized)"
    outWs.Range("H" & tlRow & ":J" & tlRow).Interior.Color = RGB(230, 230, 230)
    outWs.Range("A" & tlRow & ":J" & tlRow).HorizontalAlignment = xlCenter

    outWs.Columns("A:J").AutoFit
    outWs.Activate
    Application.ScreenUpdating = True
End Sub
