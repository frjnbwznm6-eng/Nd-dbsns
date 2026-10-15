Attribute VB_Name = "modFindOptimalPath"
Option Explicit

'===============================================================================
' FindOptimalPath
'
' Searches EVERY supported routing (all 12 below) on EVERY possible start
' date/flight in the sheet named DATA_SHEET, and writes the TOP_N best
' distinct itineraries side by side on one sheet (OUT_SHEET).
'
' A prompt at the start lets you limit the search to routes that START in
' one city: press Enter (ALL) for every route, or type a starting city such
' as AMM (routes 9 and 10) to see the best options that start there. A
' city-limited run writes to its own sheet, e.g. "Top 10 Routes - from AMM",
' so it never overwrites the all-routes results.
'
' "Best" is the same goal as before: maximize the MINIMUM slack (actual time
' spent in a city minus that city's required minimum) across the optimized
' destination cities, subject to the total trip time (first departure ->
' last arrival, time-zone adjusted) being UNDER the budget. Ties are broken
' by TOTAL slack (sum over every optimized city), then by the shorter trip.
'
'   Route 1: MPU -> CUZ -> CUN -> GIG -> FCO -> AMM -> DEL -> PEK
'   Route 2: MPU -> CUZ -> GIG -> CUN -> FCO -> AMM -> DEL -> PEK
'   Route 3: GIG -> CUZ -> MPU -> CUZ -> CUN -> FCO -> AMM -> DEL -> PEK
'   Route 4: CUN -> CUZ -> MPU -> CUZ -> GIG -> FCO -> AMM -> DEL -> PEK
'   Route 5: PEK -> DEL -> AMM -> FCO -> GIG -> CUN -> CUZ -> MPU
'   Route 6: PEK -> DEL -> AMM -> FCO -> CUN -> GIG -> CUZ -> MPU
'   Route 7: PEK -> DEL -> AMM -> FCO -> GIG -> CUZ -> MPU -> CUZ -> CUN
'   Route 8: PEK -> DEL -> AMM -> FCO -> CUN -> CUZ -> MPU -> CUZ -> GIG
'   Route 9: AMM -> DEL -> PEK -> FCO -> GIG -> CUZ -> MPU -> CUZ -> CUN
'  Route 10: AMM -> DEL -> PEK -> FCO -> CUN -> CUZ -> MPU -> CUZ -> GIG
'  Route 11: GIG -> CUZ -> MPU -> CUZ -> CUN -> FCO -> PEK -> DEL -> AMM
'  Route 12: CUN -> CUZ -> MPU -> CUZ -> GIG -> FCO -> PEK -> DEL -> AMM
' (See RouteText to add or change routes.)
'
' THE DATA
'   DATA_SHEET lists every individual flight (and the CUZ <-> MPU trains)
'   separately: a trip from CUN to CUZ might be CUN -> LIM, then LIM -> CUZ.
'   So the macro builds those connections itself. Rules for a connection at
'   an in-between airport (one that isn't a main destination):
'     - at least MIN_CONNECTION_MINUTES (60) between landing and the next
'       departure, at the SAME airport (no airport changes mid-connection)
'     - connections never pass through a main-destination city; a leg from
'       one destination to the next only stops at in-between airports
'   Main destinations are cities, not airports: AirportCity maps PKX
'   (Beijing Daxing) to PEK, so landing at either Beijing airport counts as
'   being in Beijing (the itinerary notes when you land at one and leave from
'   the other).
'
' MINIMUM STAYS AND OPEN HOURS (unchanged)
'   Every main destination you pass THROUGH has a required minimum stay
'   (RequiredMinutes). Some also have daily open hours (OpenWindow and
'   OpenBuffers): the part of the stay from (arrival + an after-arrival
'   buffer) to (next departure - a before-departure buffer) must overlap the
'   open hours by MORE than zero minutes, on the local day(s) you're there.
'   The first and last city of a route have neither rule. CUZ has a 60-minute
'   minimum that is enforced but doesn't count toward the ranking.
'
' ALGORITHM
'   1. Door-to-door legs. For every pair of consecutive destinations used by
'      any route (e.g. CUN -> CUZ), BuildCompositeLegs works out, for every
'      flight leaving the first city, the EARLIEST possible arrival at the
'      second city via any chain of connecting flights (each connection >= 60
'      min). It does this with one backward scan over all flights per
'      destination (sorted latest departure first), so the whole 64k-row
'      sheet is processed a handful of times, not once per flight. Options
'      that are beaten outright (another option leaves later and still
'      arrives no later) are dropped. What's left is a short timetable per
'      pair, shaped just like the old sheet's rows.
'   2. Per route, per start. The one truly free choice is which option you
'      start on. For a slack floor S, Propagate walks the route taking, at
'      each city, the earliest-arriving option that leaves no sooner than the
'      minimum stay (+ S, for optimized cities) and the open-hours rule
'      allow. Arriving earlier can never hurt anything downstream, so this
'      greedy choice is optimal, and BestForStart binary-searches the
'      largest S that still fits the budget for that start.
'   3. Ranking and grouping. Every (route, start) result is sorted by min
'      slack, then total slack, then shorter trip. Results that are the same
'      route through the same airports in the same order are ONE option: the
'      best-ranked one is shown, and the other start dates it works on are
'      listed in that option's "Also works starting" note (marked when the
'      flight times on that date differ). The first TOP_N options are written.
'
' REQUIRED SHEET LAYOUT
'   Row 1 of DATA_SHEET must have headers ORIGIN, DESTINATION,
'   DEPARTURE_DATETIME, ARRIVAL_DATETIME (any order, any case).
'   FLIGHT_DURATION_MINUTES is optional and only used for display.
'   Times are local to each airport. Your source sheet is never modified.
'
' OUTPUT
'   A sheet named OUT_SHEET is (re)created each run with one "card" per
'   option, side by side: rank, route, worst-case slack, total slack, total
'   trip time, start/end, the other start dates it works on, and a
'   step-by-step table of every flight, connection and destination stay.
'   The tightest destination (the one that sets the worst-case slack) is
'   gold, other counted destinations green, CUZ gray, connections light gray.
'
' HOW TO USE
'   Alt+F11 -> select the old modFindOptimalPath module -> File -> Remove ->
'   then File -> Import File... and pick this .bas -> F5 or Run >
'   FindOptimalPath. Adjust the constants below if yours differ.
'===============================================================================

' --------------------------- CONFIGURATION ------------------------------------
Private Const DATA_SHEET As String = "FULL FLIGHT DATA"   ' <-- change if needed
Private Const OUT_SHEET As String = "Top 10 Routes"
Private Const TOP_N As Long = 10

Private Const MIN_CONNECTION_MINUTES As Long = 60   ' minimum layover between two flights at an in-between airport

Private Const BUDGET_DAYS As Long = 5
Private Const BUDGET_HOURS As Long = 14
Private Const BUDGET_MINUTES As Long = 0

Private Const MAX_S_CAP As Long = 43200   ' safety cap on the slack search, minutes (30 days)

Private Const N_ROUTES As Long = 12
Private Const N_CITIES As Long = 8        ' MPU, CUZ, CUN, GIG, FCO, AMM, DEL, PEK -- see CityCode
Private Const CARD_COLS As Long = 10      ' columns per output card (plus one blank spacer column)
' -------------------------------------------------------------------------------

' Airports, indexed 1..nAp. apCity is the main-destination city index
' (1..N_CITIES) for airports that belong to one, else 0.
Private nAp As Long
Private apCode() As String
Private apCity() As Long
Private apLookup As Collection

' Individual flights/trains from DATA_SHEET, indexed 1..nF. Times are whole
' minutes since VBA's day 0 (local time at that airport).
Private nF As Long
Private fOrig() As Long, fDest() As Long
Private fDep() As Long, fArr() As Long
Private fDur() As Long, fRow() As Long

' Door-to-door options ("composite legs") between consecutive destinations.
' Pair p's options are cDep/cArr(pStart(p) .. pStart(p) + pCount(p) - 1),
' sorted by departure, with departure AND arrival strictly increasing. Each
' option's flights are segFlight(cSegStart(c) .. cSegStart(c) + cSegCount(c) - 1).
Private pairIdx() As Long          ' (1 To N_CITIES, 1 To N_CITIES) -> pair number, 0 if unused
Private nPairs As Long
Private pFrom() As Long, pTo() As Long, pStart() As Long, pCount() As Long
Private nC As Long
Private cDep() As Long, cArr() As Long, cSegStart() As Long, cSegCount() As Long
Private nSeg As Long
Private segFlight() As Long

' Routes: routeCity(r, 0 .. routeLegs(r)) are the city indexes in order.
' routeOn(r) is False for routes excluded by the starting-city choice.
Private routeCity() As Long
Private routeLegs() As Long
Private routeOn() As Boolean
Private startFilter As String      ' "" = all routes, else the required starting city
Private nRoutesOn As Long

' Every feasible (route, start) result, indexed 1..nCand.
Private nCand As Long
Private candRoute() As Long
Private candMin() As Long, candTot() As Long, candTime() As Long
Private candChosen() As Variant    ' each a (0 To legs-1) Long array of composite-leg indexes

' Scratch for the flight sort (sorted by fDep ascending).
Private order() As Long


Public Sub FindOptimalPath()
    Dim budget As Long, t0 As Single
    Dim i As Long, nShown As Long
    Dim rankIdx() As Long
    Dim shown() As Long, alsoNote() As String
    Dim missingPairs As String
    Dim prevCalc As Long

    Call InitRoutes
    If Not AskStartCity() Then Exit Sub

    On Error GoTo Fail
    t0 = Timer
    budget = BudgetMinutes()
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Application.StatusBar = "Reading '" & DATA_SHEET & "'..."
    If Not LoadFlights() Then GoTo Done

    Application.StatusBar = "Building connections between destinations (" & nF & " flights)..."
    Call BuildCompositeLegs(missingPairs)

    Application.StatusBar = "Searching " & ScopeText() & " on every start date..."
    Call SearchAllRoutes(budget)

    If nCand = 0 Then
        Application.Calculation = prevCalc
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "No itinerary on " & ScopeText() & " fits within the " & _
               FormatDuration(CDbl(budget)) & " budget while meeting every minimum stay, " & _
               "open-hours rule and the " & MIN_CONNECTION_MINUTES & "-minute connection minimum." & _
               IIf(missingPairs <> "", vbCrLf & vbCrLf & "No connections at all were found for: " & missingPairs, ""), _
               vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "Ranking " & nCand & " itineraries..."
    Call RankCandidates(rankIdx)
    Call GroupTopOptions(rankIdx, shown, alsoNote, nShown)

    Application.StatusBar = "Writing '" & OutSheetName() & "'..."
    Call WriteTopSheet(shown, alsoNote, nShown, budget, missingPairs)

    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Searched " & ScopeText() & ", " & nF & " flights, " & nCand & " complete itineraries under budget." & vbCrLf & vbCrLf & _
           "Best: Route " & candRoute(shown(1)) & " -- worst-case slack " & FormatDuration(CDbl(candMin(shown(1)))) & _
           ", trip time " & FormatDuration(CDbl(candTime(shown(1)))) & " (budget " & FormatDuration(CDbl(budget)) & ")." & vbCrLf & vbCrLf & _
           "The top " & nShown & " options are side by side on '" & OutSheetName() & "'." & vbCrLf & _
           "(" & Format(Timer - t0, "0.0") & " s)", vbInformation, "Find Optimal Path"
    Exit Sub

Done:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

Fail:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "FindOptimalPath stopped with an error: " & Err.Description, vbCritical
End Sub


' ===================== ROUTES AND CITIES =====================
Private Function RouteText(ByVal r As Long) As String
    Select Case r
        Case 1: RouteText = "MPU CUZ CUN GIG FCO AMM DEL PEK"
        Case 2: RouteText = "MPU CUZ GIG CUN FCO AMM DEL PEK"
        Case 3: RouteText = "GIG CUZ MPU CUZ CUN FCO AMM DEL PEK"
        Case 4: RouteText = "CUN CUZ MPU CUZ GIG FCO AMM DEL PEK"
        Case 5: RouteText = "PEK DEL AMM FCO GIG CUN CUZ MPU"
        Case 6: RouteText = "PEK DEL AMM FCO CUN GIG CUZ MPU"
        Case 7: RouteText = "PEK DEL AMM FCO GIG CUZ MPU CUZ CUN"
        Case 8: RouteText = "PEK DEL AMM FCO CUN CUZ MPU CUZ GIG"
        Case 9: RouteText = "AMM DEL PEK FCO GIG CUZ MPU CUZ CUN"
        Case 10: RouteText = "AMM DEL PEK FCO CUN CUZ MPU CUZ GIG"
        Case 11: RouteText = "GIG CUZ MPU CUZ CUN FCO PEK DEL AMM"
        Case 12: RouteText = "CUN CUZ MPU CUZ GIG FCO PEK DEL AMM"
    End Select
End Function


Private Function RouteLabel(ByVal r As Long) As String
    RouteLabel = Replace(RouteText(r), " ", " -> ")
End Function


Private Function CityCode(ByVal c As Long) As String
    CityCode = Choose(c, "MPU", "CUZ", "CUN", "GIG", "FCO", "AMM", "DEL", "PEK")
End Function


Private Function CityIndex(ByVal code As String) As Long
    Dim c As Long
    For c = 1 To N_CITIES
        If CityCode(c) = code Then CityIndex = c: Exit Function
    Next c
    CityIndex = 0
End Function


' The main-destination city an airport belongs to ("" if none). Add a line
' here if another airport should count as one of the destinations.
Private Function AirportCity(ByVal code As String) As String
    Select Case code
        Case "PKX": AirportCity = "PEK"   ' Beijing Daxing counts as Beijing
        Case Else: AirportCity = code
    End Select
End Function


Private Sub InitRoutes()
    Dim r As Long, j As Long, parts() As String
    ReDim routeCity(1 To N_ROUTES, 0 To 20)
    ReDim routeLegs(1 To N_ROUTES)
    For r = 1 To N_ROUTES
        parts = Split(RouteText(r), " ")
        routeLegs(r) = UBound(parts)
        For j = 0 To UBound(parts)
            routeCity(r, j) = CityIndex(parts(j))
        Next j
    Next r
    ReDim routeOn(1 To N_ROUTES)
    For r = 1 To N_ROUTES: routeOn(r) = True: Next r
    startFilter = ""
    nRoutesOn = N_ROUTES
End Sub


' Asks which starting city to search from. Enter (ALL) = every route; a
' city code = only routes that start there. Returns False on Cancel/blank.
Private Function AskStartCity() As Boolean
    Dim r As Long, resp As String, starts As String, code As String, msg As String
    starts = "|"
    For r = 1 To N_ROUTES
        code = CityCode(routeCity(r, 0))
        If InStr(starts, "|" & code & "|") = 0 Then starts = starts & code & "|"
    Next r

    msg = "Which starting city should the search use?" & vbCrLf & vbCrLf & _
          "ALL  = every route (press Enter)" & vbCrLf
    For r = 1 To N_ROUTES
        msg = msg & vbCrLf & "Route " & r & ":  " & RouteLabel(r)
    Next r
    msg = msg & vbCrLf & vbCrLf & "Or type a starting city: " & Replace(Mid(starts, 2, Len(starts) - 2), "|", ", ") & _
          vbCrLf & "(Cancel or blank to stop.)"

    Do
        resp = UCase(Trim(InputBox(msg, "Choose Starting City", "ALL")))
        If resp = "" Then Exit Function
        If resp = "ALL" Then
            startFilter = ""
            Exit Do
        End If
        If InStr(starts, "|" & resp & "|") > 0 Then
            startFilter = resp
            Exit Do
        End If
        MsgBox "No route starts in '" & resp & "'. Enter ALL or one of: " & _
               Replace(Mid(starts, 2, Len(starts) - 2), "|", ", "), vbExclamation
    Loop

    nRoutesOn = 0
    For r = 1 To N_ROUTES
        routeOn(r) = (startFilter = "" Or CityCode(routeCity(r, 0)) = startFilter)
        If routeOn(r) Then nRoutesOn = nRoutesOn + 1
    Next r
    AskStartCity = True
End Function


' e.g. "all 12 routes" or "the 2 routes starting in AMM (9, 10)"
Private Function ScopeText() As String
    Dim r As Long, routeList As String
    If startFilter = "" Then
        ScopeText = "all " & N_ROUTES & " routes"
    Else
        For r = 1 To N_ROUTES
            If routeOn(r) Then
                If routeList <> "" Then routeList = routeList & ", "
                routeList = routeList & r
            End If
        Next r
        ScopeText = "the " & nRoutesOn & " route" & IIf(nRoutesOn = 1, "", "s") & " starting in " & startFilter & " (" & routeList & ")"
    End If
End Function


Private Function OutSheetName() As String
    If startFilter = "" Then OutSheetName = OUT_SHEET Else OutSheetName = OUT_SHEET & " - from " & startFilter
End Function


' ===================== CONSTRAINTS (unchanged from the per-route version) =====================

' Fixed UTC offsets for the main destinations (no DST in Nov/Dec for any of
' them; Rome is on CET after late October). Only used to time-zone adjust
' the total trip time between the first and last city.
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


Private Function TZCorrectionMinutes(ByVal r As Long) As Long
    TZCorrectionMinutes = UTCOffsetMinutes(CityCode(routeCity(r, routeLegs(r)))) - _
                          UTCOffsetMinutes(CityCode(routeCity(r, 0)))
End Function


' Minimum stay at each destination you pass through.
Private Function RequiredMinutes(city As String) As Long
    Select Case UCase(city)
        Case "MPU": RequiredMinutes = 2 * 60
        Case "CUZ": RequiredMinutes = 60
        Case "CUN": RequiredMinutes = 8 * 60 + 5
        Case "GIG": RequiredMinutes = 3 * 60 + 1
        Case "FCO": RequiredMinutes = 3 * 60 + 0
        Case "AMM": RequiredMinutes = 7 * 60 + 38
        Case "DEL": RequiredMinutes = 9 * 60 + 0
        Case "PEK": RequiredMinutes = 3 * 60 + 0
        Case Else: RequiredMinutes = 0
    End Select
End Function


' Destinations whose slack counts toward the ranking. CUZ's 60-minute
' minimum is still enforced; its extra time just isn't counted.
Private Function IsOptimizedCity(city As String) As Boolean
    Select Case UCase(city)
        Case "MPU", "CUN", "GIG", "FCO", "AMM", "DEL", "PEK": IsOptimizedCity = True
        Case Else: IsOptimizedCity = False
    End Select
End Function


' --------------------------- OPEN HOURS ---------------------------------------
' Buffers for each city that has open hours: the usable part of a stay
' starts afterArr minutes after you land and ends beforeDep minutes before
' your next departure. Returns False for a city with no open-hours rule.
Private Function OpenBuffers(ByVal city As String, ByRef afterArr As Long, ByRef beforeDep As Long) As Boolean
    OpenBuffers = True
    Select Case UCase(city)
        Case "DEL": afterArr = 2 * 60 + 30: beforeDep = 3 * 60
        Case "PEK": afterArr = 1 * 60:      beforeDep = 2 * 60
        Case "AMM": afterArr = 2 * 60 + 30: beforeDep = 3 * 60
        Case "GIG": afterArr = 1 * 60:      beforeDep = 2 * 60 + 30
        Case "MPU": afterArr = 40:          beforeDep = 40
        Case "CUN": afterArr = 2 * 60:      beforeDep = 3 * 60
        Case Else: OpenBuffers = False     ' FCO, CUZ: no open-hours rule
    End Select
End Function


' The k-th open window (k = 1, 2, ...) for city on weekday wd (VBA Weekday:
' vbSunday = 1 ... vbSaturday = 7), as minutes after local midnight.
' Returns False when there is no k-th window that day (e.g. DEL on Friday).
' Windows within a day must be listed in time order.
Private Function OpenWindow(ByVal city As String, ByVal wd As Long, ByVal k As Long, _
                            ByRef startMin As Long, ByRef endMin As Long) As Boolean
    OpenWindow = False
    Select Case UCase(city)
        Case "DEL"   ' 6:15 AM - 5:00 PM, closed Friday
            If k = 1 And wd <> vbFriday Then startMin = 6 * 60 + 15: endMin = 17 * 60: OpenWindow = True
        Case "PEK"   ' 7:30 AM - 6:00 PM weekdays, 7:30 AM - 6:30 PM Sat/Sun
            If k = 1 Then
                startMin = 7 * 60 + 30
                If wd = vbSaturday Or wd = vbSunday Then endMin = 18 * 60 + 30 Else endMin = 18 * 60
                OpenWindow = True
            End If
        Case "AMM"   ' 6:30 AM - 5:00 PM daily, plus 8:30 PM - 10:30 PM except Fri/Sat
            If k = 1 Then
                startMin = 6 * 60 + 30: endMin = 17 * 60: OpenWindow = True
            ElseIf k = 2 And wd <> vbFriday And wd <> vbSaturday Then
                startMin = 20 * 60 + 30: endMin = 22 * 60 + 30: OpenWindow = True
            End If
        Case "GIG"   ' 7:30 AM - 6:00 PM daily
            If k = 1 Then startMin = 7 * 60 + 30: endMin = 18 * 60: OpenWindow = True
        Case "MPU"   ' 8:00 AM - 5:30 PM daily
            If k = 1 Then startMin = 8 * 60: endMin = 17 * 60 + 30: OpenWindow = True
        Case "CUN"   ' 8:00 AM - 4:00 PM daily
            If k = 1 Then startMin = 8 * 60: endMin = 16 * 60: OpenWindow = True
    End Select
End Function


' The earliest minute >= tMin at which city is open: tMin itself if it falls
' inside an open window, otherwise the start of the next window. Returns -1
' if nothing opens within 8 days. Never decreases as tMin increases, which is
' what keeps the earliest-arrival search optimal.
Private Function EarliestOpenMin(ByVal city As String, ByVal tMin As Long) As Long
    Dim dayN As Long, k As Long, ws As Long, we As Long, sMin As Long, eMin As Long
    For dayN = tMin \ 1440 To tMin \ 1440 + 8
        k = 1
        Do While OpenWindow(city, CLng(Weekday(CDate(dayN))), k, sMin, eMin)
            ws = dayN * 1440 + sMin
            we = dayN * 1440 + eMin
            If tMin < we Then
                If tMin > ws Then EarliestOpenMin = tMin Else EarliestOpenMin = ws
                Exit Function
            End If
            k = k + 1
        Loop
    Next dayN
    EarliestOpenMin = -1
End Function


' The parts of [startAt, endAt) that fall inside city's open windows, as
' clock times, e.g. "9:00 AM - 12:00 PM". If the stay covers more than one
' open window, each is listed with its weekday, e.g.
' "Thu 8:30 PM - 10:30 PM; Fri 6:30 AM - 9:15 AM". "" if there is no overlap.
Private Function OpenOverlapText(ByVal city As String, ByVal startAt As Long, ByVal endAt As Long) As String
    Dim dayN As Long, k As Long, ws As Long, we As Long, sMin As Long, eMin As Long
    Dim lo As Long, hi As Long, n As Long
    Dim segLo() As Long, segHi() As Long, txt As String
    n = 0
    If endAt > startAt Then
        For dayN = startAt \ 1440 To endAt \ 1440
            k = 1
            Do While OpenWindow(city, CLng(Weekday(CDate(dayN))), k, sMin, eMin)
                ws = dayN * 1440 + sMin
                we = dayN * 1440 + eMin
                If startAt > ws Then lo = startAt Else lo = ws
                If endAt < we Then hi = endAt Else hi = we
                If hi > lo Then
                    n = n + 1
                    ReDim Preserve segLo(1 To n)
                    ReDim Preserve segHi(1 To n)
                    segLo(n) = lo
                    segHi(n) = hi
                End If
                k = k + 1
            Loop
        Next dayN
    End If

    For k = 1 To n
        If k > 1 Then txt = txt & "; "
        If n > 1 Then txt = txt & Format(CDate(segLo(k) / 1440#), "ddd ")
        txt = txt & Format(CDate(segLo(k) / 1440#), "h:mm AM/PM") & " - " & _
                    Format(CDate(segHi(k) / 1440#), "h:mm AM/PM")
    Next k
    OpenOverlapText = txt
End Function
' -------------------------------------------------------------------------------


Private Function BudgetMinutes() As Long
    BudgetMinutes = BUDGET_DAYS * 1440 + BUDGET_HOURS * 60 + BUDGET_MINUTES
End Function


Private Function FormatDuration(totalMin As Double) As String
    Dim t As Long, d As Long, h As Long, m As Long, neg As Boolean
    t = CLng(Round(totalMin, 0))
    If t < 0 Then neg = True: t = -t
    d = t \ 1440
    h = (t Mod 1440) \ 60
    m = t Mod 60
    FormatDuration = IIf(neg, "-", "") & d & "d " & h & "h " & m & "m"
End Function


Private Function MinToDate(ByVal m As Long) As Date
    MinToDate = CDate(m / 1440#)
End Function


' ===================== LOADING THE DATA =====================

' A date/time cell (a real Date, a day-number Double, or date text) as whole
' minutes since VBA's day 0.
Private Function TryMinutes(v As Variant, ByRef outMin As Long) As Boolean
    On Error GoTo Fail
    Dim d As Double
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim(v) = "" Then Exit Function
        d = CDbl(CDate(v))
    Else
        d = CDbl(v)
    End If
    outMin = CLng(Round(d * 1440#, 0))
    TryMinutes = True
    Exit Function
Fail:
    TryMinutes = False
End Function


Private Function ApIndex(ByVal code As String) As Long
    Dim idx As Long
    idx = 0
    On Error Resume Next
    idx = apLookup.Item(code)
    On Error GoTo 0
    If idx = 0 Then
        nAp = nAp + 1
        If nAp > UBound(apCode) Then
            ReDim Preserve apCode(1 To 2 * nAp)
            ReDim Preserve apCity(1 To 2 * nAp)
        End If
        apCode(nAp) = code
        apCity(nAp) = CityIndex(AirportCity(code))
        apLookup.Add nAp, code
        idx = nAp
    End If
    ApIndex = idx
End Function


Private Function LoadFlights() As Boolean
    Dim ws As Worksheet, data As Variant
    Dim lastRow As Long, lastCol As Long, c As Long, i As Long, n As Long
    Dim colO As Long, colD As Long, colDep As Long, colArr As Long, colDur As Long
    Dim o As String, d As String, depM As Long, arrM As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(DATA_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Could not find a sheet named '" & DATA_SHEET & "'." & vbCrLf & _
               "Update the DATA_SHEET constant at the top of the module.", vbCritical
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then
        MsgBox "'" & DATA_SHEET & "' has no data rows below the header.", vbCritical
        Exit Function
    End If
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value

    For c = 1 To UBound(data, 2)
        Select Case UCase(Trim(CStr(data(1, c))))
            Case "ORIGIN": colO = c
            Case "DESTINATION": colD = c
            Case "DEPARTURE_DATETIME": colDep = c
            Case "ARRIVAL_DATETIME": colArr = c
            Case "FLIGHT_DURATION_MINUTES": colDur = c   ' optional, display only
        End Select
    Next c
    If colO = 0 Or colD = 0 Or colDep = 0 Or colArr = 0 Then
        MsgBox "Could not find all required headers (ORIGIN, DESTINATION, DEPARTURE_DATETIME, " & _
               "ARRIVAL_DATETIME) in row 1 of '" & DATA_SHEET & "'.", vbCritical
        Exit Function
    End If

    n = UBound(data, 1) - 1
    nAp = 0
    ReDim apCode(1 To 256)
    ReDim apCity(1 To 256)
    Set apLookup = New Collection
    nF = 0
    ReDim fOrig(1 To n): ReDim fDest(1 To n)
    ReDim fDep(1 To n): ReDim fArr(1 To n)
    ReDim fDur(1 To n): ReDim fRow(1 To n)

    For i = 2 To UBound(data, 1)
        If Not IsEmpty(data(i, colO)) And Not IsEmpty(data(i, colD)) Then
            o = UCase(Trim(CStr(data(i, colO))))
            d = UCase(Trim(CStr(data(i, colD))))
            If o <> "" And d <> "" And o <> d Then
                If TryMinutes(data(i, colDep), depM) And TryMinutes(data(i, colArr), arrM) Then
                    nF = nF + 1
                    fOrig(nF) = ApIndex(o)
                    fDest(nF) = ApIndex(d)
                    fDep(nF) = depM
                    fArr(nF) = arrM
                    fRow(nF) = i
                    fDur(nF) = -1
                    If colDur > 0 Then
                        If IsNumeric(data(i, colDur)) Or IsDate(data(i, colDur)) Then
                            On Error Resume Next
                            fDur(nF) = CLng(CDbl(data(i, colDur)))
                            On Error GoTo 0
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If nF = 0 Then
        MsgBox "No usable flight rows were found on '" & DATA_SHEET & "'.", vbCritical
        Exit Function
    End If
    LoadFlights = True
End Function


' ===================== DOOR-TO-DOOR LEGS =====================

' Sorts order(lo..hi) by fDep ascending.
Private Sub SortOrderByDep(ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Long, tmp As Long
    Do While lo < hi
        i = lo: j = hi
        pivot = fDep(order((lo + hi) \ 2))
        Do While i <= j
            Do While fDep(order(i)) < pivot: i = i + 1: Loop
            Do While fDep(order(j)) > pivot: j = j - 1: Loop
            If i <= j Then
                tmp = order(i): order(i) = order(j): order(j) = tmp
                i = i + 1: j = j - 1
            End If
        Loop
        ' recurse into the smaller half, loop on the larger (keeps the stack shallow)
        If j - lo < hi - i Then
            If lo < j Then SortOrderByDep lo, j
            lo = i
        Else
            If i < hi Then SortOrderByDep i, hi
            hi = j
        End If
    Loop
End Sub


' For every pair of consecutive destinations on any route, builds the list
' of door-to-door options (see the ALGORITHM notes at the top).
'
' For one destination city B at a time, flights are scanned latest departure
' first. For flight i, best(i) = the earliest arrival at B possible if you
' board flight i: its own arrival if it lands in B, otherwise the best
' option among flights leaving its arrival airport at least
' MIN_CONNECTION_MINUTES after it lands. Those later flights were all
' scanned already, and each airport keeps a running "best from here if you
' leave at or after time t" list (lDep/lBest/lVia, in scan order, so
' departures are decreasing and a binary search finds time t). Flights
' landing at any other main destination are never used as connections.
'
' (If a short westbound hop lands earlier on the local clock than it left,
' a connection right after it can be missed, never invented -- the result
' is always a real, valid itinerary.)
Private Sub BuildCompositeLegs(ByRef missingPairs As String)
    Dim i As Long, k As Long, r As Long, j As Long, p As Long, B As Long
    Dim apStart() As Long, apFill() As Long, apTot() As Long
    Dim lDep() As Long, lBest() As Long, lVia() As Long
    Dim best() As Long, via() As Long
    Dim o As Long, d As Long, v As Long, vv As Long, t As Long
    Dim lo As Long, hi As Long, m As Long, blk As Long, cnt As Long, pos As Long
    Dim usedAsTarget() As Boolean

    ' --- pairs used by the routes ---
    ReDim pairIdx(1 To N_CITIES, 1 To N_CITIES)
    ReDim usedAsTarget(1 To N_CITIES)
    nPairs = 0
    For r = 1 To N_ROUTES
        For j = 0 To routeLegs(r) - 1
            If pairIdx(routeCity(r, j), routeCity(r, j + 1)) = 0 Then
                nPairs = nPairs + 1
                pairIdx(routeCity(r, j), routeCity(r, j + 1)) = nPairs
            End If
            usedAsTarget(routeCity(r, j + 1)) = True
        Next j
    Next r
    ReDim pFrom(1 To nPairs): ReDim pTo(1 To nPairs)
    ReDim pStart(1 To nPairs): ReDim pCount(1 To nPairs)
    For o = 1 To N_CITIES
        For d = 1 To N_CITIES
            If pairIdx(o, d) > 0 Then pFrom(pairIdx(o, d)) = o: pTo(pairIdx(o, d)) = d
        Next d
    Next o

    ' --- flights sorted by departure ---
    ReDim order(1 To nF)
    For i = 1 To nF: order(i) = i: Next i
    SortOrderByDep 1, nF

    ' --- one slot block per departure airport ---
    ReDim apStart(1 To nAp): ReDim apFill(1 To nAp): ReDim apTot(1 To nAp)
    For i = 1 To nF: apTot(fOrig(i)) = apTot(fOrig(i)) + 1: Next i
    pos = 1
    For o = 1 To nAp: apStart(o) = pos: pos = pos + apTot(o): Next o
    ReDim lDep(1 To nF): ReDim lBest(1 To nF): ReDim lVia(1 To nF)
    ReDim best(1 To nF): ReDim via(1 To nF)

    nC = 0: nSeg = 0
    ReDim cDep(1 To 1024): ReDim cArr(1 To 1024)
    ReDim cSegStart(1 To 1024): ReDim cSegCount(1 To 1024)
    ReDim segFlight(1 To 4096)

    For B = 1 To N_CITIES
        If usedAsTarget(B) Then
            For o = 1 To nAp: apFill(o) = 0: Next o
            For k = nF To 1 Step -1
                i = order(k)
                best(i) = 0
                o = fOrig(i): d = fDest(i)
                If apCity(o) <> B Then
                    v = 0
                    If apCity(d) = B Then
                        v = fArr(i): via(i) = 0
                    ElseIf apCity(d) = 0 Then
                        cnt = apFill(d)
                        If cnt > 0 Then
                            blk = apStart(d)
                            t = fArr(i) + MIN_CONNECTION_MINUTES
                            lo = 0: hi = cnt   ' -> number of entries with departure >= t
                            Do While lo < hi
                                m = (lo + hi) \ 2
                                If lDep(blk + m) >= t Then lo = m + 1 Else hi = m
                            Loop
                            If lo > 0 Then v = lBest(blk + lo - 1): via(i) = lVia(blk + lo - 1)
                        End If
                    End If
                    If v > 0 Then
                        best(i) = v
                        pos = apStart(o) + apFill(o)
                        lDep(pos) = fDep(i)
                        If apFill(o) > 0 Then
                            If lBest(pos - 1) <= v Then
                                lBest(pos) = lBest(pos - 1): lVia(pos) = lVia(pos - 1)
                            Else
                                lBest(pos) = v: lVia(pos) = i
                            End If
                        Else
                            lBest(pos) = v: lVia(pos) = i
                        End If
                        apFill(o) = apFill(o) + 1
                    End If
                End If
            Next k

            ' --- extract each pair (A -> B) ---
            For p = 1 To nPairs
                If pTo(p) = B Then Call ExtractPair(p, best, via)
            Next p
        End If
    Next B

    missingPairs = ""
    For p = 1 To nPairs
        If pCount(p) = 0 Then
            If missingPairs <> "" Then missingPairs = missingPairs & ", "
            missingPairs = missingPairs & CityCode(pFrom(p)) & " -> " & CityCode(pTo(p))
        End If
    Next p
End Sub


' Keeps only the options for pair p that nothing beats (no other option
' leaves at the same time or later and arrives at the same time or earlier),
' then stores them sorted by departure along with their flight chains.
Private Sub ExtractPair(ByVal p As Long, best() As Long, via() As Long)
    Dim k As Long, i As Long, nk As Long, runMin As Long, q As Long, f As Long
    Dim kFirst() As Long
    ReDim kFirst(1 To nF)
    nk = 0
    For k = nF To 1 Step -1          ' latest departure first
        i = order(k)
        If apCity(fOrig(i)) = pFrom(p) And best(i) > 0 Then
            If nk = 0 Then
                nk = 1: kFirst(1) = i: runMin = best(i)
            ElseIf best(i) < runMin Then
                If fDep(kFirst(nk)) = fDep(i) Then
                    kFirst(nk) = i         ' same departure, earlier arrival: replace
                Else
                    nk = nk + 1: kFirst(nk) = i
                End If
                runMin = best(i)
            End If
        End If
    Next k

    pStart(p) = nC + 1
    pCount(p) = nk
    For q = nk To 1 Step -1          ' store earliest departure first
        i = kFirst(q)
        nC = nC + 1
        If nC > UBound(cDep) Then
            ReDim Preserve cDep(1 To 2 * nC): ReDim Preserve cArr(1 To 2 * nC)
            ReDim Preserve cSegStart(1 To 2 * nC): ReDim Preserve cSegCount(1 To 2 * nC)
        End If
        cDep(nC) = fDep(i)
        cArr(nC) = best(i)
        cSegStart(nC) = nSeg + 1
        cSegCount(nC) = 0
        f = i
        Do While f > 0
            nSeg = nSeg + 1
            If nSeg > UBound(segFlight) Then
                ReDim Preserve segFlight(1 To 2 * nSeg)
            End If
            segFlight(nSeg) = f
            cSegCount(nC) = cSegCount(nC) + 1
            If apCity(fDest(f)) = pTo(p) Then Exit Do
            f = via(f)
        Loop
    Next q
End Sub


' First option of pair p leaving at or after minute t (it also arrives
' earliest, since arrivals increase with departures), or 0 if none.
Private Function FirstOptionAtOrAfter(ByVal p As Long, ByVal t As Long) As Long
    Dim lo As Long, hi As Long, m As Long
    lo = pStart(p): hi = pStart(p) + pCount(p)
    Do While lo < hi
        m = (lo + hi) \ 2
        If cDep(m) < t Then lo = m + 1 Else hi = m
    Loop
    If lo < pStart(p) + pCount(p) Then FirstOptionAtOrAfter = lo Else FirstOptionAtOrAfter = 0
End Function


' ===================== SEARCH =====================

' Walks route r from starting option startC, requiring every stay to be at
' least its minimum (+ S at optimized cities) and, where a city has open
' hours, the buffered stay to overlap them. Returns False if the chain
' can't be completed.
Private Function Propagate(ByVal r As Long, ByVal startC As Long, ByVal S As Long, _
                           ByRef chosen() As Long, ByRef total As Long) As Boolean
    Dim j As Long, idx As Long, req As Long, arr As Long, minDep As Long
    Dim city As String, afterArr As Long, beforeDep As Long, openAt As Long

    ReDim chosen(0 To routeLegs(r) - 1)
    chosen(0) = startC
    arr = cArr(startC)
    For j = 1 To routeLegs(r) - 1
        city = CityCode(routeCity(r, j))
        req = RequiredMinutes(city)
        If IsOptimizedCity(city) Then req = req + S
        minDep = arr + req
        If OpenBuffers(city, afterArr, beforeDep) Then
            openAt = EarliestOpenMin(city, arr + afterArr)
            If openAt < 0 Then Exit Function
            If openAt + beforeDep + 1 > minDep Then minDep = openAt + beforeDep + 1
        End If
        idx = FirstOptionAtOrAfter(pairIdx(routeCity(r, j), routeCity(r, j + 1)), minDep)
        If idx = 0 Then Exit Function
        chosen(j) = idx
        arr = cArr(idx)
    Next j
    total = arr - cDep(startC) - TZCorrectionMinutes(r)
    Propagate = True
End Function


' A fresh copy of a Long array (explicit, so results never share storage
' with the scratch array Propagate keeps re-filling).
Private Function CopyLongs(src() As Long) As Variant
    Dim dst() As Long, i As Long
    ReDim dst(LBound(src) To UBound(src))
    For i = LBound(src) To UBound(src): dst(i) = src(i): Next i
    CopyLongs = dst
End Function

' For one route and starting option: the largest slack floor S that still
' completes under budget, and that itinerary's min/total slack and time.
Private Function BestForStart(ByVal r As Long, ByVal startC As Long, ByVal budget As Long, _
                              ByRef outChosen() As Long, ByRef outMin As Long, _
                              ByRef outTot As Long, ByRef outTime As Long) As Boolean
    Dim lo As Long, hi As Long, mid As Long, total As Long
    Dim chosen() As Long, bestChosen() As Long, bestTotal As Long

    If Not Propagate(r, startC, 0, chosen, total) Then Exit Function
    If Not (total < budget) Then Exit Function
    bestChosen = CopyLongs(chosen): bestTotal = total

    lo = 0: hi = 1
    Do While hi <= MAX_S_CAP
        If Propagate(r, startC, hi, chosen, total) Then
            If total < budget Then
                lo = hi: bestChosen = CopyLongs(chosen): bestTotal = total: hi = hi * 2
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
        If Propagate(r, startC, mid, chosen, total) Then
            If total < budget Then
                lo = mid: bestChosen = CopyLongs(chosen): bestTotal = total
            Else
                hi = mid
            End If
        Else
            hi = mid
        End If
    Loop

    outChosen = CopyLongs(bestChosen)
    outTime = bestTotal
    Dim j As Long, sl As Long, first As Boolean, city As String
    first = True: outTot = 0: outMin = 0
    For j = 1 To routeLegs(r) - 1
        city = CityCode(routeCity(r, j))
        If IsOptimizedCity(city) Then
            sl = cDep(outChosen(j)) - cArr(outChosen(j - 1)) - RequiredMinutes(city)
            outTot = outTot + sl
            If first Then
                outMin = sl: first = False
            ElseIf sl < outMin Then
                outMin = sl
            End If
        End If
    Next j
    BestForStart = True
End Function


Private Sub SearchAllRoutes(ByVal budget As Long)
    Dim r As Long, j As Long, p0 As Long, c As Long, ok As Boolean
    Dim chosen() As Long, mn As Long, tt As Long, tm As Long

    nCand = 0
    ReDim candRoute(1 To 1024): ReDim candMin(1 To 1024)
    ReDim candTot(1 To 1024): ReDim candTime(1 To 1024)
    ReDim candChosen(1 To 1024)

    For r = 1 To N_ROUTES
        ok = routeOn(r)
        For j = 0 To routeLegs(r) - 1
            If pCount(pairIdx(routeCity(r, j), routeCity(r, j + 1))) = 0 Then ok = False
        Next j
        If ok Then
            Application.StatusBar = "Searching route " & r & " of " & N_ROUTES & "..."
            p0 = pairIdx(routeCity(r, 0), routeCity(r, 1))
            For c = pStart(p0) To pStart(p0) + pCount(p0) - 1
                If BestForStart(r, c, budget, chosen, mn, tt, tm) Then
                    nCand = nCand + 1
                    If nCand > UBound(candRoute) Then
                        ReDim Preserve candRoute(1 To 2 * nCand): ReDim Preserve candMin(1 To 2 * nCand)
                        ReDim Preserve candTot(1 To 2 * nCand): ReDim Preserve candTime(1 To 2 * nCand)
                        ReDim Preserve candChosen(1 To 2 * nCand)
                    End If
                    candRoute(nCand) = r
                    candMin(nCand) = mn
                    candTot(nCand) = tt
                    candTime(nCand) = tm
                    candChosen(nCand) = CopyLongs(chosen)
                End If
            Next c
        End If
    Next r
End Sub


' ===================== RANKING AND GROUPING =====================

' True if candidate a ranks ahead of b: higher min slack, then higher total
' slack, then shorter trip, then found first.
Private Function CandBefore(ByVal a As Long, ByVal b As Long) As Boolean
    If candMin(a) <> candMin(b) Then CandBefore = candMin(a) > candMin(b): Exit Function
    If candTot(a) <> candTot(b) Then CandBefore = candTot(a) > candTot(b): Exit Function
    If candTime(a) <> candTime(b) Then CandBefore = candTime(a) < candTime(b): Exit Function
    CandBefore = a < b
End Function


Private Sub SortCandidates(idx() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Long, tmp As Long
    Do While lo < hi
        i = lo: j = hi
        pivot = idx((lo + hi) \ 2)
        Do While i <= j
            Do While CandBefore(idx(i), pivot): i = i + 1: Loop
            Do While CandBefore(pivot, idx(j)): j = j - 1: Loop
            If i <= j Then
                tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
                i = i + 1: j = j - 1
            End If
        Loop
        If j - lo < hi - i Then
            If lo < j Then SortCandidates idx, lo, j
            lo = i
        Else
            If i < hi Then SortCandidates idx, i, hi
            hi = j
        End If
    Loop
End Sub


Private Sub RankCandidates(ByRef rankIdx() As Long)
    Dim i As Long
    ReDim rankIdx(1 To nCand)
    For i = 1 To nCand: rankIdx(i) = i: Next i
    SortCandidates rankIdx, 1, nCand
End Sub


' Route number plus every airport-to-airport hop in order: two itineraries
' with the same key are "the same route" (possibly on a different date or at
' slightly different times).
Private Function GroupKey(ByVal cand As Long) As String
    Dim ch As Variant, j As Long, s As Long, f As Long, key As String
    ch = candChosen(cand)
    key = "R" & candRoute(cand)
    For j = 0 To UBound(ch)
        For s = cSegStart(ch(j)) To cSegStart(ch(j)) + cSegCount(ch(j)) - 1
            f = segFlight(s)
            key = key & "|" & apCode(fOrig(f)) & ">" & apCode(fDest(f))
        Next s
    Next j
    GroupKey = key
End Function


' Same as GroupKey but also with every departure/arrival time measured from
' midnight of the first day: equal only if it's the identical itinerary
' shifted by whole days.
Private Function ExactKey(ByVal cand As Long) As String
    Dim ch As Variant, j As Long, s As Long, f As Long, key As String, day0 As Long
    ch = candChosen(cand)
    day0 = (cDep(ch(0)) \ 1440) * 1440
    key = "R" & candRoute(cand)
    For j = 0 To UBound(ch)
        For s = cSegStart(ch(j)) To cSegStart(ch(j)) + cSegCount(ch(j)) - 1
            f = segFlight(s)
            key = key & "|" & apCode(fOrig(f)) & ">" & apCode(fDest(f)) & "@" & (fDep(f) - day0) & "-" & (fArr(f) - day0)
        Next s
    Next j
    ExactKey = key
End Function


Private Function StartDay(ByVal cand As Long) As Long
    Dim ch As Variant
    ch = candChosen(cand)
    StartDay = cDep(ch(0)) \ 1440
End Function


' Walks the ranked list; the first itinerary of each new group becomes a
' shown option (up to TOP_N). Later members of a shown group add their start
' date to that option's note (once per date).
Private Sub GroupTopOptions(rankIdx() As Long, ByRef shown() As Long, ByRef alsoNote() As String, ByRef nShown As Long)
    Dim lookup As New Collection
    Dim i As Long, c As Long, g As Long, key As String, dayKey As String
    Dim seenDays() As String, exact() As String

    ReDim shown(1 To TOP_N)
    ReDim alsoNote(1 To TOP_N)
    ReDim seenDays(1 To TOP_N)
    ReDim exact(1 To TOP_N)
    nShown = 0

    For i = 1 To nCand
        c = rankIdx(i)
        key = GroupKey(c)
        g = 0
        On Error Resume Next
        g = lookup.Item(key)
        On Error GoTo 0
        If g = 0 Then
            If nShown < TOP_N Then
                nShown = nShown + 1
                lookup.Add nShown, key
                shown(nShown) = c
                exact(nShown) = ExactKey(c)
                seenDays(nShown) = "|" & StartDay(c) & "|"
            End If
        Else
            dayKey = "|" & StartDay(c) & "|"
            If InStr(seenDays(g), dayKey) = 0 Then
                seenDays(g) = seenDays(g) & StartDay(c) & "|"
                If alsoNote(g) <> "" Then alsoNote(g) = alsoNote(g) & ";  "
                alsoNote(g) = alsoNote(g) & Format(CDate(StartDay(c)), "ddd mmm d")
                If ExactKey(c) = exact(g) Then
                    If candMin(c) <> candMin(shown(g)) Then
                        alsoNote(g) = alsoNote(g) & " (min slack " & FormatDuration(CDbl(candMin(c))) & ")"
                    End If
                Else
                    alsoNote(g) = alsoNote(g) & " (other flight times: min slack " & _
                                  FormatDuration(CDbl(candMin(c))) & ", trip " & FormatDuration(CDbl(candTime(c))) & ")"
                End If
            End If
        End If
    Next i
End Sub


' ===================== OUTPUT =====================

Private Function TravelType(ByVal f As Long) As String
    If apCode(fOrig(f)) = "MPU" Or apCode(fDest(f)) = "MPU" Then TravelType = "Train" Else TravelType = "Flight"
End Function


Private Sub WriteTopSheet(shown() As Long, alsoNote() As String, ByVal nShown As Long, _
                          ByVal budget As Long, ByVal missingPairs As String)
    Dim outWs As Worksheet
    Dim g As Long, col0 As Long, lastCol As Long, maxRow As Long, endRow As Long
    Dim longestNote As Long

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(OutSheetName()).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set outWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    outWs.Name = OutSheetName()
    ActiveWindow.DisplayGridlines = False

    lastCol = nShown * (CARD_COLS + 1) - 1
    With outWs.Range(outWs.Cells(1, 1), outWs.Cells(1, lastCol))
        .Merge
        .Value = "TOP " & nShown & " ROUTES  --  " & ScopeText() & ", every start date  " & _
                 "|  ranked by worst-case slack, then total slack, then shortest trip"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
    End With
    With outWs.Range(outWs.Cells(2, 1), outWs.Cells(2, lastCol))
        .Merge
        .Value = nF & " flights/trains from '" & DATA_SHEET & "'  |  " & nCand & _
                 " complete itineraries under the " & FormatDuration(CDbl(budget)) & " budget  |  every connection >= " & _
                 MIN_CONNECTION_MINUTES & " min at the same airport  |  times are local" & _
                 IIf(missingPairs <> "", "  |  no connections found for: " & missingPairs, "")
        .Font.Italic = True
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
    End With

    maxRow = 0
    For g = 1 To nShown
        col0 = (g - 1) * (CARD_COLS + 1) + 1
        endRow = WriteCard(outWs, g, shown(g), alsoNote(g), col0, 4, budget)
        If endRow > maxRow Then maxRow = endRow
        If Len(alsoNote(g)) > longestNote Then longestNote = Len(alsoNote(g))
        outWs.Columns(col0 + CARD_COLS).ColumnWidth = 3
    Next g

    ' "Also works starting" row (row 11): tall enough for the longest note
    outWs.Rows(11).RowHeight = Application.WorksheetFunction.Min(409, 15 * (1 + longestNote \ 140))
    outWs.Range("A4").Select
End Sub


' Writes one option's card with its top-left corner at (top, col0) and
' returns the last row used.
Private Function WriteCard(outWs As Worksheet, ByVal rank As Long, ByVal cand As Long, ByVal note As String, _
                           ByVal col0 As Long, ByVal topRow As Long, ByVal budget As Long) As Long
    Dim r As Long, ch As Variant, j As Long, s As Long, f As Long, prevF As Long
    Dim rt As Long, stepNo As Long, city As String, req As Long, sl As Long
    Dim stayStart As Long, stayEnd As Long, afterArr As Long, beforeDep As Long
    Dim txt As String, bindingRow As Long, bindingSlack As Long, haveBinding As Boolean
    Dim firstDep As Long, lastArr As Long

    rt = candRoute(cand)
    ch = candChosen(cand)
    firstDep = cDep(ch(0))
    lastArr = cArr(ch(UBound(ch)))

    ' --- column widths ---
    outWs.Columns(col0 + 0).ColumnWidth = 5     ' step
    outWs.Columns(col0 + 1).ColumnWidth = 11    ' type
    outWs.Columns(col0 + 2).ColumnWidth = 7     ' from
    outWs.Columns(col0 + 3).ColumnWidth = 7     ' to
    outWs.Columns(col0 + 4).ColumnWidth = 22    ' start
    outWs.Columns(col0 + 5).ColumnWidth = 22    ' end
    outWs.Columns(col0 + 6).ColumnWidth = 11    ' duration
    outWs.Columns(col0 + 7).ColumnWidth = 11    ' minimum
    outWs.Columns(col0 + 8).ColumnWidth = 11    ' slack
    outWs.Columns(col0 + 9).ColumnWidth = 55    ' notes

    ' --- title ---
    With outWs.Range(outWs.Cells(topRow, col0), outWs.Cells(topRow, col0 + CARD_COLS - 1))
        .Merge
        .Value = "#" & rank & "   Route " & rt
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = IIf(rank = 1, RGB(46, 125, 50), RGB(31, 78, 121))
    End With
    With outWs.Range(outWs.Cells(topRow + 1, col0), outWs.Cells(topRow + 1, col0 + CARD_COLS - 1))
        .Merge
        .Value = RouteLabel(rt)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With

    ' --- summary ---
    Call SummaryRow(outWs, topRow + 2, col0, "Worst-case slack", FormatDuration(CDbl(candMin(cand))) & _
                    "  (" & candMin(cand) & " min) -- every counted destination gets at least this much extra time")
    Call SummaryRow(outWs, topRow + 3, col0, "Total slack", FormatDuration(CDbl(candTot(cand))) & _
                    "  (" & candTot(cand) & " min) -- extra time summed over all counted destinations")
    Call SummaryRow(outWs, topRow + 4, col0, "Total trip time", FormatDuration(CDbl(candTime(cand))) & _
                    "  (time-zone adjusted; budget " & FormatDuration(CDbl(budget)) & ")")
    Call SummaryRow(outWs, topRow + 5, col0, "Trip starts", Format(MinToDate(firstDep), "ddd mmm d, yyyy h:mm AM/PM") & _
                    "  from " & CityCode(routeCity(rt, 0)))
    Call SummaryRow(outWs, topRow + 6, col0, "Trip ends", Format(MinToDate(lastArr), "ddd mmm d, yyyy h:mm AM/PM") & _
                    "  in " & CityCode(routeCity(rt, routeLegs(rt))))
    Call SummaryRow(outWs, topRow + 7, col0, "Also works starting", IIf(note = "", "(no other start dates)", note))
    outWs.Cells(topRow + 7, col0 + 2).WrapText = True
    outWs.Cells(topRow + 7, col0 + 2).VerticalAlignment = xlTop
    outWs.Cells(topRow + 7, col0).VerticalAlignment = xlTop

    ' --- step table ---
    r = topRow + 9
    outWs.Cells(r, col0).Resize(1, CARD_COLS).Value = Array("Step", "Type", "From", "To", "Start (local)", _
        "End (local)", "Duration", "Min Required", "Slack", "Notes")
    With outWs.Cells(r, col0).Resize(1, CARD_COLS)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With

    stepNo = 0
    For j = 0 To UBound(ch)
        prevF = 0
        For s = cSegStart(ch(j)) To cSegStart(ch(j)) + cSegCount(ch(j)) - 1
            f = segFlight(s)
            If prevF > 0 Then
                ' connection at an in-between airport
                r = r + 1: stepNo = stepNo + 1
                Call StepRow(outWs, r, col0, stepNo, "Connection", apCode(fDest(prevF)), apCode(fOrig(f)), _
                             fArr(prevF), fDep(f), MIN_CONNECTION_MINUTES, "", "Layover between flights")
                outWs.Cells(r, col0).Resize(1, CARD_COLS).Interior.Color = RGB(242, 242, 242)
                outWs.Cells(r, col0).Resize(1, CARD_COLS).Font.Italic = True
            End If
            r = r + 1: stepNo = stepNo + 1
            Call StepRow(outWs, r, col0, stepNo, TravelType(f), apCode(fOrig(f)), apCode(fDest(f)), _
                         fDep(f), fArr(f), -1, "", DATA_SHEET & " row " & fRow(f))
            If fDur(f) >= 0 Then outWs.Cells(r, col0 + 6).Value = FormatDuration(CDbl(fDur(f)))
            prevF = f
        Next s

        If j < UBound(ch) Then
            ' stay at a destination
            city = CityCode(routeCity(rt, j + 1))
            req = RequiredMinutes(city)
            stayStart = cArr(ch(j))
            stayEnd = cDep(ch(j + 1))
            sl = stayEnd - stayStart - req
            txt = ""
            If OpenBuffers(city, afterArr, beforeDep) Then
                txt = "Open-hours time available: " & OpenOverlapText(city, stayStart + afterArr, stayEnd - beforeDep) & _
                      "  (buffers " & FormatDuration(CDbl(afterArr)) & " after landing, " & _
                      FormatDuration(CDbl(beforeDep)) & " before leaving)"
            End If
            f = segFlight(cSegStart(ch(j + 1)))
            If apCode(fDest(prevF)) <> apCode(fOrig(f)) Then
                txt = "Arrive " & apCode(fDest(prevF)) & ", leave from " & apCode(fOrig(f)) & IIf(txt = "", "", ".  ") & txt
            End If
            If Not IsOptimizedCity(city) Then txt = "Not counted toward the ranking." & IIf(txt = "", "", "  ") & txt

            r = r + 1: stepNo = stepNo + 1
            Call StepRow(outWs, r, col0, stepNo, "STAY", apCode(fDest(prevF)), apCode(fOrig(f)), stayStart, stayEnd, req, CStr(sl), txt)
            outWs.Cells(r, col0).Resize(1, CARD_COLS).Font.Bold = True
            If IsOptimizedCity(city) Then
                outWs.Cells(r, col0).Resize(1, CARD_COLS).Interior.Color = RGB(226, 239, 218)
                If Not haveBinding Or sl < bindingSlack Then
                    haveBinding = True: bindingSlack = sl: bindingRow = r
                End If
            Else
                outWs.Cells(r, col0).Resize(1, CARD_COLS).Interior.Color = RGB(217, 217, 217)
            End If
        End If
    Next j

    If haveBinding Then
        outWs.Cells(bindingRow, col0).Resize(1, CARD_COLS).Interior.Color = RGB(255, 230, 153)
        outWs.Cells(bindingRow, col0 + 9).Value = Trim("TIGHTEST STAY (sets the worst-case slack).  " & outWs.Cells(bindingRow, col0 + 9).Value)
    End If

    With outWs.Range(outWs.Cells(topRow + 9, col0), outWs.Cells(r, col0 + CARD_COLS - 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    outWs.Range(outWs.Cells(topRow + 10, col0 + 9), outWs.Cells(r, col0 + 9)).WrapText = True
    outWs.Range(outWs.Cells(topRow + 10, col0), outWs.Cells(r, col0 + CARD_COLS - 1)).VerticalAlignment = xlTop

    WriteCard = r
End Function


Private Sub SummaryRow(outWs As Worksheet, ByVal r As Long, ByVal col0 As Long, ByVal labelText As String, ByVal valueText As String)
    With outWs.Range(outWs.Cells(r, col0), outWs.Cells(r, col0 + 1))
        .Merge
        .Value = labelText
        .Font.Bold = True
    End With
    With outWs.Range(outWs.Cells(r, col0 + 2), outWs.Cells(r, col0 + CARD_COLS - 1))
        .Merge
        .NumberFormat = "@"          ' keep text like "Fri Dec 4" from being turned into a date
        .Value = valueText
        .HorizontalAlignment = xlLeft
    End With
End Sub


' One row of the step table. minReq < 0 leaves Min Required blank.
Private Sub StepRow(outWs As Worksheet, ByVal r As Long, ByVal col0 As Long, ByVal stepNo As Long, _
                    ByVal kind As String, ByVal fromCode As String, ByVal toCode As String, _
                    ByVal startM As Long, ByVal endM As Long, ByVal minReq As Long, _
                    ByVal slackText As String, ByVal note As String)
    outWs.Cells(r, col0 + 0).Value = stepNo
    outWs.Cells(r, col0 + 1).Value = kind
    outWs.Cells(r, col0 + 2).Value = fromCode
    outWs.Cells(r, col0 + 3).Value = toCode
    outWs.Cells(r, col0 + 4).Value = MinToDate(startM)
    outWs.Cells(r, col0 + 4).NumberFormat = "ddd mmm d, h:mm AM/PM"
    outWs.Cells(r, col0 + 5).Value = MinToDate(endM)
    outWs.Cells(r, col0 + 5).NumberFormat = "ddd mmm d, h:mm AM/PM"
    outWs.Cells(r, col0 + 6).Value = FormatDuration(CDbl(endM - startM))
    If minReq >= 0 Then outWs.Cells(r, col0 + 7).Value = FormatDuration(CDbl(minReq))
    If slackText <> "" Then outWs.Cells(r, col0 + 8).Value = FormatDuration(CDbl(CLng(slackText)))
    outWs.Cells(r, col0 + 9).Value = note
    outWs.Cells(r, col0).Resize(1, 4).HorizontalAlignment = xlCenter
End Sub
