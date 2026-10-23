Attribute VB_Name = "modFindSpecialPath"
Option Explicit

'===============================================================================
' FindSpecialPath
'
' A companion to FindOptimalPath (modFindOptimalPath) -- it doesn't replace
' it, and the two modules can sit side by side in the same workbook.
'
' Same goal and rules as FindOptimalPath (maximize the worst-case extra time
' across the counted destinations, then the total extra time, then the
' shorter trip; same minimum stays, open hours, 60-minute connections, AMM /
' ETM / AQJ and PEK / PKX groups, and the DEL -> TAJ -> DEL side trip), with
' two differences:
'
'   1. These flights are REQUIRED (see PinnedLegText):
'        FCO -> IST  Fri Dec 11 2026  7:25 PM   (FULL FLIGHT DATA row 30331)
'        IST -> AMM  Sat Dec 12 2026  2:10 AM   (row 36510)
'        AMM -> DXB  Sat Dec 12 2026 11:25 PM   (row 3092)
'        DXB -> DEL  Sun Dec 13 2026  5:05 AM   (row 24997)
'        DEL -> PEK  Sun Dec 13 2026 10:50 PM   (row 23227)
'      The macro chooses everything else: the South America / Cancun part
'      and the flights to FCO, plus the DEL -> TAJ -> DEL transfers.
'
'   2. The four orders for the start of the trip (the same as routes 1-4 of
'      the main macro, which also visit MPU through CUZ):
'        Order 1: MPU -> CUZ -> CUN -> GIG -> FCO -> AMM -> DEL -> PEK
'        Order 2: MPU -> CUZ -> GIG -> CUN -> FCO -> AMM -> DEL -> PEK
'        Order 3: GIG -> CUZ -> MPU -> CUZ -> CUN -> FCO -> AMM -> DEL -> PEK
'        Order 4: CUN -> CUZ -> MPU -> CUZ -> GIG -> FCO -> AMM -> DEL -> PEK
'      and a total time limit of BUDGET (5d 16h).
'
' ALGORITHM
'   Door-to-door options between destinations are built exactly as in the
'   main macro, then the three required legs are replaced by the required
'   flights. Because the end of the trip is fixed, the number of possible
'   itineraries is small, so instead of the main macro's per-start shortcut
'   this tries EVERY combination of options (SearchPinned): a backward pass
'   first marks which options can still reach the required flights, then a
'   depth-first search walks every complete itinerary under the time limit.
'   Every result is ranked; itineraries through the same airports in the
'   same order count as one option (the best one is shown, and the note says
'   how many other timings of it there are).
'
' OUTPUT
'   A sheet named OUT_SHEET ("Top 10 Special Routes") laid out exactly like
'   the main macro's, with the required flights marked REQUIRED.
'
' HOW TO USE
'   Alt+F11 -> File -> Import File... -> pick this .bas (keep
'   modFindOptimalPath as it is) -> run FindSpecialPath.
'===============================================================================

' --------------------------- CONFIGURATION ------------------------------------
Private Const DATA_SHEET As String = "FULL FLIGHT DATA"   ' <-- change if needed
Private Const OUT_SHEET As String = "Top 10 Special Routes"
Private Const TOP_N As Long = 10

Private Const MIN_CONNECTION_MINUTES As Long = 60   ' minimum layover between two flights at an in-between airport

Private Const DEL_BEFORE_TAJ_MINUTES As Long = 90   ' in DEL: landing -> DEL to TAJ departure
Private Const DEL_AFTER_TAJ_MINUTES As Long = 150   ' in DEL: back from TAJ -> next departure

Private Const BUDGET_DAYS As Long = 5
Private Const BUDGET_HOURS As Long = 16
Private Const BUDGET_MINUTES As Long = 0

Private Const N_ROUTES As Long = 4
Private Const N_PINNED As Long = 3       ' required legs -- see PinnedLegText
Private Const N_CITIES As Long = 9        ' MPU, CUZ, CUN, GIG, FCO, AMM, DEL, PEK, TAJ -- see CityCode
Private Const CARD_COLS As Long = 10      ' columns per output card (plus one blank spacer column)
' -------------------------------------------------------------------------------

' Airports, indexed 1..nAp. apCity is the main-destination city index
' (1..N_CITIES) for airports that belong to one, else 0. apAdj is the
' minutes added to that airport's local clock to put it on its city's
' clock (e.g. +60 for ETM -> Amman time); 0 for everything else.
Private nAp As Long
Private apCode() As String
Private apCity() As Long
Private apAdj() As Long
Private apLookup As Collection

' Individual flights/trains from DATA_SHEET, indexed 1..nF. Times are whole
' minutes since VBA's day 0: local time at that airport, except at a
' main-destination airport, where they're on that city's clock (see apAdj).
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
Private routeCity() As Long
Private routeLegs() As Long
Private routeOn() As Boolean
Private nRoutesOn As Long

' Every feasible (route, start) result, indexed 1..nCand.
Private nCand As Long
Private candRoute() As Long
Private candMin() As Long, candTot() As Long, candTime() As Long
Private candChosen() As Variant    ' each a (0 To legs-1) Long array of composite-leg indexes

' Scratch for the flight sort (sorted by fDep ascending).
Private order() As Long

' Required flights (pinnedFlag(f) = True) and the exhaustive search's state.
Private pinnedFlag() As Boolean
Private okFrom() As Long           ' okFrom(j, k): first option >= k on leg j that can still finish, else 0
Private dfsChosen() As Long        ' composite-leg index chosen on each leg so far
Private dfsBudget As Long
Private dfsMinEnd As Long          ' earliest possible final arrival for the current route


Public Sub FindSpecialPath()
    Dim budget As Long, t0 As Single
    Dim i As Long, nShown As Long
    Dim rankIdx() As Long
    Dim shown() As Long, alsoNote() As String
    Dim missingPairs As String
    Dim prevCalc As Long

    Call InitRoutes

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
    If Not ApplyPinnedFlights() Then GoTo Done

    Application.StatusBar = "Trying every combination for " & ScopeText() & "..."
    Call SearchPinned(budget)

    If nCand = 0 Then
        Application.Calculation = prevCalc
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "No itinerary on " & ScopeText() & " fits within the " & _
               FormatDuration(CDbl(budget)) & " budget while meeting every minimum stay, " & _
               "open-hours rule and the " & MIN_CONNECTION_MINUTES & "-minute connection minimum " & _
               "while making the required flights." & _
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
           "Best: Order " & candRoute(shown(1)) & " -- worst-case slack " & FormatDuration(CDbl(candMin(shown(1)))) & _
           ", trip time " & FormatDuration(CDbl(candTime(shown(1)))) & " (budget " & FormatDuration(CDbl(budget)) & ")." & vbCrLf & vbCrLf & _
           "The top " & nShown & " options are side by side on '" & OutSheetName() & "'." & vbCrLf & _
           "(" & Format(Timer - t0, "0.0") & " s)", vbInformation, "Find Special Path"
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
    MsgBox "FindSpecialPath stopped with an error: " & Err.Description, vbCritical
End Sub


' ===================== ROUTES AND CITIES =====================
Private Function RouteText(ByVal r As Long) As String
    Select Case r
        Case 1: RouteText = "MPU CUZ CUN GIG FCO AMM DEL PEK"
        Case 2: RouteText = "MPU CUZ GIG CUN FCO AMM DEL PEK"
        Case 3: RouteText = "GIG CUZ MPU CUZ CUN FCO AMM DEL PEK"
        Case 4: RouteText = "CUN CUZ MPU CUZ GIG FCO AMM DEL PEK"
    End Select
End Function


' The required flights, one door-to-door leg per entry: "ORIGIN DEST
' yyyy-mm-dd hh:mm" (local departure time), several flights joined by ";".
' Each leg must run between two consecutive destinations on the routes.
Private Function PinnedLegText(ByVal k As Long) As String
    Select Case k
        Case 1: PinnedLegText = "FCO IST 2026-12-11 19:25; IST AMM 2026-12-12 02:10"
        Case 2: PinnedLegText = "AMM DXB 2026-12-12 23:25; DXB DEL 2026-12-13 05:05"
        Case 3: PinnedLegText = "DEL PEK 2026-12-13 22:50"
    End Select
End Function


' The route as it's actually searched, with the DEL -> TAJ -> DEL side trip.
Private Function RouteLabel(ByVal r As Long) As String
    Dim j As Long, txt As String
    For j = 0 To routeLegs(r)
        If j > 0 Then txt = txt & " -> "
        txt = txt & CityCode(routeCity(r, j))
    Next j
    RouteLabel = txt
End Function


Private Function CityCode(ByVal c As Long) As String
    CityCode = Choose(c, "MPU", "CUZ", "CUN", "GIG", "FCO", "AMM", "DEL", "PEK", "TAJ")
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
        Case "PKX": AirportCity = "PEK"          ' Beijing Daxing counts as Beijing
        Case "ETM", "AQJ": AirportCity = "AMM"   ' Eilat Ramon and Aqaba count as Amman
        Case Else: AirportCity = code
    End Select
End Function


' UTC offset of an airport in a destination group, when it differs from its
' city's (UTCOffsetMinutes). ETM is in Israel (UTC+2 in Nov/Dec), an hour
' behind Amman; AQJ and PKX share their city's clock.
Private Function AirportUTCOffset(ByVal code As String) As Long
    Select Case code
        Case "ETM": AirportUTCOffset = 2 * 60
        Case Else: AirportUTCOffset = UTCOffsetMinutes(AirportCity(code))
    End Select
End Function


' A city-clock time at airport ap, back on that airport's own clock.
Private Function LocalMin(ByVal t As Long, ByVal ap As Long) As Long
    LocalMin = t - apAdj(ap)
End Function


' e.g. "  (ETM clocks are 1h 0m behind AMM)", or "" if both airports share
' their city's clock.
Private Function ClockNote(ByVal ap1 As Long, ByVal ap2 As Long) As String
    Dim a As Variant, txt As String
    For Each a In Array(ap1, ap2)
        If apAdj(a) <> 0 And InStr(txt, apCode(a) & " ") = 0 Then
            txt = txt & "  (" & apCode(a) & " clocks are " & (Abs(apAdj(a)) \ 60) & "h " & (Abs(apAdj(a)) Mod 60) & "m " & _
                  IIf(apAdj(a) > 0, "behind ", "ahead of ") & CityCode(apCity(a)) & "; open hours are on " & CityCode(apCity(a)) & " time)"
        End If
    Next a
    ClockNote = txt
End Function


Private Sub InitRoutes()
    Dim r As Long, j As Long, n As Long, parts() As String
    ReDim routeCity(1 To N_ROUTES, 0 To 30)
    ReDim routeLegs(1 To N_ROUTES)
    For r = 1 To N_ROUTES
        parts = Split(RouteText(r), " ")
        n = 0
        For j = 0 To UBound(parts)
            routeCity(r, n) = CityIndex(parts(j))
            ' passing through DEL means the TAJ side trip: DEL -> TAJ -> DEL
            If parts(j) = "DEL" And j > 0 And j < UBound(parts) Then
                routeCity(r, n + 1) = CityIndex("TAJ")
                routeCity(r, n + 2) = CityIndex("DEL")
                n = n + 2
            End If
            n = n + 1
        Next j
        routeLegs(r) = n - 1
    Next r
    ReDim routeOn(1 To N_ROUTES)
    For r = 1 To N_ROUTES: routeOn(r) = True: Next r
    nRoutesOn = N_ROUTES
End Sub


Private Function ScopeText() As String
    ScopeText = "all " & N_ROUTES & " route orders"
End Function


Private Function OutSheetName() As String
    OutSheetName = OUT_SHEET
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
        Case "DEL", "TAJ": UTCOffsetMinutes = 5 * 60 + 30
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
        Case "TAJ": RequiredMinutes = 1 * 60 + 15
        Case "PEK": RequiredMinutes = 3 * 60 + 0
        Case Else: RequiredMinutes = 0
    End Select
End Function


' Destinations whose slack counts toward the ranking. CUZ's 60-minute
' minimum is still enforced; its extra time just isn't counted.
Private Function IsOptimizedCity(city As String) As Boolean
    Select Case UCase(city)
        Case "MPU", "CUN", "GIG", "FCO", "AMM", "TAJ", "PEK": IsOptimizedCity = True
        Case Else: IsOptimizedCity = False   ' CUZ, and the DEL waits around the TAJ trip
    End Select
End Function


' Minimum time at stop j of route r. Same as RequiredMinutes, except the
' two DEL waits around the TAJ side trip.
Private Function StayRequired(ByVal r As Long, ByVal j As Long) As Long
    Dim city As String
    city = CityCode(routeCity(r, j))
    If city = "DEL" And CityCode(routeCity(r, j + 1)) = "TAJ" Then
        StayRequired = DEL_BEFORE_TAJ_MINUTES
    ElseIf city = "DEL" And CityCode(routeCity(r, j - 1)) = "TAJ" Then
        StayRequired = DEL_AFTER_TAJ_MINUTES
    Else
        StayRequired = RequiredMinutes(city)
    End If
End Function


' --------------------------- OPEN HOURS ---------------------------------------
' Buffers for each city that has open hours: the usable part of a stay
' starts afterArr minutes after you land and ends beforeDep minutes before
' your next departure. Returns False for a city with no open-hours rule.
Private Function OpenBuffers(ByVal city As String, ByRef afterArr As Long, ByRef beforeDep As Long) As Boolean
    OpenBuffers = True
    Select Case UCase(city)
        Case "TAJ": afterArr = 0:           beforeDep = 0
        Case "PEK": afterArr = 1 * 60:      beforeDep = 2 * 60
        Case "AMM": afterArr = 2 * 60 + 30: beforeDep = 3 * 60
        Case "GIG": afterArr = 1 * 60:      beforeDep = 2 * 60 + 30
        Case "MPU": afterArr = 40:          beforeDep = 40
        Case "CUN": afterArr = 2 * 60:      beforeDep = 3 * 60
        Case Else: OpenBuffers = False     ' FCO, CUZ, DEL: no open-hours rule
    End Select
End Function


' The k-th open window (k = 1, 2, ...) for city on weekday wd (VBA Weekday:
' vbSunday = 1 ... vbSaturday = 7), as minutes after local midnight.
' Returns False when there is no k-th window that day (e.g. TAJ on Friday).
' Windows within a day must be listed in time order.
Private Function OpenWindow(ByVal city As String, ByVal wd As Long, ByVal k As Long, _
                            ByRef startMin As Long, ByRef endMin As Long) As Boolean
    OpenWindow = False
    Select Case UCase(city)
        Case "TAJ"   ' 6:15 AM - 5:00 PM, closed Friday
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
            ReDim Preserve apAdj(1 To 2 * nAp)
        End If
        apCode(nAp) = code
        apCity(nAp) = CityIndex(AirportCity(code))
        If apCity(nAp) > 0 Then
            apAdj(nAp) = UTCOffsetMinutes(AirportCity(code)) - AirportUTCOffset(code)
        Else
            apAdj(nAp) = 0
        End If
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
    ReDim apAdj(1 To 256)
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
                    fDep(nF) = depM + apAdj(fOrig(nF))   ' city clock at grouped airports (see apAdj)
                    fArr(nF) = arrM + apAdj(fDest(nF))
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

' A fresh copy of a Long array (explicit, so results never share storage
' with the scratch array the search keeps re-filling).
Private Function CopyLongs(src() As Long) As Variant
    Dim dst() As Long, i As Long
    ReDim dst(LBound(src) To UBound(src))
    For i = LBound(src) To UBound(src): dst(i) = src(i): Next i
    CopyLongs = dst
End Function


' Replaces each required leg's options with the single required chain of
' flights from PinnedLegText. Returns False (after a message) if a flight
' isn't in the data or the legs don't fit the routes.
Private Function ApplyPinnedFlights() As Boolean
    Dim k As Long, i As Long, f As Long, parts() As String, tok() As String
    Dim chain() As Long, nch As Long, depM As Long, p As Long, msg As String
    Dim yy As Long, mo As Long, dd As Long

    ReDim pinnedFlag(1 To nF)
    For k = 1 To N_PINNED
        parts = Split(PinnedLegText(k), ";")
        ReDim chain(1 To UBound(parts) + 1)
        nch = 0
        For i = 0 To UBound(parts)
            tok = Split(Trim(parts(i)), " ")
            yy = CLng(Left(tok(2), 4)): mo = CLng(Mid(tok(2), 6, 2)): dd = CLng(Right(tok(2), 2))
            depM = CLng(Round(CDbl(DateSerial(yy, mo, dd) + TimeSerial(CLng(Left(tok(3), 2)), CLng(Right(tok(3), 2)), 0)) * 1440#, 0))
            For f = 1 To nF
                If apCode(fOrig(f)) = tok(0) And apCode(fDest(f)) = tok(1) Then
                    If LocalMin(fDep(f), fOrig(f)) = depM Then Exit For
                End If
            Next f
            If f > nF Then
                MsgBox "Required flight not found on '" & DATA_SHEET & "': " & Trim(parts(i)) & vbCrLf & _
                       "Check PinnedLegText at the top of modFindSpecialPath.", vbCritical
                Exit Function
            End If
            nch = nch + 1
            chain(nch) = f
            pinnedFlag(f) = True
            If nch > 1 Then
                If fDest(chain(nch - 1)) <> fOrig(f) Or fDep(f) - fArr(chain(nch - 1)) < MIN_CONNECTION_MINUTES Then
                    MsgBox "Required flights don't connect (same airport, at least " & MIN_CONNECTION_MINUTES & _
                           " minutes): " & PinnedLegText(k), vbCritical
                    Exit Function
                End If
            End If
        Next i

        p = 0
        If apCity(fOrig(chain(1))) > 0 And apCity(fDest(chain(nch))) > 0 Then
            p = pairIdx(apCity(fOrig(chain(1))), apCity(fDest(chain(nch))))
        End If
        If p = 0 Then
            MsgBox "Required leg doesn't go between two consecutive destinations on the routes: " & _
                   PinnedLegText(k), vbCritical
            Exit Function
        End If

        nC = nC + 1
        If nC > UBound(cDep) Then
            ReDim Preserve cDep(1 To 2 * nC): ReDim Preserve cArr(1 To 2 * nC)
            ReDim Preserve cSegStart(1 To 2 * nC): ReDim Preserve cSegCount(1 To 2 * nC)
        End If
        cDep(nC) = fDep(chain(1))
        cArr(nC) = fArr(chain(nch))
        cSegStart(nC) = nSeg + 1
        cSegCount(nC) = nch
        For i = 1 To nch
            nSeg = nSeg + 1
            If nSeg > UBound(segFlight) Then
                ReDim Preserve segFlight(1 To 2 * nSeg)
            End If
            segFlight(nSeg) = chain(i)
        Next i
        pStart(p) = nC
        pCount(p) = 1
    Next k
    ApplyPinnedFlights = True
End Function


Private Function LegPair(ByVal r As Long, ByVal j As Long) As Long
    LegPair = pairIdx(routeCity(r, j), routeCity(r, j + 1))
End Function


' Earliest you may leave stop j of route r after arriving at minute arr:
' the minimum stay and, where the city has open hours, the buffered stay
' overlapping them. -1 if nothing opens in time.
Private Function EarliestLeave(ByVal r As Long, ByVal j As Long, ByVal arr As Long) As Long
    Dim city As String, minDep As Long, afterArr As Long, beforeDep As Long, openAt As Long
    city = CityCode(routeCity(r, j))
    minDep = arr + StayRequired(r, j)
    If OpenBuffers(city, afterArr, beforeDep) Then
        openAt = EarliestOpenMin(city, arr + afterArr)
        If openAt < 0 Then
            EarliestLeave = -1
            Exit Function
        End If
        If openAt + beforeDep + 1 > minDep Then minDep = openAt + beforeDep + 1
    End If
    EarliestLeave = minDep
End Function


' Tries every combination of door-to-door options on every route order and
' records each complete itinerary under the time limit.
Private Sub SearchPinned(ByVal budget As Long)
    Dim r As Long, j As Long, k As Long, p As Long, p2 As Long, c As Long, c2 As Long
    Dim n As Long, maxCnt As Long, t As Long, ok As Boolean

    nCand = 0
    ReDim candRoute(1 To 1024): ReDim candMin(1 To 1024)
    ReDim candTot(1 To 1024): ReDim candTime(1 To 1024)
    ReDim candChosen(1 To 1024)
    dfsBudget = budget

    For r = 1 To N_ROUTES
        n = routeLegs(r)
        ok = True
        maxCnt = 1
        For j = 0 To n - 1
            If pCount(LegPair(r, j)) = 0 Then ok = False
            If pCount(LegPair(r, j)) > maxCnt Then maxCnt = pCount(LegPair(r, j))
        Next j
        If ok Then
            Application.StatusBar = "Trying every combination for order " & r & " of " & N_ROUTES & "..."
            ' backward pass: which options can still reach the end of the route
            ReDim okFrom(0 To n - 1, 1 To maxCnt + 1)
            For j = n - 1 To 0 Step -1
                p = LegPair(r, j)
                okFrom(j, pCount(p) + 1) = 0
                For k = pCount(p) To 1 Step -1
                    c = pStart(p) + k - 1
                    ok = (j = n - 1)
                    If Not ok Then
                        t = EarliestLeave(r, j + 1, cArr(c))
                        If t >= 0 Then
                            p2 = LegPair(r, j + 1)
                            c2 = FirstOptionAtOrAfter(p2, t)
                            If c2 > 0 Then ok = (okFrom(j + 1, c2 - pStart(p2) + 1) > 0)
                        End If
                    End If
                    If ok Then okFrom(j, k) = k Else okFrom(j, k) = okFrom(j, k + 1)
                Next k
            Next j

            p = LegPair(r, n - 1)
            If okFrom(n - 1, 1) > 0 And okFrom(0, 1) > 0 Then
                dfsMinEnd = cArr(pStart(p) + okFrom(n - 1, 1) - 1)
                ReDim dfsChosen(0 To n - 1)
                Call Dfs(r, 0, 0)
            End If
        End If
    Next r
End Sub


' Depth-first over leg j of route r, having arrived at its start at minute
' arrPrev (ignored for j = 0).
Private Sub Dfs(ByVal r As Long, ByVal j As Long, ByVal arrPrev As Long)
    Dim p As Long, k As Long, c As Long, t As Long
    p = LegPair(r, j)
    If j = 0 Then
        k = okFrom(0, 1)
    Else
        t = EarliestLeave(r, j, arrPrev)
        If t < 0 Then Exit Sub
        c = FirstOptionAtOrAfter(p, t)
        If c = 0 Then Exit Sub
        k = okFrom(j, c - pStart(p) + 1)
    End If

    Do While k > 0
        c = pStart(p) + k - 1
        If j > 0 Or dfsMinEnd - cDep(c) - TZCorrectionMinutes(r) < dfsBudget Then
            dfsChosen(j) = c
            If j = routeLegs(r) - 1 Then
                Call RecordItinerary(r)
            Else
                Call Dfs(r, j + 1, cArr(c))
            End If
        End If
        k = okFrom(j, k + 1)
    Loop
End Sub


Private Sub RecordItinerary(ByVal r As Long)
    Dim j As Long, total As Long, sl As Long, mn As Long, tt As Long, first As Boolean, city As String
    total = cArr(dfsChosen(routeLegs(r) - 1)) - cDep(dfsChosen(0)) - TZCorrectionMinutes(r)
    If total >= dfsBudget Then Exit Sub
    first = True
    For j = 1 To routeLegs(r) - 1
        city = CityCode(routeCity(r, j))
        If IsOptimizedCity(city) Then
            sl = cDep(dfsChosen(j)) - cArr(dfsChosen(j - 1)) - StayRequired(r, j)
            tt = tt + sl
            If first Then
                mn = sl: first = False
            ElseIf sl < mn Then
                mn = sl
            End If
        End If
    Next j

    nCand = nCand + 1
    If nCand > UBound(candRoute) Then
        ReDim Preserve candRoute(1 To 2 * nCand): ReDim Preserve candMin(1 To 2 * nCand)
        ReDim Preserve candTot(1 To 2 * nCand): ReDim Preserve candTime(1 To 2 * nCand)
        ReDim Preserve candChosen(1 To 2 * nCand)
    End If
    candRoute(nCand) = r
    candMin(nCand) = mn
    candTot(nCand) = tt
    candTime(nCand) = total
    candChosen(nCand) = CopyLongs(dfsChosen)
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
    Dim seenDays() As String, exact() As String, nOther() As Long

    ReDim nOther(1 To TOP_N)
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
        ElseIf g <= nShown Then
            nOther(g) = nOther(g) + 1
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

    For g = 1 To nShown
        If nOther(g) > 0 Then
            alsoNote(g) = IIf(alsoNote(g) = "", "", "Other start dates: " & alsoNote(g) & ".  ") & _
                          nOther(g) & " other timing combination" & IIf(nOther(g) = 1, "", "s") & _
                          " through these same airports rank" & IIf(nOther(g) = 1, "s", "") & " lower."
        End If
    Next g
End Sub


' ===================== OUTPUT =====================

Private Function TravelType(ByVal f As Long) As String
    If apCode(fOrig(f)) = "MPU" Or apCode(fDest(f)) = "MPU" Then
        TravelType = "Train"
    ElseIf apCode(fOrig(f)) = "TAJ" Or apCode(fDest(f)) = "TAJ" Then
        TravelType = "Transfer"
    Else
        TravelType = "Flight"
    End If
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
        .Value = "TOP " & nShown & " SPECIAL ROUTES  --  required FCO-IST-AMM, AMM-DXB-DEL, DEL-PEK flights, " & _
                 ScopeText() & "  |  ranked by worst-case slack, then total slack, then shortest trip"
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

    ' "Other versions" row (row 11): tall enough for the longest note
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
    firstDep = LocalMin(cDep(ch(0)), fOrig(segFlight(cSegStart(ch(0)))))
    f = segFlight(cSegStart(ch(UBound(ch))) + cSegCount(ch(UBound(ch))) - 1)
    lastArr = LocalMin(cArr(ch(UBound(ch))), fDest(f))

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
        .Value = "#" & rank & "   Order " & rt
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
                    "  from " & apCode(fOrig(segFlight(cSegStart(ch(0))))))
    Call SummaryRow(outWs, topRow + 6, col0, "Trip ends", Format(MinToDate(lastArr), "ddd mmm d, yyyy h:mm AM/PM") & _
                    "  in " & apCode(fDest(f)))
    Call SummaryRow(outWs, topRow + 7, col0, "Other versions", IIf(note = "", "(none)", note))
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
                Call StepRow(outWs, r, col0, stepNo, "Connection", fDest(prevF), fOrig(f), _
                             fArr(prevF), fDep(f), MIN_CONNECTION_MINUTES, "", "Layover between flights")
                outWs.Cells(r, col0).Resize(1, CARD_COLS).Interior.Color = RGB(242, 242, 242)
                outWs.Cells(r, col0).Resize(1, CARD_COLS).Font.Italic = True
            End If
            r = r + 1: stepNo = stepNo + 1
            Call StepRow(outWs, r, col0, stepNo, TravelType(f), fOrig(f), fDest(f), _
                         fDep(f), fArr(f), -1, "", IIf(pinnedFlag(f), "REQUIRED flight.  ", "") & DATA_SHEET & " row " & fRow(f))
            If pinnedFlag(f) Then outWs.Cells(r, col0 + 1).Resize(1, CARD_COLS - 1).Font.Bold = True
            If fDur(f) >= 0 Then outWs.Cells(r, col0 + 6).Value = FormatDuration(CDbl(fDur(f)))
            prevF = f
        Next s

        If j < UBound(ch) Then
            ' stay at a destination
            city = CityCode(routeCity(rt, j + 1))
            req = StayRequired(rt, j + 1)
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
                txt = "Arrive " & apCode(fDest(prevF)) & ", leave from " & apCode(fOrig(f)) & "." & IIf(txt = "", "", "  ") & txt
            End If
            txt = txt & ClockNote(fDest(prevF), fOrig(f))
            If city = "DEL" Then
                If CityCode(routeCity(rt, j + 2)) = "TAJ" Then
                    txt = "Wait before the trip to TAJ." & IIf(txt = "", "", "  ") & txt
                Else
                    txt = "Wait after getting back from TAJ." & IIf(txt = "", "", "  ") & txt
                End If
            End If
            If Not IsOptimizedCity(city) Then txt = "Not counted toward the ranking." & IIf(txt = "", "", "  ") & txt

            r = r + 1: stepNo = stepNo + 1
            Call StepRow(outWs, r, col0, stepNo, "STAY", fDest(prevF), fOrig(f), stayStart, stayEnd, req, CStr(sl), Trim(txt))
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


' One row of the step table. minReq < 0 leaves Min Required blank. startM and
' endM are internal times at airports fromAp / toAp; they're shown on each
' airport's own clock, and Duration is the true elapsed time between them.
Private Sub StepRow(outWs As Worksheet, ByVal r As Long, ByVal col0 As Long, ByVal stepNo As Long, _
                    ByVal kind As String, ByVal fromAp As Long, ByVal toAp As Long, _
                    ByVal startM As Long, ByVal endM As Long, ByVal minReq As Long, _
                    ByVal slackText As String, ByVal note As String)
    outWs.Cells(r, col0 + 0).Value = stepNo
    outWs.Cells(r, col0 + 1).Value = kind
    outWs.Cells(r, col0 + 2).Value = apCode(fromAp)
    outWs.Cells(r, col0 + 3).Value = apCode(toAp)
    outWs.Cells(r, col0 + 4).Value = MinToDate(LocalMin(startM, fromAp))
    outWs.Cells(r, col0 + 4).NumberFormat = "ddd mmm d, h:mm AM/PM"
    outWs.Cells(r, col0 + 5).Value = MinToDate(LocalMin(endM, toAp))
    outWs.Cells(r, col0 + 5).NumberFormat = "ddd mmm d, h:mm AM/PM"
    outWs.Cells(r, col0 + 6).Value = FormatDuration(CDbl(endM - startM))
    If minReq >= 0 Then outWs.Cells(r, col0 + 7).Value = FormatDuration(CDbl(minReq))
    If slackText <> "" Then outWs.Cells(r, col0 + 8).Value = FormatDuration(CDbl(CLng(slackText)))
    outWs.Cells(r, col0 + 9).Value = note
    outWs.Cells(r, col0).Resize(1, 4).HorizontalAlignment = xlCenter
End Sub
