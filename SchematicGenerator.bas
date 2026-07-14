Option Explicit
'===========================================================================
' SchematicGenerator  v14
' Room-section layout | CAL chain tracing | DeviceNames classification
'
' clsCircuit: Cable, CableType, FromKey, ToKey, FromPort, ToPort,
'             Cal, Length, ConnectorA, TxA, ConnectorB, TxB
' clsNode:    Key, Room, Rack, Equip, IsPassive
' DeviceNames sheet col A = generic device name stems
'
' v14 changes:
'   * Connection lines now draw for EVERY segment (previous "toX > fromX"
'     guard silently dropped continuation / NIS links that routed right-to-
'     left, leaving multi-room circuits unconnected).
'   * DrawCktLine routes orthogonally (L-shaped) when the two endpoints are
'     not at the same height, so a link always joins the two device edges
'     instead of leaving a stray diagonal.
'   * Page width grows with the number of room sections in a circuit.
'===========================================================================
Private Const MID_GAP As Double = 4   ' routing zone between same-room and other-room sections

'-- SHEET NAMES ------------------------------------------------------------
Private Const SHT_FIBR As String = "Cable Schedule - Fibres"
Private Const SHT_DAC  As String = "Cable Schedule - DAC"
Private Const SHT_COP  As String = "Cable Schedule - Copper"
Private Const SHT_FRT  As String = "Front Sheet"
Private Const SHT_DN   As String = "DeviceNames"

'-- COLUMN INDICES ---------------------------------------------------------
Private Const CC_REF    As Long = 1   ' A
Private Const CC_CABLE  As Long = 2   ' B
Private Const CC_A_CONN As Long = 4   ' D
Private Const CC_A_ROOM As Long = 5   ' E
Private Const CC_A_RACK As Long = 6   ' F
Private Const CC_A_EQ   As Long = 7   ' G
Private Const CC_A_PORT As Long = 8   ' H
Private Const CC_A_TX   As Long = 9   ' I
Private Const CC_B_CONN As Long = 11  ' K
Private Const CC_B_ROOM As Long = 12  ' L
Private Const CC_B_RACK As Long = 13  ' M
Private Const CC_B_EQ   As Long = 14  ' N
Private Const CC_B_PORT As Long = 15  ' O
Private Const CC_B_TX   As Long = 16  ' P
Private Const CC_CAL_F  As Long = 20  ' T  Fibre
Private Const CC_CAL_D  As Long = 22  ' V  DAC
Private Const CC_LEN    As Long = 23  ' W

'-- PAGE -------------------------------------------------------------------
Private Const MG      As Double = 0.3    ' body margin
Private Const HDRI    As Double = 0.12   ' page top inset
Private Const HDR_H   As Double = 0.78   ' header strip
Private Const SITE_H  As Double = 0.42   ' site banner
Private Const ROOM_H  As Double = 0.38   ' room name strip

'-- SOURCE DEVICE ----------------------------------------------------------
Private Const SRC_W   As Double = 2.7
Private Const SRC_PX  As Double = 0.12   ' rack-box X padding
Private Const SRC_PY  As Double = 0.1    ' rack-box Y padding
Private Const SRC_RKH As Double = 0.26   ' rack header height

'-- ROOM SECTION DEVICES ---------------------------------------------------
Private Const RM_W    As Double = 2.7
Private Const RM_PX   As Double = 0.12
Private Const RM_PY   As Double = 0.08
Private Const RM_RKH  As Double = 0.22
Private Const INNER_G As Double = 0.25   ' gap between L and R sub-cols
Private Const SECT_G  As Double = 0.5    ' gap between sections
Private Const ROUTE_W As Double = 0.8    ' min routing zone

'-- PORT GEOMETRY ----------------------------------------------------------
Private Const DEV_H   As Double = 0.3
Private Const PORT_H  As Double = 0.32

'-- SEPARATOR SLOT COSTS ---------------------------------------------------
Private Const SEP_DEV As Long = 2.5
Private Const SEP_RK  As Long = 3
Private Const SEP_RM  As Long = 4

'-- COLOURS ----------------------------------------------------------------
Private Const C_DF  As String = "THEMEGUARD(RGB(248,250,253))"
Private Const C_DH  As String = "THEMEGUARD(RGB(175,200,225))"
Private Const C_OF  As String = "THEMEGUARD(RGB(246,250,247))"
Private Const C_OH  As String = "THEMEGUARD(RGB(168,205,190))"
Private Const C_RKB As String = "THEMEGUARD(RGB(244,247,251))"
Private Const C_RKH As String = "THEMEGUARD(RGB(200,213,228))"
Private Const C_BDR As String = "THEMEGUARD(RGB(100,110,125))"
Private Const C_BRK As String = "THEMEGUARD(RGB(135,158,185))"
Private Const C_SD  As String = "THEMEGUARD(RGB(170,185,205))"
Private Const C_SR  As String = "THEMEGUARD(RGB(90,120,160))"
Private Const C_SRM As String = "THEMEGUARD(RGB(55,85,130))"
Private Const C_SCH As String = "THEMEGUARD(RGB(170,192,218))"
Private Const C_SCB As String = "THEMEGUARD(RGB(105,138,175))"
Private Const C_SCS As String = "THEMEGUARD(RGB(240,244,250))"

'-- CABLE LINE STYLES ------------------------------------------------------
Private Const F_COL As String = "THEMEGUARD(RGB(0,80,200))"
Private Const F_PAT As String = "1"
Private Const F_WT  As String = "1.5pt"
Private Const D_COL As String = "THEMEGUARD(RGB(40,40,40))"
Private Const D_PAT As String = "1"
Private Const D_WT  As String = "1pt"
Private Const P_COL As String = "THEMEGUARD(RGB(40,40,40))"
Private Const P_PAT As String = "4"
Private Const P_WT  As String = "1pt"
Private Const N_COL As String = "THEMEGUARD(RGB(0,80,200))"
Private Const N_PAT As String = "4"
Private Const N_WT  As String = "1pt"

'-- MODULE STATE -----------------------------------------------------------
Private gApp   As Object
Private gPage  As Object
Private gPageH As Double
Private gPageW As Double
Private gDN()  As String
Private gDNCnt As Long

'===========================================================================
' DIAGNOSTIC
'===========================================================================
Public Sub TestColumns()
    Dim sheets(2) As String
    sheets(0) = SHT_FIBR: sheets(1) = SHT_DAC: sheets(2) = SHT_COP
    Dim si As Long
    For si = 0 To 2
        Dim ws As Worksheet: Set ws = FindSheet(sheets(si))
        If ws Is Nothing Then GoTo NextSht
        Dim lr As Long: lr = ws.Cells(ws.rows.count, CC_REF).End(xlUp).Row
        Dim r As Long
        For r = 1 To lr
            If IsCircRow(ws, r) Then
                MsgBox "Sheet: " & sheets(si) & "  Row " & r & vbLf & _
                    "A_CONN  col " & CC_A_CONN & ": [" & SC(ws, r, CC_A_CONN) & "]" & vbLf & _
                    "A_ROOM  col " & CC_A_ROOM & ": [" & SC(ws, r, CC_A_ROOM) & "]" & vbLf & _
                    "A_RACK  col " & CC_A_RACK & ": [" & SC(ws, r, CC_A_RACK) & "]" & vbLf & _
                    "A_EQ    col " & CC_A_EQ & ": [" & SC(ws, r, CC_A_EQ) & "]" & vbLf & _
                    "A_PORT  col " & CC_A_PORT & ": [" & SC(ws, r, CC_A_PORT) & "]" & vbLf & _
                    "A_TX    col " & CC_A_TX & ": [" & SC(ws, r, CC_A_TX) & "]" & vbLf & _
                    "B_EQ    col " & CC_B_EQ & ": [" & SC(ws, r, CC_B_EQ) & "]" & vbLf & _
                    "B_PORT  col " & CC_B_PORT & ": [" & SC(ws, r, CC_B_PORT) & "]" & vbLf & _
                    "CAL_F   col " & CC_CAL_F & ": [" & SC(ws, r, CC_CAL_F) & "]" & vbLf & _
                    "CAL_D   col " & CC_CAL_D & ": [" & SC(ws, r, CC_CAL_D) & "]" & vbLf & _
                    "LEN     col " & CC_LEN & ": [" & SC(ws, r, CC_LEN) & "]", _
                    vbInformation, "TestColumns"
                Exit Sub
            End If
        Next r
NextSht:
    Next si
    MsgBox "No circuit rows found.", vbExclamation
End Sub

'===========================================================================
' DEVICE NAMES
'===========================================================================
Private Sub LoadDeviceNames()
    gDNCnt = 0
    Dim ws As Worksheet: Set ws = FindSheet(SHT_DN)
    If ws Is Nothing Then Exit Sub
    Dim lr As Long: lr = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    ReDim gDN(lr)
    Dim r As Long
    For r = 1 To lr
        Dim v As String: v = Trim(CStr(ws.Cells(r, 1).value))
        If v <> "" Then gDN(gDNCnt) = LCase(v): gDNCnt = gDNCnt + 1
    Next r
End Sub

' Strip site prefix (everything up to and including first "-").
' Then check if remainder starts with any DeviceNames stem followed
' by "-" or end-of-string. e.g. "site1-device-1a" -> "device-1a" matches "device".
Private Function IsSourceDev(nm As String) As Boolean
    Dim s As String: s = LCase(Trim(nm))
    Dim i As Long
    For i = 0 To gDNCnt - 1
        Dim stm As String: stm = gDN(i)
        Dim sln As Long:   sln = Len(stm)
        If Len(s) >= sln Then
            If Left(s, sln) = stm Then
                If Len(s) = sln Then IsSourceDev = True: Exit Function
                If Mid(s, sln + 1, 1) = "-" Then IsSourceDev = True: Exit Function
            End If
        End If
    Next i
End Function

'===========================================================================
' DATA LOADING
'===========================================================================
Private Sub LoadSheet(ws As Worksheet, cabTyp As String, calCol As Long, _
    nodes As Object, circs As Collection, fibreByCAL As Object)
    Dim lr As Long: lr = ws.Cells(ws.rows.count, CC_REF).End(xlUp).Row
    Dim r As Long
    For r = 1 To lr
        If Not IsCircRow(ws, r) Then GoTo Skip
        Dim kA As String: kA = NK(SC(ws, r, CC_A_ROOM), SC(ws, r, CC_A_RACK), SC(ws, r, CC_A_EQ))
        Dim kB As String: kB = NK(SC(ws, r, CC_B_ROOM), SC(ws, r, CC_B_RACK), SC(ws, r, CC_B_EQ))
        EnsureNode nodes, kA, SC(ws, r, CC_A_ROOM), SC(ws, r, CC_A_RACK), SC(ws, r, CC_A_EQ)
        EnsureNode nodes, kB, SC(ws, r, CC_B_ROOM), SC(ws, r, CC_B_RACK), SC(ws, r, CC_B_EQ)
        Dim c As clsCircuit: Set c = New clsCircuit
        c.cable = SC(ws, r, CC_CABLE):    c.cableType = cabTyp
        c.FromKey = kA:                   c.ToKey = kB
        c.FromPort = SC(ws, r, CC_A_PORT): c.ToPort = SC(ws, r, CC_B_PORT)
        c.ConnectorA = SC(ws, r, CC_A_CONN): c.TxA = SC(ws, r, CC_A_TX)
        c.ConnectorB = SC(ws, r, CC_B_CONN): c.TxB = SC(ws, r, CC_B_TX)
        c.Length = SC(ws, r, CC_LEN)
        If calCol > 0 Then c.cal = SC(ws, r, calCol)
        circs.Add c
        If cabTyp = "Fibre" And c.cal <> "" Then
            If Not fibreByCAL.Exists(c.cal) Then fibreByCAL.Add c.cal, New Collection
            fibreByCAL(c.cal).Add c
        End If
Skip:
    Next r
End Sub

Private Sub EnsureNode(nodes As Object, k As String, _
    rm As String, rk As String, eq As String)
    If nodes.Exists(k) Then Exit Sub
    Dim n As clsNode: Set n = New clsNode
    n.key = k: n.Room = rm: n.Rack = rk: n.Equip = eq: n.IsPassive = IsPassive(eq)
    nodes.Add k, n
End Sub

'===========================================================================
' CHAIN EXTENSION  (recursive)
' Traces downstream from currentEndB using remaining continuation circuits.
' Inserts NIS implied-tie segments where a gap is detected.
'===========================================================================
Private Sub ExtendChain(currentEndB As String, contWork As Collection, _
    gSlot As Long, allNodes As Object, _
    seg_aKey() As String, seg_bKey() As String, _
    seg_aPort() As String, seg_bPort() As String, _
    seg_aConn() As String, seg_bConn() As String, _
    seg_aTx() As String, seg_bTx() As String, _
    seg_cal() As String, seg_len() As String, _
    seg_typ() As String, seg_slot() As Long, nSeg As Long)

    If contWork.count = 0 Then Exit Sub

    ' 1. Look for direct link: a continuation circuit whose End-A = currentEndB
    Dim i As Long
    For i = 1 To contWork.count
        Dim cc As clsCircuit: Set cc = contWork(i)
        If cc.FromKey = currentEndB Then
            seg_aKey(nSeg) = cc.FromKey:   seg_bKey(nSeg) = cc.ToKey
            seg_aPort(nSeg) = cc.FromPort: seg_bPort(nSeg) = cc.ToPort
            seg_aConn(nSeg) = cc.ConnectorA: seg_bConn(nSeg) = cc.ConnectorB
            seg_aTx(nSeg) = cc.TxA:        seg_bTx(nSeg) = cc.TxB
            seg_cal(nSeg) = cc.cal:         seg_len(nSeg) = cc.Length
            seg_typ(nSeg) = cc.cableType:   seg_slot(nSeg) = gSlot
            nSeg = nSeg + 1
            contWork.Remove i
            If allNodes.Exists(cc.ToKey) Then
                Dim nd As clsNode: Set nd = allNodes(cc.ToKey)
                If Not IsSourceDev(nd.Equip) And contWork.count > 0 Then
                    ExtendChain cc.ToKey, contWork, gSlot, allNodes, _
                        seg_aKey, seg_bKey, seg_aPort, seg_bPort, _
                        seg_aConn, seg_bConn, seg_aTx, seg_bTx, _
                        seg_cal, seg_len, seg_typ, seg_slot, nSeg
                End If
            End If
            Exit Sub
        End If
    Next i

    ' 2. No direct link - gap exists. Pick best next: hop End-B before destination End-B.
    Dim bestIdx As Long: bestIdx = 1
    For i = 1 To contWork.count
        Set cc = contWork(i)
        If allNodes.Exists(cc.ToKey) Then
            Dim nd2 As clsNode: Set nd2 = allNodes(cc.ToKey)
            If Not IsSourceDev(nd2.Equip) Then bestIdx = i: Exit For
        End If
    Next i
    Set cc = contWork(bestIdx)

    ' NIS implied tie: currentEndB -> cc.FromKey
    seg_aKey(nSeg) = currentEndB: seg_bKey(nSeg) = cc.FromKey
    seg_aPort(nSeg) = "":         seg_bPort(nSeg) = ""
    seg_aConn(nSeg) = "":         seg_bConn(nSeg) = ""
    seg_aTx(nSeg) = "":           seg_bTx(nSeg) = ""
    seg_cal(nSeg) = "":           seg_len(nSeg) = ""
    seg_typ(nSeg) = "NIS":        seg_slot(nSeg) = gSlot
    nSeg = nSeg + 1

    ' Continuation segment
    seg_aKey(nSeg) = cc.FromKey:   seg_bKey(nSeg) = cc.ToKey
    seg_aPort(nSeg) = cc.FromPort: seg_bPort(nSeg) = cc.ToPort
    seg_aConn(nSeg) = cc.ConnectorA: seg_bConn(nSeg) = cc.ConnectorB
    seg_aTx(nSeg) = cc.TxA:        seg_bTx(nSeg) = cc.TxB
    seg_cal(nSeg) = cc.cal:         seg_len(nSeg) = cc.Length
    seg_typ(nSeg) = cc.cableType:   seg_slot(nSeg) = gSlot
    nSeg = nSeg + 1
    contWork.Remove bestIdx

    If allNodes.Exists(cc.ToKey) And contWork.count > 0 Then
        Dim nd3 As clsNode: Set nd3 = allNodes(cc.ToKey)
        If Not IsSourceDev(nd3.Equip) Then
            ExtendChain cc.ToKey, contWork, gSlot, allNodes, _
                seg_aKey, seg_bKey, seg_aPort, seg_bPort, _
                seg_aConn, seg_bConn, seg_aTx, seg_bTx, _
                seg_cal, seg_len, seg_typ, seg_slot, nSeg
        End If
    End If
End Sub

'===========================================================================
' MAIN
'===========================================================================
Public Sub CreateSchematic()
    On Error GoTo ErrH

    LoadDeviceNames
    If gDNCnt = 0 Then MsgBox "DeviceNames sheet missing or empty.", vbCritical: Exit Sub

    ' Connect to Visio
    On Error Resume Next
    Set gApp = GetObject(, "Visio.Application")
    If Err.Number <> 0 Then Err.Clear: Set gApp = CreateObject("Visio.Application")
    On Error GoTo ErrH
    If gApp Is Nothing Then MsgBox "Cannot start Visio.", vbCritical: Exit Sub
    gApp.Visible = True
    Dim vDoc As Object: Set vDoc = gApp.Documents.Add("")

    ' Load all circuits
    Dim allNodes As Object:   Set allNodes = CreateObject("Scripting.Dictionary")
    Dim allCircs  As Collection: Set allCircs = New Collection
    Dim fibreByCAL As Object: Set fibreByCAL = CreateObject("Scripting.Dictionary")

    Dim wF As Worksheet: Set wF = FindSheet(SHT_FIBR)
    Dim wD As Worksheet: Set wD = FindSheet(SHT_DAC)
    Dim wC As Worksheet: Set wC = FindSheet(SHT_COP)
    If Not wF Is Nothing Then LoadSheet wF, "Fibre", CC_CAL_F, allNodes, allCircs, fibreByCAL
    If Not wD Is Nothing Then LoadSheet wD, "DAC", CC_CAL_D, allNodes, allCircs, fibreByCAL
    If Not wC Is Nothing Then LoadSheet wC, "Copper", 0, allNodes, allCircs, fibreByCAL
    If allCircs.count = 0 Then MsgBox "No circuits loaded.", vbExclamation: Exit Sub

    ' Identify source devices
    Dim srcList(199) As String: Dim nSrc As Long: nSrc = 0
    Dim srcSet As Object: Set srcSet = CreateObject("Scripting.Dictionary")
    Dim c As clsCircuit
    For Each c In allCircs
        If Not srcSet.Exists(c.FromKey) Then
            Dim nd0 As clsNode: Set nd0 = allNodes(c.FromKey)
            If IsSourceDev(nd0.Equip) Then
                srcSet.Add c.FromKey, True
                srcList(nSrc) = c.FromKey: nSrc = nSrc + 1
            End If
        End If
    Next c
    If nSrc = 0 Then MsgBox "No source devices matched DeviceNames.", vbExclamation: Exit Sub

    Dim usedNames As Object: Set usedNames = CreateObject("Scripting.Dictionary")
    Dim pageN As Long: pageN = 0

    ' -- Per source device -------------------------------------------------
    Dim si As Long
    For si = 0 To nSrc - 1
        Dim srcKey As String: srcKey = srcList(si)
        Dim srcNd  As clsNode: Set srcNd = allNodes(srcKey)

        ' Segment arrays
        Const MSEG As Long = 800
        Dim s_aK(MSEG) As String, s_bK(MSEG) As String
        Dim s_aP(MSEG) As String, s_bP(MSEG) As String
        Dim s_aC(MSEG) As String, s_bC(MSEG) As String
        Dim s_aT(MSEG) As String, s_bT(MSEG) As String
        Dim s_cl(MSEG) As String, s_ln(MSEG) As String
        Dim s_ty(MSEG) As String, s_sl(MSEG) As Long
        Dim nSeg As Long: nSeg = 0

        ' Slot separator info (indexed by global slot)
        Dim sl_sp(MSEG) As Integer  ' 0=none 1=dev 2=rack 3=room
        Dim nSlots As Long: nSlots = 0

        ' Collect direct circuits (src = End-A), sort by firsthop room|rack|eq|port
        Dim dirs(299) As Object: Dim nDir As Long: nDir = 0
        For Each c In allCircs
            If c.FromKey = srcKey Then Set dirs(nDir) = c: nDir = nDir + 1
        Next c
        If nDir = 0 Then GoTo NextSrc

        Dim di As Long, dj As Long
        For di = 0 To nDir - 2
            For dj = 0 To nDir - 2 - di
                Dim cA As clsCircuit: Set cA = dirs(dj)
                Dim cB As clsCircuit: Set cB = dirs(dj + 1)
                Dim nA As clsNode: Set nA = allNodes(cA.ToKey)
                Dim nB As clsNode: Set nB = allNodes(cB.ToKey)
                Dim sA As String: sA = LCase(nA.Room & "|" & nA.Rack & "|" & nA.Equip & "|" & cA.ToPort)
                Dim sB As String: sB = LCase(nB.Room & "|" & nB.Rack & "|" & nB.Equip & "|" & cB.ToPort)
                If sA > sB Then
                    Dim tmpC As Object: Set tmpC = dirs(dj)
                    Set dirs(dj) = dirs(dj + 1): Set dirs(dj + 1) = tmpC
                End If
            Next dj
        Next di

        ' Build segment list with slot assignment
        Dim prevRoom As String: prevRoom = ""
        Dim prevRack As String: prevRack = ""
        Dim prevEq   As String: prevEq = ""

        For di = 0 To nDir - 1
            Dim dc As clsCircuit: Set dc = dirs(di)
            Dim bNd As clsNode:   Set bNd = allNodes(dc.ToKey)

            ' Separator before this slot
            Dim sepCost As Long: sepCost = 0
            Dim sepType As Integer: sepType = 0
            If nSlots > 0 Then
                If LCase(bNd.Room) <> LCase(prevRoom) Then
                    sepCost = SEP_RM: sepType = 3
                ElseIf LCase(bNd.Rack) <> LCase(prevRack) Then
                    sepCost = SEP_RK: sepType = 2
                ElseIf LCase(bNd.Equip) <> LCase(prevEq) Then
                    sepCost = SEP_DEV: sepType = 1
                End If
            End If

            Dim gSlot As Long: gSlot = nSlots + sepCost
            sl_sp(gSlot) = sepType
            nSlots = gSlot + 1
            prevRoom = bNd.Room: prevRack = bNd.Rack: prevEq = bNd.Equip

            ' Direct segment
            s_aK(nSeg) = srcKey:      s_bK(nSeg) = dc.ToKey
            s_aP(nSeg) = dc.FromPort: s_bP(nSeg) = dc.ToPort
            s_aC(nSeg) = dc.ConnectorA: s_bC(nSeg) = dc.ConnectorB
            s_aT(nSeg) = dc.TxA:      s_bT(nSeg) = dc.TxB
            s_cl(nSeg) = dc.cal:       s_ln(nSeg) = dc.Length
            s_ty(nSeg) = dc.cableType: s_sl(nSeg) = gSlot
            nSeg = nSeg + 1

            ' CAL chain tracing - Fibre only
            If dc.cableType = "Fibre" And dc.cal <> "" Then
                If fibreByCAL.Exists(dc.cal) Then
                    Dim calSrc As Collection: Set calSrc = fibreByCAL(dc.cal)
                    ' Build working continuation set: exclude any circuit where End-A = srcKey
                    Dim contW As Collection: Set contW = New Collection
                    Dim cc As clsCircuit
                    For Each cc In calSrc
                        If cc.FromKey <> srcKey Then contW.Add cc
                    Next cc
                    ExtendChain dc.ToKey, contW, gSlot, allNodes, _
                        s_aK, s_bK, s_aP, s_bP, s_aC, s_bC, s_aT, s_bT, _
                        s_cl, s_ln, s_ty, s_sl, nSeg
                End If
            End If
        Next di

        ' -- Build room section & room node data ---------------------------
        '
        ' Room nodes (rn): unique (deviceKey, sub-col) entries.
        ' Sub-col placement:
        '   NIS End-B  -> LEFT only  (entry into room via implied tie)
        '   non-NIS End-A (not srcKey) -> LEFT (continuation circuit origin)
        '   non-NIS End-B -> RIGHT
        '
        Const MAXRS As Long = 20
        Dim rs_room(MAXRS) As String
        Dim rs_hasL(MAXRS) As Boolean   ' section needs a LEFT sub-col
        Dim rs_leftX(MAXRS) As Double   ' device box left X in LEFT sub-col
        Dim rs_rightX(MAXRS) As Double  ' device box left X in RIGHT sub-col
        Dim rs_swimL(MAXRS) As Double
        Dim rs_swimR(MAXRS) As Double
        Dim nRS As Long: nRS = 0
        Dim rsMap As Object: Set rsMap = CreateObject("Scripting.Dictionary")

        Const MAXRN As Long = 200
        Dim rn_key(MAXRN)  As String
        Dim rn_room(MAXRN) As String
        Dim rn_rack(MAXRN) As String
        Dim rn_eq(MAXRN)   As String
        Dim rn_pass(MAXRN) As Boolean
        Dim rn_isL(MAXRN)  As Boolean   ' True = LEFT sub-col
        Dim rn_sec(MAXRN)  As Long
        Dim rn_fsl(MAXRN)  As Long      ' first global slot
        Dim rn_lsl(MAXRN)  As Long      ' last global slot
        Dim rn_pc(MAXRN)   As Long      ' real port count
        Dim nRN As Long: nRN = 0
        Dim rnMap As Object: Set rnMap = CreateObject("Scripting.Dictionary")

        Dim sg As Long
        For sg = 0 To nSeg - 1
            Dim sTyp As String: sTyp = s_ty(sg)
            Dim aK   As String: aK = s_aK(sg)
            Dim bK   As String: bK = s_bK(sg)
            Dim gsl  As Long:   gsl = s_sl(sg)
            If bK = "" Then GoTo SkipSeg

            Dim bNode As clsNode: Set bNode = allNodes(bK)

            ' Ensure room section exists for bNode's room
            Dim rmKey As String: rmKey = LCase(bNode.Room)
            If Not rsMap.Exists(rmKey) Then
                rsMap.Add rmKey, nRS
                rs_room(nRS) = bNode.Room: nRS = nRS + 1
            End If
            Dim secIdxB As Long: secIdxB = CLng(rsMap(rmKey))

            If sTyp = "NIS" Then
                ' NIS End-B -> LEFT sub-col only (no port)
                Dim rnKL As String: rnKL = bK & "|L"
                If Not rnMap.Exists(rnKL) Then
                    rnMap.Add rnKL, nRN
                    rn_key(nRN) = bK: rn_room(nRN) = bNode.Room
                    rn_rack(nRN) = bNode.Rack: rn_eq(nRN) = bNode.Equip
                    rn_pass(nRN) = bNode.IsPassive: rn_isL(nRN) = True
                    rn_sec(nRN) = secIdxB
                    rn_fsl(nRN) = gsl: rn_lsl(nRN) = gsl: rn_pc(nRN) = 0
                    rs_hasL(secIdxB) = True: nRN = nRN + 1
                Else
                    Dim rnI0 As Long: rnI0 = CLng(rnMap(rnKL))
                    If gsl < rn_fsl(rnI0) Then rn_fsl(rnI0) = gsl
                    If gsl > rn_lsl(rnI0) Then rn_lsl(rnI0) = gsl
                End If
                GoTo SkipSeg   ' NIS has no End-B port or End-A port to add elsewhere
            End If

            ' Non-NIS: End-B -> RIGHT sub-col
            Dim rnKR As String: rnKR = bK & "|R"
            If Not rnMap.Exists(rnKR) Then
                rnMap.Add rnKR, nRN
                rn_key(nRN) = bK: rn_room(nRN) = bNode.Room
                rn_rack(nRN) = bNode.Rack: rn_eq(nRN) = bNode.Equip
                rn_pass(nRN) = bNode.IsPassive: rn_isL(nRN) = False
                rn_sec(nRN) = secIdxB
                rn_fsl(nRN) = gsl: rn_lsl(nRN) = gsl: rn_pc(nRN) = 1
                nRN = nRN + 1
            Else
                Dim rnI1 As Long: rnI1 = CLng(rnMap(rnKR))
                If gsl < rn_fsl(rnI1) Then rn_fsl(rnI1) = gsl
                If gsl > rn_lsl(rnI1) Then rn_lsl(rnI1) = gsl
                rn_pc(rnI1) = rn_pc(rnI1) + 1
            End If

            ' Non-NIS with End-A not srcKey -> End-A also in LEFT sub-col
            If aK <> srcKey And allNodes.Exists(aK) Then
                Dim aNode As clsNode: Set aNode = allNodes(aK)
                Dim rmKeyA As String: rmKeyA = LCase(aNode.Room)
                If Not rsMap.Exists(rmKeyA) Then
                    rsMap.Add rmKeyA, nRS
                    rs_room(nRS) = aNode.Room: nRS = nRS + 1
                End If
                Dim secIdxA As Long: secIdxA = CLng(rsMap(rmKeyA))
                Dim rnKLA As String: rnKLA = aK & "|L"
                If Not rnMap.Exists(rnKLA) Then
                    rnMap.Add rnKLA, nRN
                    rn_key(nRN) = aK: rn_room(nRN) = aNode.Room
                    rn_rack(nRN) = aNode.Rack: rn_eq(nRN) = aNode.Equip
                    rn_pass(nRN) = aNode.IsPassive: rn_isL(nRN) = True
                    rn_sec(nRN) = secIdxA
                    rn_fsl(nRN) = gsl: rn_lsl(nRN) = gsl: rn_pc(nRN) = 1
                    rs_hasL(secIdxA) = True: nRN = nRN + 1
                Else
                    Dim rni2 As Long: rni2 = CLng(rnMap(rnKLA))
                    If gsl < rn_fsl(rni2) Then rn_fsl(rni2) = gsl
                    If gsl > rn_lsl(rni2) Then rn_lsl(rni2) = gsl
                    rn_pc(rni2) = rn_pc(rni2) + 1
                End If
            End If
SkipSeg:
        Next sg
        ' Guarantee source room is the leftmost destination section
        Dim srcRmKey As String: srcRmKey = LCase(srcNd.Room)
        If rsMap.Exists(srcRmKey) Then
            Dim srcRsI As Long: srcRsI = CLng(rsMap(srcRmKey))
            If srcRsI <> 0 Then
                Dim swpRoom As String:  swpRoom = rs_room(0)
                Dim swpHasL As Boolean: swpHasL = rs_hasL(0)
                rs_room(0) = rs_room(srcRsI): rs_hasL(0) = rs_hasL(srcRsI)
                rs_room(srcRsI) = swpRoom:   rs_hasL(srcRsI) = swpHasL
                rsMap(srcRmKey) = 0
                rsMap(LCase(swpRoom)) = srcRsI
                Dim rniS As Long
                For rniS = 0 To nRN - 1
                    If rn_sec(rniS) = srcRsI Then
                        rn_sec(rniS) = 0
                    ElseIf rn_sec(rniS) = 0 Then
                        rn_sec(rniS) = srcRsI
                    End If
                Next rniS
            End If
        End If
        If nRS = 0 Then GoTo NextSrc

        ' -- Section widths ------------------------------------------------
        Dim totSecW As Double: totSecW = 0
        Dim secW(MAXRS) As Double
        Dim rs As Long
        For rs = 0 To nRS - 1
            If rs_hasL(rs) Then
                secW(rs) = 2 * RM_PX + 2 * RM_W + INNER_G
            Else
                secW(rs) = 2 * RM_PX + RM_W
            End If
            totSecW = totSecW + secW(rs)
        Next rs

        ' -- Split sections: same room as source (left group) vs others ----
        Dim srcRmLow As String: srcRmLow = LCase(srcNd.Room)
        Dim sameW As Double: sameW = 0:  Dim nSame As Long: nSame = 0
        Dim othW  As Double: othW = 0:   Dim nOth  As Long: nOth = 0
        For rs = 0 To nRS - 1
            If LCase(rs_room(rs)) = srcRmLow Then
                If nSame > 0 Then sameW = sameW + SECT_G
                sameW = sameW + secW(rs): nSame = nSame + 1
            Else
                If nOth > 0 Then othW = othW + SECT_G
                othW = othW + secW(rs): nOth = nOth + 1
            End If
        Next rs

        ' -- Page width: grows with the number of room sections ------------
        Dim srcColW As Double: srcColW = 2 * SRC_PX + SRC_W
        Dim minPW As Double
        minPW = MG + srcColW + ROUTE_W + sameW + MID_GAP + othW + MG
        If minPW < 16# Then minPW = 16#   ' A3 landscape minimum
        gPageW = minPW

        ' -- Section X positions: same-room packed left, others packed right
        Dim curL As Double: curL = MG + srcColW + ROUTE_W
        Dim curR As Double: curR = gPageW - MG
        Dim rs2 As Long
        For rs2 = 0 To nRS - 1
            If LCase(rs_room(rs2)) = srcRmLow Then
                rs_swimL(rs2) = curL
                rs_swimR(rs2) = curL + secW(rs2)
                curL = rs_swimR(rs2) + SECT_G
            Else
                rs_swimR(rs2) = curR
                rs_swimL(rs2) = curR - secW(rs2)
                curR = rs_swimL(rs2) - SECT_G
            End If
            If rs_hasL(rs2) Then
                rs_leftX(rs2) = rs_swimL(rs2) + RM_PX
                rs_rightX(rs2) = rs_swimL(rs2) + RM_PX + RM_W + INNER_G
            Else
                rs_rightX(rs2) = rs_swimL(rs2) + RM_PX
                rs_leftX(rs2) = rs_rightX(rs2)
            End If
        Next rs2

        ' Source device X and height
        Dim lX  As Double: lX = MG + SRC_PX
        Dim lH  As Double: lH = DEV_H + nSlots * PORT_H
        Dim lRkH As Double: lRkH = SRC_PY + SRC_RKH + lH + SRC_PY

        ' Page height
        Dim pH As Double: pH = HDRI + HDR_H + SITE_H + ROOM_H + MG + lRkH + MG
        If pH < 11 Then pH = 11
        gPageH = pH

        ' lTop = Y of top of source device box
        Dim lTop As Double
        lTop = gPageH - HDRI - HDR_H - SITE_H - ROOM_H - MG - SRC_PY - SRC_RKH

        ' -- Create Visio page ---------------------------------------------
        Dim vPg As Object
        If pageN = 0 Then Set vPg = vDoc.Pages(1) Else Set vPg = vDoc.Pages.Add()
        Set gPage = vPg
        Dim pBase As String: pBase = Left(srcNd.Rack & " - " & srcNd.Equip, 31)
        Dim pName As String: pName = pBase
        Dim pSuf As Long: pSuf = 2
        Do While usedNames.Exists(LCase(pName))
            pName = Left(pBase, 27) & " (" & pSuf & ")": pSuf = pSuf + 1
        Loop
        usedNames.Add LCase(pName), True
        gPage.name = pName
        gPage.PageSheet.Cells("PageWidth").Formula = gPageW & " in"
        gPage.PageSheet.Cells("PageHeight").Formula = gPageH & " in"
        pageN = pageN + 1

        ' -- DRAW ----------------------------------------------------------

        ' 1. Swim lanes per room section
        For rs = 0 To nRS - 1
            ' Find first and last global slot in this section
            Dim fsInSec As Long: fsInSec = 999999
            Dim lsInSec As Long: lsInSec = 0
            Dim rni As Long
            For rni = 0 To nRN - 1
                If rn_sec(rni) = rs Then
                    If rn_fsl(rni) < fsInSec Then fsInSec = rn_fsl(rni)
                    If rn_lsl(rni) > lsInSec Then lsInSec = rn_lsl(rni)
                End If
            Next rni
            If fsInSec = 999999 Then GoTo NextRS
            Dim swT As Double: swT = lTop - fsInSec * PORT_H + RM_PY + RM_RKH
            Dim swB As Double: swB = lTop - (lsInSec + 1) * PORT_H - DEV_H - RM_PY

            ' Swim lane background
            Dim swSh As Object
            Set swSh = gPage.DrawRectangle(rs_swimL(rs), swB, rs_swimR(rs), swT)
            swSh.Cells("FillForegnd").Formula = C_SCS
            swSh.Cells("FillBkgnd").Formula = C_SCS
            swSh.Cells("FillPattern").Formula = "1"
            swSh.Cells("LineColor").Formula = C_SCB
            swSh.Cells("LineWeight").Formula = "0.75pt"

            ' Room name banner in the ROOM_H strip
            Dim bnT As Double: bnT = gPageH - HDRI - HDR_H - SITE_H
            Dim bnB As Double: bnB = bnT - ROOM_H
            Dim bnSh As Object
            Set bnSh = MkBox(rs_swimL(rs), bnB, rs_swimR(rs), bnT, 222, 230, 240)
            bnSh.Cells("LineWeight").Formula = "1pt"
            bnSh.text = StrConv(rs_room(rs), vbProperCase)
            bnSh.Cells("Char.Size").Formula = "10pt"
            bnSh.Cells("Char.Style").Formula = "1"
            bnSh.Cells("VerticalAlign").Formula = "1"
            bnSh.Cells("Para.HorzAlign").Formula = "1"
NextRS:
        Next rs

        ' 2. Room section rack boxes and device boxes
        For rni2 = 0 To nRN - 1
            If rn_pc(rni2) = 0 Then GoTo SkipRN   ' no real ports
            Dim devX As Double
            If rn_isL(rni2) Then
                devX = rs_leftX(rn_sec(rni2))
            Else
                devX = rs_rightX(rn_sec(rni2))
            End If
            Dim devTopY As Double: devTopY = lTop - rn_fsl(rni2) * PORT_H
            Dim devHt As Double
            devHt = DEV_H + (rn_lsl(rni2) - rn_fsl(rni2) + 1) * PORT_H

            ' Rack box (wraps device box)
            Dim rkX1 As Double: rkX1 = devX - RM_PX
            Dim rkX2 As Double: rkX2 = devX + RM_W + RM_PX
            Dim rkY1 As Double: rkY1 = devTopY - devHt - RM_PY
            Dim rkY2 As Double: rkY2 = devTopY + RM_RKH + RM_PY
            Dim rkSh As Object
            Set rkSh = gPage.DrawRectangle(rkX1, rkY1, rkX2, rkY2)
            rkSh.Cells("FillForegnd").Formula = C_RKB
            rkSh.Cells("FillBkgnd").Formula = C_RKB
            rkSh.Cells("FillPattern").Formula = "1"
            rkSh.Cells("LineColor").Formula = C_BRK
            rkSh.Cells("LineWeight").Formula = "0.75pt"
            Dim rkHSh As Object
            Set rkHSh = gPage.DrawRectangle(rkX1, rkY2 - RM_RKH, rkX2, rkY2)
            rkHSh.Cells("FillForegnd").Formula = C_RKH
            rkHSh.Cells("FillBkgnd").Formula = C_RKH
            rkHSh.Cells("FillPattern").Formula = "1"
            rkHSh.Cells("LineColor").Formula = C_BRK
            rkHSh.Cells("LineWeight").Formula = "0.5pt"
            rkHSh.text = rn_rack(rni2)
            rkHSh.Cells("Char.Size").Formula = "8pt"
            rkHSh.Cells("Char.Style").Formula = "1"
            rkHSh.Cells("VerticalAlign").Formula = "1"
            rkHSh.Cells("Para.HorzAlign").Formula = "1"

            ' Device box (body + header)
            Dim dvSh As Object
            Set dvSh = gPage.DrawRectangle(devX, devTopY - devHt, devX + RM_W, devTopY)
            dvSh.Cells("FillForegnd").Formula = IIf(rn_pass(rni2), C_OF, C_DF)
            dvSh.Cells("FillBkgnd").Formula = dvSh.Cells("FillForegnd").Formula
            dvSh.Cells("FillPattern").Formula = "1"
            dvSh.Cells("LineColor").Formula = C_BDR
            dvSh.Cells("LineWeight").Formula = "0.75pt"
            Dim dvHSh As Object
            Set dvHSh = gPage.DrawRectangle(devX, devTopY - DEV_H, devX + RM_W, devTopY)
            dvHSh.Cells("FillForegnd").Formula = IIf(rn_pass(rni2), C_OH, C_DH)
            dvHSh.Cells("FillBkgnd").Formula = dvHSh.Cells("FillForegnd").Formula
            dvHSh.Cells("FillPattern").Formula = "1"
            dvHSh.Cells("LineColor").Formula = C_BDR
            dvHSh.Cells("LineWeight").Formula = "0.5pt"
            dvHSh.text = rn_eq(rni2)
            dvHSh.Cells("Char.Size").Formula = "8pt"
            dvHSh.Cells("Char.Style").Formula = "1"
            dvHSh.Cells("VerticalAlign").Formula = "1"
            dvHSh.Cells("Para.HorzAlign").Formula = "1"
SkipRN:
        Next rni2

        ' 3. Source device rack box + device box
        DrawRackBox lX - SRC_PX, lTop - lH - SRC_PY, _
                    lX + SRC_W + SRC_PX, lTop + SRC_RKH + SRC_PY, srcNd.Rack
        ' Source body
        Dim srcSh As Object
        Set srcSh = gPage.DrawRectangle(lX, lTop - lH, lX + SRC_W, lTop)
        srcSh.Cells("FillForegnd").Formula = IIf(srcNd.IsPassive, C_OF, C_DF)
        srcSh.Cells("FillBkgnd").Formula = srcSh.Cells("FillForegnd").Formula
        srcSh.Cells("FillPattern").Formula = "1"
        srcSh.Cells("LineColor").Formula = C_BDR
        srcSh.Cells("LineWeight").Formula = "0.75pt"
        ' Source header
        Dim srcHSh As Object
        Set srcHSh = gPage.DrawRectangle(lX, lTop - DEV_H, lX + SRC_W, lTop)
        srcHSh.Cells("FillForegnd").Formula = IIf(srcNd.IsPassive, C_OH, C_DH)
        srcHSh.Cells("FillBkgnd").Formula = srcHSh.Cells("FillForegnd").Formula
        srcHSh.Cells("FillPattern").Formula = "1"
        srcHSh.Cells("LineColor").Formula = C_BDR
        srcHSh.Cells("LineWeight").Formula = "0.5pt"
        srcHSh.text = srcNd.Equip
        srcHSh.Cells("Char.Size").Formula = "9pt"
        srcHSh.Cells("Char.Style").Formula = "1"
        srcHSh.Cells("VerticalAlign").Formula = "1"
        srcHSh.Cells("Para.HorzAlign").Formula = "1"

        ' 4. Source device room banner
        Dim bnT2 As Double: bnT2 = gPageH - HDRI - HDR_H - SITE_H
        Dim bnB2 As Double: bnB2 = bnT2 - ROOM_H
        Dim srcBnSh As Object
        Set srcBnSh = MkBox(MG, bnB2, MG + srcColW + 0.1, bnT2, 222, 230, 240)
        srcBnSh.Cells("LineWeight").Formula = "1pt"
        srcBnSh.text = StrConv(srcNd.Room, vbProperCase)
        srcBnSh.Cells("Char.Size").Formula = "10pt"
        srcBnSh.Cells("Char.Style").Formula = "1"
        srcBnSh.Cells("VerticalAlign").Formula = "1"
        srcBnSh.Cells("Para.HorzAlign").Formula = "1"

        ' 5. Connection lines, port labels, source separators
        Dim prevSlot As Long: prevSlot = -1
        Dim sg2 As Long
        For sg2 = 0 To nSeg - 1
            Dim sTyp2 As String: sTyp2 = s_ty(sg2)
            Dim aK2   As String: aK2 = s_aK(sg2)
            Dim bK2   As String: bK2 = s_bK(sg2)
            Dim gsl2  As Long:   gsl2 = s_sl(sg2)
            If bK2 = "" Then GoTo SkipSg

            Dim lineY As Double: lineY = lTop - DEV_H - (gsl2 + 0.5) * PORT_H

            ' -- Source-side separator and port label ----------------------
            If aK2 = srcKey Then
                ' Separator before this slot (only once per slot)
                If prevSlot <> gsl2 Then
                    Dim sp2 As Integer: sp2 = sl_sp(gsl2)
                    Dim sepY As Double: sepY = lTop - DEV_H - gsl2 * PORT_H
                    If sp2 = 3 Then DrawSepLine lX + 0.05, sepY, lX + SRC_W - 0.05, C_SRM, "2.5pt"
                    If sp2 = 2 Then DrawSepLine lX + 0.05, sepY, lX + SRC_W - 0.05, C_SR, "1.75pt"
                    If sp2 = 1 Then DrawSepLine lX + 0.05, sepY, lX + SRC_W - 0.05, C_SD, "0.75pt"
                    prevSlot = gsl2
                End If
                ' Source port label (right-aligned within source box)
                If sTyp2 <> "NIS" Then
                    DrawLbl lX, SRC_W, lTop, gsl2, _
                        s_aP(sg2), s_cl(sg2), sTyp2, s_aC(sg2), s_aT(sg2), True
                End If
            End If

            ' -- Continuation End-A port label (LEFT sub-col) --------------
            If aK2 <> srcKey And sTyp2 <> "NIS" Then
                Dim aRnK As String: aRnK = aK2 & "|L"
                If rnMap.Exists(aRnK) Then
                    Dim aRnI As Long: aRnI = CLng(rnMap(aRnK))
                    Dim aTopY As Double: aTopY = lTop - rn_fsl(aRnI) * PORT_H
                    Dim aLocSl As Long:  aLocSl = gsl2 - rn_fsl(aRnI)
                    Dim aDX As Double: aDX = rs_leftX(rn_sec(aRnI))
                    DrawLbl aDX, RM_W, aTopY, aLocSl, _
                        s_aP(sg2), s_cl(sg2), sTyp2, s_aC(sg2), s_aT(sg2), True
                End If
            End If

            ' -- Destination port label (RIGHT sub-col) --------------------
            If sTyp2 <> "NIS" Then
                Dim dRnK As String: dRnK = bK2 & "|R"
                If rnMap.Exists(dRnK) Then
                    Dim dRnI As Long: dRnI = CLng(rnMap(dRnK))
                    Dim dTopY As Double: dTopY = lTop - rn_fsl(dRnI) * PORT_H
                    Dim dLocSl As Long:  dLocSl = gsl2 - rn_fsl(dRnI)
                    Dim dDX As Double: dDX = rs_rightX(rn_sec(dRnI))
                    DrawLbl dDX, RM_W, dTopY, dLocSl, _
                        s_bP(sg2), s_cl(sg2), sTyp2, s_bC(sg2), s_bT(sg2), False
                End If
            End If

            ' -- Connection line -------------------------------------------
            ' Compute the exact anchor point (edge + vertical centre of the
            ' segment's slot row) at BOTH ends, then join them. The line is
            ' drawn for every segment regardless of left/right ordering so
            ' multi-room continuation and NIS links are never dropped.
            Dim fromX As Double, fromY As Double
            Dim toX As Double,   toY As Double
            Dim haveFrom As Boolean: haveFrom = False
            Dim haveTo   As Boolean: haveTo = False

            ' --- End-A anchor (right edge of the End-A device) ---
            If aK2 = srcKey Then
                fromX = lX + SRC_W
                fromY = lineY
                haveFrom = True
            Else
                Dim aRnK2 As String: aRnK2 = aK2 & "|L"
                If Not rnMap.Exists(aRnK2) Then aRnK2 = aK2 & "|R"
                If rnMap.Exists(aRnK2) Then
                    Dim aRnI2 As Long: aRnI2 = CLng(rnMap(aRnK2))
                    Dim aDevX2 As Double
                    aDevX2 = IIf(rn_isL(aRnI2), rs_leftX(rn_sec(aRnI2)), rs_rightX(rn_sec(aRnI2)))
                    fromX = aDevX2 + RM_W
                    fromY = PortRowY(lTop, rn_fsl(aRnI2), gsl2)
                    haveFrom = True
                End If
            End If

            ' --- End-B anchor (left edge of the End-B device) ---
            If sTyp2 = "NIS" Then
                Dim nisBK As String: nisBK = bK2 & "|L"
                If rnMap.Exists(nisBK) Then
                    Dim nisBRnI As Long: nisBRnI = CLng(rnMap(nisBK))
                    toX = rs_leftX(rn_sec(nisBRnI))
                    toY = PortRowY(lTop, rn_fsl(nisBRnI), gsl2)
                    haveTo = True
                End If
            Else
                Dim dRnK2 As String: dRnK2 = bK2 & "|R"
                If rnMap.Exists(dRnK2) Then
                    Dim dRnI2 As Long: dRnI2 = CLng(rnMap(dRnK2))
                    toX = rs_rightX(rn_sec(dRnI2))
                    toY = PortRowY(lTop, rn_fsl(dRnI2), gsl2)
                    haveTo = True
                End If
            End If

            If haveFrom And haveTo Then
                DrawCktLine fromX, fromY, toX, toY, sTyp2

                ' Length label (direct circuits only, mid-span)
                If aK2 = srcKey And sTyp2 <> "NIS" And Trim(s_ln(sg2)) <> "" Then
                    Dim lenTxt As String: lenTxt = Trim(s_ln(sg2))
                    If LCase(Right(lenTxt, 1)) <> "m" Then lenTxt = lenTxt & "m"
                    Dim mx As Double: mx = (fromX + toX) / 2
                    Dim my As Double: my = (fromY + toY) / 2
                    TxtBx mx - 0.3, my + 0.02, mx + 0.3, my + PORT_H * 0.42, lenTxt, 7, 1
                End If

                ' NIS label
                If sTyp2 = "NIS" Then
                    Dim mx2 As Double: mx2 = (fromX + toX) / 2
                    Dim my2 As Double: my2 = (fromY + toY) / 2
                    TxtBx mx2 - 0.25, my2 + 0.02, mx2 + 0.25, my2 + PORT_H * 0.42, "NIS", 7, 1
                End If
            End If
SkipSg:
        Next sg2

        ' 6. Page header
        DrawPageHeader
NextSrc:
    Next si

    MsgBox "Done - " & pageN & " page(s) created.", vbInformation
    Exit Sub
ErrH:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbLf & _
           Erl, vbCritical
End Sub

'===========================================================================
' Vertical centre (Y) of a segment's slot row within a device box.
' devTopY = Y of top of the source device box (lTop).
' firstSlot = the device box's first global slot (rn_fsl).
' globalSlot = the segment's global slot.
'===========================================================================
Private Function PortRowY(devTopY As Double, firstSlot As Long, globalSlot As Long) As Double
    Dim localSlot As Long: localSlot = globalSlot - firstSlot
    PortRowY = devTopY - firstSlot * PORT_H - DEV_H - (localSlot + 0.5) * PORT_H
End Function

'===========================================================================
' DRAW - PORT LABEL
' devTopY = Y of top of device box (= lTop for source, lTop-fsl*PORT_H for room nodes)
' localSlot = slot index relative to this device (0-based)
' rightAligned: True = text right-aligned (source / LEFT sub-col devices)
'               False = text left-aligned  (RIGHT sub-col devices)
'===========================================================================
Private Sub DrawLbl(devX As Double, devWd As Double, devTopY As Double, _
    localSlot As Long, portTxt As String, calTxt As String, cabTyp As String, _
    connTxt As String, txTxt As String, rightAligned As Boolean)
    If portTxt = "" And calTxt = "" Then Exit Sub
    Dim rowT As Double: rowT = devTopY - DEV_H - localSlot * PORT_H
    Dim rowB As Double: rowB = rowT - PORT_H
    Dim sh As Object
    Set sh = gPage.DrawRectangle(devX + 0.04, rowB + 0.02, devX + devWd - 0.04, rowT - 0.02)
    Dim lbl As String
    lbl = Trim(portTxt)
    If calTxt <> "" Then lbl = lbl & "  " & calTxt
    If cabTyp = "Fibre" Then
        If connTxt <> "" Then lbl = lbl & "  " & connTxt
        If txTxt <> "" Then lbl = lbl & "  " & txTxt
    End If
    sh.text = Trim(lbl)
    sh.Cells("Char.Size").Formula = "7pt"
    sh.Cells("VerticalAlign").Formula = "1"
    sh.Cells("Para.HorzAlign").Formula = IIf(rightAligned, "2", "0")
    sh.Cells("LinePattern").Formula = "0"
    sh.Cells("FillPattern").Formula = "0"
End Sub

'===========================================================================
' DRAW - CONNECTION LINES
' Joins (x1,y1) to (x2,y2). When the two anchors share a height the link is a
' single horizontal line; otherwise it is routed orthogonally (L-shaped) via
' a mid-span vertical so the link always lands on both device edges instead
' of leaving a stray diagonal.
'===========================================================================
Private Sub DrawCktLine(x1 As Double, y1 As Double, x2 As Double, y2 As Double, cabTyp As String)
    If Abs(x2 - x1) < 0.0001 And Abs(y2 - y1) < 0.0001 Then Exit Sub
    Dim col As String, pat As String, wt As String
    Select Case LCase(cabTyp)
        Case "fibre":  col = F_COL: pat = F_PAT: wt = F_WT
        Case "copper": col = P_COL: pat = P_PAT: wt = P_WT
        Case "nis":    col = N_COL: pat = N_PAT: wt = N_WT
        Case Else:     col = D_COL: pat = D_PAT: wt = D_WT  ' DAC
    End Select
    If Abs(y2 - y1) < 0.0001 Then
        DrawStyledLine x1, y1, x2, y2, col, pat, wt
    Else
        Dim midX As Double: midX = (x1 + x2) / 2
        DrawStyledLine x1, y1, midX, y1, col, pat, wt
        DrawStyledLine midX, y1, midX, y2, col, pat, wt
        DrawStyledLine midX, y2, x2, y2, col, pat, wt
    End If
End Sub

Private Sub DrawStyledLine(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
    col As String, pat As String, wt As String)
    Dim sh As Object: Set sh = gPage.DrawLine(x1, y1, x2, y2)
    sh.Cells("LineColor").Formula = col
    sh.Cells("LinePattern").Formula = pat
    sh.Cells("LineWeight").Formula = wt
    sh.Cells("BeginArrow").Formula = "0"
    sh.Cells("EndArrow").Formula = "0"
End Sub

'===========================================================================
' DRAW - SEPARATOR LINES ON SOURCE DEVICE
'===========================================================================
Private Sub DrawSepLine(x1 As Double, y As Double, x2 As Double, _
    col As String, wt As String)
    Dim sh As Object: Set sh = gPage.DrawLine(x1, y, x2, y)
    sh.Cells("LineColor").Formula = col
    sh.Cells("LinePattern").Formula = "1"
    sh.Cells("LineWeight").Formula = wt
    sh.Cells("BeginArrow").Formula = "0"
    sh.Cells("EndArrow").Formula = "0"
End Sub

'===========================================================================
' DRAW - SOURCE RACK BOX
'===========================================================================
Private Sub DrawRackBox(x1 As Double, y1 As Double, x2 As Double, y2 As Double, rk As String)
    Dim sh As Object
    Set sh = gPage.DrawRectangle(x1, y1, x2, y2)
    sh.Cells("FillForegnd").Formula = C_RKB
    sh.Cells("FillBkgnd").Formula = C_RKB
    sh.Cells("FillPattern").Formula = "1"
    sh.Cells("LineColor").Formula = C_BRK
    sh.Cells("LineWeight").Formula = "1pt"
    Dim hdrY As Double: hdrY = y2 - SRC_RKH
    Dim hSh As Object
    Set hSh = gPage.DrawRectangle(x1, hdrY, x2, y2)
    hSh.Cells("FillForegnd").Formula = C_RKH
    hSh.Cells("FillBkgnd").Formula = C_RKH
    hSh.Cells("FillPattern").Formula = "1"
    hSh.Cells("LineColor").Formula = C_BRK
    hSh.Cells("LineWeight").Formula = "0.5pt"
    hSh.text = rk
    hSh.Cells("Char.Size").Formula = "9pt"
    hSh.Cells("Char.Style").Formula = "1"
    hSh.Cells("VerticalAlign").Formula = "1"
    hSh.Cells("Para.HorzAlign").Formula = "1"
End Sub

'===========================================================================
' DRAW - PAGE HEADER
'===========================================================================
Private Sub DrawPageHeader()
    Dim hT As Double: hT = gPageH - HDRI
    Dim hB As Double: hB = hT - HDR_H
    Dim sT As Double: sT = hB
    Dim sB As Double: sB = hB - SITE_H

    ' Left panel: author / filename
    Dim lSh As Object
    Set lSh = MkBox(MG, hB, 3.9, hT, 245, 248, 252)
    lSh.text = "Author:" & Chr(10) & "Filename:"
    lSh.Cells("Char.Size").Formula = "8pt"
    lSh.Cells("VerticalAlign").Formula = "0"
    lSh.Cells("Para.HorzAlign").Formula = "0"

    ' Centre panel: cable type key
    Dim kH As Double: kH = hT - hB
    Dim y1 As Double: y1 = hB + kH * 0.75
    Dim y2 As Double: y2 = hB + kH * 0.5
    Dim y3 As Double: y3 = hB + kH * 0.25
    MkBox 4#, hB, 7.5, hT, 255, 255, 255
    TxtBx 4.05, y1 - 0.1, 4.6, y1 + 0.1, "SMF Fibre", 8, 0
    TxtBx 4.05, y2 - 0.1, 4.6, y2 + 0.1, "DAC", 8, 0
    TxtBx 4.05, y3 - 0.1, 4.6, y3 + 0.1, "Copper", 8, 0
    DrawSeg 4.62, y1, 5.8, y1, F_COL, F_PAT, F_WT
    DrawSeg 4.62, y2, 5.8, y2, D_COL, D_PAT, D_WT
    DrawSeg 4.62, y3, 5.8, y3, P_COL, P_PAT, P_WT
    TxtBx 5.85, y1 - 0.1, 6.3, y1 + 0.1, "NIS ->", 8, 0
    DrawSeg 6.35, y1, 7.4, y1, N_COL, N_PAT, N_WT

    ' Right panel: classification
    Dim rSh As Object
    Set rSh = MkBox(gPageW - 5#, hB, gPageW - MG, hT, 245, 248, 252)
    rSh.text = "Classification = Internal"
    rSh.Cells("Char.Size").Formula = "10pt"
    rSh.Cells("Char.Style").Formula = "1"
    rSh.Cells("VerticalAlign").Formula = "1"
    rSh.Cells("Para.HorzAlign").Formula = "1"

    ' Site banner
    Dim siteName As String: siteName = GetCell("D9")
    If siteName = "" Then siteName = "Schematic"
    Dim acronym  As String: acronym = GetCell("B9")
    Dim siteCode As String: siteCode = GetCell("C9")
    Dim sSh As Object
    Set sSh = MkBox(MG, sB, gPageW - MG, sT, 255, 255, 255)
    sSh.Cells("LineWeight").Formula = "1.2pt"
    sSh.text = siteName
    sSh.Cells("Char.Size").Formula = "13pt"
    sSh.Cells("Char.Style").Formula = "1"
    sSh.Cells("VerticalAlign").Formula = "1"
    sSh.Cells("Para.HorzAlign").Formula = "1"
    If acronym <> "" Then TxtBx MG + 0.1, sB + 0.04, MG + 2#, sT - 0.04, acronym, 9, 0
    If siteCode <> "" Then TxtBx gPageW - MG - 2#, sB + 0.04, gPageW - MG - 0.1, sT - 0.04, siteCode, 9, 2
End Sub

'===========================================================================
' PRIMITIVES
'===========================================================================
Private Sub DrawSeg(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
    col As String, pat As String, wt As String)
    Dim sh As Object: Set sh = gPage.DrawLine(x1, y1, x2, y2)
    sh.Cells("LineColor").Formula = col
    sh.Cells("LinePattern").Formula = pat
    sh.Cells("LineWeight").Formula = wt
    sh.Cells("BeginArrow").Formula = "0"
    sh.Cells("EndArrow").Formula = "0"
End Sub

Private Function MkBox(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
    r As Long, g As Long, b As Long) As Object
    Set MkBox = gPage.DrawRectangle(x1, y1, x2, y2)
    MkBox.Cells("FillForegnd").Formula = "THEMEGUARD(RGB(" & r & "," & g & "," & b & "))"
    MkBox.Cells("FillBkgnd").Formula = MkBox.Cells("FillForegnd").Formula
    MkBox.Cells("FillPattern").Formula = "1"
    MkBox.Cells("LineColor").Formula = C_BDR
End Function

Private Sub TxtBx(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
    txt As String, pts As Integer, hAlign As Integer)
    If Trim(txt) = "" Then Exit Sub
    Dim sh As Object: Set sh = gPage.DrawRectangle(x1, y1, x2, y2)
    sh.text = txt
    sh.Cells("Char.Size").Formula = pts & "pt"
    sh.Cells("VerticalAlign").Formula = "1"
    sh.Cells("Para.HorzAlign").Formula = CStr(hAlign)
    sh.Cells("LinePattern").Formula = "0"
    sh.Cells("FillPattern").Formula = "0"
End Sub

'===========================================================================
' DATA HELPERS
'===========================================================================
Private Function IsCircRow(ws As Worksheet, r As Long) As Boolean
    Dim v As Variant: v = ws.Cells(r, CC_REF).value
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    Dim d As Double
    On Error Resume Next: d = CDbl(v)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    IsCircRow = (d > 0) And (Int(d) = d)
End Function

Private Function SC(ws As Worksheet, r As Long, c As Long) As String
    Dim v As Variant: v = ws.Cells(r, c).value
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then SC = "": Exit Function
    SC = Trim(CStr(v))
End Function

Private Function NK(rm As String, rk As String, eq As String) As String
    NK = LCase(Trim(rm)) & "|" & LCase(Trim(rk)) & "|" & LCase(Trim(eq))
End Function

Private Function IsPassive(eq As String) As Boolean
    Dim s As String: s = LCase(Trim(eq))
    IsPassive = (InStr(s, "odf") > 0 Or InStr(s, "patch") > 0 Or _
                 InStr(s, "panel") > 0 Or InStr(s, "cass") > 0)
End Function

Private Function FindSheet(nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase(ws.name) = LCase(nm) Then Set FindSheet = ws: Exit Function
    Next ws
End Function

Private Function GetCell(addr As String) As String
    Dim ws As Worksheet: Set ws = FindSheet(SHT_FRT)
    If ws Is Nothing Then GetCell = "": Exit Function
    Dim v As Variant: v = ws.Range(addr).value
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then GetCell = "": Exit Function
    GetCell = Trim(CStr(v))
End Function
