Option Explicit
'===========================================================================
' SchematicGenerator  v15
' Hop-ordered room layout | CAL chain tracing | DeviceNames classification
'
' clsCircuit: Cable, CableType, FromKey, ToKey, FromPort, ToPort,
'             Cal, Length, ConnectorA, TxA, ConnectorB, TxB
' clsNode:    Key, Room, Rack, Equip, IsPassive
' DeviceNames sheet col A = generic device name stems
'
' v15 layout model:
'   * One page per source device; each circuit / CAL chain is one horizontal
'     row (slot). Devices shared across circuits are drawn once as a tall box.
'   * Each device is placed left-to-right by its "hop distance" from the source.
'     The A-end (source) is pinned to the far left; the terminal B-end is pinned
'     to the far right.
'   * Rooms divide the page into EQUAL-WIDTH zones (dynamic: 1 room -> hops
'     spread wide and centred; N rooms -> page split into N zones with dividers).
'   * Because every hop of a chain shares the same row Y, links are clean
'     straight horizontal lines that join the device edges.
'===========================================================================

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
Private Const MIN_PW  As Double = 16#    ' A3 landscape minimum page width

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

'-- HORIZONTAL SPACING -----------------------------------------------------
Private Const HOP_GAP  As Double = 1#    ' target gap between hops within a room
Private Const ZONE_PAD As Double = 0.4   ' inset from a room zone's edges

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
Private Const C_SCB As String = "THEMEGUARD(RGB(105,138,175))"
Private Const C_SCS As String = "THEMEGUARD(RGB(240,244,250))"
Private Const C_DIV As String = "THEMEGUARD(RGB(150,165,185))"   ' room divider line

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

        '===================================================================
        ' Build the hop-ordered node model
        '   rn_*  : one entry per distinct device (node key)
        '   rs_*  : one entry per distinct room
        '===================================================================
        Const MAXRN As Long = 300
        Dim rn_key(MAXRN)  As String
        Dim rn_rack(MAXRN) As String
        Dim rn_eq(MAXRN)   As String
        Dim rn_pass(MAXRN) As Boolean
        Dim rn_src(MAXRN)  As Boolean
        Dim rn_sec(MAXRN)  As Long      ' room section index
        Dim rn_fsl(MAXRN)  As Long      ' first slot
        Dim rn_lsl(MAXRN)  As Long      ' last slot
        Dim rn_hop(MAXRN)  As Long      ' hop distance from source (-1 = unset)
        Dim rn_col(MAXRN)  As Long      ' column within its room
        Dim rn_x(MAXRN)    As Double    ' left edge X of device box
        Dim nRN As Long: nRN = 0
        Dim rnMap As Object: Set rnMap = CreateObject("Scripting.Dictionary")

        Const MAXRS As Long = 40
        Dim rs_room(MAXRS)   As String
        Dim rs_minHop(MAXRS) As Long
        Dim rs_maxHop(MAXRS) As Long
        Dim rs_rank(MAXRS)   As Long
        Dim nRS As Long: nRS = 0
        Dim rsMap As Object: Set rsMap = CreateObject("Scripting.Dictionary")

        ' Pre-create the source node as hop 0
        Dim srcIdx As Long
        srcIdx = EnsureRoomNode(srcKey, srcNd, True, nodesInit:=True, _
            rn_key:=rn_key, rn_rack:=rn_rack, rn_eq:=rn_eq, rn_pass:=rn_pass, _
            rn_src:=rn_src, rn_sec:=rn_sec, rn_fsl:=rn_fsl, rn_lsl:=rn_lsl, _
            rn_hop:=rn_hop, nRN:=nRN, rnMap:=rnMap, _
            rs_room:=rs_room, nRS:=nRS, rsMap:=rsMap)
        rn_hop(srcIdx) = 0

        Dim sg As Long
        For sg = 0 To nSeg - 1
            Dim aK As String: aK = s_aK(sg)
            Dim bK As String: bK = s_bK(sg)
            Dim gsl As Long:  gsl = s_sl(sg)
            If bK = "" Then GoTo SkipSeg

            Dim aIdx As Long, bIdx As Long
            Dim aNode As clsNode: Set aNode = allNodes(aK)
            Dim bNode As clsNode: Set bNode = allNodes(bK)
            aIdx = EnsureRoomNode(aK, aNode, (aK = srcKey), nodesInit:=False, _
                rn_key:=rn_key, rn_rack:=rn_rack, rn_eq:=rn_eq, rn_pass:=rn_pass, _
                rn_src:=rn_src, rn_sec:=rn_sec, rn_fsl:=rn_fsl, rn_lsl:=rn_lsl, _
                rn_hop:=rn_hop, nRN:=nRN, rnMap:=rnMap, _
                rs_room:=rs_room, nRS:=nRS, rsMap:=rsMap)
            bIdx = EnsureRoomNode(bK, bNode, False, nodesInit:=False, _
                rn_key:=rn_key, rn_rack:=rn_rack, rn_eq:=rn_eq, rn_pass:=rn_pass, _
                rn_src:=rn_src, rn_sec:=rn_sec, rn_fsl:=rn_fsl, rn_lsl:=rn_lsl, _
                rn_hop:=rn_hop, nRN:=nRN, rnMap:=rnMap, _
                rs_room:=rs_room, nRS:=nRS, rsMap:=rsMap)

            ' Hop distance: B is one hop past A. Use the LONGEST path so each
            ' device settles at its deepest column and links never run backwards.
            If rn_hop(aIdx) < 0 Then rn_hop(aIdx) = 0
            Dim h As Long: h = rn_hop(aIdx) + 1
            If h > rn_hop(bIdx) Then rn_hop(bIdx) = h

            UpdateSlot aIdx, gsl, rn_fsl, rn_lsl
            UpdateSlot bIdx, gsl, rn_fsl, rn_lsl
SkipSeg:
        Next sg
        If nRN = 0 Then GoTo NextSrc

        ' Any device never reached keeps hop 0 (normally only the source)
        Dim rni As Long
        For rni = 0 To nRN - 1
            If rn_hop(rni) < 0 Then rn_hop(rni) = 0
        Next rni

        '===================================================================
        ' Horizontal layout: ONE COLUMN PER HOP DEPTH.
        ' Every device the same distance from the source shares a column, so
        ' all the ODFs in a room line up vertically and links stay horizontal.
        ' Column 0 is the source (far left); the deepest hop sits far right.
        '===================================================================
        Dim maxHop As Long: maxHop = 0
        For rni = 0 To nRN - 1
            If rn_hop(rni) > maxHop Then maxHop = rn_hop(rni)
        Next rni

        ' Column left-edge X for each hop (colX(0) = source)
        Dim colX(400) As Double
        Dim fixedBoxes As Double: fixedBoxes = SRC_W + maxHop * RM_W
        Dim gap As Double: gap = HOP_GAP
        Dim contentW As Double: contentW = fixedBoxes + maxHop * gap
        gPageW = MG + SRC_PX + contentW + SRC_PX + MG
        If gPageW < MIN_PW Then
            gPageW = MIN_PW
            If maxHop > 0 Then gap = (gPageW - 2 * MG - 2 * SRC_PX - fixedBoxes) / maxHop
            If gap < HOP_GAP Then gap = HOP_GAP
        End If
        Dim hcx As Double: hcx = MG + SRC_PX
        Dim hh As Long
        For hh = 0 To maxHop
            colX(hh) = hcx
            hcx = hcx + IIf(hh = 0, SRC_W, RM_W) + gap
        Next hh

        ' Assign each device its column X by hop depth
        For rni = 0 To nRN - 1
            rn_x(rni) = colX(rn_hop(rni))
        Next rni

        ' Per-room hop-column range (drives the room bands + dividers)
        Dim rs As Long
        For rs = 0 To nRS - 1: rs_minHop(rs) = 999999: rs_maxHop(rs) = -1: Next rs
        For rni = 0 To nRN - 1
            Dim scn As Long: scn = rn_sec(rni)
            If rn_hop(rni) < rs_minHop(scn) Then rs_minHop(scn) = rn_hop(rni)
            If rn_hop(rni) > rs_maxHop(scn) Then rs_maxHop(scn) = rn_hop(rni)
        Next rni

        ' Rank rooms left-to-right by minimum hop (source room = 0)
        Dim ord(MAXRS) As Long
        For rs = 0 To nRS - 1: ord(rs) = rs: Next rs
        Dim oi As Long, oj As Long
        For oi = 0 To nRS - 2
            For oj = 0 To nRS - 2 - oi
                If rs_minHop(ord(oj)) > rs_minHop(ord(oj + 1)) Then
                    Dim ot As Long: ot = ord(oj): ord(oj) = ord(oj + 1): ord(oj + 1) = ot
                End If
            Next oj
        Next oi
        For rs = 0 To nRS - 1: rs_rank(ord(rs)) = rs: Next rs

        ' Room band boundaries (midpoints between adjacent rooms, tiled edge-to-edge)
        Dim bnd(MAXRS) As Double
        bnd(0) = MG
        bnd(nRS) = gPageW - MG
        Dim bb As Long
        For bb = 1 To nRS - 1
            Dim lRoom As Long: lRoom = ord(bb - 1)
            Dim rRoom As Long: rRoom = ord(bb)
            Dim lEdge As Double: lEdge = colX(rs_maxHop(lRoom)) + BoxWAtHop(rs_maxHop(lRoom))
            Dim rEdge As Double: rEdge = colX(rs_minHop(rRoom))
            bnd(bb) = (lEdge + rEdge) / 2
        Next bb

        '===================================================================
        ' Vertical layout & page
        '===================================================================
        Dim lH  As Double: lH = DEV_H + nSlots * PORT_H
        Dim lRkH As Double: lRkH = SRC_PY + SRC_RKH + lH + SRC_PY
        Dim pH As Double: pH = HDRI + HDR_H + SITE_H + ROOM_H + MG + lRkH + MG
        If pH < 11 Then pH = 11
        gPageH = pH
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

        '===================================================================
        ' DRAW
        '===================================================================
        Dim bandT As Double: bandT = gPageH - HDRI - HDR_H - SITE_H
        Dim bandB As Double: bandB = bandT - ROOM_H

        ' 1. Room zones: swim-lane background + header band (tiled edge-to-edge)
        For rs = 0 To nRS - 1
            Dim ra As Long: ra = rs_rank(rs)
            Dim zL As Double: zL = bnd(ra)
            Dim zR As Double: zR = bnd(ra + 1)

            ' vertical extent of this room's devices
            Dim fsInSec As Long: fsInSec = 999999
            Dim lsInSec As Long: lsInSec = 0
            For rni = 0 To nRN - 1
                If rn_sec(rni) = rs Then
                    If rn_fsl(rni) < fsInSec Then fsInSec = rn_fsl(rni)
                    If rn_lsl(rni) > lsInSec Then lsInSec = rn_lsl(rni)
                End If
            Next rni
            If fsInSec = 999999 Then GoTo NextZone
            Dim swT As Double: swT = lTop - fsInSec * PORT_H + RM_PY + RM_RKH
            Dim swB As Double: swB = lTop - (lsInSec + 1) * PORT_H - DEV_H - RM_PY

            Dim swSh As Object
            Set swSh = gPage.DrawRectangle(zL, swB, zR, swT)
            swSh.Cells("FillForegnd").Formula = C_SCS
            swSh.Cells("FillBkgnd").Formula = C_SCS
            swSh.Cells("FillPattern").Formula = "1"
            swSh.Cells("LineColor").Formula = C_SCB
            swSh.Cells("LineWeight").Formula = "0.75pt"

            Dim bnSh As Object
            Set bnSh = MkBox(zL, bandB, zR, bandT, 222, 230, 240)
            bnSh.Cells("LineWeight").Formula = "1pt"
            bnSh.text = StrConv(rs_room(rs), vbProperCase)
            bnSh.Cells("Char.Size").Formula = "10pt"
            bnSh.Cells("Char.Style").Formula = "1"
            bnSh.Cells("VerticalAlign").Formula = "1"
            bnSh.Cells("Para.HorzAlign").Formula = "1"
NextZone:
        Next rs

        ' 2. Room divider lines at the band boundaries
        Dim contTop As Double: contTop = lTop + SRC_RKH + SRC_PY
        Dim contBot As Double: contBot = lTop - nSlots * PORT_H - DEV_H - SRC_PY
        Dim dv As Long
        For dv = 1 To nRS - 1
            Dim dvX As Double: dvX = bnd(dv)
            Dim dvSh0 As Object: Set dvSh0 = gPage.DrawLine(dvX, contBot, dvX, contTop)
            dvSh0.Cells("LineColor").Formula = C_DIV
            dvSh0.Cells("LinePattern").Formula = "2"
            dvSh0.Cells("LineWeight").Formula = "1pt"
            dvSh0.Cells("BeginArrow").Formula = "0"
            dvSh0.Cells("EndArrow").Formula = "0"
        Next dv

        ' 3. Device boxes (rack wrapper + body + header) for every hop
        For rni = 0 To nRN - 1
            DrawDeviceBox rn_x(rni), lTop, rn_fsl(rni), rn_lsl(rni), _
                rn_src(rni), rn_pass(rni), rn_rack(rni), rn_eq(rni)
        Next rni

        ' 4. Source separators (room / rack / device bands on the source box)
        Dim srcX As Double: srcX = rn_x(srcIdx)
        Dim gs As Long
        For gs = 0 To nSlots - 1
            Dim sp As Integer: sp = sl_sp(gs)
            If sp = 0 Then GoTo NextSep
            Dim sepY As Double: sepY = lTop - DEV_H - gs * PORT_H
            If sp = 3 Then DrawSepLine srcX + 0.05, sepY, srcX + SRC_W - 0.05, C_SRM, "2.5pt"
            If sp = 2 Then DrawSepLine srcX + 0.05, sepY, srcX + SRC_W - 0.05, C_SR, "1.75pt"
            If sp = 1 Then DrawSepLine srcX + 0.05, sepY, srcX + SRC_W - 0.05, C_SD, "0.75pt"
NextSep:
        Next gs

        ' 5. Connection lines + port labels
        Dim sg2 As Long
        For sg2 = 0 To nSeg - 1
            Dim aK2 As String: aK2 = s_aK(sg2)
            Dim bK2 As String: bK2 = s_bK(sg2)
            Dim gsl2 As Long:  gsl2 = s_sl(sg2)
            Dim typ2 As String: typ2 = s_ty(sg2)
            If bK2 = "" Then GoTo SkipDraw
            If Not (rnMap.Exists(aK2) And rnMap.Exists(bK2)) Then GoTo SkipDraw

            Dim ai As Long: ai = CLng(rnMap(aK2))
            Dim bi As Long: bi = CLng(rnMap(bK2))
            Dim lineY As Double: lineY = lTop - DEV_H - (gsl2 + 0.5) * PORT_H

            ' Outgoing (End-A) port label, right-aligned near the right edge
            If typ2 <> "NIS" Then
                Dim aTopY As Double: aTopY = lTop - rn_fsl(ai) * PORT_H
                DrawLbl rn_x(ai), BoxW(rn_src(ai)), aTopY, gsl2 - rn_fsl(ai), _
                    s_aP(sg2), s_cl(sg2), typ2, s_aC(sg2), s_aT(sg2), True
                ' Incoming (End-B) port label, left-aligned near the left edge
                Dim bTopY As Double: bTopY = lTop - rn_fsl(bi) * PORT_H
                DrawLbl rn_x(bi), BoxW(rn_src(bi)), bTopY, gsl2 - rn_fsl(bi), _
                    s_bP(sg2), s_cl(sg2), typ2, s_bC(sg2), s_bT(sg2), False
            End If

            ' Connection line: A right edge -> B left edge (same Y => straight)
            Dim fromX As Double: fromX = rn_x(ai) + BoxW(rn_src(ai))
            Dim toX As Double:   toX = rn_x(bi)
            DrawCktLine fromX, lineY, toX, lineY, typ2

            ' Length label (direct circuits only, mid-span)
            If aK2 = srcKey And typ2 <> "NIS" And Trim(s_ln(sg2)) <> "" Then
                Dim lenTxt As String: lenTxt = Trim(s_ln(sg2))
                If LCase(Right(lenTxt, 1)) <> "m" Then lenTxt = lenTxt & "m"
                Dim mx As Double: mx = (fromX + toX) / 2
                TxtBx mx - 0.3, lineY + 0.02, mx + 0.3, lineY + PORT_H * 0.42, lenTxt, 7, 1
            End If

            ' NIS label mid-span
            If typ2 = "NIS" Then
                Dim mx2 As Double: mx2 = (fromX + toX) / 2
                TxtBx mx2 - 0.25, lineY + 0.02, mx2 + 0.25, lineY + PORT_H * 0.42, "NIS", 7, 1
            End If
SkipDraw:
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
' Ensure a room-node entry exists; ensure its room section exists.
' Returns the rn index.
'===========================================================================
Private Function EnsureRoomNode(k As String, node As clsNode, isSrc As Boolean, _
    nodesInit As Boolean, _
    rn_key() As String, rn_rack() As String, rn_eq() As String, _
    rn_pass() As Boolean, rn_src() As Boolean, rn_sec() As Long, _
    rn_fsl() As Long, rn_lsl() As Long, rn_hop() As Long, _
    nRN As Long, rnMap As Object, _
    rs_room() As String, nRS As Long, rsMap As Object) As Long

    If rnMap.Exists(k) Then EnsureRoomNode = CLng(rnMap(k)): Exit Function

    ' Room section
    Dim rmKey As String: rmKey = LCase(node.Room)
    If Not rsMap.Exists(rmKey) Then
        rsMap.Add rmKey, nRS
        rs_room(nRS) = node.Room: nRS = nRS + 1
    End If

    Dim idx As Long: idx = nRN
    rnMap.Add k, idx
    rn_key(idx) = k
    rn_rack(idx) = node.Rack
    rn_eq(idx) = node.Equip
    rn_pass(idx) = node.IsPassive
    rn_src(idx) = isSrc
    rn_sec(idx) = CLng(rsMap(rmKey))
    rn_fsl(idx) = 999999
    rn_lsl(idx) = -1
    rn_hop(idx) = -1
    nRN = nRN + 1
    EnsureRoomNode = idx
End Function

Private Sub UpdateSlot(idx As Long, gsl As Long, rn_fsl() As Long, rn_lsl() As Long)
    If gsl < rn_fsl(idx) Then rn_fsl(idx) = gsl
    If gsl > rn_lsl(idx) Then rn_lsl(idx) = gsl
End Sub

Private Function BoxW(isSrc As Boolean) As Double
    BoxW = IIf(isSrc, SRC_W, RM_W)
End Function

' Box width for the device that sits at a given hop column (hop 0 = source).
Private Function BoxWAtHop(hop As Long) As Double
    BoxWAtHop = IIf(hop = 0, SRC_W, RM_W)
End Function

'===========================================================================
' DRAW - ONE DEVICE (rack wrapper + body + header)
'===========================================================================
Private Sub DrawDeviceBox(devX As Double, lTop As Double, fsl As Long, lsl As Long, _
    isSrc As Boolean, isPass As Boolean, rackTxt As String, eqTxt As String)
    If lsl < fsl Then Exit Sub
    Dim bw As Double: bw = BoxW(isSrc)
    Dim px As Double: px = IIf(isSrc, SRC_PX, RM_PX)
    Dim py As Double: py = IIf(isSrc, SRC_PY, RM_PY)
    Dim rkh As Double: rkh = IIf(isSrc, SRC_RKH, RM_RKH)
    Dim fpt As String: fpt = IIf(isSrc, "9pt", "8pt")

    Dim devTopY As Double: devTopY = lTop - fsl * PORT_H
    Dim devHt As Double: devHt = DEV_H + (lsl - fsl + 1) * PORT_H

    ' Rack box
    Dim rkX1 As Double: rkX1 = devX - px
    Dim rkX2 As Double: rkX2 = devX + bw + px
    Dim rkY1 As Double: rkY1 = devTopY - devHt - py
    Dim rkY2 As Double: rkY2 = devTopY + rkh + py
    Dim rkSh As Object: Set rkSh = gPage.DrawRectangle(rkX1, rkY1, rkX2, rkY2)
    rkSh.Cells("FillForegnd").Formula = C_RKB
    rkSh.Cells("FillBkgnd").Formula = C_RKB
    rkSh.Cells("FillPattern").Formula = "1"
    rkSh.Cells("LineColor").Formula = C_BRK
    rkSh.Cells("LineWeight").Formula = IIf(isSrc, "1pt", "0.75pt")
    Dim rkHSh As Object: Set rkHSh = gPage.DrawRectangle(rkX1, rkY2 - rkh, rkX2, rkY2)
    rkHSh.Cells("FillForegnd").Formula = C_RKH
    rkHSh.Cells("FillBkgnd").Formula = C_RKH
    rkHSh.Cells("FillPattern").Formula = "1"
    rkHSh.Cells("LineColor").Formula = C_BRK
    rkHSh.Cells("LineWeight").Formula = "0.5pt"
    rkHSh.text = rackTxt
    rkHSh.Cells("Char.Size").Formula = fpt
    rkHSh.Cells("Char.Style").Formula = "1"
    rkHSh.Cells("VerticalAlign").Formula = "1"
    rkHSh.Cells("Para.HorzAlign").Formula = "1"

    ' Device body
    Dim dvSh As Object: Set dvSh = gPage.DrawRectangle(devX, devTopY - devHt, devX + bw, devTopY)
    dvSh.Cells("FillForegnd").Formula = IIf(isPass, C_OF, C_DF)
    dvSh.Cells("FillBkgnd").Formula = dvSh.Cells("FillForegnd").Formula
    dvSh.Cells("FillPattern").Formula = "1"
    dvSh.Cells("LineColor").Formula = C_BDR
    dvSh.Cells("LineWeight").Formula = "0.75pt"
    ' Device header
    Dim dvHSh As Object: Set dvHSh = gPage.DrawRectangle(devX, devTopY - DEV_H, devX + bw, devTopY)
    dvHSh.Cells("FillForegnd").Formula = IIf(isPass, C_OH, C_DH)
    dvHSh.Cells("FillBkgnd").Formula = dvHSh.Cells("FillForegnd").Formula
    dvHSh.Cells("FillPattern").Formula = "1"
    dvHSh.Cells("LineColor").Formula = C_BDR
    dvHSh.Cells("LineWeight").Formula = "0.5pt"
    dvHSh.text = eqTxt
    dvHSh.Cells("Char.Size").Formula = fpt
    dvHSh.Cells("Char.Style").Formula = "1"
    dvHSh.Cells("VerticalAlign").Formula = "1"
    dvHSh.Cells("Para.HorzAlign").Formula = "1"
End Sub

'===========================================================================
' DRAW - PORT LABEL
' devTopY = Y of top of the device box; localSlot = row index within the box.
' rightAligned: True = text hugs right edge (outgoing / End-A)
'               False = text hugs left edge (incoming / End-B)
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
' Joins (x1,y1) to (x2,y2). Same height => single horizontal line; otherwise
' routed orthogonally (L-shaped) via a mid-span vertical.
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
