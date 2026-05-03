Option Explicit

' =========================================================
' POL1 – SQL BUS SCRIPT – PRODUCTION FINAL
' =========================================================

Dim gTagError
Dim zatopVal

' =========================================================
' 1. HELPER FUNCTIONS
' =========================================================

Function TagGet(ByVal devName, ByVal tagPath)
    Dim v
    v = Null
    On Error Resume Next
    Err.Clear
    v = RTag.GetTagValue(devName, tagPath)
    If Err.Number <> 0 Then
        gTagError = True
        v = Null
    End If
    On Error GoTo 0
    TagGet = v
End Function

Sub TagSet(ByVal devName, ByVal tagPath, ByVal val)
    On Error Resume Next
    Err.Clear
    RTag.SetTagValue devName, tagPath, val
    If Err.Number <> 0 Then
        gTagError = True
    End If
    On Error GoTo 0
End Sub

Function ToIntSafe(ByVal v, ByVal fallbackVal)
    Dim x
    If IsNull(v) Or IsEmpty(v) Then
        ToIntSafe = fallbackVal
        Exit Function
    End If
    On Error Resume Next
    x = CInt(v)
    If Err.Number <> 0 Then
        ToIntSafe = fallbackVal
    Else
        ToIntSafe = x
    End If
    On Error GoTo 0
End Function

Function BoolToInt(ByVal v)
    BoolToInt = ToBitSafe(v, 0)
End Function

Function ToBitSafe(ByVal v, ByVal fallbackVal)
    Dim txt, x
    If IsNull(v) Or IsEmpty(v) Then
        ToBitSafe = fallbackVal
        Exit Function
    End If

    txt = UCase(Trim(CStr(v)))
    If (txt = "TRUE") Or (txt = "1") Or (txt = "-1") Then
        ToBitSafe = 1
        Exit Function
    End If
    If (txt = "FALSE") Or (txt = "0") Then
        ToBitSafe = 0
        Exit Function
    End If

    On Error Resume Next
    x = CInt(v)
    If Err.Number <> 0 Then
        ToBitSafe = fallbackVal
    Else
        If x = 0 Then
            ToBitSafe = 0
        Else
            ToBitSafe = 1
        End If
    End If
    On Error GoTo 0
End Function

Sub SetTopAndSystemZatop(ByVal val)
    Const OPC = "OPC1"
    Const DEV = "System"
    Const TAG_TOP = "TOP"
    Const TAG_SYS_ZATOP = "zatop"
    Dim bitVal
    bitVal = ToBitSafe(val, 0)
    TagSet OPC, TAG_TOP, bitVal
    TagSet DEV, TAG_SYS_ZATOP, bitVal
End Sub

Sub ReconcileTopFromSystemZatop(ByRef zatopVal)
    Const OPC = "OPC1"
    Const DEV = "System"
    Const TAG_TOP = "TOP"
    Const TAG_SYS_ZATOP = "zatop"
    Dim desiredVal, actualTop

    ' System.zatop je primarni desired state; při chybě čtení držíme poslední validní hodnotu.
    desiredVal = ToBitSafe(TagGet(DEV, TAG_SYS_ZATOP), ToBitSafe(zatopVal, 0))
    actualTop = ToBitSafe(TagGet(OPC, TAG_TOP), -999)

    If actualTop <> desiredVal Then
        TagSet OPC, TAG_TOP, desiredVal
    End If

    zatopVal = desiredVal
End Sub

Function NzStr(ByVal v, ByVal fallbackVal)
    If IsNull(v) Then
        NzStr = fallbackVal
    Else
        NzStr = CStr(v)
    End If
End Function

Sub DiagDb(ByVal okBit, ByVal errText)
    Const DEV = "System"
    Const TAG_DB_OK = "pol1_db_ok"
    Const TAG_DB_ERR = "pol1_db_last_err"
    TagSet DEV, TAG_DB_OK, CInt(okBit)
    TagSet DEV, TAG_DB_ERR, CStr(errText)
End Sub

Sub WriteHeartbeat()
    Const DEV = "System"
    Const TAG_HEARTBEAT = "POL1_script_last_run"
    Const TAG_HB_COUNTER = "pol1_hb_counter"
    Dim c
    gTagError = False
    TagSet DEV, TAG_HEARTBEAT, CStr(Now)
    c = ToIntSafe(TagGet(DEV, TAG_HB_COUNTER), 0)
    c = c + 1
    TagSet DEV, TAG_HB_COUNTER, c
End Sub

Function IsIntervalValid(hZap, mZap, hVyp, mVyp, denNazev, amVal, ByRef errOut)
    IsIntervalValid = True

    ' Pokud je den vypnutý (-1), validace končí (je to v pořádku)
    If hZap = -1 Or hVyp = -1 Then Exit Function

    Dim totalZap, totalVyp, delka, typText, minDelka
    totalZap = (hZap * 60) + mZap
    totalVyp = (hVyp * 60) + mVyp

    ' Zjištění režimu pro limity a texty
    If amVal <> 0 Then
        typText = "Směna"
        minDelka = 60 ' Automat vyžaduje hodinu
    Else
        typText = "Topení"
        minDelka = 3  ' Manuál pustí i krátké 3minutové testy
    End If

' 1. ZAP nesmí být větší nebo rovno VYP
    If totalZap >= totalVyp Then
        If amVal <> 0 Then
            ' Hláška pro AUTOMAT (Směny)
            errOut = "CHYBA V HARMONOGRAMU (Den: " & denNazev & ")" & vbCrLf & _
                     "Konec směny je zadaný dříve než její začátek!" & vbCrLf & vbCrLf & _
                     "JAK TO OPRAVIT:" & vbCrLf & _
                     "1) Opravte překlep (začátek směny musí být dřív než konec směny)." & vbCrLf & _
                     "2) Pokud jde o noční směnu, rozdělte ji do dvou a více dní." & vbCrLf & _
                     "   (Příklad: ÚT 22:00 - 23:59  |  ST 00:00 - 06:00)"
        Else
            ' Hláška pro MANUÁL (Topení)
            errOut = "CHYBA V HARMONOGRAMU (Den: " & denNazev & ")" & vbCrLf & _
                     "Čas vypnutí je zadaný dříve než čas zapnutí!" & vbCrLf & vbCrLf & _
                     "JAK TO OPRAVIT:" & vbCrLf & _
                     "1) Opravte překlep (zapnutí topení musí být dřív než konec topení)." & vbCrLf & _
                     "2) Pokud topíte přes půlnoc, rozdělte to do dvou a více dní." & vbCrLf & _
                     "   (Příklad: ÚT 22:00 - 23:59  |  ST 00:00 - 06:00)"
        End If
        IsIntervalValid = False : Exit Function
    End If

    ' 2. Dynamická minimální délka
    delka = totalVyp - totalZap
    If delka < minDelka Then
        If amVal <> 0 Then
            errOut = "Chyba " & denNazev & ": Směna musí trvat minimálně 60 minut!"
        Else
            errOut = "Chyba " & denNazev & ": Topení musí být zapnuté minimálně " & minDelka & " minuty!"
        End If
        IsIntervalValid = False : Exit Function
    End If
End Function

Function OpenDbConnection()
    ' UPRAVENO: Tvoje pripojeni
    'Const CONN_STR = "Provider=SQLOLEDB.1;Data Source=XXXXXXX;Initial Catalog=XXXXX;User ID=XXXXX;Password=XXXXX;TrustServerCertificate=Yes;"
    Const CONN_STR = "Provider=SQLOLEDB.1;Data Source=XXXXXXX;Initial Catalog=XXXXXX;User ID=XXXX;Password=XXXX;TrustServerCertificate=Yes;"
    Dim cn
    Set cn = CreateObject("ADODB.Connection")
    cn.CursorLocation = 3
    On Error Resume Next
    Err.Clear
    cn.Open CONN_STR
    If Err.Number <> 0 Then
        DiagDb 0, "CONNECT FAIL: " & Err.Description
        Set cn = Nothing
    End If
    On Error GoTo 0
    Set OpenDbConnection = cn
End Function

Function CheckPythonWatchdog(ByVal cn)
    Dim sql, rs
    CheckPythonWatchdog = False
    sql = "SELECT 1 FROM dbo.POL1_State WHERE Id=1 AND PythonHeartbeat > DATEADD(minute, -3, SYSDATETIME())"

    On Error Resume Next
    Err.Clear
    Set rs = cn.Execute(sql)
    If Err.Number <> 0 Then
        DiagDb 0, "WATCHDOG SQL ERR: " & Err.Description
        Exit Function
    End If
    On Error GoTo 0

    If rs Is Nothing Then Exit Function
    If Not rs.EOF Then
        CheckPythonWatchdog = True
    Else
        CheckPythonWatchdog = False
    End If
    rs.Close
    Set rs = Nothing
End Function


Function ClaimNextCommand(ByVal cn, ByRef cmdId, ByRef cmdType, ByRef valueBit)
    Const QUEUE_TABLE = "dbo.POL1_CommandQueue"
    Dim rs, sqlSel, sqlUpd
    ClaimNextCommand = False
    cmdId = 0 : cmdType = "" : valueBit = 0
    On Error Resume Next
    Err.Clear
    cn.BeginTrans
    If Err.Number <> 0 Then Exit Function
    sqlSel = "SET NOCOUNT ON; SELECT TOP 1 CmdId, CmdType, ValueBit FROM " & QUEUE_TABLE & " WITH (UPDLOCK, READPAST, ROWLOCK) WHERE Status='NEW' ORDER BY TsCreated"
    Set rs = cn.Execute(sqlSel)
    If Err.Number <> 0 Or rs Is Nothing Then
        cn.RollbackTrans
        Exit Function
    End If
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Err.Clear
        cn.CommitTrans
        If Err.Number <> 0 Then cn.RollbackTrans
        Exit Function
    End If
    cmdId = CLng(rs.Fields("CmdId").Value)
    cmdType = CStr(rs.Fields("CmdType").Value)
    valueBit = ToBitSafe(rs.Fields("ValueBit").Value, 0)
    rs.Close: Set rs = Nothing
    Err.Clear
    sqlUpd = "UPDATE " & QUEUE_TABLE & " SET Status='INPROGRESS', TsStarted=SYSDATETIME() WHERE CmdId=" & CStr(cmdId)
    cn.Execute sqlUpd
    If Err.Number <> 0 Then
        cn.RollbackTrans
        Exit Function
    End If
    Err.Clear
    cn.CommitTrans
    If Err.Number <> 0 Then
        cn.RollbackTrans
        Exit Function
    End If
    ClaimNextCommand = True
End Function

Sub AckCommand(ByVal cn, ByVal cmdId, ByVal status, ByVal errText)
    Const QUEUE_TABLE = "dbo.POL1_CommandQueue"
    Dim sql, cmd
    sql = "UPDATE " & QUEUE_TABLE & " SET Status=?, TsFinished=SYSDATETIME(), ErrText=? WHERE CmdId=?"
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter(, 202, 1, 16, CStr(status))
    cmd.Parameters.Append cmd.CreateParameter(, 202, 1, 255, NzStr(errText, ""))
    cmd.Parameters.Append cmd.CreateParameter(, 20, 1, , CLng(cmdId))
    On Error Resume Next
    cmd.Execute
    On Error GoTo 0
End Sub

Sub ProcessOneCommand(ByVal cn)
    Const OPC = "OPC1"
    Const DEV = "System"
    Const TAG_ZATOP = "TOP"
    Const TAG_SYS_ZATOP = "zatop"
    Const MAX_READBACK_RETRIES = 6
    Const READBACK_DELAY_MS = 400
    Const REWRITE_AFTER_ATTEMPT = 3

    Dim cmdId, cmdType, val, okClaim
    Dim writeVal, topRead, sysRead, errTxt
    Dim rawTop, rawSys, attempt, okReadback

    okClaim = ClaimNextCommand(cn, cmdId, cmdType, val)
    If Not okClaim Then Exit Sub

    If UCase(CStr(cmdType)) <> "SET_ZATOP" Then
        AckCommand cn, cmdId, "ERROR", "Unknown CmdType: " & CStr(cmdType)
        Exit Sub
    End If

    writeVal = ToBitSafe(val, 0)

    ' 1) Zapis do fyzickeho tagu
    SetTopAndSystemZatop writeVal

    ' 2) Readback check s retry:
    ' - primarni je fyzicky OPC.TOP
    ' - sekundarni je System.zatop (mirror)
    ' Ack DONE davam az po potvrzeni z jednoho zdroje.
    okReadback = False
    topRead = -999
    sysRead = -999
    For attempt = 1 To MAX_READBACK_RETRIES
        rawTop = TagGet(OPC, TAG_ZATOP)
        topRead = ToBitSafe(rawTop, -999)
        rawSys = TagGet(DEV, TAG_SYS_ZATOP)
        sysRead = ToBitSafe(rawSys, -999)

        If (topRead = writeVal) Or (sysRead = writeVal) Then
            okReadback = True
            Exit For
        End If

        ' Jedno opakovane prepsani pro pripad kratkeho prepisu od jine logiky.
        If attempt = REWRITE_AFTER_ATTEMPT Then
            SetTopAndSystemZatop writeVal
        End If

        If attempt < MAX_READBACK_RETRIES Then
            WScript.Sleep READBACK_DELAY_MS
        End If
    Next

    If okReadback Then
        ' 3) Interni mirror + ACK DONE
        zatopVal = writeVal
        AckCommand cn, cmdId, "DONE", ""
        DiagDb 1, "CMD OK: " & CStr(cmdId) & " SET_ZATOP=" & CStr(writeVal)
    Else
        errTxt = "Write/read mismatch; cmd=" & CStr(writeVal) & ", read_top=" & CStr(topRead) & ", read_sys=" & CStr(sysRead)
        AckCommand cn, cmdId, "ERROR", errTxt
        DiagDb 0, "CMD MISMATCH: " & CStr(cmdId) & " " & errTxt
    End If
End Sub

' =========================================================
' 2. DB STATE FUNCTIONS
' =========================================================

Function DbReadState(ByVal cn, _
    ByRef poZap, ByRef poZapMin, ByRef poVyp, ByRef poVypMin, _
    ByRef utZap, ByRef utZapMin, ByRef utVyp, ByRef utVypMin, _
    ByRef stZap, ByRef stZapMin, ByRef stVyp, ByRef stVypMin, _
    ByRef ctZap, ByRef ctZapMin, ByRef ctVyp, ByRef ctVypMin, _
    ByRef paZap, ByRef paZapMin, ByRef paVyp, ByRef paVypMin, _
    ByRef soZap, ByRef soZapMin, ByRef soVyp, ByRef soVypMin, _
    ByRef neZap, ByRef neZapMin, ByRef neVyp, ByRef neVypMin, _
    ByRef zatopVal, ByRef bufVal, ByRef amVal, ByRef coolingVal)

    Const STATE_TABLE = "dbo.POL1_State"
    Dim rs, sql
    DbReadState = False

    sql = "SELECT po_zap, po_zap_min, po_vyp, po_vyp_min, " & _
          "ut_zap, ut_zap_min, ut_vyp, ut_vyp_min, " & _
          "st_zap, st_zap_min, st_vyp, st_vyp_min, " & _
          "ct_zap, ct_zap_min, ct_vyp, ct_vyp_min, " & _
          "pa_zap, pa_zap_min, pa_vyp, pa_vyp_min, " & _
          "so_zap, so_zap_min, so_vyp, so_vyp_min, " & _
          "ne_zap, ne_zap_min, ne_vyp, ne_vyp_min, " & _
          "zatop, BufferMinutes, auto_man, CoolingEnabled FROM " & STATE_TABLE & " WHERE Id=1"

    On Error Resume Next
    Set rs = cn.Execute(sql)
    If Err.Number <> 0 Then
        DiagDb 0, "READ STATE FAIL"
        Exit Function
    End If
    On Error GoTo 0
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If

    poZap = ToIntSafe(rs.Fields("po_zap").Value, -999): poZapMin = ToIntSafe(rs.Fields("po_zap_min").Value, 0)
    poVyp = ToIntSafe(rs.Fields("po_vyp").Value, -999): poVypMin = ToIntSafe(rs.Fields("po_vyp_min").Value, 0)

    utZap = ToIntSafe(rs.Fields("ut_zap").Value, -999): utZapMin = ToIntSafe(rs.Fields("ut_zap_min").Value, 0)
    utVyp = ToIntSafe(rs.Fields("ut_vyp").Value, -999): utVypMin = ToIntSafe(rs.Fields("ut_vyp_min").Value, 0)

    stZap = ToIntSafe(rs.Fields("st_zap").Value, -999): stZapMin = ToIntSafe(rs.Fields("st_zap_min").Value, 0)
    stVyp = ToIntSafe(rs.Fields("st_vyp").Value, -999): stVypMin = ToIntSafe(rs.Fields("st_vyp_min").Value, 0)

    ctZap = ToIntSafe(rs.Fields("ct_zap").Value, -999): ctZapMin = ToIntSafe(rs.Fields("ct_zap_min").Value, 0)
    ctVyp = ToIntSafe(rs.Fields("ct_vyp").Value, -999): ctVypMin = ToIntSafe(rs.Fields("ct_vyp_min").Value, 0)

    paZap = ToIntSafe(rs.Fields("pa_zap").Value, -999): paZapMin = ToIntSafe(rs.Fields("pa_zap_min").Value, 0)
    paVyp = ToIntSafe(rs.Fields("pa_vyp").Value, -999): paVypMin = ToIntSafe(rs.Fields("pa_vyp_min").Value, 0)

    soZap = ToIntSafe(rs.Fields("so_zap").Value, -999): soZapMin = ToIntSafe(rs.Fields("so_zap_min").Value, 0)
    soVyp = ToIntSafe(rs.Fields("so_vyp").Value, -999): soVypMin = ToIntSafe(rs.Fields("so_vyp_min").Value, 0)

    neZap = ToIntSafe(rs.Fields("ne_zap").Value, -999): neZapMin = ToIntSafe(rs.Fields("ne_zap_min").Value, 0)
    neVyp = ToIntSafe(rs.Fields("ne_vyp").Value, -999): neVypMin = ToIntSafe(rs.Fields("ne_vyp_min").Value, 0)

    zatopVal = ToBitSafe(rs.Fields("zatop").Value, -999)
    On Error Resume Next
    bufVal = ToIntSafe(rs.Fields("BufferMinutes").Value, 30)
    amVal = BoolToInt(rs.Fields("auto_man").Value)
    coolingVal = BoolToInt(rs.Fields("CoolingEnabled").Value)
    On Error GoTo 0
    rs.Close: Set rs = Nothing: DbReadState = True
End Function

Function StateDiffers( _
    a_poZap, a_poZapMin, a_poVyp, a_poVypMin, _
    a_utZap, a_utZapMin, a_utVyp, a_utVypMin, _
    a_stZap, a_stZapMin, a_stVyp, a_stVypMin, _
    a_ctZap, a_ctZapMin, a_ctVyp, a_ctVypMin, _
    a_paZap, a_paZapMin, a_paVyp, a_paVypMin, _
    a_soZap, a_soZapMin, a_soVyp, a_soVypMin, _
    a_neZap, a_neZapMin, a_neVyp, a_neVypMin, _
    a_zatopVal, a_buf, a_am, a_cool, _
    b_poZap, b_poZapMin, b_poVyp, b_poVypMin, _
    b_utZap, b_utZapMin, b_utVyp, b_utVypMin, _
    b_stZap, b_stZapMin, b_stVyp, b_stVypMin, _
    b_ctZap, b_ctZapMin, b_ctVyp, b_ctVypMin, _
    b_paZap, b_paZapMin, b_paVyp, b_paVypMin, _
    b_soZap, b_soZapMin, b_soVyp, b_soVypMin, _
    b_neZap, b_neZapMin, b_neVyp, b_neVypMin, _
    b_zatopVal, b_buf, b_am, b_cool)

    StateDiffers = False
    If a_poZap <> b_poZap Then StateDiffers = True
    If a_poZapMin <> b_poZapMin Then StateDiffers = True
    If a_poVyp <> b_poVyp Then StateDiffers = True
    If a_poVypMin <> b_poVypMin Then StateDiffers = True

    If a_utZap <> b_utZap Then StateDiffers = True
    If a_utZapMin <> b_utZapMin Then StateDiffers = True
    If a_utVyp <> b_utVyp Then StateDiffers = True
    If a_utVypMin <> b_utVypMin Then StateDiffers = True

    If a_stZap <> b_stZap Then StateDiffers = True
    If a_stZapMin <> b_stZapMin Then StateDiffers = True
    If a_stVyp <> b_stVyp Then StateDiffers = True
    If a_stVypMin <> b_stVypMin Then StateDiffers = True

    If a_ctZap <> b_ctZap Then StateDiffers = True
    If a_ctZapMin <> b_ctZapMin Then StateDiffers = True
    If a_ctVyp <> b_ctVyp Then StateDiffers = True
    If a_ctVypMin <> b_ctVypMin Then StateDiffers = True

    If a_paZap <> b_paZap Then StateDiffers = True
    If a_paZapMin <> b_paZapMin Then StateDiffers = True
    If a_paVyp <> b_paVyp Then StateDiffers = True
    If a_paVypMin <> b_paVypMin Then StateDiffers = True

    If a_soZap <> b_soZap Then StateDiffers = True
    If a_soZapMin <> b_soZapMin Then StateDiffers = True
    If a_soVyp <> b_soVyp Then StateDiffers = True
    If a_soVypMin <> b_soVypMin Then StateDiffers = True

    If a_neZap <> b_neZap Then StateDiffers = True
    If a_neZapMin <> b_neZapMin Then StateDiffers = True
    If a_neVyp <> b_neVyp Then StateDiffers = True
    If a_neVypMin <> b_neVypMin Then StateDiffers = True

    If a_zatopVal <> b_zatopVal Then StateDiffers = True
    If a_buf <> b_buf Then StateDiffers = True
    If a_am <> b_am Then StateDiffers = True
    If a_cool <> b_cool Then StateDiffers = True
End Function

Function WriteState(ByVal cn, _
    ByVal poZap, ByVal poZapMin, ByVal poVyp, ByVal poVypMin, _
    ByVal utZap, ByVal utZapMin, ByVal utVyp, ByVal utVypMin, _
    ByVal stZap, ByVal stZapMin, ByVal stVyp, ByVal stVypMin, _
    ByVal ctZap, ByVal ctZapMin, ByVal ctVyp, ByVal ctVypMin, _
    ByVal paZap, ByVal paZapMin, ByVal paVyp, ByVal paVypMin, _
    ByVal soZap, ByVal soZapMin, ByVal soVyp, ByVal soVypMin, _
    ByVal neZap, ByVal neZapMin, ByVal neVyp, ByVal neVypMin, _
    ByVal zatopVal, ByVal bufVal, ByVal amVal, ByVal coolingVal, ByVal srcOk, ByVal srcErr)

    Dim sql, cmd
    WriteState = False

    sql = "UPDATE dbo.POL1_State SET Ts=SYSDATETIME(), " & _
          "po_zap=?, po_zap_min=?, po_vyp=?, po_vyp_min=?, " & _
          "ut_zap=?, ut_zap_min=?, ut_vyp=?, ut_vyp_min=?, " & _
          "st_zap=?, st_zap_min=?, st_vyp=?, st_vyp_min=?, " & _
          "ct_zap=?, ct_zap_min=?, ct_vyp=?, ct_vyp_min=?, " & _
          "pa_zap=?, pa_zap_min=?, pa_vyp=?, pa_vyp_min=?, " & _
          "so_zap=?, so_zap_min=?, so_vyp=?, so_vyp_min=?, " & _
          "ne_zap=?, ne_zap_min=?, ne_vyp=?, ne_vyp_min=?, " & _
          "zatop=?, BufferMinutes=?, auto_man=?, CoolingEnabled=?, src_ok=?, src_err=? WHERE Id=1"

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandText = sql

    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , poZap)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , poZapMin)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , poVyp)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , poVypMin)

    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , utZap)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , utZapMin)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , utVyp)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , utVypMin)

    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , stZap)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , stZapMin)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , stVyp)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , stVypMin)

    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , ctZap)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , ctZapMin)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , ctVyp)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , ctVypMin)

    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , paZap)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , paZapMin)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , paVyp)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , paVypMin)

    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , soZap)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , soZapMin)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , soVyp)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , soVypMin)

    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , neZap)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , neZapMin)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , neVyp)
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , neVypMin)

    cmd.Parameters.Append cmd.CreateParameter(, 11, 1, , CInt(zatopVal))
    cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , CInt(bufVal))
    cmd.Parameters.Append cmd.CreateParameter(, 11, 1, , CInt(amVal))
    cmd.Parameters.Append cmd.CreateParameter(, 11, 1, , CInt(coolingVal))
    cmd.Parameters.Append cmd.CreateParameter(, 11, 1, , CInt(srcOk))
    cmd.Parameters.Append cmd.CreateParameter(, 202, 1, 255, NzStr(srcErr, ""))

    On Error Resume Next
    cmd.Execute
    If Err.Number <> 0 Then
        DiagDb 0, "WRITE STATE FAIL"
        Exit Function
    End If
    On Error GoTo 0
    WriteState = True
End Function

' =========================================================
' 3. MAIN LOGIC
' =========================================================
Sub Main()
    Const DEV = "System"
    Const OPC = "OPC1"    ' TADY PIDÁNA CONST DEV ABY VIDĚL
    ' Z DIAGNOSTIKY VIME: Tagy se jmenuji presne takto (bez slozky Harmonogram.)
    ' Protoze RTag.GetTagValue("System", "Po zap min") vratilo 0, nikoliv Error.

    ' Tagy pro Hodiny
    Const TAG_PO_ZAP = "Po zap1": Const TAG_PO_VYP = "Po vyp1"
    Const TAG_UT_ZAP = "Ut zap1": Const TAG_UT_VYP = "Ut vyp1"
    Const TAG_ST_ZAP = "St zap1": Const TAG_ST_VYP = "St vyp1"
    Const TAG_CT_ZAP = "Ct zap1": Const TAG_CT_VYP = "Ct vyp1"
    Const TAG_PA_ZAP = "Pa zap1": Const TAG_PA_VYP = "Pa vyp1"
    Const TAG_SO_ZAP = "So zap1": Const TAG_SO_VYP = "So vyp1"
    Const TAG_NE_ZAP = "Ne zap1": Const TAG_NE_VYP = "Ne vyp1"


    ' Tagy pro Minuty
    Const TAG_PO_ZAP_MIN = "Po zap min": Const TAG_PO_VYP_MIN = "Po vyp min"
    Const TAG_UT_ZAP_MIN = "Ut zap min": Const TAG_UT_VYP_MIN = "Ut vyp min"
    Const TAG_ST_ZAP_MIN = "St zap min": Const TAG_ST_VYP_MIN = "St vyp min"
    Const TAG_CT_ZAP_MIN = "Ct zap min": Const TAG_CT_VYP_MIN = "Ct vyp min"
    Const TAG_PA_ZAP_MIN = "Pa zap min": Const TAG_PA_VYP_MIN = "Pa vyp min"
    Const TAG_SO_ZAP_MIN = "So zap min": Const TAG_SO_VYP_MIN = "So vyp min"
    Const TAG_NE_ZAP_MIN = "Ne zap min": Const TAG_NE_VYP_MIN = "Ne vyp min"

    Const TAG_COOLING_ENA = "CoolingEnabled"
    Const TAG_ZATOP = "TOP"
    'Const TAG_ZATOP = "zatop"
    Const TAG_BUFFER = "buffer"
    Const TAG_AUTO_MAN = "auto_man"
    Const TAG_PYTHON_DEAD = "ALARM_Python_is_Dead"

    ' Proměnné
    Dim poZap, poZapMin, poVyp, poVypMin
    Dim utZap, utZapMin, utVyp, utVypMin
    Dim stZap, stZapMin, stVyp, stVypMin
    Dim ctZap, ctZapMin, ctVyp, ctVypMin
    Dim paZap, paZapMin, paVyp, paVypMin
    Dim soZap, soZapMin, soVyp, soVypMin
    Dim neZap, neZapMin, neVyp, neVypMin

    Dim bufVal, amVal
    Dim coolingVal, db_coolingVal

    ' DB Variables
    Dim db_poZap, db_poZapMin, db_poVyp, db_poVypMin
    Dim db_utZap, db_utZapMin, db_utVyp, db_utVypMin
    Dim db_stZap, db_stZapMin, db_stVyp, db_stVypMin
    Dim db_ctZap, db_ctZapMin, db_ctVyp, db_ctVypMin
    Dim db_paZap, db_paZapMin, db_paVyp, db_paVypMin
    Dim db_soZap, db_soZapMin, db_soVyp, db_soVypMin
    Dim db_neZap, db_neZapMin, db_neVyp, db_neVypMin
    Dim db_zatopVal, db_bufVal, db_amVal

    Dim cn, okRead, differs, okWrite, srcOk, srcErr, isPythonAlive

    gTagError = False
    WriteHeartbeat

   ' 1. DB Connect
    Set cn = OpenDbConnection()
    If cn Is Nothing Then Exit Sub

    ' 2. Read Tags (Aktuální stav v Relianci)
    poZap = ToIntSafe(TagGet(DEV, TAG_PO_ZAP), -1): poZapMin = ToIntSafe(TagGet(DEV, TAG_PO_ZAP_MIN), 0)
    poVyp = ToIntSafe(TagGet(DEV, TAG_PO_VYP), -1): poVypMin = ToIntSafe(TagGet(DEV, TAG_PO_VYP_MIN), 0)

    utZap = ToIntSafe(TagGet(DEV, TAG_UT_ZAP), -1): utZapMin = ToIntSafe(TagGet(DEV, TAG_UT_ZAP_MIN), 0)
    utVyp = ToIntSafe(TagGet(DEV, TAG_UT_VYP), -1): utVypMin = ToIntSafe(TagGet(DEV, TAG_UT_VYP_MIN), 0)

    stZap = ToIntSafe(TagGet(DEV, TAG_ST_ZAP), -1): stZapMin = ToIntSafe(TagGet(DEV, TAG_ST_ZAP_MIN), 0)
    stVyp = ToIntSafe(TagGet(DEV, TAG_ST_VYP), -1): stVypMin = ToIntSafe(TagGet(DEV, TAG_ST_VYP_MIN), 0)

    ctZap = ToIntSafe(TagGet(DEV, TAG_CT_ZAP), -1): ctZapMin = ToIntSafe(TagGet(DEV, TAG_CT_ZAP_MIN), 0)
    ctVyp = ToIntSafe(TagGet(DEV, TAG_CT_VYP), -1): ctVypMin = ToIntSafe(TagGet(DEV, TAG_CT_VYP_MIN), 0)

    paZap = ToIntSafe(TagGet(DEV, TAG_PA_ZAP), -1): paZapMin = ToIntSafe(TagGet(DEV, TAG_PA_ZAP_MIN), 0)
    paVyp = ToIntSafe(TagGet(DEV, TAG_PA_VYP), -1): paVypMin = ToIntSafe(TagGet(DEV, TAG_PA_VYP_MIN), 0)

    soZap = ToIntSafe(TagGet(DEV, TAG_SO_ZAP), -1): soZapMin = ToIntSafe(TagGet(DEV, TAG_SO_ZAP_MIN), 0)
    soVyp = ToIntSafe(TagGet(DEV, TAG_SO_VYP), -1): soVypMin = ToIntSafe(TagGet(DEV, TAG_SO_VYP_MIN), 0)

    neZap = ToIntSafe(TagGet(DEV, TAG_NE_ZAP), -1): neZapMin = ToIntSafe(TagGet(DEV, TAG_NE_ZAP_MIN), 0)
    neVyp = ToIntSafe(TagGet(DEV, TAG_NE_VYP), -1): neVypMin = ToIntSafe(TagGet(DEV, TAG_NE_VYP_MIN), 0)

    zatopVal = ToBitSafe(TagGet(DEV, "zatop"), 0)
    bufVal = ToIntSafe(TagGet(DEV, TAG_BUFFER), 30)
    amVal = BoolToInt(TagGet(DEV, TAG_AUTO_MAN))
    coolingVal = BoolToInt(TagGet(DEV, TAG_COOLING_ENA))

    ' 3. Read DB State (NAČÍTÁME HNED ZDE, ABYCHOM POZNALI ZMĚNU!)
    okRead = DbReadState(cn, _
        db_poZap, db_poZapMin, db_poVyp, db_poVypMin, _
        db_utZap, db_utZapMin, db_utVyp, db_utVypMin, _
        db_stZap, db_stZapMin, db_stVyp, db_stVypMin, _
        db_ctZap, db_ctZapMin, db_ctVyp, db_ctVypMin, _
        db_paZap, db_paZapMin, db_paVyp, db_paVypMin, _
        db_soZap, db_soZapMin, db_soVyp, db_soVypMin, _
        db_neZap, db_neZapMin, db_neVyp, db_neVypMin, _
        db_zatopVal, db_bufVal, db_amVal, db_coolingVal)

   ' =========================================================
' 4. DETEKCE PŘEPNUTÍ DO MANUÁLU (RESET HARMONOGRAMU)
' =========================================================
' Pokud je v DB Auto (1) a v Tagu Manual (0), došlo právě k přepnutí.
' Nebo pokud je DB nedostupná, ale jsme v Manualu, preventivně resetujeme.
' =========================================================
' 4. DETEKCE PŘEPNUTÍ DO MANUÁLU (RESET HARMONOGRAMU + QUEUE)
' =========================================================
' Pokud je v DB Auto (1) a v Tagu Manual (0), došlo právě k přepnutí.

If okRead And (db_amVal <> 0) And (amVal = 0) Then
    DiagDb 1, "MANUAL SWITCH DETECTED -> RESETTING ALL"

    ' --- Zrušení čekajících příkazů v Python frontě (aby nečekaly na návrat do AUTO) ---
    On Error Resume Next
    cn.Execute "UPDATE dbo.POL1_CommandQueue SET Status='CANCELLED', TsFinished=SYSDATETIME(), ErrText='Manual mode override' WHERE Status='NEW'"
    If Err.Number <> 0 Then DiagDb 0, "QUEUE CANCEL FAIL: " & Err.Description
    On Error GoTo 0

    ' --- Fyzický reset a výchozí hodnoty ---
    coolingVal = 0
    TagSet DEV, TAG_COOLING_ENA, 0

    ' Pondělí - Start: 00:00, Konec: z DB
    poZap=0: poZapMin=0
    poVyp=db_poVyp: poVypMin=db_poVypMin
    TagSet DEV, TAG_PO_ZAP, 0: TagSet DEV, TAG_PO_ZAP_MIN, 0
    TagSet DEV, TAG_PO_VYP, db_poVyp: TagSet DEV, TAG_PO_VYP_MIN, db_poVypMin

    ' Úterý - Start: 00:00, Konec: z DB
    utZap=0: utZapMin=0
    utVyp=db_utVyp: utVypMin=db_utVypMin
    TagSet DEV, TAG_UT_ZAP, 0: TagSet DEV, TAG_UT_ZAP_MIN, 0
    TagSet DEV, TAG_UT_VYP, db_utVyp: TagSet DEV, TAG_UT_VYP_MIN, db_utVypMin

    ' Středa - Start: 00:00, Konec: z DB
    stZap=0: stZapMin=0
    stVyp=db_stVyp: stVypMin=db_stVypMin
    TagSet DEV, TAG_ST_ZAP, 0: TagSet DEV, TAG_ST_ZAP_MIN, 0
    TagSet DEV, TAG_ST_VYP, db_stVyp: TagSet DEV, TAG_ST_VYP_MIN, db_stVypMin

    ' Čtvrtek - Start: 00:00, Konec: z DB
    ctZap=0: ctZapMin=0
    ctVyp=db_ctVyp: ctVypMin=db_ctVypMin
    TagSet DEV, TAG_CT_ZAP, 0: TagSet DEV, TAG_CT_ZAP_MIN, 0
    TagSet DEV, TAG_CT_VYP, db_ctVyp: TagSet DEV, TAG_CT_VYP_MIN, db_ctVypMin

    ' Pátek - Start: 00:00, Konec: z DB
    paZap=0: paZapMin=0
    paVyp=db_paVyp: paVypMin=db_paVypMin
    TagSet DEV, TAG_PA_ZAP, 0: TagSet DEV, TAG_PA_ZAP_MIN, 0
    TagSet DEV, TAG_PA_VYP, db_paVyp: TagSet DEV, TAG_PA_VYP_MIN, db_paVypMin

    ' Sobota - Start: -1 (Vypnuto), Konec: z DB
    soZap=-1: soZapMin=0
    soVyp=db_soVyp: soVypMin=db_soVypMin
    TagSet DEV, TAG_SO_ZAP, -1: TagSet DEV, TAG_SO_ZAP_MIN, 0
    TagSet DEV, TAG_SO_VYP, db_soVyp: TagSet DEV, TAG_SO_VYP_MIN, db_soVypMin

    ' Neděle - Start: 17:00, Konec: 23:59
    neZap=17: neZapMin=0
    neVyp=23: neVypMin=59
    TagSet DEV, TAG_NE_ZAP, 17: TagSet DEV, TAG_NE_ZAP_MIN, 0
    TagSet DEV, TAG_NE_VYP, 23: TagSet DEV, TAG_NE_VYP_MIN, 59

    ' Okamžité vypnutí topení (aby se nečekalo na sekci 6)
    zatopVal = 0
    SetTopAndSystemZatop 0
    'TagSet DEV, TAG_ZATOP, 0

End If
' =========================================================
    ' 4b. CHYTRÝ RESET PŘI PŘEPNUTÍ NA AUTO (MODIFIED)
    ' =========================================================
    ' Logika:
    ' 1. Kontrolujeme jen START. Pokud je 0 (systémová hodnota z manuálu), resetujeme ho na 6.
    ' 2. KONEC vždy bereme z DB (db_...vyp). Tím zachováme případné změny konce, které operátor udělal.

    If okRead And (db_amVal = 0) And (amVal <> 0) Then
        DiagDb 1, "AUTO SWITCH DETECTED -> SMART START RESET"

        ' --- PONDĚLÍ ---
        ' Pokud je Start 0 (nedotčený), nastavíme 6:00. Konec potvrdíme z DB.
        If (poZap = 0) Then
            poZap=6: poZapMin=0
            TagSet DEV, TAG_PO_ZAP, 6: TagSet DEV, TAG_PO_ZAP_MIN, 0

            ' Konec explicitně načteme z DB (aby se nepřepsal na default, pokud byl změněn)
            poVyp = db_poVyp: poVypMin = db_poVypMin
            TagSet DEV, TAG_PO_VYP, db_poVyp: TagSet DEV, TAG_PO_VYP_MIN, db_poVypMin
        End If

        ' --- ÚTERÝ ---
        If (utZap = 0) Then
            utZap=6: utZapMin=0
            TagSet DEV, TAG_UT_ZAP, 6: TagSet DEV, TAG_UT_ZAP_MIN, 0

            utVyp = db_utVyp: utVypMin = db_utVypMin
            TagSet DEV, TAG_UT_VYP, db_utVyp: TagSet DEV, TAG_UT_VYP_MIN, db_utVypMin
        End If

        ' --- STŘEDA ---
        If (stZap = 0) Then
            stZap=6: stZapMin=0
            TagSet DEV, TAG_ST_ZAP, 6: TagSet DEV, TAG_ST_ZAP_MIN, 0

            stVyp = db_stVyp: stVypMin = db_stVypMin
            TagSet DEV, TAG_ST_VYP, db_stVyp: TagSet DEV, TAG_ST_VYP_MIN, db_stVypMin
        End If

        ' --- ČTVRTEK ---
        If (ctZap = 0) Then
            ctZap=6: ctZapMin=0
            TagSet DEV, TAG_CT_ZAP, 6: TagSet DEV, TAG_CT_ZAP_MIN, 0

            ctVyp = db_ctVyp: ctVypMin = db_ctVypMin
            TagSet DEV, TAG_CT_VYP, db_ctVyp: TagSet DEV, TAG_CT_VYP_MIN, db_ctVypMin
        End If

        ' --- PÁTEK ---
        If (paZap = 0) Then
            paZap=6: paZapMin=0
            TagSet DEV, TAG_PA_ZAP, 6: TagSet DEV, TAG_PA_ZAP_MIN, 0

            paVyp = db_paVyp: paVypMin = db_paVypMin
            TagSet DEV, TAG_PA_VYP, db_paVyp: TagSet DEV, TAG_PA_VYP_MIN, db_paVypMin
        End If

        ' --- SOBOTA (Pokud je default -1, nastavíme komplet -1/-1) ---
        If (soZap = -1) Then
            soZap=-1: soZapMin=0: soVyp=-1: soVypMin=0
            TagSet DEV, TAG_SO_ZAP, -1: TagSet DEV, TAG_SO_ZAP_MIN, 0
            TagSet DEV, TAG_SO_VYP, -1: TagSet DEV, TAG_SO_VYP_MIN, 0
        End If

        ' --- NEDĚLE (Pokud je default 17, nastavíme komplet -1/-1) ---
        If (neZap = 17) Then
            neZap=-1: neZapMin=0: neVyp=-1: neVypMin=0
            TagSet DEV, TAG_NE_ZAP, -1: TagSet DEV, TAG_NE_ZAP_MIN, 0
            TagSet DEV, TAG_NE_VYP, -1: TagSet DEV, TAG_NE_VYP_MIN, 0
        End If
    End If


' =========================================================
' 5. WATCHDOG CONTROL (Fail-Safe)
' =========================================================
isPythonAlive = CheckPythonWatchdog(cn)

If isPythonAlive Then
    ' === A. PYTHON ALIVE ===

    ' Příkazy zpracujeme pouze v automatickém režimu
    If amVal <> 0 Then
        ProcessOneCommand cn
        DiagDb 1, "PYTHON ALIVE - AUTO MODE"
    Else
        ' V manuálu frontu ignorujeme (příkazy byly zrušeny v sekci 4)
        DiagDb 1, "PYTHON ALIVE - MANUAL MODE (QUEUE IGNORED)"
    End If

    TagSet DEV, TAG_PYTHON_DEAD, 0
Else
    ' === B. PYTHON DEAD (WATCHDOG FAIL) ===
    TagSet DEV, TAG_PYTHON_DEAD, 1
    coolingVal = 0
    TagSet DEV, TAG_COOLING_ENA, 0

    ' Pokud jsme byli v AUTO, vynutíme MANUÁL pro bezpečnost
    If amVal <> 0 Then
        DiagDb 0, "WATCHDOG: PYTHON DEAD -> FORCE MANUAL"
        amVal = 0
        TagSet DEV, TAG_AUTO_MAN, 0
        ' Topení nevypínáme natvrdo zde, o stavu rozhodne hned v zápětí
        ' Sekce 6 na základě fixního harmonogramu.
    Else
        DiagDb 0, "WATCHDOG: PYTHON DEAD (ALREADY MANUAL)"
    End If
End If

' =========================================================
    ' 6. + 7. MANUÁLNÍ REŽIM (Řídí Reliance pouze v amVal = 0)
    ' =========================================================
    Dim den, startH, startM, stopH, stopM, aktH, aktMin
    Dim aktTotalMins, startTotalMins, stopTotalMins, maTopit
    Dim dalsiDen, pristiStartH, pristiStartM, pristiStartTotal
    Dim mezera, dnesVypH, dnesVypM

    ' HLAVNÍ OBAL - Pouze pokud je manuál
    If amVal = 0 Then
        den = Weekday(Now, 2)
        aktH = Hour(Now)
        aktMin = Minute(Now)

        startH = -1 : startM = 0 : stopH = -1 : stopM = 0

        Select Case den
            Case 1: startH = poZap : startM = poZapMin : stopH = poVyp : stopM = poVypMin
            Case 2: startH = utZap : startM = utZapMin : stopH = utVyp : stopM = utVypMin
            Case 3: startH = stZap : startM = stZapMin : stopH = stVyp : stopM = stVypMin
            Case 4: startH = ctZap : startM = ctZapMin : stopH = ctVyp : stopM = ctVypMin
            Case 5: startH = paZap : startM = paZapMin : stopH = paVyp : stopM = paVypMin
            Case 6: startH = soZap : startM = soZapMin : stopH = soVyp : stopM = soVypMin
            Case 7: startH = neZap : startM = neZapMin : stopH = neVyp : stopM = neVypMin
        End Select

        If startH <> -1 And stopH <> -1 Then
            aktTotalMins = (aktH * 60) + aktMin
            startTotalMins = (startH * 60) + startM
            stopTotalMins = (stopH * 60) + stopM

            maTopit = False

            ' Standardní kontrola směny
            If startTotalMins < stopTotalMins Then
                If (aktTotalMins >= startTotalMins) And (aktTotalMins < stopTotalMins) Then maTopit = True
            Else
                If (aktTotalMins >= startTotalMins) Or (aktTotalMins < stopTotalMins) Then maTopit = True
            End If

            ' PŘEKLENUTÍ PŮLNOCI
            If (Not maTopit) Or (aktTotalMins >= stopTotalMins - 5) Then
                dalsiDen = den + 1
                If dalsiDen > 7 Then dalsiDen = 1

                Select Case dalsiDen
                    Case 1: pristiStartH = poZap : pristiStartM = poZapMin
                    Case 2: pristiStartH = utZap : pristiStartM = utZapMin
                    Case 3: pristiStartH = stZap : pristiStartM = stZapMin
                    Case 4: pristiStartH = ctZap : pristiStartM = ctZapMin
                    Case 5: pristiStartH = paZap : pristiStartM = paZapMin
                    Case 6: pristiStartH = soZap : pristiStartM = soZapMin
                    Case 7: pristiStartH = neZap : pristiStartM = neZapMin
                End Select

                If pristiStartH <> -1 Then
                    pristiStartTotal = (pristiStartH * 60) + pristiStartM
                    mezera = (1440 - stopTotalMins) + pristiStartTotal

                    If mezera <= 5 Then
                        If (aktTotalMins >= stopTotalMins - 5) Then
                            maTopit = True
                        End If
                    End If
                End If
            End If

            ' Zápis stavu
            If maTopit Then
                SetTopAndSystemZatop 1
                'TagSet DEV, TAG_ZATOP, 1
                zatopVal = 1
            Else
                SetTopAndSystemZatop 0
                'TagSet DEV, TAG_ZATOP, 0
                zatopVal = 0
            End If

            ' Sekce 7 - Pojistka (vypnutí v přesný čas konce)
            If (aktH = stopH) And (aktMin = stopM) And (Second(Now) < 10) Then
                If Not maTopit Then
                    If zatopVal <> 0 Then
                        SetTopAndSystemZatop 0
                        'TagSet DEV, TAG_ZATOP, 0
                        zatopVal = 0
                    End If
                End If
            End If

        Else
            ' Pokud den nemá nastavenou směnu, vypni topení
            SetTopAndSystemZatop 0
            'TagSet DEV, TAG_ZATOP, 0
            zatopVal = 0
        End If ' Konec podmínky (startH <> -1)

    End If ' KONEC HLAVNÍHO OBALU (amVal = 0)


' =========================================================
    ' 7b. RECONCILE TOP <- System.zatop (kazdy cyklus)
    ' =========================================================
    ReconcileTopFromSystemZatop zatopVal

' =========================================================
    ' 8. SYNCHRONIZACE S DB + VALIDACE S CHYTRÝM OKNEM
    ' =========================================================

    ' 1. Nejdříve zjistíme, zda se data v Relianci liší od dat v DB
    differs = True
    If okRead Then
        differs = StateDiffers( _
            poZap, poZapMin, poVyp, poVypMin, _
            utZap, utZapMin, utVyp, utVypMin, _
            stZap, stZapMin, stVyp, stVypMin, _
            ctZap, ctZapMin, ctVyp, ctVypMin, _
            paZap, paZapMin, paVyp, paVypMin, _
            soZap, soZapMin, soVyp, soVypMin, _
            neZap, neZapMin, neVyp, neVypMin, _
            zatopVal, bufVal, amVal, coolingVal, _
            db_poZap, db_poZapMin, db_poVyp, db_poVypMin, _
            db_utZap, db_utZapMin, db_utVyp, db_utVypMin, _
            db_stZap, db_stZapMin, db_stVyp, db_stVypMin, _
            db_ctZap, db_ctZapMin, db_ctVyp, db_ctVypMin, _
            db_paZap, db_paZapMin, db_paVyp, db_paVypMin, _
            db_soZap, db_soZapMin, db_soVyp, db_soVypMin, _
            db_neZap, db_neZapMin, db_neVyp, db_neVypMin, _
            db_zatopVal, db_bufVal, db_amVal, db_coolingVal)
    End If

    ' 2. VALIDACE INTERVALŮ
    Dim errMsg, valid
    valid = True : errMsg = ""

    ' Předáváme parametr amVal do kontroly
    If Not IsIntervalValid(poZap, poZapMin, poVyp, poVypMin, "PONDĚLÍ", amVal, errMsg) Then valid = False
    If valid And Not IsIntervalValid(utZap, utZapMin, utVyp, utVypMin, "ÚTERÝ", amVal, errMsg) Then valid = False
    If valid And Not IsIntervalValid(stZap, stZapMin, stVyp, stVypMin, "STŘEDA", amVal, errMsg) Then valid = False
    If valid And Not IsIntervalValid(ctZap, ctZapMin, ctVyp, ctVypMin, "ČTVRTEK", amVal, errMsg) Then valid = False
    If valid And Not IsIntervalValid(paZap, paZapMin, paVyp, paVypMin, "PÁTEK", amVal, errMsg) Then valid = False
    If valid And Not IsIntervalValid(soZap, soZapMin, soVyp, soVypMin, "SOBOTA", amVal, errMsg) Then valid = False
    If valid And Not IsIntervalValid(neZap, neZapMin, neVyp, neVypMin, "NEDĚLE", amVal, errMsg) Then valid = False

    ' --- CHYTRÁ REAKCE NA CHYBU ---
    If Not valid Then
        Dim newErrText, lastErrText, nadpisOkna
        newErrText = "CHYBA: " & errMsg

        ' Načteme poslední zapsanou chybu
        lastErrText = NzStr(TagGet(DEV, "pol1_db_last_err"), "")

        ' Nastavíme nadpis okna podle režimu
        If amVal <> 0 Then
            nadpisOkna = "Validace harmonogramu POL1 - AUTOMAT"
        Else
            nadpisOkna = "Validace harmonogramu POL1 - MANUÁL"
        End If

        ' 1. OKAMŽITĚ zapíšeme chybu do tagu (zastaví to spamování oken v dalším cyklu)
        DiagDb 0, newErrText

        ' 2. Okno vyskočí JEN tehdy, když je to pro skript "nová" chyba
        If differs And (lastErrText <> newErrText) Then
            Dim wsh
            Set wsh = CreateObject("WScript.Shell")

            ' Číslo 5 znamená 5 vteřin. 4096 = vždy nahoře.
            wsh.Popup "DATA NEBUDOU ZAPSÁNA DO HARMONOGRAMU!" & vbCrLf & vbCrLf & errMsg, 20, nadpisOkna, 48 + 4096

            Set wsh = Nothing
        End If

        ' 3. Čisté ukončení bez zápisu do DB
        If Not cn Is Nothing Then
            On Error Resume Next
            cn.Close
            Set cn = Nothing
        End If
        Exit Sub
    End If

    ' 3. PŘÍPRAVA DIAGNOSTIKY (pokud validace OK)
    srcOk = 1 : srcErr = ""
    If gTagError Then srcOk = 0 : srcErr = "TAG ERROR"
    If Not isPythonAlive Then srcOk = 0 : srcErr = "PYTHON DEAD"

    ' 4. ZÁPIS DO DB (Pouze pokud jsou data jiná a validní)
    If differs Then
        okWrite = WriteState(cn, _
            poZap, poZapMin, poVyp, poVypMin, _
            utZap, utZapMin, utVyp, utVypMin, _
            stZap, stZapMin, stVyp, stVypMin, _
            ctZap, ctZapMin, ctVyp, ctVypMin, _
            paZap, paZapMin, paVyp, paVypMin, _
            soZap, soZapMin, soVyp, soVypMin, _
            neZap, neZapMin, neVyp, neVypMin, _
            zatopVal, bufVal, amVal, coolingVal, srcOk, srcErr)

        If okWrite Then
            DiagDb 1, "OK - Data zapsána"
        Else
            DiagDb 0, "ZÁPIS SELHAL (DB ERROR)"
        End If
    End If

    ' Čisté ukončení spojení na konci skriptu
    On Error Resume Next
    cn.Close
    Set cn = Nothing
End Sub

Main
