Attribute VB_Name = "modGewindeModellieren"
'==============================================================================
' Kosmetische Gewinde als echtes 3D-Gewinde schneiden  -  Autodesk Inventor 2016
'
' Makro: GewindeModellieren   (per Tastenkürzel aufrufen)
'
'  * Vorauswahl (Browser oder Grafik) von Gewinde-Features, Bohrungs-Features
'    oder Gewindeflächen  -> nur diese werden modelliert
'  * ohne Vorauswahl     -> alle Gewinde / Gewindebohrungen des Bauteils,
'    die noch nicht modelliert wurden
'
' Unterstützte Gewinde:
'  - Metrisches ISO-Gewinde (Regel- und Feingewinde, 60°, DIN 13)
'  - ISO-Rohrgewinde zylindrisch  G / Rp  (ISO 228-1, EN 10226-1, 55° Whitworth)
'  - ISO-Rohrgewinde kegelig      R / Rc  (ISO 7-1, EN 10226-2, Kegel 1:16)
'  Sonstige Gewinde werden mit 60°-Profil angenähert, NPT/NPS übersprungen.
'
' Ergebnis je Gewindefläche: Arbeitsebene + Skizze (Profil) + Arbeitsachse +
' Spirale (Schneiden). Das kosmetische Gewinde bleibt erhalten (Zeichnung!).
' Alles in einer Transaktion -> ein Strg+Z macht alles rückgängig.
' Diagnose: je Gewinde eine Zeile im Direktfenster des VBA-Editors (Strg+G).
'==============================================================================
Option Explicit

' Rechts-/Linksgewinde vertauscht -> True
Private Const INVERT_ROTATION As Boolean = False
' Kegelige Rohrgewinde (R/Rc) werden in die falsche Richtung kegelig -> True
Private Const INVERT_TAPER As Boolean = False
' True = nach jedem Gewinde eine Meldung mit den ermittelten Werten
Private Const SHOW_DEBUG As Boolean = False

Private Const ATTR_SET As String = "ModGewinde"
Private Const K_ISO As Double = 0.541265877365274      ' Gewindetiefe ISO: 5/8 * H (H = 0,866025 P)
Private Const K_WW As Double = 0.640327                 ' Gewindetiefe Whitworth 55°
Private Const TOL As Double = 0.0005                    ' cm

Private Type TSpec
    P As Double          ' Steigung [cm]
    DMaj As Double       ' Außendurchmesser [cm] (kegelig: in der Bezugsebene)
    Depth As Double      ' Profiltiefe [cm]
    HwMaj As Double      ' halbe Lückenbreite am Außendurchmesser [cm]
    HwMin As Double      ' halbe Lückenbreite am Kerndurchmesser [cm]
    Taper As Boolean     ' kegelig 1:16
    GaugeLen As Double   ' Abstand Bezugsebene vom Gewindeende (R außen) [cm]
    Kind As String       ' "ISO", "G", "R", "Rc" ...
    Des As String        ' Bezeichnung
End Type

Private m_TG As TransientGeometry

'------------------------------------------------------------------------------
Public Sub GewindeModellieren()
    Dim oDoc As Document
    Set oDoc = ThisApplication.ActiveEditDocument
    If oDoc Is Nothing Then Exit Sub
    If oDoc.DocumentType <> kPartDocumentObject Then
        MsgBox "Bitte ein Bauteil (.ipt) öffnen bzw. bearbeiten.", vbExclamation, "Gewinde modellieren"
        Exit Sub
    End If

    Dim oPart As PartDocument: Set oPart = oDoc
    Dim oDef As PartComponentDefinition: Set oDef = oPart.ComponentDefinition
    Set m_TG = ThisApplication.TransientGeometry

    ' ---- zu bearbeitende Features sammeln ----
    Dim colFeat As New Collection
    CollectSelection oDef, colFeat
    Dim bSel As Boolean: bSel = (colFeat.Count > 0)

    If Not bSel Then
        Dim oTF As ThreadFeature
        For Each oTF In oDef.Features.ThreadFeatures
            If Not oTF.Suppressed Then AddUnique colFeat, oTF
        Next
        Dim oHF As HoleFeature
        For Each oHF In oDef.Features.HoleFeatures
            If Not oHF.Suppressed Then
                If oHF.Tapped Then AddUnique colFeat, oHF
            End If
        Next
    End If

    If colFeat.Count = 0 Then
        MsgBox "Keine Gewinde gefunden.", vbInformation, "Gewinde modellieren"
        Exit Sub
    End If

    ' ---- modellieren ----
    Dim oTrans As Transaction
    Set oTrans = ThisApplication.TransactionManager.StartTransaction(oPart, "Gewinde modellieren")
    Dim bSilent As Boolean: bSilent = ThisApplication.SilentOperation
    Dim nOk As Long, nSkip As Long, sLog As String
    Dim oFeat As Object

    On Error GoTo Fehler
    ThisApplication.SilentOperation = True
    For Each oFeat In colFeat
        If IsAlreadyModeled(oDef, oFeat.Name) Then
            nSkip = nSkip + 1
        Else
            ProcessFeature oDef, oFeat, nOk, sLog
        End If
    Next
    ThisApplication.SilentOperation = bSilent
    oTrans.End
    oPart.Update

    Dim s As String
    s = nOk & " Gewinde modelliert"
    If nSkip > 0 Then s = s & ", " & nSkip & " bereits vorhanden (übersprungen)"
    ThisApplication.StatusBarText = s
    If Len(sLog) > 0 Or nOk = 0 Then
        If Len(sLog) > 0 Then s = s & vbCrLf & vbCrLf & "Hinweise:" & vbCrLf & sLog
        MsgBox s, vbInformation, "Gewinde modellieren"
    End If
    Exit Sub

Fehler:
    ThisApplication.SilentOperation = bSilent
    oTrans.Abort
    MsgBox "Fehler: " & Err.Description, vbCritical, "Gewinde modellieren"
End Sub

'------------------------------------------------------------------------------
' Vorauswahl auswerten
Private Sub CollectSelection(ByVal oDef As PartComponentDefinition, colFeat As Collection)
    Dim oSS As SelectSet, o As Object, i As Long
    Set oSS = ThisApplication.ActiveDocument.SelectSet
    For i = 1 To oSS.Count
        Set o = oSS.Item(i)
        If TypeOf o Is FaceProxy Then Set o = o.NativeObject
        If TypeOf o Is Face Then Set o = FindFeatureForFace(oDef, o)
        If Not o Is Nothing Then
            If TypeOf o Is ThreadFeature Or TypeOf o Is HoleFeature Then AddUnique colFeat, o
        End If
    Next
End Sub

Private Function FindFeatureForFace(ByVal oDef As PartComponentDefinition, ByVal oFace As Face) As Object
    Dim oTF As ThreadFeature, oX As Object, oF As Face
    On Error Resume Next
    For Each oTF In oDef.Features.ThreadFeatures
        Set oX = Nothing
        Set oX = oTF.ThreadedFace
        If Not oX Is Nothing Then
            If TypeOf oX Is Faces Then
                For Each oF In oX
                    If oF.TransientKey = oFace.TransientKey Then Set FindFeatureForFace = oTF: Exit Function
                Next
            ElseIf TypeOf oX Is Face Then
                If oX.TransientKey = oFace.TransientKey Then Set FindFeatureForFace = oTF: Exit Function
            End If
        End If
    Next
    Set FindFeatureForFace = oFace.CreatedByFeature
End Function

Private Sub AddUnique(col As Collection, o As Object)
    On Error Resume Next
    col.Add o, o.Name
End Sub

' Wurde zu diesem Feature schon eine Spirale erzeugt (und existiert sie noch)?
Private Function IsAlreadyModeled(ByVal oDef As PartComponentDefinition, ByVal sName As String) As Boolean
    Dim oC As CoilFeature
    On Error Resume Next
    For Each oC In oDef.Features.CoilFeatures
        If oC.AttributeSets.NameIsUsed(ATTR_SET) Then
            If oC.AttributeSets.Item(ATTR_SET).Item("Quelle").Value = sName Then
                IsAlreadyModeled = True
                Exit Function
            End If
        End If
    Next
End Function

'------------------------------------------------------------------------------
' Ein Feature (Gewinde oder Gewindebohrung) verarbeiten
Private Sub ProcessFeature(ByVal oDef As PartComponentDefinition, ByVal oFeat As Object, _
                           ByRef nOk As Long, ByRef sLog As String)
    Dim oInfo As Object, oX As Object, oF As Face
    Dim colFaces As New Collection
    Dim bFull As Boolean, bHaveFull As Boolean, bHole As Boolean
    Dim dLen As Double, dOff As Double, rMin As Double

    On Error GoTo EH

    If TypeOf oFeat Is ThreadFeature Then
        Set oInfo = oFeat.ThreadInfo
        Set oX = oFeat.ThreadedFace
        If TypeOf oX Is Faces Then
            For Each oF In oX
                colFaces.Add oF
            Next
        Else
            colFaces.Add oX
        End If
        bHaveFull = TryGetBool(oFeat, "FullLength", bFull)
        TryGetDouble oFeat, "ThreadLength", dLen
        TryGetDouble oFeat, "ThreadOffset", dOff

    ElseIf TypeOf oFeat Is HoleFeature Then
        bHole = True
        If Not oFeat.Tapped Then Exit Sub
        Set oInfo = oFeat.TapInfo
        bHaveFull = TryGetBool(oInfo, "FullTapDepth", bFull)
        TryGetDouble oInfo, "ThreadDepth", dLen
        ' kegelige Gewindebohrung: flache Kegelflächen (< ~6°)
        For Each oF In oFeat.Faces
            If oF.SurfaceType = kConeSurface Then
                If Abs(oF.Geometry.HalfAngle) < 0.1 Then colFaces.Add oF
            End If
        Next
        ' sonst: kleinster Zylinder der Bohrung (Senkungen sind größer)
        If colFaces.Count = 0 Then
            rMin = 1E+30
            For Each oF In oFeat.Faces
                If oF.SurfaceType = kCylinderSurface Then
                    If oF.Geometry.Radius < rMin Then rMin = oF.Geometry.Radius
                End If
            Next
            For Each oF In oFeat.Faces
                If oF.SurfaceType = kCylinderSurface Then
                    If Abs(oF.Geometry.Radius - rMin) < TOL Then colFaces.Add oF
                End If
            Next
        End If
    Else
        Exit Sub
    End If

    If Not bHaveFull Then bFull = (dLen <= TOL)

    Dim oFace As Object, oCoil As CoilFeature, sMsg As String
    For Each oFace In colFaces
        sMsg = ""
        Set oCoil = ModelThread(oDef, oFace, oInfo, bFull, dLen, dOff, bHole, oFeat.Name, sMsg)
        If oCoil Is Nothing Then
            sLog = sLog & oFeat.Name & ": " & sMsg & vbCrLf
        Else
            nOk = nOk + 1
            oCoil.AttributeSets.Add(ATTR_SET).Add "Quelle", kStringType, oFeat.Name
        End If
    Next
    Exit Sub

EH:
    sLog = sLog & oFeat.Name & ": Fehler - " & Err.Description & vbCrLf
End Sub

'------------------------------------------------------------------------------
' Gewinde auf einer Zylinder- oder Kegelfläche modellieren
Private Function ModelThread(ByVal oDef As PartComponentDefinition, ByVal oFace As Face, _
        ByVal oInfo As Object, ByVal bFull As Boolean, ByVal dLenSpec As Double, _
        ByVal dOffset As Double, ByVal bHole As Boolean, ByVal sFeat As String, _
        ByRef sMsg As String) As CoilFeature

    Dim oG As Object
    Select Case oFace.SurfaceType
        Case kCylinderSurface, kConeSurface
            Set oG = oFace.Geometry
        Case Else
            sMsg = "Gewindefläche ist weder Zylinder noch Kegel": Exit Function
    End Select
    Dim O As Point: Set O = oG.BasePoint
    Dim a As UnitVector: Set a = oG.AxisVector

    ' axiale / radiale Ausdehnung der Fläche
    Dim tMin As Double, tMax As Double, rFMin As Double, rFMax As Double
    GetExtents oFace, O, a, tMin, tMax, rFMin, rFMax
    If tMax - tMin < TOL Then sMsg = "Länge der Gewindefläche nicht ermittelbar": Exit Function

    ' Innen/Außen, Rechts/Links
    Dim bInternal As Boolean, bRight As Boolean
    If Not TryGetBool(oInfo, "Internal", bInternal) Then
        If bHole Then bInternal = True Else bInternal = IsInternalFace(oFace, O, a)
    End If
    If Not TryGetBool(oInfo, "RightHanded", bRight) Then bRight = True

    ' Gewindedaten
    Dim sp As TSpec
    If Not GetThreadSpec(oInfo, IIf(bInternal, rFMin, rFMax), bInternal, sp, sMsg) Then Exit Function

    ' Gewindeanfang und Richtung
    Dim sStart As Double, dirSign As Double, dLenBase As Double, bBase As Boolean
    Dim bOpenMin As Boolean, bOpenMax As Boolean, sSrc As String
    GetOpenEnds oFace, O, a, tMin, tMax, bOpenMin, bOpenMax
    bBase = GetThreadBase(oInfo, O, a, tMin, tMax, sStart, dirSign, dLenBase)
    If bBase Then
        ' Plausibilität: API-Richtung muss vom offenen Ende ins Material zeigen
        If (dirSign > 0 And bOpenMax And Not bOpenMin) Or (dirSign < 0 And bOpenMin And Not bOpenMax) Then bBase = False
    End If
    If bBase Then
        sSrc = "API"
    Else
        sSrc = "Geometrie"
        If bOpenMin And Not bOpenMax Then
            sStart = tMin: dirSign = 1
        Else
            sStart = tMax: dirSign = -1
        End If
        sStart = sStart + dirSign * dOffset
    End If

    ' Gewindelänge
    Dim dAvail As Double, L As Double
    If dirSign > 0 Then dAvail = tMax - sStart Else dAvail = sStart - tMin
    If bFull Then
        L = dAvail
    ElseIf dLenSpec > TOL Then
        L = dLenSpec
    ElseIf bBase And dLenBase > TOL Then
        L = dLenBase
    Else
        L = dAvail
    End If
    If L > dAvail Then L = dAvail
    If L < sp.P Then sMsg = "Gewindelänge kürzer als eine Steigung": Exit Function

    Dim dP As Double: dP = sp.P

    ' Außendurchmesser an der Profilposition (eine Steigung vor dem Anfang)
    Dim rMaj As Double, dTaper As Double
    If sp.Taper Then
        If bInternal Then
            rMaj = (sp.DMaj + dP / 16) / 2                      ' Rc: Bezugsebene = Stirnfläche
            dTaper = -Atn(1 / 32)
        Else
            rMaj = (sp.DMaj - (sp.GaugeLen + dP) / 16) / 2      ' R: Bezugsebene im Abstand GaugeLen
            dTaper = Atn(1 / 32)
        End If
        If INVERT_TAPER Then dTaper = -dTaper
    Else
        rMaj = sp.DMaj / 2
    End If
    Dim rB As Double: rB = rMaj - sp.Depth

    ' Prüfen, ob die Fläche überhaupt im Profilbereich liegt
    Dim dCut As Double
    If sp.Taper Then
        dCut = 0.5 * sp.Depth
    ElseIf bInternal Then
        dCut = rMaj - IIf(rFMin > rB, rFMin, rB)
    Else
        dCut = IIf(rFMax < rMaj, rFMax, rMaj) - rB
    End If
    If dCut < 0.05 * sp.Depth Then
        sMsg = "Flächendurchmesser " & Format(IIf(bInternal, rFMin, rFMax) * 20, "0.###") & _
               " mm passt nicht zu " & sp.Des
        Exit Function
    End If

    ' ---- Profil (Sechseck, radial r / axial relativ zur Profilmitte) ----
    Dim pr(5) As Double, pa(5) As Double, rX As Double
    If bInternal Then
        rX = IIf(rFMin < rB, rFMin, rB) - 0.3 * dP              ' Überstand in die Bohrung
        If rX <= 0 Then rX = 0.5 * rB
        pr(0) = rMaj: pa(0) = sp.HwMaj
        pr(1) = rB:   pa(1) = sp.HwMin
        pr(2) = rX:   pa(2) = sp.HwMin
        pr(3) = rX:   pa(3) = -sp.HwMin
        pr(4) = rB:   pa(4) = -sp.HwMin
        pr(5) = rMaj: pa(5) = -sp.HwMaj
    Else
        If rB <= 0 Then sMsg = "Geometrie ungültig": Exit Function
        rX = IIf(rFMax > rMaj, rFMax, rMaj) + 0.3 * dP          ' Überstand nach außen
        pr(0) = rB:   pa(0) = sp.HwMin
        pr(1) = rMaj: pa(1) = sp.HwMaj
        pr(2) = rX:   pa(2) = sp.HwMaj
        pr(3) = rX:   pa(3) = -sp.HwMaj
        pr(4) = rMaj: pa(4) = -sp.HwMaj
        pr(5) = rB:   pa(5) = -sp.HwMin
    End If

    Dim t0 As Double: t0 = sStart - dirSign * dP
    Dim Hc As Double: Hc = L + 0.5 * dP

    DebugOut sFeat, sp, bInternal, bRight, rFMin, rFMax, sStart, dirSign, L, tMin, tMax, sSrc

    Dim rv As UnitVector: Set rv = PerpUnit(a)
    Dim oOrig As Point
    Set oOrig = m_TG.CreatePoint(O.X + a.X * t0, O.Y + a.Y * t0, O.Z + a.Z * t0)

    Dim oWP As WorkPlane
    Set oWP = oDef.WorkPlanes.AddFixed(oOrig, a, rv)
    Dim oSk As PlanarSketch
    Set oSk = oDef.Sketches.Add(oWP)

    Dim p2(5) As Point2d, i As Long, oMP As Point
    For i = 0 To 5
        Set oMP = m_TG.CreatePoint(oOrig.X + a.X * pa(i) + rv.X * pr(i), _
                                   oOrig.Y + a.Y * pa(i) + rv.Y * pr(i), _
                                   oOrig.Z + a.Z * pa(i) + rv.Z * pr(i))
        Set p2(i) = oSk.ModelToSketchSpace(oMP)
    Next
    Dim oFirst As SketchLine, oL As SketchLine
    Set oFirst = oSk.SketchLines.AddByTwoPoints(p2(0), p2(1))
    Set oL = oFirst
    For i = 2 To 5
        Set oL = oSk.SketchLines.AddByTwoPoints(oL.EndSketchPoint, p2(i))
    Next
    oSk.SketchLines.AddByTwoPoints oL.EndSketchPoint, oFirst.StartSketchPoint

    ' ---- Spirale schneiden ----
    Dim bCW As Boolean: bCW = (Not bRight) Xor INVERT_ROTATION
    Dim rMean As Double: rMean = (rMaj + rB) / 2
    ' Mindestabtrag = 20 % des erwarteten Volumens (2*pi*r * 0,4*P*Tiefe / P * L)
    Dim dMinCut As Double: dMinCut = 0.2 * (6.283 * rMean * 0.4 * dCut * L)
    Dim v0 As Double: v0 = TotalVolume(oDef)
    Dim oAx As WorkAxis, oCoil As CoilFeature

    Set oAx = oDef.WorkAxes.AddFixed(oOrig, m_TG.CreateUnitVector(a.X * dirSign, a.Y * dirSign, a.Z * dirSign))
    Set oCoil = TryCoil(oDef, oSk, oAx, dP, Hc, dTaper, bCW, v0, dMinCut)
    If oCoil Is Nothing Then
        ' Achse andersherum versuchen (Spirale lief aus dem Material heraus)
        oAx.Delete
        Set oAx = oDef.WorkAxes.AddFixed(oOrig, m_TG.CreateUnitVector(-a.X * dirSign, -a.Y * dirSign, -a.Z * dirSign))
        Set oCoil = TryCoil(oDef, oSk, oAx, dP, Hc, dTaper, bCW, v0, dMinCut)
    End If
    If oCoil Is Nothing Then
        On Error Resume Next
        oSk.Delete: oAx.Delete: oWP.Delete
        sMsg = "Spirale konnte nicht erzeugt werden (" & sp.Des & ")"
        Exit Function
    End If

    ' aufräumen / benennen
    On Error Resume Next
    oSk.Visible = False: oWP.Visible = False: oAx.Visible = False
    SafeName oCoil, "Gewinde " & sp.Des
    SafeName oSk, "Gewindeprofil " & sp.Des
    SafeName oWP, "Gewindeebene " & sp.Des
    SafeName oAx, "Gewindeachse " & sp.Des
    On Error GoTo 0

    Set ModelThread = oCoil
End Function

Private Function TryCoil(ByVal oDef As PartComponentDefinition, ByVal oSk As PlanarSketch, _
        ByVal oAx As WorkAxis, ByVal dP As Double, ByVal Hc As Double, ByVal dTaper As Double, _
        ByVal bCW As Boolean, ByVal v0 As Double, ByVal dMinCut As Double) As CoilFeature
    Dim oProf As Profile, oC As CoilFeature
    On Error Resume Next
    Set oProf = oSk.Profiles.AddForSolid
    If oProf Is Nothing Then Exit Function
    ' Profile, Axis, Pitch, Height, Operation, TaperAngle, AxisDirectionReversed, RotateClockwise
    Set oC = oDef.Features.CoilFeatures.AddByPitchAndHeight(oProf, oAx, dP, Hc, kCutOperation, dTaper, False, bCW)
    If Err.Number <> 0 Then Err.Clear: Set oC = Nothing
    If oC Is Nothing Then Exit Function
    If oC.HealthStatus <> kUpToDateHealth Or (v0 - TotalVolume(oDef)) < dMinCut Then
        oC.Delete True, True, True    ' Skizze behalten für zweiten Versuch
        Exit Function
    End If
    Set TryCoil = oC
End Function

'------------------------------------------------------------------------------
' Gewindedaten ermitteln
Private Function GetThreadSpec(ByVal oInfo As Object, ByVal R As Double, ByVal bInternal As Boolean, _
                               ByRef sp As TSpec, ByRef sMsg As String) As Boolean
    Dim sDes As String, sType As String, sAll As String, re As Object, m As Object
    Dim v As Double, bPipe As Boolean

    sDes = Trim$(TryGetString(oInfo, "ThreadDesignation"))
    sType = TryGetString(oInfo, "ThreadType")
    sAll = UCase$(sType & " " & sDes)
    sp.Taper = (TypeName(oInfo) = "TaperedThreadInfo")
    sp.Des = sDes

    If InStr(sAll, "NPT") > 0 Or InStr(sAll, "NPS") > 0 Then
        sMsg = "NPT/NPS-Gewinde wird nicht unterstützt (" & sDes & ")"
        Exit Function
    End If

    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^\s*(G|RP|RC|R)\s*\d"
    bPipe = sp.Taper Or InStr(sAll, "PIPE") > 0 Or InStr(sAll, "ROHR") > 0 Or _
            InStr(sAll, "WHITWORTH") > 0 Or InStr(sAll, "BSP") > 0 Or re.Test(sDes)

    If bPipe Then
        '---------------- ISO-Rohrgewinde (55° Whitworth) ----------------
        Dim dSize As Double, tpi As Double, dMajMM As Double, gaugeMM As Double
        dSize = ParsePipeSize(sDes)
        If Not PipeTable(dSize, tpi, dMajMM, gaugeMM) Then
            sMsg = "Rohrgewindegröße unbekannt (" & sDes & ")"
            Exit Function
        End If
        If sp.Taper And gaugeMM <= 0 Then
            sMsg = "Kegeliges Rohrgewinde " & sDes & " ist in ISO 7-1 nicht genormt"
            Exit Function
        End If
        sp.P = 2.54 / tpi
        sp.DMaj = dMajMM / 10
        sp.GaugeLen = gaugeMM / 10
        sp.Depth = K_WW * sp.P
        If bInternal Then
            sp.HwMaj = sp.P / 12: sp.HwMin = 5 * sp.P / 12
        Else
            sp.HwMaj = 5 * sp.P / 12: sp.HwMin = sp.P / 12
        End If
        If sp.Taper Then
            If bInternal Then sp.Kind = "Rc" Else sp.Kind = "R"
        Else
            If bInternal Then sp.Kind = "G/Rp" Else sp.Kind = "G"
        End If
        If sp.Des = "" Then sp.Des = sp.Kind & " " & PipeSizeText(dSize)
    Else
        '---------------- metrisch ISO (60°) ----------------
        Dim dPdes As Double, dPapi As Double
        re.Pattern = "(^|[^A-Za-z])M\s*(\d+(?:[.,]\d+)?)(?:\s*[xX]\s*(\d+(?:[.,]\d+)?))?"
        If re.Test(sDes) Then
            Set m = re.Execute(sDes)(0)
            sp.DMaj = Val(Replace(m.SubMatches(1), ",", ".")) / 10
            If Len(m.SubMatches(2) & "") > 0 Then dPdes = Val(Replace(m.SubMatches(2), ",", ".")) / 10
        End If
        If sp.DMaj <= 0 Then
            v = Val(Replace(TryGetString(oInfo, "NominalSize"), ",", "."))
            If v > 0 Then sp.DMaj = v / 10
        End If
        If TryGetDouble(oInfo, "Pitch", v) Then
            If v > 0.005 And v < 1 Then dPapi = v
        End If
        ' Priorität: Steigung aus Bezeichnung > Steigung aus API > Regelgewinde
        If dPdes > 0 Then
            sp.P = dPdes
        ElseIf dPapi > 0 Then
            sp.P = dPapi
        ElseIf sp.DMaj > 0 Then
            sp.P = CoarsePitchMM(sp.DMaj * 10) / 10
        End If
        If sp.P <= 0 Then sMsg = "Steigung nicht ermittelbar (" & sDes & ")": Exit Function
        If sp.DMaj <= 0 Then
            If bInternal Then sp.DMaj = 2 * (R + K_ISO * sp.P) Else sp.DMaj = 2 * R
        End If
        sp.Depth = K_ISO * sp.P
        If bInternal Then
            sp.HwMaj = sp.P / 16: sp.HwMin = 3 * sp.P / 8
        Else
            sp.HwMaj = 7 * sp.P / 16: sp.HwMin = sp.P / 8
        End If
        sp.Kind = "ISO"
        If sp.Des = "" Then sp.Des = "M" & Format(sp.DMaj * 10, "0.##") & "x" & Format(sp.P * 10, "0.###")
    End If
    GetThreadSpec = True
End Function

' Rohrgewindegröße in Zoll aus der Bezeichnung ("G 1 1/4", "1/2-14", "Rp3/8" ...)
Private Function ParsePipeSize(ByVal s As String) As Double
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d+)\s*[- ]\s*(\d+)\s*/\s*(\d+)"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        ParsePipeSize = Val(m.SubMatches(0)) + Val(m.SubMatches(1)) / Val(m.SubMatches(2))
        Exit Function
    End If
    re.Pattern = "(\d+)\s*/\s*(\d+)"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        ParsePipeSize = Val(m.SubMatches(0)) / Val(m.SubMatches(1))
        Exit Function
    End If
    re.Pattern = "(\d+(?:[.,]\d+)?)"
    If re.Test(s) Then ParsePipeSize = Val(Replace(re.Execute(s)(0).SubMatches(0), ",", "."))
End Function

Private Function PipeSizeText(ByVal d As Double) As String
    Dim w As Long, f As Double, sf As String
    w = Int(d + 0.0001): f = d - w
    Select Case Round(f * 16)
        Case 1: sf = "1/16"
        Case 2: sf = "1/8"
        Case 4: sf = "1/4"
        Case 6: sf = "3/8"
        Case 8: sf = "1/2"
        Case 10: sf = "5/8"
        Case 12: sf = "3/4"
        Case 14: sf = "7/8"
    End Select
    If w > 0 And sf <> "" Then
        PipeSizeText = w & " " & sf
    ElseIf w > 0 Then
        PipeSizeText = CStr(w)
    Else
        PipeSizeText = sf
    End If
End Function

' ISO 228-1 / ISO 7-1: Gangzahl je Zoll, Außendurchmesser [mm], Bezugslänge ISO 7-1 [mm]
Private Function PipeTable(ByVal dSize As Double, ByRef tpi As Double, ByRef dMaj As Double, _
                           ByRef gauge As Double) As Boolean
    Dim vS As Variant, vT As Variant, vD As Variant, vG As Variant, i As Long
    vS = Array(0.0625, 0.125, 0.25, 0.375, 0.5, 0.625, 0.75, 0.875, 1, 1.125, 1.25, 1.5, _
               1.75, 2, 2.25, 2.5, 2.75, 3, 3.5, 4, 4.5, 5, 5.5, 6)
    vT = Array(28, 28, 19, 19, 14, 14, 14, 14, 11, 11, 11, 11, _
               11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11)
    vD = Array(7.723, 9.728, 13.157, 16.662, 20.955, 22.911, 26.441, 30.201, 33.249, 37.897, 41.91, 47.803, _
               53.746, 59.614, 65.71, 75.184, 81.534, 87.884, 100.33, 113.03, 125.73, 138.43, 151.13, 163.83)
    vG = Array(4, 4, 6, 6.4, 8.2, 0, 9.5, 0, 10.4, 0, 12.7, 12.7, _
               0, 15.9, 0, 17.5, 0, 20.6, 0, 25.4, 0, 28.6, 0, 28.6)
    For i = LBound(vS) To UBound(vS)
        If Abs(dSize - vS(i)) < 0.001 Then
            tpi = vT(i): dMaj = vD(i): gauge = vG(i)
            PipeTable = True
            Exit Function
        End If
    Next
End Function

' ISO-Regelgewinde (DIN 13-1), Steigung in mm
Private Function CoarsePitchMM(ByVal d As Double) As Double
    Select Case Round(d, 1)
        Case 1, 1.2: CoarsePitchMM = 0.25
        Case 1.4: CoarsePitchMM = 0.3
        Case 1.6, 1.8: CoarsePitchMM = 0.35
        Case 2: CoarsePitchMM = 0.4
        Case 2.5: CoarsePitchMM = 0.45
        Case 3: CoarsePitchMM = 0.5
        Case 3.5: CoarsePitchMM = 0.6
        Case 4: CoarsePitchMM = 0.7
        Case 5: CoarsePitchMM = 0.8
        Case 6, 7: CoarsePitchMM = 1
        Case 8: CoarsePitchMM = 1.25
        Case 10: CoarsePitchMM = 1.5
        Case 12: CoarsePitchMM = 1.75
        Case 14, 16: CoarsePitchMM = 2
        Case 18, 20, 22: CoarsePitchMM = 2.5
        Case 24, 27: CoarsePitchMM = 3
        Case 30, 33: CoarsePitchMM = 3.5
        Case 36, 39: CoarsePitchMM = 4
        Case 42, 45: CoarsePitchMM = 4.5
        Case 48, 52: CoarsePitchMM = 5
        Case 56, 60: CoarsePitchMM = 5.5
        Case 64: CoarsePitchMM = 6
        Case Else: CoarsePitchMM = 0
    End Select
End Function

'------------------------------------------------------------------------------
' Geometrie-Hilfsfunktionen
Private Sub GetExtents(ByVal oFace As Face, ByVal O As Point, ByVal a As UnitVector, _
                       ByRef tMin As Double, ByRef tMax As Double, ByRef rMin As Double, ByRef rMax As Double)
    Dim oE As Edge, pMin As Double, pMax As Double, i As Long
    Dim prm(0) As Double, pts() As Double, t As Double, r As Double
    Dim dx As Double, dy As Double, dz As Double
    tMin = 1E+30: tMax = -1E+30: rMin = 1E+30: rMax = -1E+30
    For Each oE In oFace.Edges
        oE.Evaluator.GetParamExtents pMin, pMax
        For i = 0 To 8
            prm(0) = pMin + (pMax - pMin) * i / 8
            oE.Evaluator.GetPointAtParam prm, pts
            dx = pts(0) - O.X: dy = pts(1) - O.Y: dz = pts(2) - O.Z
            t = dx * a.X + dy * a.Y + dz * a.Z
            r = Sqr((dx - a.X * t) ^ 2 + (dy - a.Y * t) ^ 2 + (dz - a.Z * t) ^ 2)
            If t < tMin Then tMin = t
            If t > tMax Then tMax = t
            If r < rMin Then rMin = r
            If r > rMax Then rMax = r
        Next
    Next
End Sub

' Welches Flächenende ist "offen" (Stirnfläche/Fase zeigt vom Gewinde weg)?
Private Sub GetOpenEnds(ByVal oFace As Face, ByVal O As Point, ByVal a As UnitVector, _
        ByVal tMin As Double, ByVal tMax As Double, ByRef bMin As Boolean, ByRef bMax As Boolean)
    Dim oE As Edge, oF As Face, oOther As Face
    Dim pMin As Double, pMax As Double, prm(0) As Double, pts() As Double, n() As Double
    Dim t As Double, na As Double, dTol As Double
    dTol = 0.02 * (tMax - tMin) + TOL
    On Error Resume Next
    For Each oE In oFace.Edges
        oE.Evaluator.GetParamExtents pMin, pMax
        prm(0) = (pMin + pMax) / 2
        oE.Evaluator.GetPointAtParam prm, pts
        t = (pts(0) - O.X) * a.X + (pts(1) - O.Y) * a.Y + (pts(2) - O.Z) * a.Z
        If Abs(t - tMax) < dTol Or Abs(t - tMin) < dTol Then
            Set oOther = Nothing
            For Each oF In oE.Faces
                If oF.TransientKey <> oFace.TransientKey Then Set oOther = oF
            Next
            If Not oOther Is Nothing Then
                Err.Clear
                Erase n
                oOther.Evaluator.GetNormalAtPoint pts, n
                If Err.Number = 0 Then
                    na = n(0) * a.X + n(1) * a.Y + n(2) * a.Z
                    If Abs(t - tMax) < dTol Then
                        If na > 0.2 Then bMax = True
                    Else
                        If na < -0.2 Then bMin = True
                    End If
                End If
            End If
        End If
    Next
End Sub

' Basispunkt/Richtung aus der Gewindeinfo (falls die API sie liefert)
Private Function GetThreadBase(ByVal oInfo As Object, ByVal O As Point, ByVal a As UnitVector, _
        ByVal tMin As Double, ByVal tMax As Double, ByRef sStart As Double, _
        ByRef dirSign As Double, ByRef dLen As Double) As Boolean
    Dim bp As Object, td As Object, dDot As Double
    On Error Resume Next
    Set bp = oInfo.ThreadBasePoint
    Set td = oInfo.ThreadDirection
    On Error GoTo 0
    If bp Is Nothing Or td Is Nothing Then Exit Function
    dLen = Sqr(td.X ^ 2 + td.Y ^ 2 + td.Z ^ 2)
    dDot = td.X * a.X + td.Y * a.Y + td.Z * a.Z
    If dLen < 0.000000001 Then Exit Function
    If Abs(dDot) < 0.5 * dLen Then Exit Function
    sStart = (bp.X - O.X) * a.X + (bp.Y - O.Y) * a.Y + (bp.Z - O.Z) * a.Z
    If sStart < tMin - 0.05 Or sStart > tMax + 0.05 Then Exit Function
    dirSign = Sgn(dDot)
    If dirSign > 0 And sStart > tMax - TOL Then Exit Function
    If dirSign < 0 And sStart < tMin + TOL Then Exit Function
    GetThreadBase = True
End Function

Private Function IsInternalFace(ByVal oFace As Face, ByVal O As Point, ByVal a As UnitVector) As Boolean
    Dim p As Point, pts(2) As Double, n() As Double, t As Double
    Set p = oFace.PointOnFace
    pts(0) = p.X: pts(1) = p.Y: pts(2) = p.Z
    oFace.Evaluator.GetNormalAtPoint pts, n
    t = (p.X - O.X) * a.X + (p.Y - O.Y) * a.Y + (p.Z - O.Z) * a.Z
    IsInternalFace = ((p.X - O.X - a.X * t) * n(0) + (p.Y - O.Y - a.Y * t) * n(1) + _
                      (p.Z - O.Z - a.Z * t) * n(2)) < 0
End Function

Private Function PerpUnit(ByVal a As UnitVector) As UnitVector
    Dim vRef As Vector, c As Vector
    If Abs(a.X) < 0.9 Then Set vRef = m_TG.CreateVector(1, 0, 0) Else Set vRef = m_TG.CreateVector(0, 1, 0)
    Set c = a.AsVector.CrossProduct(vRef)
    c.Normalize
    Set PerpUnit = c.AsUnitVector
End Function

Private Function TotalVolume(ByVal oDef As PartComponentDefinition) As Double
    Dim b As SurfaceBody, v As Double
    For Each b In oDef.SurfaceBodies
        v = v + b.Volume(0.000001)
    Next
    TotalVolume = v
End Function

'------------------------------------------------------------------------------
' Diagnose: Zeile im Direktfenster (Strg+G), optional als Meldung
Private Sub DebugOut(ByVal sFeat As String, sp As TSpec, ByVal bInt As Boolean, ByVal bRight As Boolean, _
        ByVal rFMin As Double, ByVal rFMax As Double, ByVal sStart As Double, ByVal dirSign As Double, _
        ByVal L As Double, ByVal tMin As Double, ByVal tMax As Double, ByVal sSrc As String)
    Dim s As String
    s = sFeat & " | " & sp.Des & " | " & sp.Kind & IIf(sp.Taper, " kegelig", "") & _
        " | " & IIf(bInt, "innen", "außen") & IIf(bRight, " rechts", " links") & _
        " | P=" & Format(sp.P * 10, "0.000") & " D=" & Format(sp.DMaj * 10, "0.000") & _
        " Tiefe=" & Format(sp.Depth * 10, "0.000") & _
        " | Fläche D=" & Format(rFMin * 20, "0.000") & ".." & Format(rFMax * 20, "0.000") & _
        " Länge=" & Format((tMax - tMin) * 10, "0.00") & _
        " | Start " & Format((sStart - tMin) * 10, "0.00") & " mm, Richtung " & IIf(dirSign > 0, "+", "-") & _
        " (" & sSrc & ") L=" & Format(L * 10, "0.00") & " mm"
    Debug.Print s
    If SHOW_DEBUG Then MsgBox s, vbInformation, "Gewinde modellieren - Diagnose"
End Sub

'------------------------------------------------------------------------------
' API-Eigenschaften fehlertolerant lesen (nicht jede Eigenschaft existiert in 2016)
Private Function TryGetDouble(ByVal obj As Object, ByVal sProp As String, ByRef v As Double) As Boolean
    Dim o As Object, x As Variant
    On Error Resume Next
    Set o = CallByName(obj, sProp, VbGet)
    If Err.Number = 0 Then
        If Not o Is Nothing Then
            x = o.Value                       ' Parameter-Objekt
            If Err.Number = 0 Then v = CDbl(x): TryGetDouble = True
        End If
        Exit Function
    End If
    Err.Clear
    x = CallByName(obj, sProp, VbGet)
    If Err.Number = 0 Then
        If IsNumeric(x) Then v = CDbl(x): TryGetDouble = True
    End If
End Function

Private Function TryGetBool(ByVal obj As Object, ByVal sProp As String, ByRef b As Boolean) As Boolean
    Dim x As Variant
    On Error Resume Next
    x = CallByName(obj, sProp, VbGet)
    If Err.Number = 0 Then
        If VarType(x) = vbBoolean Then b = x: TryGetBool = True
    End If
End Function

Private Function TryGetString(ByVal obj As Object, ByVal sProp As String) As String
    Dim x As Variant
    On Error Resume Next
    x = CallByName(obj, sProp, VbGet)
    If Err.Number = 0 Then TryGetString = CStr(x)
End Function

Private Sub SafeName(ByVal o As Object, ByVal sBase As String)
    Dim i As Long
    On Error Resume Next
    For i = 1 To 99
        Err.Clear
        If i = 1 Then o.Name = sBase Else o.Name = sBase & " (" & i & ")"
        If Err.Number = 0 Then Exit Sub
    Next
End Sub
