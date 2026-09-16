' ModelCheck - a tool to grade a SOLIDWORKS part by both robustness and conventions.
' Robustness governs checks that could have mechanical consequence/errors down the line.
' Conventions governs checks that are not explicitly required for the part to work well, but will cause issues for the next reader.
' A "rubric" is made for these two scores to give both newer SOLIDWORKS users and experienced helpful outlook.

' For this code, variable/dim names use the sw prefix for SOLIDWORKS, the m prefix for module, and no prefix for local variables.

Option Explicit

' Section Identifiers
Private Const SEC_ROBUST As Long = 1
Private Const SEC_CONV   As Long = 2

' Robustness Weights
Private Const W_SKETCH As Double = 100

' Convention Weights - Totaling 100
Private Const W_NAMING   As Double = 40
Private Const W_ORPHAN   As Double = 20
Private Const W_MATERIAL As Double = 20
Private Const W_PROPS    As Double = 20

' SOLIDWORKS Document Types
Private Const SW_DOC_PART     As Long = 1
Private Const SW_DOC_ASSEMBLY As Long = 2
Private Const SW_DOC_DRAWING  As Long = 3

' Sketch Constraint Status: 2 = Underdefined, 3 = Fully Defined
Private Const SKETCH_UNDER As Long = 2
Private Const SKETCH_FULL  As Long = 3

Private Const MAX_FINDINGS As Long = 6
Private Const MAX_NAMED    As Long = 6

Private Type RubricItem
    Label    As String
    Earned   As Double
    Possible As Double
    Assessed As Boolean
    Section  As Long
End Type

Private swApp   As Object
Private swModel As Object
Private swPart  As Object

Private mRubric(1 To 5) As RubricItem
Private mFindings       As Collection

' Raw Facts
Private mMaterial     As String
Private mMaterialReq  As Boolean
Private mSketchFull   As Long
Private mFeatTotal    As Long
Private mFeatDefault  As Long
Private mPropsFound   As Long
Private mPropsTotal   As Long
Private mMissingProps As String
Private mFaceCheckOk  As Boolean    ' False if GetParents is unusable

Private mVisited      As Collection
Private mAbsorbed     As Collection
Private mSketchNames  As Collection
Private mUnderDefined As Collection
Private mOverDefined  As Collection
Private mOrphans      As Collection
Private mFaceBased    As Collection


Sub Main()

    Set swApp = Application.SldWorks

    If Not AcquirePart() Then Exit Sub

    ResetState
    ScanModel

    mMaterialReq = AskMaterialRequired()

    BuildRubric

    MsgBox FormatReport(), vbInformation, "Model check"

End Sub

Private Function AcquirePart() As Boolean

    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "Open a part, then run the macro again.", _
               vbExclamation, "Nothing to check"
        Exit Function
    End If

    If swModel.GetType <> SW_DOC_PART Then
        MsgBox "This tool checks parts. The active document is a " & _
               DocTypeName(swModel.GetType) & ".", _
               vbExclamation, "Wrong document type"
        Exit Function
    End If

    Set swPart = swModel
    AcquirePart = True

End Function

Private Function DocTypeName(ByVal t As Long) As String
    Select Case t
        Case SW_DOC_ASSEMBLY: DocTypeName = "assembly"
        Case SW_DOC_DRAWING:  DocTypeName = "drawing"
        Case Else:            DocTypeName = "unknown document"
    End Select
End Function

' Not all students and CAD designers are working with materials.
' So, if a material is not found, the program will not flag you for it being missing and first ask if it is required.

Private Function AskMaterialRequired() As Boolean

    If Len(mMaterial) > 0 Then AskMaterialRequired = True: Exit Function

    Dim answer As Long
    answer = MsgBox("This part has no material assigned." & vbCrLf & vbCrLf & _
                    "Does it need one?" & vbCrLf & vbCrLf & _
                    "Choose No for a concept model or sketch study, and the " & _
                    "material check will be left out of the report.", _
                    vbYesNo + vbQuestion, "Does this part need a material?")

    AskMaterialRequired = (answer = vbYes)

End Function

Private Sub ResetState()

    Dim i As Long
    For i = 1 To 5
        mRubric(i).Label = ""
        mRubric(i).Earned = 0
        mRubric(i).Possible = 0
        mRubric(i).Assessed = False
        mRubric(i).Section = 0
    Next i

    Set mFindings = New Collection
    Set mVisited = New Collection
    Set mAbsorbed = New Collection
    Set mSketchNames = New Collection
    Set mUnderDefined = New Collection
    Set mOverDefined = New Collection
    Set mOrphans = New Collection
    Set mFaceBased = New Collection

    mMaterial = ""
    mMaterialReq = False
    mSketchFull = 0
    mFeatTotal = 0
    mFeatDefault = 0
    mPropsFound = 0
    mPropsTotal = 0
    mMissingProps = ""
    mFaceCheckOk = True

End Sub

' Scanning of the Model which involves no judging or scoring
Private Sub ScanModel()
    ReadMaterial
    ReadProperties
    WalkTree
    FindOrphans
End Sub

Private Sub ReadMaterial()
    Dim db As Variant
    On Error Resume Next
    mMaterial = swPart.GetMaterialPropertyName2( _
                    swModel.ConfigurationManager.ActiveConfiguration.Name, db)
    If Err.Number <> 0 Then Err.Clear: mMaterial = ""
    On Error GoTo 0
    mMaterial = Trim$(mMaterial)
End Sub

Private Sub ReadProperties()

    Dim wanted As Variant
    wanted = Array("Description", "PartNo", "Revision")
    mPropsTotal = UBound(wanted) - LBound(wanted) + 1

    Dim mgr As Object
    Set mgr = swModel.Extension.CustomPropertyManager("")

    Dim i As Long, raw As Variant, resolved As Variant
    For i = LBound(wanted) To UBound(wanted)
        raw = "": resolved = ""
        On Error Resume Next
        mgr.Get4 CStr(wanted(i)), False, raw, resolved
        If Err.Number <> 0 Then Err.Clear: resolved = ""
        On Error GoTo 0

        If Len(Trim$(CStr(resolved))) > 0 Then
            mPropsFound = mPropsFound + 1
        Else
            mMissingProps = mMissingProps & _
                            IIf(Len(mMissingProps) > 0, ", ", "") & CStr(wanted(i))
        End If
    Next i

End Sub

' GetNextFeature returns both absorbed sketches at the top level and the underneath sketches, making the same sketch visible twice.
' With that, we could not determine if a sketch is used.
' Therefore, ownership is established by reaching a sketch as a sub-feature.

Private Sub WalkTree()

    Dim swFeat As Object
    Set swFeat = swModel.FirstFeature

    Do While Not swFeat Is Nothing
        Inspect swFeat, True
        WalkSubFeatures swFeat
        Set swFeat = swFeat.GetNextFeature
    Loop

End Sub

Private Sub WalkSubFeatures(ByVal parent As Object)

    Dim swSub As Object
    Set swSub = parent.GetFirstSubFeature

    Do While Not swSub Is Nothing
        Inspect swSub, False
        WalkSubFeatures swSub
        Set swSub = swSub.GetNextSubFeature
    Loop

End Sub

Private Sub Inspect(ByVal swFeat As Object, ByVal isTopLevel As Boolean)

    Dim featType As String, featName As String
    featType = swFeat.GetTypeName2
    featName = swFeat.Name

    If IsSystemFeature(featType) Then Exit Sub

    Dim isSketch As Boolean
    isSketch = (featType = "ProfileFeature" Or featType = "3DProfileFeature")

    If isSketch And Not isTopLevel Then SetAdd mAbsorbed, featName

    If SetHas(mVisited, featName) Then Exit Sub
    SetAdd mVisited, featName

    If isSketch Then
        InspectSketch swFeat, featName
        Exit Sub
    End If

    mFeatTotal = mFeatTotal + 1
    If IsDefaultName(featName) Then mFeatDefault = mFeatDefault + 1

End Sub

Private Sub InspectSketch(ByVal swFeat As Object, ByVal featName As String)

    Dim swSketch As Object
    Set swSketch = swFeat.GetSpecificFeature2
    If swSketch Is Nothing Then Exit Sub

    If IsSketchEmpty(swSketch) Then Exit Sub

    mSketchNames.Add featName

    Dim state As Long
    state = ConstraintState(swSketch)

    Select Case state
        Case SKETCH_FULL
            mSketchFull = mSketchFull + 1
        Case SKETCH_UNDER
            mUnderDefined.Add featName
        Case Else
            mOverDefined.Add featName & " (status " & state & ")"
    End Select

    If mFaceCheckOk Then
        If IsFaceBased(swFeat) Then mFaceBased.Add featName
    End If

End Sub


Private Function IsFaceBased(ByVal swFeat As Object) As Boolean

    Dim parents As Variant, i As Long

    On Error Resume Next
    parents = swFeat.GetParents
    If Err.Number <> 0 Then
        Err.Clear
        mFaceCheckOk = False
        Exit Function
    End If
    On Error GoTo 0

    If IsEmpty(parents) Then Exit Function
    If Not IsArray(parents) Then Exit Function
    If UBound(parents) < LBound(parents) Then Exit Function

    For i = LBound(parents) To UBound(parents)
        On Error Resume Next
        If parents(i).GetTypeName2 = "RefPlane" Then
            If Err.Number = 0 Then Exit Function     ' plane found: stable
        End If
        Err.Clear
        On Error GoTo 0
    Next i

    IsFaceBased = True

End Function

Private Function IsSketchEmpty(ByVal swSketch As Object) As Boolean

    Dim segs As Variant

    On Error Resume Next
    segs = swSketch.GetSketchSegments
    If Err.Number <> 0 Then
        Err.Clear
        IsSketchEmpty = False   ' fail safe where we judge rather than skipping
        Exit Function
    End If
    On Error GoTo 0

    If IsEmpty(segs) Then IsSketchEmpty = True: Exit Function
    If Not IsArray(segs) Then IsSketchEmpty = True: Exit Function
    If UBound(segs) < LBound(segs) Then IsSketchEmpty = True

End Function

Private Function ConstraintState(ByVal swSketch As Object) As Long
    On Error Resume Next
    ConstraintState = swSketch.GetConstrainedStatus
    If Err.Number <> 0 Then Err.Clear: ConstraintState = SKETCH_FULL
    On Error GoTo 0
End Function

Private Sub FindOrphans()
    Dim i As Long
    For i = 1 To mSketchNames.Count
        If Not SetHas(mAbsorbed, mSketchNames(i)) Then mOrphans.Add mSketchNames(i)
    Next i
End Sub

Private Sub SetAdd(ByRef c As Collection, ByVal key As String)
    On Error Resume Next
    c.Add True, key
    Err.Clear
    On Error GoTo 0
End Sub

Private Function SetHas(ByVal c As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c(key)
    SetHas = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


Private Function IsSystemFeature(ByVal featType As String) As Boolean

    If Right$(featType, 6) = "Folder" Then IsSystemFeature = True: Exit Function
    If Right$(featType, 5) = "Light" Then IsSystemFeature = True: Exit Function

    Select Case featType
        Case "OriginProfileFeature", "DetailCabinet", _
             "RefPlane", "RefAxis", "RefPoint", "CoordSys", _
             "Camera", "Scene", "Environment"
            IsSystemFeature = True
    End Select

End Function

' Known Limitation: This section only compares Feature Names vs. their defaults.
' If someone renames features to a1, a2, a3... They will be given full marks.
' Effectively, this only grades IF someone renames it, not grading them on the naming itself.
Private Function IsDefaultName(ByVal nm As String) As Boolean

    Static prefixes As Variant
    If IsEmpty(prefixes) Then
        prefixes = Array( _
            "Boss-Extrude", "Cut-Extrude", "Boss-Revolve", "Cut-Revolve", _
            "Boss-Sweep", "Cut-Sweep", "Boss-Loft", "Cut-Loft", _
            "Boss-Thicken", "Cut-Thicken", "Boss-Extrude-Thin", _
            "Fillet", "Chamfer", "Shell", "Draft", "Rib", "Dome", "Wrap", _
            "Hole", "M_Hole", "CBORE", "CSK", "Thread", _
            "LPattern", "CirPattern", "SketchPattern", "CurvePattern", _
            "Mirror", "MirrorPattern", "Move-Copy Body", _
            "Split Line", "Scale", "Combine", "Indent", "Deform", _
            "Sheet-Metal", "Base-Flange", "Edge-Flange", "Sketched Bend", _
            "Miter Flange", "Hem", "Jog", "Flat-Pattern", _
            "Surface-Extrude", "Surface-Trim", "Surface-Knit", "Surface-Fill", _
            "Curve", "Helix/Spiral", "Sketch", "3DSketch")
    End If

    Dim i As Long, p As String
    For i = LBound(prefixes) To UBound(prefixes)
        p = prefixes(i)
        If Len(nm) > Len(p) Then
            If StrComp(Left$(nm, Len(p)), p, vbTextCompare) = 0 Then
                If IsAllDigits(Mid$(nm, Len(p) + 1)) Then
                    IsDefaultName = True
                    Exit Function
                End If
            End If
        End If
    Next i

End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function


Private Sub BuildRubric()

    Dim i As Long
    Dim sketchTotal As Long
    sketchTotal = mSketchNames.Count

    'Robustness
    mRubric(1).Label = "Sketches fully defined"
    mRubric(1).Possible = W_SKETCH
    mRubric(1).Section = SEC_ROBUST
    mRubric(1).Assessed = (sketchTotal > 0)
    If mRubric(1).Assessed Then
        mRubric(1).Earned = W_SKETCH * (mSketchFull / sketchTotal)
    End If
    For i = 1 To mOverDefined.Count
        AddFinding "fail", mOverDefined(i) & " is not fully defined"
    Next i
    For i = 1 To mUnderDefined.Count
        AddFinding "warn", mUnderDefined(i) & " is under defined - it will " & _
                   "move when an upstream dimension changes"
    Next i

    'Convention 1 - Feature Naming
    mRubric(2).Label = "Features renamed"
    mRubric(2).Possible = W_NAMING
    mRubric(2).Section = SEC_CONV
    mRubric(2).Assessed = (mFeatTotal > 0)
    If mRubric(2).Assessed Then
        mRubric(2).Earned = W_NAMING * ((mFeatTotal - mFeatDefault) / mFeatTotal)
        If mFeatDefault > 0 Then
            AddFinding "warn", mFeatDefault & " of " & mFeatTotal & _
                       " features still have default names"
        End If
    End If

    ' Convention 2 - Orphan/Unused Sketches
    mRubric(3).Label = "No unused sketches"
    mRubric(3).Possible = W_ORPHAN
    mRubric(3).Section = SEC_CONV
    mRubric(3).Assessed = (sketchTotal > 0)
    If mRubric(3).Assessed Then
        mRubric(3).Earned = W_ORPHAN * (1 / (1 + mOrphans.Count))
    End If
    For i = 1 To mOrphans.Count
        AddFinding "warn", mOrphans(i) & " is not used by any feature"
    Next i

    ' Convention 3 - Material Properties
    mRubric(4).Label = "Material assigned"
    mRubric(4).Possible = W_MATERIAL
    mRubric(4).Section = SEC_CONV
    mRubric(4).Assessed = mMaterialReq
    If mRubric(4).Assessed Then
        If Len(mMaterial) = 0 Then
            AddFinding "warn", "No material assigned - mass and CG will be wrong"
        Else
            mRubric(4).Earned = W_MATERIAL
        End If
    End If

    ' Convention 4 - File Properties
    mRubric(5).Label = "File properties filled"
    mRubric(5).Possible = W_PROPS
    mRubric(5).Section = SEC_CONV
    mRubric(5).Assessed = True
    mRubric(5).Earned = W_PROPS * (mPropsFound / mPropsTotal)
    If Len(mMissingProps) > 0 Then
        AddFinding "warn", "Missing properties: " & mMissingProps
    End If

End Sub

Private Sub AddFinding(ByVal state As String, ByVal text As String)
    mFindings.Add state & "|" & text
End Sub

' SCORING
Private Function SectionEarned(ByVal sec As Long) As Double
    Dim i As Long
    For i = 1 To 5
        If mRubric(i).Section = sec And mRubric(i).Assessed Then _
            SectionEarned = SectionEarned + mRubric(i).Earned
    Next i
End Function

Private Function SectionPossible(ByVal sec As Long) As Double
    Dim i As Long
    For i = 1 To 5
        If mRubric(i).Section = sec And mRubric(i).Assessed Then _
            SectionPossible = SectionPossible + mRubric(i).Possible
    Next i
End Function

Private Function SectionScore(ByVal sec As Long) As Double
    If SectionPossible(sec) = 0 Then Exit Function
    SectionScore = (SectionEarned(sec) / SectionPossible(sec)) * 100
End Function

Private Function RobustnessVerdict() As String
    If SectionPossible(SEC_ROBUST) = 0 Then
        RobustnessVerdict = "No sketches to assess."
        Exit Function
    End If
    Select Case True
        Case SectionScore(SEC_ROBUST) >= 100
            RobustnessVerdict = "Every sketch is locked down."
        Case SectionScore(SEC_ROBUST) >= 60
            RobustnessVerdict = "Some geometry can still drift on edit."
        Case Else
            RobustnessVerdict = "Most of this model will move if anything upstream changes."
    End Select
End Function


' PRESENT
' MsgBox truncates at roughly 1024 characters, which is why the findings list and the CAUTION name list are both capped.

Private Function FormatReport() As String

    Dim s As String, i As Long

    s = swModel.GetTitle & vbCrLf & String$(46, "-") & vbCrLf & vbCrLf

    ' Robustness
    s = s & "ROBUSTNESS   " & ScoreText(SEC_ROBUST) & vbCrLf
    s = s & "  " & RobustnessVerdict() & vbCrLf
    For i = 1 To 5
        If mRubric(i).Section = SEC_ROBUST Then s = s & RubricLine(i)
    Next i
    If mSketchNames.Count > 0 Then
        s = s & "  (" & mSketchFull & " of " & mSketchNames.Count & _
                " sketches)" & vbCrLf
    End If

    ' Conventions
    s = s & vbCrLf & "CONVENTIONS   " & ScoreText(SEC_CONV) & vbCrLf
    For i = 1 To 5
        If mRubric(i).Section = SEC_CONV Then s = s & RubricLine(i)
    Next i

    ' Caution
    s = s & vbCrLf & "CAUTION" & vbCrLf & CautionBlock()

    ' Findings
    s = s & vbCrLf & "WHAT TO FIX" & vbCrLf
    If mFindings.Count = 0 Then
        s = s & "  Nothing flagged." & vbCrLf
    Else
        Dim shown As Long, parts As Variant
        For i = 1 To mFindings.Count
            If shown >= MAX_FINDINGS Then Exit For
            parts = Split(mFindings(i), "|")
            s = s & "  " & IIf(parts(0) = "fail", "[X] ", "[!] ") & _
                parts(1) & vbCrLf
            shown = shown + 1
        Next i
        If mFindings.Count > MAX_FINDINGS Then
            s = s & "  ... and " & (mFindings.Count - MAX_FINDINGS) & _
                    " more" & vbCrLf
        End If
    End If

    FormatReport = s

End Function

Private Function CautionBlock() As String

    If Not mFaceCheckOk Then
        CautionBlock = "  Sketch anchoring could not be checked on this " & _
                       "version." & vbCrLf
        Exit Function
    End If

    If mSketchNames.Count = 0 Then
        CautionBlock = "  No sketches to check." & vbCrLf
        Exit Function
    End If

    If mFaceBased.Count = 0 Then
        CautionBlock = "  All sketches are anchored to reference planes." & vbCrLf
        Exit Function
    End If

    Dim s As String, i As Long, names As String
    s = "  " & mFaceBased.Count & " of " & mSketchNames.Count & _
        " sketches are built on model geometry." & vbCrLf
    s = s & "  These break if the geometry they reference moves or is " & _
        "consumed." & vbCrLf

    For i = 1 To mFaceBased.Count
        If i > MAX_NAMED Then
            names = names & ", +" & (mFaceBased.Count - MAX_NAMED) & " more"
            Exit For
        End If
        names = names & IIf(Len(names) > 0, ", ", "") & mFaceBased(i)
    Next i
    s = s & "    " & names & vbCrLf

    CautionBlock = s

End Function

Private Function ScoreText(ByVal sec As Long) As String
    If SectionPossible(sec) = 0 Then
        ScoreText = "n/a"
    Else
        ScoreText = Format$(SectionScore(sec), "0") & " / 100"
    End If
End Function

Private Function RubricLine(ByVal i As Long) As String
    If mRubric(i).Assessed Then
        RubricLine = "  " & Pad(Format$(mRubric(i).Earned, "0") & "/" & _
                     Format$(mRubric(i).Possible, "0"), 8) & _
                     mRubric(i).Label & vbCrLf
    Else
        RubricLine = "  " & Pad("n/a", 8) & mRubric(i).Label & vbCrLf
    End If
End Function

Private Function Pad(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then Pad = s & " " Else Pad = s & Space$(n - Len(s))
End Function

