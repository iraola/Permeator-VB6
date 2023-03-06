VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CINPPermeator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private myContainer As ExtnUnitOperationContainer
    Private dynContainer As ExtnDynUnitOpContainer
    '***************************************************************************'
    '                             VB Variables                                  '
    '***************************************************************************'
    ' Indices of permeated components: in this order: H2, HD, HT, D2, DT, T2
    Private nComp As Integer
    Private nPerm As Integer          ' number of permeating species (6 if all hydrogens)
    Private nHeteroNuclear As Integer
    Private nPermAtom As Integer      ' number of permeating atoms (3 if H, D and T)
    Private permCoeffs() As Variant   ' array of coefficients to compute partial pressures
                                      ' contribution per diatomic molecule
    Private permIndicesHomo() As Variant
    Private permIndicesHetero() As Variant ' array of indices to locate heteronuclear molecules
    ' for each atom in HYSYS, with the shape: {{-1, HD, HT}, {HD, -1, DT}, {HT, DT, -1}}
    Private permIndices() As Variant  ' array of indices to adress each diatomic species in
    ' hysys component list
    Private IsForgetting As Boolean

    ' Plot
    Dim myPlotNameH As InternalTextVariable
    Dim myPlotH As TwoDimensionalPlot
    Dim myPlotNameT As InternalTextVariable
    Dim myPlotT As TwoDimensionalPlot
    Dim GrafL As Variant

    '***************************************************************************'
    '                            EDF Variables                                  '
    '***************************************************************************'
    Private edfInlet As ProcessStream
    Private edfPermeate As ProcessStream
    Private edfRetentate As ProcessStream
    Private edfPermIn As ProcessStream
    Private edfThick As InternalRealVariable
    Private edfDiam As InternalRealVariable
    Private edfLen As InternalRealVariable
    Private edfAperm As InternalRealVariable
    Private edfNtubes As InternalRealVariable
    Private edfk As InternalRealVariable
    Private edfPH As InternalRealVariable
    Private edfPD As InternalRealVariable
    Private edfPT As InternalRealVariable
    Private edfTotalPermMolFlow As InternalRealVariable
    'Private edfComposition As InternalRealFlexVariable
    'Private edfCompName As InternalTextFlexVariable
    'Private edfStreamName As InternalTextFlexVariable
    'Private edfVapFrac As InternalRealFlexVariable
    'Private edfTemperature As InternalRealFlexVariable
    'Private edfPressure As InternalRealFlexVariable
    'Private edfMolFlow As InternalRealFlexVariable
    'Private edfMassFlow As InternalRealFlexVariable
    Private edfPressDrop As InternalRealVariable
    Private edfPermPressDrop As InternalRealVariable
    'Private edfDiamExt As InternalRealVariable
    Private edfLengthPos As InternalRealFlexVariable
    Private edfCompT As InternalRealFlexVariable
    'Private edfCompH As InternalRealFlexVariable
    Private edfNpoints As InternalRealVariable

    '***************************************************************************'
    '                            Physical Variables                             '
    '***************************************************************************'
    ' Geometry
    Dim totalLen As Double, thick As Double, Aperm As Double
    ' Streams
    Dim fluidList() As fluid
    Dim StreamList() As ProcessStream
    ' Permeation flow
    Dim GlobalPermComponentMolarFlow() As Double
    Dim FpermCellsPlot() As Double

    ' Constants
    Private Const pi As Double = 3.14159265359
    Private Const R As Double = 8.314472    ' kJ/kmol-K
    Private Const MOLFLOW_UNITS As String = "kgmole/s"
    Private A() As Variant
    Private E() As Variant


Implements HYSYS.ExtnUnitOperation
Implements HYSYS.ExtensionObject


Private Function ExtnUnitOperation_Initialize(ByVal Container As HYSYS.ExtnUnitOperationContainer, ByVal IsRecalling As Boolean) As Long
    Dim IRV As InternalRealVariable
    On Error GoTo ErrorTrap
    ' Initiali  ze container
    Set myContainer = Container
    ' Initialize EDF variables
    Call pointedEDFVariables
    ' Recall check
    If Not IsRecalling Then
        'Step 3 - set the NumberOfPoints variable to 0.
    Else
        ' Visibility controllers
        ' I PREFER TO SEE THE LAST USED VIEW INSTEAD OF THE MAIN ONE
        'IRV = myContainer.FindVariable("Design_enum").Variable
        'IRV.Value = 0
    End If
    ' Global variables: first calculation
    nPerm = 6
    nHeteroNuclear = 3
    nPermAtom = 3
    A = Array(0.0000000558, 0.0000000343, 2.63207547169811E-08)
    E = Array(-6304, -6156, -6304)
    permCoeffs = Array(Array(1, 0.5, 0.5, 0, 0, 0), Array(0, 0.5, 0, 1, 0.5, 0), Array(0, 0, 0.5, 0, 0.5, 1))
    ' Geometry initializations
    thick = edfThick.GetValue()
    Aperm = edfAperm.GetValue()
    totalLen = edfLen.GetValue() * edfNtubes.GetValue()
    Call CalcAperm
    ' Plotting initializations
    If edfLengthPos Is Nothing Or edfCompT Is Nothing Then
        Dim auxZeros() As Double
        If edfNpoints Is Nothing Or edfNpoints.Value = -32767 Then
            ReDim auxZeros(0)
        Else
            ReDim auxZeros(edfNpoints.Value - 1)
        End If
        Call edfLengthPos.SetValues(auxZeros)
        Call edfCompT.SetValues(auxZeros)
    End If
    Call CreatePlot
    If Not edfLen Is Nothing And Not edfNpoints Is Nothing And edfNpoints.Value <> -32767 Then
        GrafL = LengthVector(edfLen.Value, edfNpoints.Value)
    End If
    ' Loop for setting index for components of interest (based on Inlet's basis manager)
    Call IniCompIndex
    ' Return Initialize
    ExtnUnitOperation_Initialize = CurrentExtensionVersion_enum.extnCurrentVersion
    Exit Function

ErrorTrap:
    MsgBox ("Initialize Error: " + Err.GetException().ToString)
End Function

Private Sub ExtnUnitOperation_Execute(ByVal Forgetting As Boolean)
  ' execute gets hit twice, once on a forgetting pass and then on _
    'a calculate pass
    Dim flashPerm As Double, flashNoPerm As Double
    Dim Tin As Double, retPressure As Double, permPressDrop As Double
    Dim i As Long
    On Error GoTo ErrorTrap

    ' Step 1 - Forgetting check
    If Forgetting Then
        IsForgetting = True
        Exit Sub
    End If

    ' Step 2 - Check that we have enough information to Calculate
    Call CheckExtnRequirements

    ' Step 3 - Build vector of Streams and Fluids (StreamList and fluidList)
    Call BuildStreamsAndFluids

    ' Step 4 - Calculate Permeation
    nComp = UBound(edfInlet.ComponentMolarFlowValue) + 1
    Dim permMolarFlows() As Double
    ReDim permMolarFlows(nComp - 1)
    permMolarFlows = Permeation() ' kmol/s

    ' Step 5 - Set the calculated values to the list of fluids and product streams
    Call SetPermeationFlows(StreamList, permMolarFlows)

    '' Step 6 - Pressure drop and flash specifications
    Tin = edfInlet.TemperatureValue
    retPressure = edfInlet.PressureValue - edfPressDrop.Value
    permPressDrop = edfInlet.PressureValue - edfPermeate.PressureValue
    Call edfRetentate.Temperature.Calculate(Tin)
    Call edfPermeate.Temperature.Calculate(Tin)
    Call edfRetentate.Pressure.Calculate(retPressure)
    Call edfPermPressDrop.Calculate(permPressDrop)

    ' Step 7 - Final Balance, checks and EDF visualization
    ' Now use the .Balance method of the container object to make Hysys perform a Balance
    ' The 1 parameter means that the first entry in the array is a feed stream and the rest
    ' are products. Do a total balance so that if temps are specified then we'll do a heat
    Call myContainer.Balance(BalanceType_enum.btTotalBalance, 1, StreamList)
    ' Composition and Condition functions are for visualizing in the EDF
    ' TODO: Note that the next fluidList is not updated at this point without flashes
    ' TODO: Comp = Composition(fluidList) *********
    ' TODO: cond = Condition(fluidList) ************

    ' Step 8 - If we are here it means we solved the unit properly
    ' Check if the Product streams are completely solved
    'If edfPermeate.DuplicateFluid.IsUpToDate And edfRetentate.DuplicateFluid.IsUpToDate Then
    ' Need to remove the IF statement for erratic behavior, though is good practice
    Call myContainer.SolveComplete
    ' End If
    Exit Sub
ErrorTrap:
End Sub

Private Sub ExtensionObject_StatusQuery(ByVal Status As HYSYS.ObjectStatus)
    If myContainer.ExtensionInterface.IsIgnored = True Then Exit Sub
    
    If edfInlet Is Nothing Then
        Call Status.AddStatusCondition(slMissingRequiredInformation, 1, "Feed stream is missing")
    End If
    If edfPermeate Is Nothing Then
        Call Status.AddStatusCondition(slMissingRequiredInformation, 2, "Permeated stream is missing")
    End If
    If edfRetentate Is Nothing Then
        Call Status.AddStatusCondition(slMissingRequiredInformation, 3, "Non permeated outlet stream is missing")
    End If
    If edfLen < 0 Or edfDiam < 0 Or edfThick < 0 Or edfk < 0 Or edfNtubes < 0 Or edfAperm < 0 Or edfNpoints < 1 Then
        Call Status.AddStatusCondition(slMissingRequiredInformation, 4, "Physical parameters missing or incorrect")
    End If
    If IsForgetting Then Exit Sub
End Sub

Private Function ExtensionObject_OnHelp(HelpPanel As String) As Boolean
    ExtensionObject_OnHelp = True
End Function

Private Function ExtensionObject_OnView(ViewName As String) As Boolean
    ExtensionObject_OnView = True
End Function

Private Sub ExtensionObject_Save()
End Sub

Private Sub ExtensionObject_VariableChanged(ByVal Variable As HYSYS.InternalVariableWrapper)
' Called when the user modifies any edf variable. It is required to update these variales as needed

    On Error GoTo ErrorTrap
    Dim pressureVT As Object
    Dim flowVT As Object
    Dim i As Integer
    Dim Ravg As Double
    Select Case Variable.Tag
        '%% Attachment variables
        Case "Inlet"
            Set edfInlet = myContainer.FindVariable("Inlet").Variable.object
        Case "Permeate"
            Set edfPermeate = myContainer.FindVariable("Permeate").Variable.object
        Case "Retentate"
            Set edfRetentate = myContainer.FindVariable("Retentate").Variable.object
        Case "PermIn"
            Set edfPermIn = myContainer.FindVariable("PermIn").Variable.object
        Case "UpdatePlot"
            Call edfCompT.SetBounds(edfNpoints.Value)
            Call edfLengthPos.SetBounds(edfNpoints.Value)
            edfLengthPos.Values = GrafL
            edfCompT.Values = FpermCellsPlot
            'Call edfCompH.SetBounds(edfNpoints.Value + 1)
            'edfCompH.Values = FpermCellsPlot

        '%% Geometry variables
        Case "Length"
            Set edfLen = myContainer.FindVariable("Length").Variable
            If edfLen.Value > 0 Then
                If edfNpoints.Value > 0 Then
                    GrafL = LengthVector(edfLen.Value, edfNpoints.Value)
                End If
                If edfNtubes.Value > 0 Then
                    totalLen = edfLen.Value * edfNtubes.Value
                Else
                    totalLen = -32767
                End If
            End If
            Call CalcAperm
        Case "Diam"
            Set edfDiam = myContainer.FindVariable("Diam").Variable
            Call CalcAperm
        Case "Thickness"
            Set edfThick = myContainer.FindVariable("Thickness").Variable
            thick = edfThick.GetValue()
            Call CalcAperm
        Case "Aperm"
            Set edfAperm = myContainer.FindVariable("Aperm").Variable
            Aperm = edfAperm.Value
            Call CalcLength
        Case "Ntubes"
            Set edfNtubes = myContainer.FindVariable("Ntubes").Variable
            If edfLen.Value > 0 And edfNtubes.Value > 0 Then
                totalLen = edfLen.Value * edfNtubes.Value
            Else
                totalLen = -32767
            End If
            Call CalcAperm

        '%% Pressure drop and others
        Case "PressDrop"
            Set edfPressDrop = myContainer.FindVariable("PressDrop").Variable
        Case "PermPressDrop"
            Set edfPermPressDrop = myContainer.FindVariable("PermPressDrop").Variable
        Case "k"
            Set edfk = myContainer.FindVariable("k").Variable
        Case "NumberOfPoints"
            Set edfNpoints = myContainer.FindVariable("NumberOfPoints").Variable
            If edfLen.Value > 0 And edfNpoints.Value > 0 Then
                GrafL = LengthVector(edfLen.Value, edfNpoints.Value)
            End If

    End Select
    Exit Sub
ErrorTrap:
    Call myContainer.Trace(myContainer.Name & ": Error in Variable Changed for variable " & Variable.Tag & ".", False)
    MsgBox ("Variable Changed Error")
End Sub

Private Sub CalcAperm()
    ' Re-calculate the effective permeation area when some geometry parameter changes
    ' using the logarithmic formula (see Ackerman and Koskinas 1972)
    ' Check non-defined values
    Dim Rin As Double, Rext As Double, X As Double
    If edfDiam.Value <= 0 Or edfThick.Value <= 0 Or totalLen <= 0 Then
        Aperm = -32767
        Call edfAperm.SetValue(Aperm)
        Exit Sub
    End If
    ' Calc
    Rin = edfDiam.Value / 2
    Rext = Rin + edfThick.Value
    X = Rext - Rin
    Aperm = (2 * pi * totalLen * X) / Log(Rext / Rin)
    Call edfAperm.SetValue(Aperm)
End Sub

Private Sub CalcLength()
    ' Calc length based on permeation area
    ' Check non-defined values
    If edfDiam.Value <= 0 Or edfThick.Value <= 0 Or edfAperm.Value <= 0 Or edfNtubes.Value <= 0 Then
        totalLen = -32767
        Call edfLen.SetValue(-32767)
        Exit Sub
    End If
    ' Calc
    Dim Rext As Double, Rin As Double, X As Double
    Rin = edfDiam.Value / 2
    Rext = Rin + edfThick.Value
    X = Rext - Rin
    totalLen = edfAperm.Value * Log(Rext / Rin) / (2 * pi * X)
    Call edfLen.SetValue(totalLen / edfNtubes.Value)
End Sub

Private Function ExtensionObject_VariableChanging(ByVal Variable As HYSYS.InternalVariableWrapper) As Boolean
    Select Case Variable.Tag

        Case "NumberOfPoints"
            If Variable.NewRealValue < 1 Or Variable.NewRealValue > 1000 Then
                Call MsgBox("Entered Value is out of range, must be between 1 and 1000.")
                ExtensionObject_VariableChanging = False
                Exit Function
            End If

    End Select

    ExtensionObject_VariableChanging = True
End Function

Private Sub ExtensionObject_VariableQuery(ByVal Variable As HYSYS.InternalVariableWrapper)
End Sub

Private Sub ExtnUnitOperation_BasisChanged()
End Sub

Private Sub ExtensionObject_Terminate()
    Set edfInlet = Nothing
    Set edfPermeate = Nothing
    Set edfRetentate = Nothing
    Set edfPermIn = Nothing
    Call myContainer.DeletePlot("PQPlot")
    'Set myPlot = Nothing
    'Set myPlotName = Nothing

    Set edfThick = Nothing
    Set edfLen = Nothing
    Set edfDiam = Nothing
    Set edfAperm = Nothing
    Set edfNtubes = Nothing
    Set edfk = Nothing
    Set edfPH = Nothing
    Set edfPD = Nothing
    Set edfPT = Nothing
    Set edfTotalPermMolFlow = Nothing
    'Set edfComposition = Nothing
    'Set edfCompName = Nothing
    'Set edfStreamName = Nothing
    'Set edfVapFrac = Nothing
    'Set edfTemperature = Nothing
    'Set edfPressure = Nothing
    'Set edfMolFlow = Nothing
    'Set edfMassFlow = Nothing
    Set edfPressDrop = Nothing
    Set edfPermPressDrop = Nothing
    Set edfLengthPos = Nothing
    Set edfCompT = Nothing
    'Set edfCompH = Nothing
    Set edfNpoints = Nothing
    Set myPlotH = Nothing
    Set myPlotNameH = Nothing
    Set myPlotT = Nothing
    Set myPlotNameT = Nothing
    'edfDiamExt = Nothing
End Sub

Private Sub pointedEDFVariables()
    With myContainer
        Set edfInlet = .FindVariable("Inlet").Variable.object
        Set edfRetentate = .FindVariable("Retentate").Variable.object
        Set edfPermeate = .FindVariable("Permeate").Variable.object
        Set edfPermIn = .FindVariable("PermIn").Variable.object
        'Set myPlotName = .FindVariable("PlotName").Variable
        Set edfThick = .FindVariable("Thickness").Variable
        Set edfLen = .FindVariable("Length").Variable
        Set edfDiam = .FindVariable("Diam").Variable
        Set edfAperm = .FindVariable("Aperm").Variable
        Set edfNtubes = .FindVariable("Ntubes").Variable
        Set edfk = .FindVariable("k").Variable
        Set edfPH = .FindVariable("PH").Variable
        Set edfPD = .FindVariable("PD").Variable
        Set edfPT = .FindVariable("PT").Variable
        Set edfTotalPermMolFlow = .FindVariable("TotalPermMolFlow").Variable
        'Set edfComposition = .FindVariable("Composition").Variable
        'Set edfCompName = .FindVariable("CompoName").Variable
        'Set edfStreamName = .FindVariable("StreamNames").Variable
        'Set edfVapFrac = .FitndVariable("VapFrac").Variable
        'Set edfTemperature = .FindVariable("StreamTemp").Variable
        'Set edfPressure = .FindVariable("StreamPress").Variable
        'Set edfMolFlow = .FindVariable("StreamMolFlow").Variable
        'Set edfMassFlow = .FindVariable("StreamMassFlow").Variable
        Set edfPermPressDrop = .FindVariable("PermPressDrop").Variable
        Set edfPressDrop = .FindVariable("PressDrop").Variable
        Set edfLengthPos = .FindVariable("LengthPos").Variable
        Set edfCompT = .FindVariable("CompT").Variable
        'Set edfCompH = .FindVariable("CompH").Variable
        Set edfNpoints = .FindVariable("NumberOfPoints").Variable
        Set myPlotNameH = .FindVariable("PlotNameH").Variable
        Set myPlotNameT = .FindVariable("PlotNameT").Variable
        'Set edfDiamExt = .FindVariable("DiamExtern").Variable
        ''Set molDensity = .FindVariable("molDensity").Variable
    End With
End Sub




'***************************************************************************'
'                         Permeation Functions                              '
'***************************************************************************'
Private Function Permeation()
    ' Calculate vector of permeated species in default HYSYS magnitude: "kmol/s"
    '   FpermMolec:     Double(nComp)   OUTPUT of the function. Permeated flow (per component)
    '   Ffeed:          Double(nComp)   inlet molar flow (per component)
    '   CellFpermMolec: Double(nComp)   permeated molar flow in one cell (per component)
    '   CellFfeedMolec: Double(nComp)   Cell retentate flows, and, at the same time, next cell
    '                                   feed component molar flow
    '   FpermCellsPlot: Double(Npoints) collect total permeation in each cell for plotting
    '                                   purposes.
    '   CellFpermAtom:  Double(3)       Cell output of the diffusion part of code (H, D, T mole
    '                                   flows)
    '   PermMolFrac:    Double(3)       mol fraction contribution of a specific atom contributed by
    '                                   various molecules for the diffusion
    '   X:              Double(3)       Fraction variable to take into account contribution of
    '                                   each atom to the inlet permeation pressure
    '   Y:              Double(3)       Similar to X but to balance output pressure in permeation
    On Error GoTo errorHandler
    ' Step 1 - Declarations
    Dim Fperm() As Double, FpermMolec() As Double, Ffeed() As Double, Fcell() As Double
    Dim CellFpermAtom() As Double, X() As Double, PermMolFrac() As Double, Y() As Double
    Dim cellFfeedMolec() As Double, CellFpermMolec() As Double, CellFeedMolFractions() As Double
    Dim Qfeed As Double, TfeedK As Double, PfeedPa As Double, ApermCell As Double
    Dim PfeedPerm As Double, molFrac As Double, totalPermMolFrac As Double
    Dim i As Long, j As Long, iPerm As Long, nCell As Long, iCell As Integer
    Dim flagPermMismatch As Boolean, flagConv As Boolean, flagZeroSum As Integer
    Dim Pperm As Double, PermMolFracOut() As Double, permFlow As Double
    Dim FpermMolecOld() As Double, totalFpermMolecOld As Double
    Dim iter As Integer, maxIter As Integer
    ReDim Fperm(nComp - 1), Fcell(nComp - 1), FpermMolecOld(nComp - 1)
    ReDim PermMolFrac(nPermAtom - 1), X(nPermAtom - 1), Y(nPermAtom - 1)
    ReDim CellFpermMolec(nComp - 1)
    ReDim PermMolFracOut(nPermAtom - 1)

    ' Step 2 - Initializations
    ' First check to see if feed flow is unsuitable
    If IsPermeationZero() Then
        ' Nothing permeates (FpermMolec is an array of zeros)
        Permeation = FpermMolecOld
        Exit Function
    End If
    ' Get inlet stream parameters
    PfeedPa = edfInlet.Pressure.GetValue("N/m2")    ' Pascals to match permeability formula
    Pperm = edfPermeate.Pressure.GetValue("N/m2")
    TfeedK = edfInlet.Temperature.GetValue("K")     ' K degrees to match permeability formula
    Qfeed = edfInlet.ActualVolumeFlowValue
    Ffeed = edfInlet.ComponentMolarFlowValue        ' molar flow per component
    If Not AnyNeg(edfPermeate.ComponentMolarFlowValue) Then
        ' use previous solution as estimation
        FpermMolecOld = edfPermeate.ComponentMolarFlowValue
    End If
    ' Cell discretization initializations
    nCell = edfNpoints.GetValue()
    ApermCell = Aperm / nCell                       ' permeation surface per differential cell
    ReDim FpermCellsPlot(nCell - 1)
    ' Setup Permeabilities - calculate them and write in EDF (read-only)
    Dim P() As Double
    ReDim P(nPermAtom - 1)
    ' We calculate Input units as (kmol � m-1 � s -1 � Pa-0.5), therefore the division by 1000
    For i = 0 To nPermAtom - 1
        P(i) = A(i) * Math.Exp(E(i) / (R * TfeedK)) / 1000  ' A() comes in mol, not kmol
    Next
    ' Publish permeabilities in mol � ...
    Call edfPH.SetValue(P(0) * 1000)
    Call edfPD.SetValue(P(1) * 1000)
    Call edfPT.SetValue(P(2) * 1000)

    ' ITER LOOP
    If nCell = 1 Then
        maxIter = 1  ' Avoid iterating with 1 cell due to non-convergence problems found
    Else
        maxIter = 15
    End If
    iter = 0
    Dim weight As Double, delta_weight As Double
    weight = 0.5
    delta_weight = 0.5
    Do While (Not flagConv) And (iter < maxIter)
        ' CELL LOOP
        CellFeedMolFractions = edfInlet.ComponentMolarFractionValue
        cellFfeedMolec = edfInlet.ComponentMolarFlowValue
        ReDim FpermMolec(nComp - 1)  ' to re-init components to zero
        For iCell = 0 To nCell - 1
            ' Step 3 - Calculate permeating pressure
            ' LOOPS: use "i" for atoms (H, D, T) and "j" for molecules (H2, HD, HT, D2, DT, T2)
            ' Get total inlet pressure of permeating species ONLY. Apply Dalton's law
            PfeedPerm = 0  ' p1 in thesis
            totalPermMolFrac = 0
            For j = 0 To nPerm - 1
                iPerm = permIndices(j)
                totalPermMolFrac = totalPermMolFrac + CellFeedMolFractions(iPerm)
            Next
            PfeedPerm = PfeedPa * totalPermMolFrac

            'If PfeedPerm = 0 Then
            ' TODO: I commented this IF because it broke the code when discretizing, but it
            '       would be good to prepare another piece of code to avoid extra cell
            '       iterations when there is no permeating species left
            '' No permeating species in, exit permeation
            'Permeation = FpermMolec
            '    Exit Function
            'End If

            ' Step 4 - Calculate atomic diffusion flows
            ReDim CellFpermAtom(nPermAtom - 1)
            totalFpermMolecOld = Sum(FpermMolecOld)
            For i = 0 To nPermAtom - 1
                ' Loop through nPerm (should be 6) to get contribution to partial pressure of atom "i"
                PermMolFrac(i) = 0
                PermMolFracOut(i) = 0
                For j = 0 To nPerm - 1
                    ' In this loop we use feed mol fracs to calculate X and mol flows to get Y
                    molFrac = CellFeedMolFractions(permIndices(j))
                    permFlow = FpermMolecOld(permIndices(j))
                    ' permCoeffs is 1 or 0.5 depending on molecule being homonuclear or heteronuclear
                    PermMolFrac(i) = PermMolFrac(i) + permCoeffs(i)(j) * molFrac
                    PermMolFracOut(i) = PermMolFracOut(i) + permCoeffs(i)(j) * permFlow
                Next
                If totalPermMolFrac <= 0 Then
                    X(i) = 0                    ' Avoid division by zero
                Else
                    X(i) = PermMolFrac(i) / totalPermMolFrac
                End If
                If totalFpermMolecOld <= 0 Then ' Avoid division by zero
                    Y(i) = 0
                Else
                    Y(i) = PermMolFracOut(i) / totalFpermMolecOld
                End If
                ' PERMEATION FORMULA: F = P(i) * A / t * (X(i) * sqrt(p_in) - Y(i) * sqrt(p_out))
                CellFpermAtom(i) = P(i) * ApermCell / thick * (X(i) * Sqr(PfeedPerm) - Y(i) * Math.Sqr(Pperm))
                If CellFpermAtom(i) < 0 Then
                    CellFpermAtom(i) = 0
                End If
            Next

            ' Step 5 - Distribute molecular flow depending on the calculated atomic diffusion flow
            ' Check if we have all permeating species available in the inlet in `isFeedComplete`
            ReDim CellFpermMolec(nComp - 1)
            Dim isFeedComplete As Boolean
            isFeedComplete = True
            For i = 0 To nPerm - 1
                iPerm = permIndices(i)
                If cellFfeedMolec(iPerm) = 0 Then isFeedComplete = False
            Next
            ' Select heuristic function to assign molecular distribution
            If isFeedComplete Then
                ' OPTION 1: We have all molecules in the feed and there is enough of all of them
                Call PermeateHeuristic1(CellFpermMolec, cellFfeedMolec, CellFpermAtom)
            Else
                ' OPTION 2: Some molecules in the feed are missing
                Call PermeateHeuristic2(CellFpermMolec, cellFfeedMolec, CellFpermAtom)
            End If
    
            ' Step 6 - Last checks
            ' Check that the total permeated flow is the same from both Atom and Molecular sides
            ' Use the rounded relative error to measure the closeness of both sums
            'If Not flagPermMismatch Then
            '    Dim SumMolec As Double, SumAtom As Double
            '    SumAtom = Sum(CellFpermAtom)
            '    If SumAtom <> 0 Then
            '        SumMolec = Sum(CellFpermMolec)
            '        Dim RelError As Double
            '        RelError = Math.Abs((SumAtom - SumMolec) / SumAtom)
            '        If Math.Round(RelError, 10) > 0 Then  ' 10 is an arbitrary number of decimals
            '            flagPermMismatch = True
            '        End If
            '    End If
            'End If
    
            ' Step 7 - Calculate cell retentate and accumulate permeation flow
            FpermMolec = SumArrays(FpermMolec, CellFpermMolec)
            FpermCellsPlot(iCell) = Sum(FpermMolec)
            ' Current cell Retentate will be the next cell Feed
            cellFfeedMolec = Subtract(cellFfeedMolec, CellFpermMolec)
            ' Calculate molar fractions off this new retentate
            CellFeedMolFractions = CalcMolFractions(cellFfeedMolec)
        Next
    
        ' Step 8 - Convergence test
        Dim SumNew As Double, SumOld As Double
        SumNew = Sum(FpermMolec)
        SumOld = Sum(FpermMolecOld)
        ''WATCH ONLY*********************
        'Dim FpermMolecWatch As Variant, FpermMolecOldWatch As Variant, InletWatch As Variant
        'FpermMolecWatch = ScalarMult(FpermMolec, 3600000)
        'FpermMolecOldWatch = ScalarMult(FpermMolecOld, 3600000)
        'InletWatch = ScalarMult(edfInlet.ComponentMolarFlowValue, 3600000)
        ''''*********************
        Dim RelErrorConv As Double
        If SumNew = 0 Then
            ' Nothing permeates though there are permeation species in the feed
            ' When this happens we ignore the issue and iterations will exhaust with some small
            ' permeation that will flow out or none at all
            flagZeroSum = flagZeroSum + 1
        Else
            RelErrorConv = Math.Abs((SumNew - SumOld) / SumNew)
            If RelErrorConv < 0.001 Then  ' 0.1 % as Ackerman 1972
                flagConv = True
            End If
        End If
        ' Make a copy of the output perm flow and update weighting
        Call WeightedSum(FpermMolecOld, FpermMolec, weight)
        If (iter >= maxIter / 2) And (weight < 0.95) Then
            delta_weight = delta_weight / 2
            weight = weight + delta_weight
        End If
        'If flagZeroSum >= 3 Then
        '    ' The FpermMolec is already zero: exit loop
        '    flagConv = True
        'End If
        ' Update iterations
        iter = iter + 1
    Loop

    ' Step 9 - Catch Step 6 permeation mismatch flag
    ' FIX - Remove traces to avoid to frequent and annoying warnings
    ' It does not mean there's a malfunction in the extension. It can be due to permeation
    ' exhaustion or a warning pointing that the heuristics are working at their limits
    'If flagPermMismatch Then
    '    Call myContainer.Trace(myContainer.Name + ": Permeation function did not match required" _
    '        + " atomic permeation (CellFpermAtom) with the actual output (CellFpermMolec)" _
    '        + " in some cells", False)
    'End If
    'If (iter >= maxIter) And (nCell > 1) Then
    '    Call myContainer.Trace(myContainer.Name + ": Permeation function exhausted max " _
    '                           + "iterations without convergence", False)
    'End If
    Permeation = FpermMolec
    Exit Function
errorHandler:
    Dim sMsg As String
    sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    Call MsgBox(myContainer.Name + " - Permeation: " + sMsg)
End Function

Private Function ScalarMult(vector, scalar) As Double()
    Dim i As Integer, out() As Double
    ReDim out(UBound(vector))
    For i = 0 To UBound(vector)
        out(i) = vector(i) * scalar
    Next
    ScalarMult = out
End Function

Private Function CalcMolFractions(cellFfeedMolec() As Double) As Double()
    Dim molfracs() As Double, totalflow As Double, i As Integer
    ReDim molfracs(nComp - 1)
    totalflow = Sum(cellFfeedMolec)
    ' Consider case total flow is zero
    If totalflow = 0 Then
        molfracs(0) = 1
        CalcMolFractions = molfracs
        Exit Function
    End If
    ' Normal case: calculate molar fractions
    For i = 0 To nComp - 1
        molfracs(i) = cellFfeedMolec(i) / totalflow
    Next
    CalcMolFractions = molfracs
End Function

Private Sub PermeateHeuristic1(FpermMolec() As Double, Ffeed() As Double, FpermAtom() As Double)
    ' Assign molecular molar flow (FpermMolec) based on previosly calculated diffusion atomic
    ' flow (FpermAtom)
    ' OPTION 1: Case in which we have H, D and T flow, now calculate H2, HD, HT, D2, DT, T2 flows
    Dim i As Long, j As Long, iPerm As Long, auxFpermMolec As Double
    For j = 0 To nPerm - 1
        iPerm = permIndices(j)
        ' Calculate contribution of each species
        auxFpermMolec = 0
        For i = 0 To nPermAtom - 1
            ' Dividing by 2 means e.g. 1 (H2) + 0.5 (HD) + 0.5 (HT) = 2
            auxFpermMolec = auxFpermMolec + FpermAtom(i) * (permCoeffs(i)(j) / 2)
        Next
        ' Check if we exhaust the feed
        If Ffeed(iPerm) > auxFpermMolec Then
            ' There is enough feed for the required permeation
            FpermMolec(iPerm) = auxFpermMolec
        Else
            ' Exhaust feed molecule
            FpermMolec(iPerm) = Ffeed(iPerm)
        End If
    Next
End Sub

Private Sub PermeateHeuristic2(FpermMolec() As Double, Ffeed() As Double, FpermAtom() As Double)
    ' Assign molecular molar flow (FpermMolec) based on previosly calculated diffusion atomic
    ' flow (FpermAtom)
    '
    ' OPTION 2: not all 6 species are available in the inlet
    '
    '   flagLoop:           Boolean     Used to assert if H, D or T permeation finished and break loop
    '   permIndicesSorted:  Integer(3)    Sorts indices to address FpermAtom in ascending order
    '   FpermAtomFlag:      Double(3)
    '   permIndicesHeteroCopy: Integer(3,3)

    Dim i As Long
    ' Loop in ascending order of permeation flow rate per atom
    Dim flagLoop As Boolean
    Dim permIndicesSorted() As Variant, iAtom As Integer, k As Integer, iMolec As Integer
    ' We will subtract the flows we go assigning from this array
    Dim FpermAtomFlag() As Double, permIndicesHeteroCopy() As Variant
    ReDim FpermAtomFlag(nPermAtom - 1)
    FpermAtomFlag = FpermAtom
    permIndicesHeteroCopy = permIndicesHetero

    ' Sort flows in ascending order and return the indices of the ordered list
    permIndicesSorted = SortIndices(FpermAtom)

    ' Heteronuclears loop first
    For i = 0 To nPermAtom - 1
        iAtom = permIndicesSorted(i)
        flagLoop = False
        For k = 0 To nHeteroNuclear - 1
            ' With heteronuclears we need double the flow rate bc they contribute by half
            iMolec = permIndicesHeteroCopy(iAtom)(k)
            ' Continue to next iteration if element is -1
            If iMolec >= 0 Then
                ' First check if there is feed flow rate at all for "iMolec"
                If Ffeed(iMolec) > 0 Then
                    If Ffeed(iMolec) / 2 <= FpermAtomFlag(iAtom) Then
                        ' We exhaust all inlet flow of this molecule for permeation
                        FpermMolec(iMolec) = FpermMolec(iMolec) + Ffeed(iMolec)
                    Else
                        ' Permeation of atom "i" is finished and there will be leftovers of this molecule in retentate
                        FpermMolec(iMolec) = FpermMolec(iMolec) + 2 * FpermAtomFlag(iAtom)
                        flagLoop = True
                    End If
                    ' Subtract flow from our flag permeation array for both atoms affected
                    ' without the x2!
                    FpermAtomFlag(iAtom) = FpermAtomFlag(iAtom) - FpermMolec(iMolec) / 2   ' current atom
                    FpermAtomFlag(k) = FpermAtomFlag(k) - FpermMolec(iMolec) / 2       ' the other atom affected ' TODO: this could become negative!
                    '                                                  (but it shouldn't because we go in ascending order)
                    ' Remove molecule from matrix to avoid passing through it twice
                    ' (since permIndicesHetero is upper diagonal)
                    permIndicesHeteroCopy(iAtom)(k) = -1
                    permIndicesHeteroCopy(k)(iAtom) = -1
                    ' Exit For loop in case we finished 'iAtom' permeation
                    If flagLoop Then
                        Exit For
                    End If
                End If
            End If
        Next
    Next

    ' Homonuclears loop last (simpler loop)
    For i = 0 To nPermAtom - 1
        iAtom = permIndicesSorted(i)
        iMolec = permIndicesHomo(iAtom)
        ' First check if there is feed flow rate at all for "iMolec"
        If Ffeed(iMolec) > 0 Then
            If FpermAtomFlag(iAtom) >= Ffeed(iMolec) Then
                ' We exhaust all inlet flow of this molecule in permeation
                FpermMolec(iMolec) = FpermMolec(iMolec) + Ffeed(iMolec)
            Else
                ' Permeation of atom "i" is finished
                FpermMolec(iMolec) = FpermMolec(iMolec) + FpermAtomFlag(iAtom)
            End If
            ' Subtract flow from our flag permeation array
            FpermAtomFlag(iAtom) = FpermAtomFlag(iAtom) - FpermMolec(iMolec)
        End If
    Next
End Sub


'***************************************************************************'
'                      HYSYS Flowsheet Functions                            '
'***************************************************************************'
Private Sub IniCompIndex()
    ' Loop for setting index for components that can permeate (based on Inlet's basis manager)
    Dim ComponentList As Components
    Dim isH As Integer, isD As Integer, isT As Integer, i As Long
    ' Init the array of the indices of permeating species to -1
    permIndices = Array(-1, -1, -1, -1, -1, -1)
    ' Init component list
    If edfInlet Is Nothing Then
        Set ComponentList = myContainer.Flowsheet.FluidPackage.Components
    Else
        Set ComponentList = edfInlet.DuplicateFluid.Components
    End If
    nComp = ComponentList.Count         ' Number of Components
    ' Loop over component list and search for hydrogen isotopes
    For i = 0 To nComp - 1
        Select Case ComponentList.Item(i).Name
            Case "Hydrogen"
                permIndices(0) = i
                isH = 1
            Case "Hydrogen*"
                permIndices(0) = i
                isH = 1
            Case "HD*"
                permIndices(1) = i
                isH = 1
                isD = 1
            Case "HT*"
                permIndices(2) = i
                isH = 1
                isT = 1
            Case "Deuterium*"
                permIndices(3) = i
                isD = 1
            Case "DT*"
                permIndices(4) = i
                isD = 1
                isT = 1
            Case "Tritium*"
                permIndices(5) = i
                isT = 1
        End Select
    Next i
    permIndicesHetero = Array( _
        Array(-1, permIndices(1), permIndices(2)), _
        Array(permIndices(1), -1, permIndices(4)), _
        Array(permIndices(2), permIndices(4), -1) _
    )
    permIndicesHomo = Array(permIndices(0), permIndices(3), permIndices(5))
End Sub
Private Function CheckExtnRequirements() As Boolean
    ' Check if everything is set for the extension to run
    CheckExtnRequirements = True
    ' Stream check
    If edfInlet Is Nothing Or edfPermeate Is Nothing Or edfRetentate Is Nothing Then
        CheckExtnRequirements = False
        Exit Function
    End If
    ' Geometric check
    If Not edfLen.IsKnown Or Not edfDiam.IsKnown Or Not edfThick.IsKnown Or Not edfk.IsKnown _
          Or Not edfAperm.IsKnown Or Not edfNtubes.IsKnown Or Not edfNpoints.IsKnown Then
        CheckExtnRequirements = False
        Exit Function
    End If
End Function
Private Sub BuildStreamsAndFluids()
    Dim i As Integer
    ReDim Preserve StreamList(0 To 2)
    Set StreamList(0) = edfInlet                ' In
    Set StreamList(1) = edfRetentate            ' Out non Permeated
    Set StreamList(2) = edfPermeate             ' Out Permeated
    If Not edfPermIn Is Nothing Then         ' (optional) 2nd stream in
        ReDim Preserve StreamList(0 To 3)
        Set StreamList(3) = edfPermIn
    End If
    ReDim Preserve fluidList(0 To UBound(StreamList)) ' do fluid list
    For i = 0 To UBound(StreamList)
        Set fluidList(i) = StreamList(i).DuplicateFluid
    Next i
End Sub
Private Sub SetPermeationFlows(streams() As ProcessStream, permMolarFlows() As Double)
    ' Update vector of streams (vector of 3 or 4 Fluids) and,
    ' specifically, the products: Retentate (index 1) and Permeate
    ' (index 2), with the new permeated composition.
    '
    ' To do this, update component molar flows AND total molar flow.
    '
    Dim inletMolarFlows() As Double, retMolarFlows() As Double
    Dim flagRet As Boolean, flagPerm As Boolean
    ' Calculate retentate component flows
    inletMolarFlows = streams(0).ComponentMolarFlowValue
    retMolarFlows = Subtract(inletMolarFlows, permMolarFlows)
    ' Set the fictitious molar flow vectors to the actual ones
    flagRet = SetValueComponentMolarFlow(streams(1), retMolarFlows)     ' Retentate
    flagPerm = SetValueComponentMolarFlow(streams(2), permMolarFlows)   ' Permeate
    ' Set total molar flow to each fluid
    Call streams(1).MolarFlow.Calculate(Sum(retMolarFlows), MOLFLOW_UNITS)       ' Retentate
    Call streams(2).MolarFlow.Calculate(Sum(permMolarFlows), MOLFLOW_UNITS)      ' Permeate
    ' Set traces if empty streams
    If flagRet Then
        Call myContainer.Trace(myContainer.Name + ": WARNING - " + _
                               " returns an empty retentate stream", False)
    End If
    If flagPerm Then
        Call myContainer.Trace(myContainer.Name + ": WARNING - " + _
                                      " returns an empty permeate stream", False)
    End If
End Sub
Private Sub SetPermeationFluids(fluids() As fluid, permMolarFlows() As Double)
    ' Update vector of fluids (vector of 3 or 4 Fluids) and,
    ' specifically the products: Retentate (index 1) and Permeate
    ' (index 2), with the new permeated composition.
    '
    ' To do this, update component molar flows AND total molar flow.
    '
    Dim inletMolarFlows() As Double, retMolarFlows() As Double
    Dim flagRet As Boolean, flagPerm As Boolean
    ' Calculate retentate component flows
    inletMolarFlows = fluids(0).MolarFlowsValue
    retMolarFlows = SubtractVectorsIf(inletMolarFlows, permMolarFlows)
    ' Set the fictitious molar flow vectors to the actual ones
    flagRet = SetValueComponentMolarFlowFluid(fluids(1), retMolarFlows)     ' Retentate
    flagPerm = SetValueComponentMolarFlowFluid(fluids(2), permMolarFlows)   ' Permeate

    'fluids(1).MolarFlows.SetValues(retMolarFlows, MOLFLOW_UNITS)       ' Retentate
    'fluids(2).MolarFlows.SetValues(permMolarFlows, MOLFLOW_UNITS)      ' Permeate
    ' Set total molar flow to each fluid
    Call fluids(1).MolarFlow.SetValue(Sum(retMolarFlows), MOLFLOW_UNITS)       ' Retentate
    Call fluids(2).MolarFlow.SetValue(Sum(permMolarFlows), MOLFLOW_UNITS)      ' Permeate
    ' Don't do traces since this will work in dynamics and it would be too noisy
End Sub

Private Function SetValueComponentMolarFlow(stream As ProcessStream, componentMolarFLow() As Double) As Boolean
    ' Use ComponentMolarFlow.Calculate if the provided array is nonzero, otherwise just set
    ' some molar fraction with ComponentMolarFraction.Calculate so that HYSYS can calculate the
    ' stream
    '
    ' Return flag = True if zero flow
    '
    If Sum(componentMolarFLow) = 0 Then
        SetValueComponentMolarFlow = True
        ' Stream is empty
        Dim auxMolarFractions() As Double
        ReDim auxMolarFractions(nComp - 1)
        ' Arbitrarily assign 1 to the first component
        auxMolarFractions(0) = 1
        Call stream.ComponentMolarFraction.Calculate(auxMolarFractions)
    Else
        SetValueComponentMolarFlow = False
        ' Assign flows normally
        Call stream.componentMolarFLow.Calculate(componentMolarFLow, MOLFLOW_UNITS)
    End If
End Function
Private Function SetValueComponentMolarFlowFluid(fluid As fluid, componentMolarFLow() As Double) As Boolean
    ' Use ComponentMolarFlow.Calculate if the provided array is nonzero, otherwise just set
    ' some molar fraction with ComponentMolarFraction.Calculate so that HYSYS can calculate the
    ' stream
    '
    ' Return flag = True if zero flow
    '
    If Sum(componentMolarFLow) = 0 Then
        SetValueComponentMolarFlowFluid = True
        ' Stream is empty
        Dim auxMolarFractions() As Double
        ReDim auxMolarFractions(nComp - 1)
        ' Arbitrarily assign 1 to the first component
        auxMolarFractions(0) = 1
        Call fluid.MolarFractions.SetValues(auxMolarFractions)
    Else
        SetValueComponentMolarFlowFluid = False
        ' Assign flows normally
        Call fluid.MolarFlows.SetValues(componentMolarFLow, MOLFLOW_UNITS)
    End If
End Function

Private Function IsPermeationZero() As Boolean
    ' Check if there are no permeating species at all in the feed
    Dim j As Integer, iPerm As Integer
    IsPermeationZero = True
    For j = 0 To nPerm - 1
        iPerm = permIndices(j)
        If edfInlet.ComponentMolarFlowValue(iPerm) > 0 Then
            IsPermeationZero = False
            Exit Function
        End If
    Next
End Function



'***************************************************************************'
'                         Auxiliary Functions                               '
'***************************************************************************'

Private Function LengthVector(L, n) As Double()
    ' Builds a vector of equidistant elements representing the position [m] of each cell
    Dim i As Long
    Dim lenvec() As Double, dx As Double
    ReDim lenvec(n - 1)
    dx = L / n
    lenvec(0) = dx
    For i = 1 To n - 1
        lenvec(i) = lenvec(i - 1) + dx
    Next i
    LengthVector = lenvec
End Function
Private Function Sum(myArray() As Double) As Double
    Dim i As Long
    For i = 0 To UBound(myArray)
        Sum = Sum + myArray(i)
    Next
End Function
Private Sub WeightedSum(ByRef A1() As Double, A2() As Double, w As Double)
    ' Do a weighted sum for vectors in an element-wise way with weight 'w'.
    ' A high value of w means a slower and steadier movement, while a lower one means
    '  a fastest and riskier approach to the solution.
    ' Do these operations in place for the A1 ByRef array!
    Dim i As Integer
    For i = 0 To UBound(A1, 1)
        A1(i) = w * A1(i) + (1 - w) * A2(i)
    Next i
End Sub
Private Function SumArrays(A1, A2) As Double()
    ' Sum 1D arrays
    Dim n As Long, i As Long
    Dim A() As Double
    ' Get sizes
    n = UBound(A1, 1) + 1
    ReDim A(n - 1)
    ' Do calculation
    For i = 0 To n - 1
        A(i) = A1(i) + A2(i)
    Next i
    SumArrays = A
End Function
Private Function Subtract(ByRef A1, ByRef A2) As Variant
    ' Subtract 1D arrays
    Dim n As Long, i As Long
    Dim A() As Double
    ' Get sizes
    n = UBound(A1, 1) + 1
    ReDim A(n - 1)
    ' Do calculation
    For i = 0 To n - 1
        A(i) = A1(i) - A2(i)
    Next i
    Subtract = A
End Function
Private Function SubtractVectorsIf(ByRef A1, ByRef A2) As Double()
    ' Subtract 1D arrays avoiding negative outputs.
    '
    ' e.g.  we want to calculate the operation Fret = Finlet - Fperm
    '       in a per-component basis.
    '       Thus, A1 = Finlet, A2 = Fperm
    '       If Finlet = (4) and Fperm = (10); Fret = (-6)
    '       But avoiding negative numbers, the result we should get
    '       is Fperm = (4) and Fret = (0)
    '
    ' To do this, we use ByRef arguments and modify A2 (here Fperm)
    ' if necessary.
    Dim n As Long, i As Long
    Dim A() As Double
    ' Get sizes
    n = UBound(A1, 1) + 1
    ReDim A(n - 1)
    ' Check condition that allow sum/subtract
    If (UBound(A2, 1) + 1) <> n Then
        MsgBox ("Sizes do not match in array sum function.")
        Exit Function
    End If
    ' Do calculation
    For i = 0 To n - 1
        If A1(i) >= A2(i) Then
            A(i) = A1(i) - A2(i)
        Else
            ' Molecule i permeated till its limits
            A(i) = 0  ' no flow in retentate
            A2(i) = A1(i)  ' adjust permeate flow to the exact input
        End If
    Next i
    SubtractVectorsIf = A
End Function
Private Function IsZero(myArray() As Double) As Boolean
    ' Check if the input 1D array is full of zeros
    Dim i As Long
    IsZero = True
    For i = 0 To UBound(myArray)
        If myArray(i) <> 0 Then
            IsZero = False
        End If
    Next
End Function
Private Function AnyNeg(A As Variant) As Boolean
    ' Check if any element of the input array is negative
    Dim i As Integer
    AnyNeg = False
    i = 0
    Do While (i <= UBound(A)) And (Not AnyNeg)
        If A(i) < 0 Then
            AnyNeg = True
        End If
        i = i + 1
    Loop
End Function
Private Function SortIndices(SrcArray() As Double) As Variant

    'Description: Sorts the arrays passed so that smallest values occur first in KeyArray()
    '             does the same rearrangements on OtherArray() so values still correspond
    '             - Uses a Ripple type sort (Good for smallish data sets)
    '
    On Error GoTo ErrorTrap
    'Declare Variables------------------------------------------------------------------------------

    Dim i As Integer
    Dim j As Integer 'Counters
    Dim Temp As Integer 'used to swap values
    Dim TempIndices() As Variant
    TempIndices = Array(0, 1, 2)
    
    'Procedure--------------------------------------------------------------------------------------
    For i = 0 To UBound(SrcArray) - 1
        For j = i + 1 To UBound(SrcArray)
            If SrcArray(j) < SrcArray(i) Then
                ' Swap data
                'Temp = SortedArray(J)
                'SortedArray(J) = SortedArray(I)
                'SortedArray(I) = Temp
                ' Swap indices
                Temp = TempIndices(j)
                TempIndices(j) = TempIndices(i)
                TempIndices(i) = Temp
            End If
        Next  'J
    Next  'I
    SortIndices = TempIndices
    Exit Function

ErrorTrap:
    MsgBox ("Sorting Error")
End Function



'***************************************************************************'
'                          Plotting Functions                               '
'***************************************************************************'
' Private Function Composition(myFluid() As fluid) As Boolean
' Dim iFluid As Long, i As Long
' Dim streamName() As String
' Dim X() As Double
'     ' Set bounds of final EDF vatiables
'     edfCompName.SetBounds nComp
'     edfComposition.SetBounds nFluid, nComp
'     edfStreamName.SetBounds nFluid
'     ' Construct auxiliary vectors/matrix for tables
'     ReDim Preserve X(nComp - 1, 0 To nFluid - 1)
'     ReDim streamName(0 To nFluid - 1)
'     For iFluid = 0 To nFluid - 1
'         streamName(iFluid) = myFluid(iFluid).Name
'         For i = 0 To nComp - 1
'             X(i, iFluid) = myFluid(iFluid).MolarFractions(i)
'         Next i
'     Next iFluid
'     ' Associate matrix to final EDF variables for visualizing
'     edfStreamName = streamName                          ' Columns labels (streams)
'     edfComposition.Values = X                           ' Values in table
'     edfCompName.Values = myFluid(0).Components.Names    ' Rows labels (components)
'     Composition = True
' End Function
' Private Function Condition(myFluid() As fluid) As Boolean
' Dim i As Long
' Dim vapfrac() As Double, Temp() As Double, Press() As Double, MolFlow() As Double, massFlow() As Double
'     ' Set bounds of final EDF vatiables
'     edfVapFrac.SetBounds nFluid
'     edfTemperature.SetBounds nFluid
'     edfPressure.SetBounds nFluid
'     edfMolFlow.SetBounds nFluid
'     edfMassFlow.SetBounds nFluid
'     ' Construct auxiliary vectors/matrix for tables
'     ReDim vapfrac(0 To nFluid - 1)
'     ReDim Temp(0 To nFluid - 1)
'     ReDim Press(0 To nFluid - 1)
'     ReDim MolFlow(0 To nFluid - 1)
'     ReDim massFlow(0 To nFluid - 1)
'     For i = 0 To nFluid - 1
'          vapfrac(i) = myFluid(i).VapourFractionValue
'          Temp(i) = myFluid(i).TemperatureValue
'          Press(i) = myFluid(i).PressureValue
'          massFlow(i) = myFluid(i).MassFlowValue
'          MolFlow(i) = myFluid(i).MolarFlowValue
'      Next i
'      ' Associate matrix to final EDF variables for visualizing
'      edfVapFrac = vapfrac
'      edfTemperature = Temp
'      edfPressure = Press
'      edfMolFlow = MolFlow
'      edfMassFlow = massFlow
'      Condition = True
' End Function
'
Sub CreatePlot()
'     If Not myPlotH Is Nothing Then
'         myContainer.DeletePlot "LCPlotH"
'         Set myPlotH = Nothing
'     End If
'     myContainer.BuildPlot2 "LCPlotH", myPlotH, hptTwoDimensionalPlot
'     myPlotNameH.Value = "LCPlotH"
'     With myPlotH
'         .TitleData = "Length vs Non Permeation"
'
'         .SetAxisLabelData hatXAxis, "Length (m)"
'         .SetAxisLabelData hatYAxis, "NoPerm Flow"
'         .SetAxisLabelVisible hatXAxis, True
'         .SetAxisLabelVisible hatYAxis, True
'
'         .LegendVisible = True
'
'         .CreateXYDataSet 1, "PQData"
'         .SetDataSetXData 1, edfLengthPos
'         .SetDataSetYData 1, edfCompH
'         .SetDataSetColour 1, "Red"
'     End With
    If Not myPlotT Is Nothing Then
        Call myContainer.DeletePlot("LCPlotT")
        Set myPlotT = Nothing
    End If
    Call myContainer.BuildPlot2("LCPlotT", myPlotT, hptTwoDimensionalPlot)
    myPlotNameT.Value = "LCPlotT"
    With myPlotT
        .TitleData = "Length vs Permeation"
        .SetAxisLabelData hatXAxis, "Length (m)"
        .SetAxisLabelData hatYAxis, "Perm Flow"
        .SetAxisLabelVisible hatXAxis, True
        .SetAxisLabelVisible hatYAxis, True
        .LegendVisible = True
        .CreateXYDataSet 1, "PQData"
        .SetDataSetXData 1, edfLengthPos
        .SetDataSetYData 1, edfCompT
        .SetDataSetColour 1, "Red"
    End With
End Sub







'***************************************************************************'
'                           Dynamic Methods                                 '
'***************************************************************************'
Public Sub DynInitialize(dContainer As ExtnDynUnitOpContainer, IsRecalling As Boolean, MyVersion As Long, HoldupExist As Boolean)
    ' Called to Initialize the extension.
    Set dynContainer = dContainer
    HoldupExist = False
    MyVersion = CurrentExtensionVersion_enum.extnCurrentVersion
End Sub


Public Function InitializeSystem(ForceInit As Boolean) As Boolean
    ' Called each time the integration starts
    InitializeSystem = False
    nComp = UBound(edfInlet.ComponentMolarFlowValue) + 1
    ' Set pressure modifiable status
    '  Would prefer it to be "calculated" but throws an error when using .Calculated
    Call edfPermPressDrop.SetModifyState(VariableStatus_enum.vsSpecifiedOutside)
    Call edfPressDrop.SetModifyState(VariableStatus_enum.vsCalculated)
    ' Check existing streams
    If (edfInlet Is Nothing Or edfPermeate Is Nothing Or edfRetentate Is Nothing) Then
        Exit Function
    End If
    ' Set up StreamList() and FluidList - will directly point to ProcessStreams
    Call BuildStreamsAndFluids
    InitializeSystem = True
End Function





Public Function NumberOfFlowEquations() As Long
' Called to get the number of Pressure Flow equations the extension contributes separately.
    NumberOfFlowEquations = 1
End Function

Public Function NumberOfPressBalEquations() As Long
' Called to get the number of Pressure Balance equations the extension contributes separately.
    NumberOfPressBalEquations = 2
End Function

Public Function NumberOfFlowBalEquations() As Long
' Called to get the number of Flow Balance equations the extension contributes separately.
    NumberOfFlowBalEquations = 2
End Function

Public Function NumberOfGeneralEquations() As Long
' Called to get the number of General equations the extension contributes separately.
    NumberOfGeneralEquations = 0
End Function





Public Function VariablesInPressBalanceEquations()
' Called to get the number of variables in different Pressure Balance equations extension contributes separately.
    Dim NVars(1) As Long
    
    NVars(0) = 3
    NVars(1) = 3
  
    VariablesInPressBalanceEquations = NVars
End Function

Public Function PressureBalanceEquationVars(EquationIdx As Long)
' Called to get the variables in each Pressure Balance equation.
    Dim P() As RealVariable
        
    If edfPermIn Is Nothing Then
        Select Case EquationIdx
            Case 0 'No Perm Side
                ReDim P(2)
                Set P(0) = edfInlet.Pressure
                Set P(1) = edfRetentate.Pressure
                Set P(2) = edfPressDrop
            Case 1 'Perm Side
                ReDim P(2)
                Set P(0) = edfInlet.Pressure
                Set P(1) = edfPermeate.Pressure
                Set P(2) = edfPermPressDrop
        End Select
    Else
        Select Case EquationIdx
            Case 0 'No Perm Side
                ReDim P(1)
                Set P(0) = edfInlet.Pressure
                Set P(1) = edfRetentate.Pressure
            Case 1 'Perm Side
                ReDim P(1)
                Set P(0) = edfPermIn.Pressure
                Set P(1) = edfPermeate.Pressure
        End Select
    End If
    
    PressureBalanceEquationVars = P
End Function

Public Function PressureBalanceEquationCoef(EquationIdx As Long)
' Called to get the coefficients of different variables in each Pressure Balance equation.
    Dim Coef() As Double
    
    Select Case EquationIdx
        Case 0 'Perm Side
            ReDim Coef(3)
            Coef(0) = 1 'Inlet
            Coef(1) = -1 'No Perm
            Coef(2) = -1 'dp
            Coef(3) = 0
        Case 1 'No Perm Side
            ReDim Coef(3)
            Coef(0) = 1 'Inlet
            Coef(1) = -1 'Perm
            Coef(2) = -1 'dp
            Coef(3) = 0
    End Select

    PressureBalanceEquationCoef = Coef
End Function





Public Function VariablesInFlowBalanceEquations()
' Called to get the number of variables in different Flow Balance equations extension contributes separately.
    Dim NVars(1) As Long

    If edfPermIn Is Nothing Then
        NVars(0) = 1    ' Perm Flow
        NVars(1) = 3
    Else
        Call MsgBox("FlowBalanceEqs are not programmed for Permeated Inlet")
        Exit Function
    End If

    VariablesInFlowBalanceEquations = NVars 'feed-product
End Function

Public Function FlowBalanceEquationVars(EquationIdx As Long)
' Called to get the variables in each Flow Balance equation.
    Dim FV() As RealVariable
    
    Select Case EquationIdx
        Case 0 ' Permeation balance
            ReDim FV(0)
            Set FV(0) = edfPermeate.MolarFlow
        Case 1 ' Normal balance
            ReDim FV(2)
            Set FV(0) = edfInlet.MolarFlow
            Set FV(1) = edfPermeate.MolarFlow
            Set FV(2) = edfRetentate.MolarFlow
    End Select
    
    FlowBalanceEquationVars = FV
End Function

Public Function FlowBalanceEquationCoef(EquationIdx As Long)
' Called to get the coefficients of different variables in each Flow Balance equation.
    Dim Coef() As Double
    
    Select Case EquationIdx
        Case 0 ' Permeation balance
            ReDim Coef(1)
            GlobalPermComponentMolarFlow = Permeation()
            Coef(0) = 1    'Permflow
            Coef(1) = Sum(GlobalPermComponentMolarFlow)  'Constant
        Case 1 ' Normal balance
            ReDim Coef(3)
            Coef(0) = 1
            Coef(1) = -1
            Coef(2) = -1
            Coef(3) = 0
    End Select
    
    FlowBalanceEquationCoef = Coef
End Function





Public Sub ReferencePressureInFlowEquations(p1, p2)
' Called to get the pair of pressure variables that are used in Pressure Flow equations.
    Set p1(0) = edfInlet.Pressure
    Set p2(0) = edfPressDrop
End Sub

Public Function FlowInFlowEquations()
' Called to get the flow variables used in Pressure Flow equations.
    Dim flow(0) As RealVariable
    Set flow(0) = edfInlet.MolarFlow
    FlowInFlowEquations = flow
End Function

Public Sub CoefficientsOfFlowEquations(k1, k2)
' Called at each step of the integration to update the coefficients of pressure flow equations.
    k1(0) = edfk.Value
    k2(0) = k1(0)
End Sub

Public Sub UpdateCoefficientsOfFlowEquations(Dtime As Double, k1, k2)
' Called at each step of the integration to update the coefficients of pressure flow equations.
    k1(0) = edfk.Value
    k2(0) = k1(0)
End Sub

Public Function UpdateDensities(Dtime As Double)
' Called at each step of the integration to update the densities in all pressure flow equations.
    Dim density(0) As Double
    density(0) = edfInlet.MolarDensityValue
    UpdateDensities = density
End Function





Public Function VariablesInGeneralEquations()
' Called to get the number of variables in different General equations extension contributes separately.
    'Dim NVars(1) As Long
    'NVars(0) = 2
    'VariablesInGeneralEquations = NVars
End Function

Public Function GeneralEquationVars(EquationIdx As Long)
' Called to get the variables in each General equation.
   'Dim GV(1) As RealVariable
    'Set GV(0) = edfInlet.MolarFlow
    'Set GV(1) = edfRetentate.MolarFlow
    'GeneralEquationVars = GV
End Function

Public Sub PrepareToIterateOnGeneralEqns(Dtime As Double)
' Called at each step of integration right before pressure flow solver starts iteration to solve the set of equations.
    'Fout_old = edfRetentate.MolarFlow
    ' Here we store "old" of derivatives
End Sub

Public Sub UpdateGeneralEqDerivsAndRHS(Dtime As Double, Derivs, Rhs)
' Called at each iteration of the Pressure Flow Solver to update the derivatives and right hand side of General equations.
'  Dim Fin As Double, Fout As Double, Fperm As Double
'  Dim rhoIn As Double, rhoOut As Double, rhoPerm As Double
'  Dim Qin As Double, Qout As Double, Qperm As Double
'    ' Get process variables
'    With edfInlet
'        Fin = .MolarFlowValue
'        'rhoIn = .MolarDensityValue
'        'Qin = .ActualVolumeFlow
'    End With
'    With edfRetentate
'        Fout = .MolarFlowValue
'        'rhoOut = .MolarDensityValue
'        Qout = .ActualVolumeFlow
'    End With
'    With edfPermeate
'        Fperm = .MolarFlow
'        'rhoPerm = .MolarDensityValue
'        'Qperm = .ActualVolumeFlow
'    End With
'    ' Right Hand Side and derivatives assigment
'    Rhs(0) = Volume * (Fout - Fout_old) / Qout / Dtime - Fin + Fout_old + Fperm
'    ' rho [molar] = F/Q
'    Derivs(0, 0) = -rhoIn
'    Derivs(0, 1) = Volume / Qout / Dtime + 2 * Fout / Qout'
End Sub





Public Function PreProcessStates(Dtime As Double) As Boolean
' Called at each step of the integration just before Pressure Flow Solver starts to solve the set of equations.
  On Error GoTo errorHandler
    ' Update permeation: set value it to edfTotalPermMolFlow so it updates FlowBalance coeffs
    GlobalPermComponentMolarFlow = Permeation()
    Call dynContainer.UpdateFlowBalanceEquationCoef(0, 1, Sum(GlobalPermComponentMolarFlow))
    PreProcessStates = CheckExtnRequirements()
    Exit Function
errorHandler:
    Dim sMsg As String
    sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    Call MsgBox(myContainer.Name + " - PreProcessStates: " + sMsg)
End Function

Public Function PostProcessStates(Dtime As Double) As Boolean
' Called at each step of the integration just after Pressure Flow Solver finishes solving the set of equations.
  On Error GoTo errorHandler
    'Updates P-F data


    ' Step 5 - Set the calculated values to the list of fluids and product streams
    'SetPermeationFlows(StreamList, GlobalPermComponentMolarFlow) 'TODO

    ' Step 6 - Pressure drop and flash specifications
    'Dim Tin As Double, retPressure As Double, permPressDrop As Double
    'Tin = edfInlet.TemperatureValue
    'retPressure = edfInlet.PressureValue - edfPressDrop.Value
    'permPressDrop = edfInlet.PressureValue - edfPermeate.PressureValue
    'Call edfRetentate.Temperature.Calculate(Tin) ' TODO: move to StepEnergy?
    'Call edfPermeate.Temperature.Calculate(Tin)  ' TODO: move to StepEnergy?
    'Call edfRetentate.Pressure.Calculate(retPressure)
    'Call edfPermPressDrop.Calculate(permPressDrop)

    ' ProductStream list
    Dim ProductStreams(1)
    'ReDim Preserve ProductStreams(0 To 1)
    ProductStreams(0) = edfRetentate
    ProductStreams(1) = edfPermeate
    'Call dynContainer.UpdateFluidsOfStreams(ProductStreams)          ' TODO: use only products - apparently yes 'TODO: not working currently
    'Call dynContainer.GetStreamsPropertiesFromFluids(ProductStreams)
    PostProcessStates = True
    Exit Function
errorHandler:
    Dim sMsg As String
    sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    Call MsgBox(myContainer.Name + " - PostProcessStates: " + sMsg)
End Function

Public Function StepEnergyExplicitly(Dtime As Double) As Boolean
' Called to do the energy integration on the extension. PERFORM FLASH
  On Error GoTo errorHandler
    Dim i As Long
    ' Save Streams into a vector of fluids
    ' Need to repeat every step to make sure is the duplicates of StreamList
    ReDim Preserve fluidList(0 To UBound(StreamList))
    For i = 0 To UBound(StreamList)
        Set fluidList(i) = StreamList(i).DuplicateFluid
    Next i
    ' Store values in double VB variables for flash
    Dim flashRet As FlashStatus_enum, flashPerm As FlashStatus_enum
    Dim retPressure As Double, permPressDrop As Double, Tin As RealVariable
    Set Tin = edfInlet.Temperature
    retPressure = edfInlet.PressureValue - edfPressDrop.Value
    permPressDrop = edfInlet.PressureValue - edfPermeate.PressureValue
    Call edfPermPressDrop.SetValue(permPressDrop)
    ' Perform Flash
    flashRet = fluidList(1).TPFlash(Tin, retPressure)
    flashPerm = fluidList(2).TPFlash(Tin, edfPermeate.Pressure)
    If flashRet > 0 Then
        Call myContainer.Trace(myContainer.Name & ": Retentate flash failed", False)
    End If
    If flashPerm > 0 Then
        Call myContainer.Trace(myContainer.Name & ": Permeate flash failed", False)
    End If
    ' Transfer fluid to actual streams
    ' Use False argument: dont need to make this method perform flash again
    Call dynContainer.UpdateStreamFluidFromFluid(edfPermeate, fluidList(2), False)
    Call dynContainer.UpdateStreamFluidFromFluid(edfRetentate, fluidList(1), False)     'Updates
    StepEnergyExplicitly = True
    Exit Function
errorHandler:
    Dim sMsg As String
    sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    Call MsgBox(myContainer.Name + " - StepEnergy: " + sMsg)
End Function

Public Function StepCompositionExplicitly(Dtime As Double) As Boolean
  On Error GoTo errorHandler
    ' Distribute flows per component
    Call SetPermeationFluids(fluidList, GlobalPermComponentMolarFlow)
    ' Transfer fluid to actual streams
    Call dynContainer.UpdateStreamFluidFromFluid(edfPermeate, fluidList(2), True)
    Call dynContainer.UpdateStreamFluidFromFluid(edfRetentate, fluidList(1), True)
    ' Visualization functions
    'Composition(fluidList)
    'Condition(fluidList)
    StepCompositionExplicitly = True
    Exit Function
errorHandler:
    Dim sMsg As String
    sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    Call MsgBox(myContainer.Name + " - StepComposition: " + sMsg)
End Function
