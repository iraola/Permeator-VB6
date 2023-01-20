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

'...Container...
Private myContainer As ExtnUnitOperationContainer
'Private m_InpExtnUtils As ExtnUtils_v2.CExtnUtils
'...DynContainer...
Private dyn_Container As ExtnDynUnitOpContainer

'***************************************************************************'
'                                                                           '
'                                 EDF Variables                             '
'                                                                           '
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
Private edfn As InternalRealVariable
Private edfDT As InternalRealVariable
Private edfDH As InternalRealVariable
Private edfComposition As InternalRealFlexVariable
Private edfCompName As InternalTextFlexVariable
Private edfStreamName As InternalTextFlexVariable
Private edfVapFrac As InternalRealFlexVariable
Private edfTemperature As InternalRealFlexVariable
Private edfPressure As InternalRealFlexVariable
Private edfMolFlow As InternalRealFlexVariable
Private edfMassFlow As InternalRealFlexVariable
Private edfPressDrop As InternalRealVariable
Private edfPermPressDrop As InternalRealVariable
Private edfDiamExt As InternalRealVariable
Private edfLengthPos As InternalRealFlexVariable
Private edfCompT As InternalRealFlexVariable
Private edfCompH As InternalRealFlexVariable
Private edfNpoints As InternalRealVariable
Private edfk As InternalRealVariable

'***************************************************************************'
'                                                                           '
'                                  VB Variables                             '
'                                                                           '
'***************************************************************************'

' ...CONSTANTS...
Private Const pi As Double = 3.14159265359
Private Const R As Double = 8.314472        '[kJ/kmol·K]
Private idxH2 As Long, idxHD As Long, idxHT As Long
Private idxD2 As Long, idxDT As Long, idxT2 As Long
Private nComp As Long, nFluid As Long '3 fluids
Private DT As Double
Private DH As Double
Private flashPerm As Double, flashnoperm As Double
'Plot
Dim myPlotNameH As InternalTextVariable
Dim myPlotH As TwoDimensionalPlot
Dim myPlotNameT As InternalTextVariable
Dim myPlotT As TwoDimensionalPlot
'...Flags and other boolean variables...
Private IsForgetting As Boolean
Private flagRet As Boolean, flagPerm As Boolean

'***************************************************************************'
'                                                                           '
'                               Physical Variables                           '
'                                                                           '
'***************************************************************************'
' Geometry
Dim L As Double, thick As Double, Din As Double, Aperm As Double
' Concentrations
Dim CTi As Double, CTo As Double, CHTi As Double, CHTo As Double, CHo As Double, CHi As Double
Dim CTsi As Double, CTso As Double, CHTsi As Double, CHTso As Double, CHso As Double, CHsi As Double
' Temperatures
Private Tin As Double, Tin_K As Double
' Velocities
Private Vo As Double, Vext As Double
' auxiliary variables
Private ko As Double, kext As Double
' Flows
Private PermeationT As Double, PermeationH As Double
Private QCTo_Ti As Double, QCHo_Hi As Double, QCTi_To As Double, QCHi_Ho As Double
Private flowTo As Double, flowHo As Double, flowTi As Double, flowHi As Double
Private FpermCells() As Double
' Compositions
Private CompTIni As Double, CompHIni As Double, CompTiIni As Double, CompHiIni As Double
' Densities
Private Fout_old As Double ' for GeneralEq iterations
' Volumes
Private Volume, Area As Double
' Arrays for plotting
Private GrafT() As Double, GrafH() As Double, GrafL() As Double
' Streams
Dim fluidList() As Fluid
Dim StreamList() As ProcessStream
Dim permList() As Double

Implements HYSYS.ExtnUnitOperation
Implements HYSYS.ExtensionObject


Private Function ExtnUnitOperation_Initialize(ByVal Container As HYSYS.ExtnUnitOperationContainer, ByVal IsRecalling As Boolean) As Long
 Dim IRV As InternalRealVariable '...for the controllers that only need to be initialized...
 Dim Ravg As Double
    '...Get the pointer to the container...
    Set myContainer = Container
    Call pointedEDFVariables
    
    '  m_InpExtnUtils.PutVersionInfo
    If IsRecalling Then
        edfn.SetValue 2
        edfn.SetModifyState vsDefaultedValue
    Else
        '     m_InpExtnUtils.InitType = 1
        '...visibility controlers...
        Set IRV = myContainer.FindVariable("Design_enum").Variable
        IRV.Value = 0
        'Set the initial num value: Karlsruhe experiment for WDS and TEP (Welte, 2009)
        thick = 0.0001              'm
        edfThick.SetValue thick
        Din = 0.0023                'm
        edfDiam.SetValue Din
        edfNtubes.SetValue 183      '-
        edfLen.SetValue 0.9         'm
        L = edfLen.GetValue * edfNtubes.GetValue
        Ravg = ((Din + (Din + 2 * thick)) / 2) / 2
        Aperm = L * (2 * pi * Ravg)
        edfAperm.SetValue Aperm     'm2
    End If
    ' Global variables: first calc
    thick = edfThick.GetValue
    Din = edfDiam.GetValue
    Aperm = edfAperm.GetValue
    Area = pi * (Din ^ 2) / 4
    L = edfLen.GetValue * edfNtubes.GetValue
    Volume = L * Area
    Call CreatePlot
    
    ' Loop for setting index for components of interest (based on Inlet's basis manager
    Call iniCompIndex
    
    ExtnUnitOperation_Initialize = extnCurrentVersion
    Exit Function
ErrorCatch:
    MsgBox "Error in Initialize function"
End Function

Private Sub ExtnUnitOperation_Execute(ByVal Forgetting As Boolean)
  Dim PermeationT As Double, PermeationH As Double, NonPermeationT As Double, NonPermeationH As Double
  Dim Comp As Boolean, cond As Boolean
  Dim i As Long

On Error GoTo ErrorCatch:
    If Forgetting Then
        IsForgetting = True
        Exit Sub
    End If
    IsForgetting = False    ' Personal flag
    
    ' Check if we have the streams available to calculate
    If (edfInlet Is Nothing Or edfPermeate Is Nothing Or edfRetentate Is Nothing) Then Exit Sub

    
    ' Building vector of Streams and Fluids
    ReDim Preserve StreamList(0 To 2)
    Set StreamList(0) = edfInlet            ' In
    Set StreamList(1) = edfRetentate        ' Out non Permeated
    Set StreamList(2) = edfPermeate         ' Out Permeated
    If Not edfPermIn Is Nothing Then        ' (optional) 2nd stream in
        ReDim Preserve StreamList(0 To 3)
        Set StreamList(3) = edfPermIn
    End If
    ReDim Preserve fluidList(0 To UBound(StreamList))
    For i = 0 To UBound(StreamList)
        Set fluidList(i) = StreamList(i).DuplicateFluid
    Next i
    
    ' Calculate Permeation
    Dim aux() As Double
    ReDim aux(nComp - 1)
    aux = Permeation() ' kmol/s
    PermeationT = aux(idxT2)
    PermeationH = aux(idxH2)
  
    ' Set the calculated values to the list of fluids
    fluidList = setProductFluids(fluidList, PermeationT, PermeationH)

    'Pressure Drop Calculation
    ' Para el flujo no permeado (1)
    If edfPressDrop.IsKnown And fluidList(0).Pressure.IsKnown Then
        flashnoperm = fluidList(1).TPFlash(Tin, fluidList(0).PressureValue - edfPressDrop.Value)
        edfRetentate.Pressure.Calculate fluidList(1).PressureValue
    ElseIf fluidList(0).Pressure.IsKnown And fluidList(1).Pressure.IsKnown Then
        flashnoperm = fluidList(1).TPFlash(Tin, fluidList(1).PressureValue)
        edfPressDrop.Calculate fluidList(0).PressureValue - fluidList(1).PressureValue
    Else
        Exit Sub
    End If
    ' Para el flujo permeado (2)
    If edfPermPressDrop.IsKnown And fluidList(0).Pressure.IsKnown Then
        flashPerm = fluidList(2).TPFlash(Tin, fluidList(0).PressureValue - edfPermPressDrop.Value)
        edfPermeate.Pressure.Calculate fluidList(2).PressureValue
    ElseIf fluidList(0).Pressure.IsKnown And fluidList(2).Pressure.IsKnown Then
        flashPerm = fluidList(2).TPFlash(Tin, fluidList(2).PressureValue)
        edfPressDrop.Calculate fluidList(0).PressureValue - fluidList(2).PressureValue
    Else
        Exit Sub
    End If

    ' If flashes=0 means that Flash methods has gone alright and we update Streams conditions
    If flashPerm = 0 And flashnoperm = 0 Then
        For i = 0 To UBound(StreamList)
            StreamList(i).CalculateAsFluid fluidList(i), ftTPFlash
        Next i
        ' A precautionary Balance is performed
        myContainer.Balance btTotalBalance, 1, StreamList
        ' Composition and Condition functions are for visualizing in the EDF (no physical operation here)
        Comp = Composition(fluidList)
        cond = Condition(fluidList)
        '...if we are here it is because we have solved the unit properly...
        myContainer.SolveComplete
    End If
    Exit Sub
    
ErrorCatch:
    'MsgBox "Error in Execute method. " & Err.Description
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
    If edfLen < 0 Or edfDiam < 0 Or edfThick < 0 Or edfk < 0 Or edfNtubes < 0 Or edfAperm < 0 Then
        Call Status.AddStatusCondition(slMissingRequiredInformation, 4, "Physical parameters missing or incorrect")
    End If
    If flagRet Then
       Call Status.AddStatusCondition(slError, 5, "All feed flow is being permeated")
    End If
    If flagPerm Then
        Call Status.AddStatusCondition(slError, 6, "No species permeating")
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
' Called when the user modifies any edf variable -> here is required to update this variales as needed
    Dim Ravg As Double
    Select Case Variable.Tag
        Case "Inlet"
             Set edfInlet = myContainer.FindVariable("Inlet").Variable.object
        Case "PermOut"
            Set edfPermeate = myContainer.FindVariable("PermOut").Variable.object
        Case "NoPermOut"
            Set edfRetentate = myContainer.FindVariable("NoPermOut").Variable.object
        Case "PermIn"
            Set edfPermIn = myContainer.FindVariable("PermIn").Variable.object
        Case "ActualizePlot"
            edfCompT.SetBounds (edfNpoints.Value + 1)
            edfCompH.SetBounds (edfNpoints.Value + 1)
            edfLengthPos.SetBounds (edfNpoints.Value + 1)
            edfLengthPos.Values = GrafL
            edfCompT.Values = FpermCells
            edfCompH.Values = FpermCells
        '%% Geometry variables
        Case "Length"
            Set edfLen = myContainer.FindVariable("Length").Variable
            L = edfLen.Value * edfNtubes                    ' total length of tubes
            ' Recalculation of cross-area and total volume
            Area = pi * (Din ^ 2) / 4
            Volume = L * Area
            GrafL = LengthVector(edfLen, edfNpoints)
            ' Recalculation of total permeation surface
            Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
            edfAperm.SetValue L * (2 * pi * Ravg)
            Aperm = edfAperm.Value
        Case "Diam"
            Set edfDiam = myContainer.FindVariable("Diam").Variable
            Din = edfDiam.GetValue
            ' Recalculation of cross-area and total volume
            Area = pi * (Din ^ 2) / 4
            Volume = edfLen.Value * edfNtubes * Area
            ' Recalculation of total permeation surface
            Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
            edfAperm.SetValue L * (2 * pi * Ravg)
            Aperm = edfAperm.Value
        Case "Thickness"
            Set edfThick = myContainer.FindVariable("Thickness").Variable
            thick = edfThick.Value
            ' Recalculation of total permeation surface
            Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
            edfAperm.SetValue L * (2 * pi * Ravg)
            Aperm = edfAperm.Value
        Case "Aperm"
            Set edfAperm = myContainer.FindVariable("Aperm").Variable
            Aperm = edfAperm.Value
            ' Recalculation of total permeation surface
            Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
            L = Aperm / (2 * pi * Ravg)                     ' total length of tubes
            edfLen.SetValue L / edfNtubes.GetValue
            ' Recalculation of cross-area and total volume
            Area = pi * (Din ^ 2) / 4
            Volume = L * Area
        Case "Ntubes"
            Set edfNtubes = myContainer.FindVariable("Ntubes").Variable
            ' Recalculation of total length (mantaining length/tube (edfLen))
            L = edfLen.Value * edfNtubes
            ' Recalculation of total permeation surface (last change affects area)
            Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
            edfAperm.SetValue L * (2 * pi * Ravg)
            Aperm = edfAperm.Value
            ' Recalculation of cross-area and total volume
            Area = pi * (Din ^ 2) / 4
            Volume = L * Area
        '%% Pressure drop and others
        Case "n"
            Set edfn = myContainer.FindVariable("n").Variable
        Case "PressDrop"
            Set edfPressDrop = myContainer.FindVariable("PressDrop").Variable
        Case "PermPressDrop"
            Set edfPermPressDrop = myContainer.FindVariable("PermPressDrop").Variable
        Case "k"
            Set edfk = myContainer.FindVariable("k").Variable
        Case "NumberOfPoints"
            Set edfNpoints = myContainer.FindVariable("NumberOfPoints").Variable
            GrafL = LengthVector(edfLen, edfNpoints)
    End Select
        
        ' edf Varaiables remaining
        'Set edfDT = .FindVariable("DT").Variable
        'Set edfDH = .FindVariable("DH").Variable
        'Set edfComposition = .FindVariable("Composition").Variable
        'Set edfCompName = .FindVariable("CompoName").Variable
        'Set edfStreamName = .FindVariable("StreamNames").Variable
        'Set edfVapFrac = .FindVariable("VapFrac").Variable
        'Set edfTemperature = .FindVariable("StreamTemp").Variable
        'Set edfPressure = .FindVariable("StreamPress").Variable
        'Set edfMolFlow = .FindVariable("StreamMolFlow").Variable
        'Set edfMassFlow = .FindVariable("StreamMassFlow").Variable
        'Set edfLengthPos = .FindVariable("LengthPos").Variable
        'Set edfCompT = .FindVariable("CompT").Variable
        'Set edfCompH = .FindVariable("CompH").Variable
        'Set myPlotNameH = .FindVariable("PlotNameH").Variable
        'Set myPlotNameT = .FindVariable("PlotNameT").Variable
        'Set edfDiamExt = .FindVariable("DiamExtern").Variable
    
    Exit Sub

ERROR_CATCH:
    Call myContainer.Trace(myContainer.Name & ": Error in Variable Changed for variable " & Variable.Tag & ".", False)
End Sub

Private Function ExtensionObject_VariableChanging(ByVal Variable As HYSYS.InternalVariableWrapper) As Boolean
    ExtensionObject_VariableChanging = True
    'Dim varName As String
    ' Catch the name of the variable being changed
    'varName = Variable.Tag
    Select Case Variable.Tag
    'Al final este código daba errores... y ni si quiera hacían falta con el VariableChanged funcionando
    '    Case "Inlet"
    '        If Variable.NewObjectValue Is Nothing Then
    '        End If
    '    Case "PermOut"
    '        If Variable.NewObjectValue Is Nothing Then
    '            Set edfPermeate = Nothing
    '        End If
    '    Case "NoPermOut"
    '        If myContainer.Flowsheet.Streams.Item(varName) Is Nothing Then
    '            ExtensionObject_VariableChanging = False
    '        End If
    '    Case "PermIn"
    '        If myContainer.Flowsheet.Streams.Item(varName) Is Nothing Then
    '            ExtensionObject_VariableChanging = False
    '        End If
    End Select
End Function

Private Sub ExtensionObject_VariableQuery(ByVal Variable As HYSYS.InternalVariableWrapper)
End Sub

Private Sub ExtnUnitOperation_BasisChanged()
End Sub

Private Sub ExtensionObject_Terminate()
    '  m_InpExtnUtils.ExtnUtils_Terminate
    Set edfThick = Nothing
    Set edfLen = Nothing
    Set edfDiam = Nothing
    Set edfAperm = Nothing
    Set edfNtubes = Nothing
    Set edfn = Nothing
    Set edfInlet = Nothing
    Set edfPermeate = Nothing
    Set edfRetentate = Nothing
    Set edfPermIn = Nothing
    Set edfDT = Nothing
    Set edfDH = Nothing
    Set edfComposition = Nothing
    Set edfCompName = Nothing
    Set edfStreamName = Nothing
    Set edfVapFrac = Nothing
    Set edfTemperature = Nothing
    Set edfPressure = Nothing
    Set edfMolFlow = Nothing
    Set edfMassFlow = Nothing
    Set edfPressDrop = Nothing
    Set edfPermPressDrop = Nothing
    Set edfLengthPos = Nothing
    Set edfCompT = Nothing
    Set edfCompH = Nothing
    Set edfNpoints = Nothing
    Set myPlotH = Nothing
    Set myPlotNameH = Nothing
    Set myPlotT = Nothing
    Set myPlotNameT = Nothing
    Set edfDiamExt = Nothing
    Set edfk = Nothing
End Sub

Private Sub pointedEDFVariables()
    With myContainer
        Set edfThick = .FindVariable("Thickness").Variable
        Set edfLen = .FindVariable("Length").Variable
        Set edfDiam = .FindVariable("Diam").Variable
        Set edfAperm = .FindVariable("Aperm").Variable
        Set edfNtubes = .FindVariable("Ntubes").Variable
        Set edfInlet = .FindVariable("Inlet").Variable.object
        Set edfPermeate = .FindVariable("PermOut").Variable.object
        Set edfRetentate = .FindVariable("NoPermOut").Variable.object
        Set edfPermIn = .FindVariable("PermIn").Variable.object
        Set edfn = .FindVariable("n").Variable
        Set edfDT = .FindVariable("DT").Variable
        Set edfDH = .FindVariable("DH").Variable
        Set edfComposition = .FindVariable("Composition").Variable
        Set edfCompName = .FindVariable("CompoName").Variable
        Set edfStreamName = .FindVariable("StreamNames").Variable
        Set edfVapFrac = .FindVariable("VapFrac").Variable
        Set edfTemperature = .FindVariable("StreamTemp").Variable
        Set edfPressure = .FindVariable("StreamPress").Variable
        Set edfMolFlow = .FindVariable("StreamMolFlow").Variable
        Set edfMassFlow = .FindVariable("StreamMassFlow").Variable
        Set edfPressDrop = .FindVariable("PressDrop").Variable
        Set edfPermPressDrop = .FindVariable("PermPressDrop").Variable
        Set edfLengthPos = .FindVariable("LengthPos").Variable
        Set edfCompT = .FindVariable("CompT").Variable
        Set edfCompH = .FindVariable("CompH").Variable
        Set edfNpoints = .FindVariable("NumberOfPoints").Variable
        Set myPlotNameH = .FindVariable("PlotNameH").Variable
        Set myPlotNameT = .FindVariable("PlotNameT").Variable
        Set edfDiamExt = .FindVariable("DiamExtern").Variable
        'Set molDensity = .FindVariable("molDensity").Variable
        Set edfk = .FindVariable("k").Variable
    End With
End Sub


'*******************************
'***    USER FUNCTIONS       ***
'*******************************
Private Function Composition(myFluid() As Fluid) As Boolean
Dim iFluid As Long, i As Long
Dim streamName() As String
Dim x() As Double
    ' Set bounds of final EDF vatiables
    edfCompName.SetBounds nComp
    edfComposition.SetBounds nFluid, nComp
    edfStreamName.SetBounds nFluid
    ' Construct auxiliary vectors/matrix for tables
    ReDim Preserve x(nComp - 1, 0 To nFluid - 1)
    ReDim streamName(0 To nFluid - 1)
    For iFluid = 0 To nFluid - 1
        streamName(iFluid) = myFluid(iFluid).Name
        For i = 0 To nComp - 1
            x(i, iFluid) = myFluid(iFluid).MolarFractions(i)
        Next i
    Next iFluid
    ' Associate matrix to final EDF variables for visualizing
    edfStreamName = streamName                          ' Columns labels (streams)
    edfComposition.Values = x                           ' Values in table
    edfCompName.Values = myFluid(0).Components.names    ' Rows labels (components)
    Composition = True
End Function
Private Function Condition(myFluid() As Fluid) As Boolean
Dim i As Long
Dim vapfrac() As Double, Temp() As Double, Press() As Double, MolFlow() As Double, massFlow() As Double
    ' Set bounds of final EDF vatiables
    edfVapFrac.SetBounds nFluid
    edfTemperature.SetBounds nFluid
    edfPressure.SetBounds nFluid
    edfMolFlow.SetBounds nFluid
    edfMassFlow.SetBounds nFluid
    ' Construct auxiliary vectors/matrix for tables
    ReDim vapfrac(0 To nFluid - 1)
    ReDim Temp(0 To nFluid - 1)
    ReDim Press(0 To nFluid - 1)
    ReDim MolFlow(0 To nFluid - 1)
    ReDim massFlow(0 To nFluid - 1)
    For i = 0 To nFluid - 1
         vapfrac(i) = myFluid(i).VapourFractionValue
         Temp(i) = myFluid(i).TemperatureValue
         Press(i) = myFluid(i).PressureValue
         massFlow(i) = myFluid(i).MassFlowValue
         MolFlow(i) = myFluid(i).MolarFlowValue
     Next i
     ' Associate matrix to final EDF variables for visualizing
     edfVapFrac = vapfrac
     edfTemperature = Temp
     edfPressure = Press
     edfMolFlow = MolFlow
     edfMassFlow = massFlow
     Condition = True
End Function

Sub CreatePlot()
    If Not myPlotH Is Nothing Then
        myContainer.DeletePlot "LCPlotH"
        Set myPlotH = Nothing
    End If
    myContainer.BuildPlot2 "LCPlotH", myPlotH, hptTwoDimensionalPlot
    myPlotNameH.Value = "LCPlotH"
    With myPlotH
        .TitleData = "Length vs Non Permeation"
        
        .SetAxisLabelData hatXAxis, "Length (m)"
        .SetAxisLabelData hatYAxis, "NoPerm Flow"
        .SetAxisLabelVisible hatXAxis, True
        .SetAxisLabelVisible hatYAxis, True
        
        .LegendVisible = True
        
        .CreateXYDataSet 1, "PQData"
        .SetDataSetXData 1, edfLengthPos
        .SetDataSetYData 1, edfCompH
        .SetDataSetColour 1, "Red"
    End With
    If Not myPlotT Is Nothing Then
        myContainer.DeletePlot "LCPlotT"
        Set myPlotT = Nothing
    End If
    myContainer.BuildPlot2 "LCPlotT", myPlotT, hptTwoDimensionalPlot
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

Private Sub iniCompIndex()
' Loop for setting index for components that can permeate (based on Inlet's basis manager)
  Dim ComponentList As Components
  Dim i As Long
    If edfInlet Is Nothing Then
        Set ComponentList = myContainer.Flowsheet.FluidPackage.Components
    Else
        Set ComponentList = edfInlet.DuplicateFluid.Components
    End If
    nComp = ComponentList.Count         ' Number of Components
    nFluid = 3                          ' Number of fluids attached: 1 feed, 2 products
    For i = 0 To nComp - 1
        Select Case ComponentList.Item(i).Name
            Case "Hydrogen"
                idxH2 = i
            Case "Hydrogen*"
                idxH2 = i
            Case "HD*"
                idxHD = i
            Case "HT*"
                idxHT = i
            Case "Deuterium*"
                idxD2 = i
            Case "DT*"
                idxDT = i
            Case "Tritium*"
                idxT2 = i
        End Select
    Next i
End Sub

Private Function Permeation()
' Calculate vector of permeated species in default VB HYSYS magnitude: "kmol/s"
'   Fin(nComp):         vector      inlet molar flow (per component)
'   Fperm(nComp):       vector      permeated molar flow in one cell (per component)
'   FpermTotal(nComp):  vector      aggregates molar flow permeated for each cell (per component)
'   Fcell(nComp):       vector      similar to "Fin" but for each cell calculated from previous cell
'   FpermCells(Npoints) vector      collect total permeation in each cell for plotting purposes
 Dim Fperm() As Double, FpermTotal() As Double, Fin() As Double, Fcell() As Double
 Dim Qin As Double
 Dim dx As Double, Ravg As Double, ApermCell As Double
 Dim i As Long, nCell As Long
  ReDim Fin(nComp - 1), Fperm(nComp - 1), FpermTotal(nComp - 1), Fcell(nComp - 1)
    ' Get EDF parameters into double VB variables
    nCell = edfNpoints.GetValue
  ReDim FpermCells(nCell - 1)
    Tin = edfInlet.Temperature.GetValue("C")
    Tin_K = edfInlet.Temperature.GetValue("K")
    Qin = edfInlet.ActualVolumeFlowValue
    ' Difusivity Calculation
        '    ' (Austenitic Steel 316L)
        '    DT = 0.00000059 * Exp(-51.9 * 1000 / (R * Tin_K))
        '    DH = DT * (3 ^ (1 / 2))
            ' AgPd Diffusivity on hydrogen (Serra et al., 1998)
    DH = 0.000000307 * Exp(-25902 / (R * Tin_K))
    DT = DH / Sqr(3)        ' Mass isotopic teorical classical relationship
    ' Set calculated values in edf for user's visualization
    edfDT.SetValue DT
    edfDH.SetValue DH
    ' Geometric calculations
    ApermCell = Aperm / nCell                  ' permeation surface per differential cell
    ' Initialize
    Fin = fluidList(0).MolarFlowsValue              ' molar flow per component
    Fcell = Fin
    'FpermTotal = 0 ' initialy zero with no need to initialize it as long as it was local
    ' LOOP over cells
    For i = 0 To nCell - 1
        ' Calculate concentrations [kgmole/m3]
        CHTsi = Fcell(idxHT) / Qin
        CHsi = Fcell(idxH2) / Qin
        CTsi = Fcell(idxT2) / Qin
        CTso = 0
        CHso = 0
        CHTso = 0
        If CTi = -32767 Then CTi = 0
        If CTo = -32767 Then CTo = 0
        If CHTi = -32767 Then CHTi = 0
        If CHTo = -32767 Then CHTo = 0
        If CHi = -32767 Then CHi = 0
        If CHo = -32767 Then CHo = 0
        ' Permeation calculation [at/s]. Richardson's law: [kmol/s·m2] -> multiply by "ApermCell" -> [kmol/s]
        Fperm(idxT2) = (DT / thick) * ((CHTsi + CTsi) - (CTso + CHTso)) * ApermCell
        Fperm(idxH2) = (DH / thick) * ((CHTsi + CHsi) - (CHso + CHTso)) * ApermCell
        '   Old calc. (pi * L * DH / Log(1 + (thick / (Din / 2))) * ((CHTsi + CHsi) - (CHso + CHTso)))
        ' Molar flow component vector in next cell
        Fcell = vectorSubtractIf(Fcell, Fperm)      ' Special function for dealing with undesired negarive permeate flow
        FpermTotal = vectorSum(Fperm, FpermTotal)   ' Aggregate permeation from previous cells
        FpermCells(i) = sumVectorElements(Fperm)
    Next i
    ' Check if permeate results bigger than inlet flow
    Permeation = FpermTotal
End Function

Private Function setProductFluids(fluids, permT, permH) As Variant
' Returns vector of fluids with new permeated composition
 Dim names() As String
 Dim permflow() As Double, retflow() As Double
 Dim retH As Double, retT As Double
 Dim i As Long
 ReDim names(nComp - 1), permflow(nComp - 1), retflow(nComp - 1)
    ' Calculate retentate flow
    retT = fluidList(0).MolarFlowsValue(idxT2) - permT
    retH = fluidList(0).MolarFlowsValue(idxH2) - permH
    If retT < 0 Then retT = 0
    If retH < 0 Then retH = 0
    ' Loop over all components setting the proper molar flow
    names = fluids(0).Components.names
    For i = 0 To nComp - 1
        If names(i) = "Tritium*" Then
            permflow(i) = permT
            retflow(i) = retT
        ElseIf (names(i) = "Hydrogen" Or names(i) = "Hydrogen*") Then
            permflow(i) = permH
            retflow(i) = retH
        Else
            ' Any non-permeating component
            permflow(i) = 0
            retflow(i) = fluids(0).MolarFlowsValue()(i) ' Necessary () for not error in property
        End If
    Next i
    ' Check if streams are null flow (solver will not continue). Used in StatusQuery
    flagPerm = True
    flagRet = True
    For i = 0 To nComp - 1
        If permflow(i) <> 0 Then flagPerm = False
        If retflow(i) <> 0 Then flagRet = False
    Next i
    ' Set the fictitious molar flow vectors to the actual ones
    fluids(1).MolarFlows.SetValues retflow, "kgmole/s"        ' Outlet
    fluids(2).MolarFlows.SetValues permflow, "kgmole/s"      ' Permeate
    
    setProductFluids = fluids
End Function



'***************************
'*** AUXILIARY FUNCTIONS ***
'***************************
Private Function max(n1, n2) As Double
    If n1 > n2 Then max = n1 Else: max = n2
End Function
Private Function vectorSum(A1, A2) As Double()
'' Subtract vectors
Dim n As Long, i As Long
Dim A() As Double
    ' Get sizes
    n = UBound(A1, 1) + 1
    ReDim A(n - 1)
    ' Check condition that allow sum/subtract
    If (UBound(A2, 1) + 1) <> n Then
        MsgBox "Sizes do not match in array sum function."
        Exit Function
    End If
    ' Operate
    For i = 0 To n - 1
        A(i) = A1(i) + A2(i)
    Next i
    vectorSum = A
End Function
Private Function vectorSubtract(A1, A2) As Double()
'' Subtract vectors and takes into account that output cannot be negative, modifying also a vector input if necessary
Dim n As Long, i As Long
Dim A() As Double
    ' Get sizes
    n = UBound(A1, 1) + 1
    ReDim A(n - 1)
    ' Check condition that allow sum/subtract
    If (UBound(A2, 1) + 1) <> n Then
        MsgBox "Sizes do not match in array sum function."
        Exit Function
    End If
    ' Operate
    For i = 0 To n - 1
        A(i) = A1(i) - A2(i)
    Next i
    vectorSubtract = A
End Function
Private Function vectorSubtractIf(ByRef A1, ByRef A2) As Double()
'' Subtract vectors
Dim n As Long, i As Long
Dim A() As Double
    ' Get sizes
    n = UBound(A1, 1) + 1
    ReDim A(n - 1)
    ' Check condition that allow sum/subtract
    If (UBound(A2, 1) + 1) <> n Then
        MsgBox "Sizes do not match in array sum function."
        Exit Function
    End If
    ' Operate
    For i = 0 To n - 1
        If A1(i) >= A2(i) Then
            A(i) = A1(i) - A2(i)
        Else
            ' Component i has been permeated till its limit
            A(i) = 0        'no more "i" passes to next cell
            A2(i) = A1(i)   'the amount of permeated reduces to the existing flow of "i"
        End If
    Next i
    vectorSubtractIf = A
End Function
Private Function LengthVector(L, n) As Double()
' Builds a vector of equidistant elements representing the position [m] of each cell
 Dim i As Long
 Dim lenvec() As Double, dx As Double
    ReDim lenvec(n - 1)
    ReDim LenghtVector(n - 1)
    dx = L / n
    lenvec(0) = dx
    For i = 1 To n - 1
        lenvec(i) = lenvec(i - 1) + dx
    Next i
    LengthVector = lenvec
End Function
Private Function sumVectorElements(x) As Double
' Returns the sum of elements in vector 'x'
 Dim n As Long, i As Long
 Dim aux As Double
    n = UBound(x) + 1
    aux = 0
    For i = 0 To n - 1
        aux = aux + x(i)
    Next i
    sumVectorElements = aux
End Function
'*************************************
'*** ***************************** ***
'*************************************







'-----------------------------------------------------------------------'
'                                                                       '
'                      Dynamics Methods                                 '
'                                                                       '
'-----------------------------------------------------------------------'

Public Sub DynInitialize(dContainer As ExtnDynUnitOpContainer, IsRecalling As Boolean, MyVersion As Long, HoldupExist As Boolean)
' Called to Initialize the extension.
    Set dyn_Container = dContainer
    HoldupExist = False
    'Quant es necessari
    MyVersion = extnCurrentVersion
End Sub


Public Function InitializeSystem(ForceInit As Boolean) As Boolean
' Called each time the integration starts
    Dim i As Long
    InitializeSystem = False
    'ErrorStatus = False
    
    If (edfInlet Is Nothing Or edfPermeate Is Nothing Or edfRetentate Is Nothing) Then Exit Function
    
    ' ProccesStreams list
    ReDim Preserve StreamList(0 To 2)
    Set StreamList(0) = edfInlet
    Set StreamList(1) = edfRetentate
    Set StreamList(2) = edfPermeate
    ' La StreamList apunta directamente a las variables Attachment del EDF,
    ' si se modifican los streams, también se ve reflejado en StreamList
    If Not edfPermIn Is Nothing Then
        ReDim Preserve StreamList(0 To 3)
        Set StreamList(3) = edfPermIn
    End If
    ' Fluid list
    ReDim Preserve fluidList(0 To UBound(StreamList))
    For i = 0 To UBound(StreamList)
        Set fluidList(i) = StreamList(i).DuplicateFluid
    Next i
    
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
                Set P(0) = edfInlet.Pressure
                Set P(1) = edfRetentate.Pressure
            Case 1 'Perm Side
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
        MsgBox "FlowBalanceEqs are not programmed for Permeated Inlet"
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
            Dim aux() As Double
            ReDim aux(nComp - 1)
            ' Calculate first perm value
            aux = Permeation()
            ' Define coefs
            Coef(0) = 1    'Permflow
            Coef(1) = aux(idxT2) + aux(idxH2) 'Constant
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
    Dim NVars(1) As Long
    NVars(0) = 2
    VariablesInGeneralEquations = NVars
End Function

Public Function GeneralEquationVars(EquationIdx As Long)
' Called to get the variables in each General equation.
    Dim GV(1) As RealVariable
    Set GV(0) = edfInlet.MolarFlow
    Set GV(1) = edfRetentate.MolarFlow
    GeneralEquationVars = GV
End Function

Public Sub PrepareToIterateOnGeneralEqns(Dtime As Double)
' Called at each step of integration right before pressure flow solver starts iteration to solve the set of equations.
    Fout_old = edfRetentate.MolarFlow
    ' Aquí guardamos las "old" de las variables en derivadas
End Sub

Public Sub UpdateGeneralEqDerivsAndRHS(Dtime As Double, Derivs, Rhs)
' Called at each iteration of the Pressure Flow Solver to update the derivatives and right hand side of General equations.
  Dim Fin As Double, Fout As Double, Fperm As Double
  Dim rhoIn As Double, rhoOut As Double, rhoPerm As Double
  Dim Qin As Double, Qout As Double, Qperm As Double
    ' Get process variables
    With edfInlet
        Fin = .MolarFlowValue
        'rhoIn = .MolarDensityValue
        'Qin = .ActualVolumeFlow
    End With
    With edfRetentate
        Fout = .MolarFlowValue
        'rhoOut = .MolarDensityValue
        Qout = .ActualVolumeFlow
    End With
    With edfPermeate
        Fperm = .MolarFlow
        'rhoPerm = .MolarDensityValue
        'Qperm = .ActualVolumeFlow
    End With
    ' Right Hand Side and derivatives assigment
    Rhs(0) = Volume * (Fout - Fout_old) / Qout / Dtime - Fin + Fout_old + Fperm
    ' rho [molar] = F/Q
    Derivs(0, 0) = -rhoIn
    Derivs(0, 1) = Volume / Qout / Dtime + 2 * Fout / Qout
End Sub





Public Function PreProcessStates(Dtime As Double) As Boolean
' Called at each step of the integration just before Pressure Flow Solver starts to solve the set of equations.
    Dim aux() As Double
    ReDim aux(nComp - 1)
    aux = Permeation()
    ' Save in PermeationT, PermeationH, so it will be used in StepEnergyExplicitly( )
    PermeationT = aux(idxT2)
    PermeationH = aux(idxH2)
    dyn_Container.UpdateFlowBalanceEquationCoef 0, 1, (PermeationT + PermeationH)
    If Not edfLen.IsKnown Or Not edfDiam.IsKnown Or Not edfThick.IsKnown Or Not edfk.IsKnown Or Not edfAperm.IsKnown Or Not edfNtubes.IsKnown Then
        PreProcessStates = False
    Else
        PreProcessStates = True
    End If
End Function

Public Function PostProcessStates(Dtime As Double) As Boolean
' Called at each step of the integration just after Pressure Flow Solver finishes solving the set of equations.
'Updates P-F data
   PostProcessStates = True
End Function

Public Function StepEnergyExplicitly(Dtime As Double) As Boolean
' Called to do the energy integration on the extension. PERFORM FLASH
    Dim i As Long
    ' Save Streams into a vector of fluids
    ReDim Preserve fluidList(0 To UBound(StreamList))
    For i = 0 To UBound(StreamList)
        Set fluidList(i) = StreamList(i).DuplicateFluid
    Next i
    ' Store values in double VB variables for flash
    Tin = edfInlet.Temperature
    Dim Pperm, Pout As Double
    Pperm = edfPermeate.PressureValue
    Pout = edfRetentate.PressureValue
    ' Perform Flash
    flashPerm = fluidList(2).TPFlash(Tin, Pperm)
    flashnoperm = fluidList(1).TPFlash(Tin, Pout)
    ' Transfer fluid to actual streams
    dyn_Container.UpdateStreamFluidFromFluid edfPermeate, fluidList(2), True
    dyn_Container.UpdateStreamFluidFromFluid edfRetentate, fluidList(1), True     'Updates
    
    StepEnergyExplicitly = True
End Function

Public Function StepCompositionExplicitly(Dtime As Double) As Boolean
  On Error GoTo errorHandler
    ' Distribute flows per component
    fluidList = setProductFluids(fluidList, PermeationT, PermeationH)
    ' Transfer fluid to actual streams
    dyn_Container.UpdateStreamFluidFromFluid edfPermeate, fluidList(2), True
    dyn_Container.UpdateStreamFluidFromFluid edfRetentate, fluidList(1), True
    ' Visualization functions
    Call Composition(fluidList)
    Call Condition(fluidList)
    StepCompositionExplicitly = True
    Exit Function
errorHandler:
    If flagPerm Or flagRet Then
        MsgBox "Empty streams resulting from permeator " + myContainer.Name
    Else
        MsgBox myContainer.Name + "error"
    End If
End Function
