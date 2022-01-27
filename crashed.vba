Option Explicit
Sub APC()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
ActiveSheet.DisplayPageBreaks = False
Application.EnableEvents = False
Application.Calculation = xlManual

    Call timefill

    Dim currTime As Date
    Dim elapsedtime As Date
    Dim timeremain As Date
    currTime = Now()
    
    Dim C_nim1 As Double
    Dim S_nim1 As Double
    Dim D_nm1i As Double
    Dim D_nm1i_2 As Double
    Dim E_ni As Double
    Dim P_ni As Double
    Dim R_ni As Double
    Dim species As String
    Dim TimeStep As Integer
    Dim OLIBal As Double

    Dim WaterMass As Double
    Dim LiqDensity As Double
    Dim LiqPH As Double
    Dim SolidMass As Double
    Dim LiquidMass As Double
    Dim SolidVolume As Double
    Dim DischargeVol As Double
    Dim TDS As Double
    Dim V_lnij As Double
    Dim V_tnij As Double
    Dim V_tnim1 As Double
    Dim Q As Double
    Dim C_nij As Double
    Dim S_nij As Double
    Dim S_ni As Double
    Dim C_ni As Double
    Dim F_ni As Double
    Dim D_ni As Double
    Dim H_ni As Double
    Dim SeepageRatio As Double
    Dim OutflowRatio As Double
    Dim HarvestRatio As Double
    Dim PC2toC3 As Double
    Dim TimeStepFactor As String
    Dim TimeFactor As Double
    Dim CurrentWS As String
    Dim StartVol As Double
    Dim MaxDepth As Double
    Dim MaxVol As Double
    Dim MinDisch As Double
    Dim MaxDisch As Double
    
    Dim ContinueRun As Boolean
    ContinueRun = True
    
    Dim totsteps As Integer
    Dim numponds As Integer
    Dim trtotsteps As Integer
    Dim Step As Integer
    Dim pondint As Integer
    Dim trstep As Integer
    
    Dim StartPond As Worksheet
    Dim InputOLI As Worksheet
    Dim OutputOLI As Worksheet
    Dim GlobalInputs As Worksheet
    Dim inputpond As Worksheet
    Dim Lookup As Worksheet
    Dim ws As Worksheet
    
    Dim WinterRise As Double
    Dim SeepFactor1 As Double
    Dim SeepFactor2 As Double
    Dim wintervol As Double
    Dim compstat As Double
    Dim SaltVol As Double
    Dim totalVolume As Double
    Dim targetLevel As Double
    Dim pondArea As Double
    
    Dim currmonth As String
    
    Dim totIntake As Double
    Dim totRetTails As Double
    Dim totEvap As Double
    Dim totPrecip As Double
    Dim totSeepage As Double
    Dim totDischarge As Double
    Dim totHarvest As Double
    Dim initPondMass As Double
    Dim finPondMass As Double
    
    totIntake = 0: totRetTails = 0: totEvap = 0: totPrecip = 0: totSeepage = 0:
    totDischarge = 0: totHarvest = 0: initPondMass = 0: finPondMass = 0
    
    'MC variables
    Dim mc_run As Integer
    Dim mc_run_step As Integer
    Dim mc_rand_row As Integer
    Dim MC As Worksheet
    Dim SSP As Double
    Dim saltSSP As Double
    Dim mcCDBV As Double
    Dim mcCHS As Double
    Dim mcCHSA As Double
    Dim mcCDBVA As Double
    Dim mcPC2CDBVA As Double
    Dim mcAHSKCL As Double
    Dim mcSP22Carn As Double
    Dim mcPC2Carn As Double
    Dim mcMinPC2K As Double
    Dim mcMinEvapCF As Double
    
    'Dim harvestedamt As Double    'Removing salt pond harvested amount sensitivity from this version since currently variable values based on 2020 dreging patterns have been assigned for each pond.
    Set MC = Application.Worksheets("MC")
    Dim cpv As Integer
    Dim mcruns As Integer
    
    Dim carn_max_limit As Double
    carn_max_limit = 500000000
    Dim DeadSeaRevisedIntake As Double
    Dim DeadSeaRevisedIntakePCT As Double
    Dim DeadSeaRevisedIntakeInc As Double
    Dim DeadSeaRevisedIntakeDec As Double
    DeadSeaRevisedIntakeInc = 1.02 '2% increase in adjusting Dead Sea intake
    DeadSeaRevisedIntakeDec = 0.98 '2% decrease in adjusting Dead Sea intake
    Dim ResetCount As Integer
    ResetCount = 0
    Dim ResetBool As Boolean
    ResetBool = True 'Set to False if not adjusting Dead Sea intake
       
    'call to totalsteps() function to obtain # of timesteps
    totsteps = totalsteps()
    numponds = 22
    trtotsteps = totsteps * numponds
    
    'Worksheet shortcuts
    '-------------------
    Set StartPond = Application.Worksheets("SP0B")
    Set InputOLI = Application.Worksheets("OLI-Input")
    Set OutputOLI = Application.Worksheets("OLI-Calc")
    Set GlobalInputs = Application.Worksheets("Global Inputs")
    Set Lookup = Application.Worksheets("Look Up Tables")
    
    'Winter Global Tuning Parameters
    '-------------------------------
    WinterRise = GlobalInputs.Range("winterrise").Value
    'SeepFactor1 = GlobalInputs.Range("SeepFactor1").Value
    'SeepFactor2 = GlobalInputs.Range("SeepFactor2").Value
    
    'Time Factor (Daily/Weekly/Monthly) for Feed Streams
    '---------------------------------------------------
    TimeStepFactor = GlobalInputs.Range("time").Value
    If TimeStepFactor = "Weekly" Then
        TimeFactor = 31 / 7
    ElseIf TimeStepFactor = "Daily" Then
        TimeFactor = 31
    Else
        TimeFactor = 1
    End If
    
    'Discharge limits outside winter months
    MinDisch = GlobalInputs.Range("MinDischarge").Value * 31 / TimeFactor
    MaxDisch = GlobalInputs.Range("MaxDischarge").Value * 31 / TimeFactor
    
    'Scrub PPT from OLI solid species output
    OutputOLI.Unprotect
    OutputOLI.Range("olioutputsol").Replace What:="PPT", Replacement:="", _
    SearchOrder:=xlByRows, MatchCase:=True
    OutputOLI.Protect
    
    'Collections and Dictionary Setup
    '--------------------------------
    'BrineDict_c matches item = column # with key = Brine species name in Global Inputs worksheet
    Dim BrineDict_c As New Scripting.Dictionary
    Dim b As Range
    For Each b In GlobalInputs.Range("brineinflow")
        BrineDict_c.Add Item:=b.Column, Key:=b.Value
    Next
    
    'DeadSeaMon_r matches item = row # with key = month name in Global Inputs worksheet
    Dim DeadSeaMon_r As New Collection
    Dim dsm As Range
    For Each dsm In GlobalInputs.Range("dsmonth")
        DeadSeaMon_r.Add Item:=dsm.Row, Key:=dsm.Value
    Next
    
    'ReturnMon_r matches item = row # with key = month name in Global Inputs worksheet
    Dim ReturnMon_r As New Collection
    Dim rm As Range
    For Each rm In GlobalInputs.Range("rbmonth")
        ReturnMon_r.Add Item:=rm.Row, Key:=rm.Value
    Next
    
    'solidsDict_c matches item = column # with key = solid species name in SP0B worksheet
    Dim solidsDict_c As New Scripting.Dictionary
    Dim s As Range
    For Each s In StartPond.Range("solidrange")
        solidsDict_c.Add Item:=s.Column, Key:=s.Value
    Next
    
    'HarvDict_c matches item = column # with key = solid species name in SP0B worksheet
    Dim HarvDict_c As New Scripting.Dictionary
    Dim hv As Range
    For Each hv In StartPond.Range("harvrange")
        HarvDict_c.Add Item:=hv.Column, Key:=hv.Value
    Next
    
    'liquidsColl_c matches item = column # with key = liquid species name in SP0B worksheet
    Dim liquidsColl_c As New Collection
    Dim l As Range
    For Each l In StartPond.Range("liqrange")
        liquidsColl_c.Add Item:=l.Column, Key:=l.Value
    Next
    
    'olioutliqColl_r matches item = row # with key = liquid species name in OLI-Calc worksheet
    Dim olioutliqColl_r As New Collection
    Dim ol As Range
    For Each ol In OutputOLI.Range("olioutputliq")
        olioutliqColl_r.Add Item:=ol.Row, Key:=ol.Value
    Next
    
    'olioutsolColl_r matches item = row # with key = solid species name in OLI-Calc worksheet
    Dim olioutsolColl_r As New Collection
    Dim os As Range
    For Each os In OutputOLI.Range("olioutputsol")
        olioutsolColl_r.Add Item:=os.Row, Key:=os.Value
    Next
    
    'species_r matches item = row # with key = liquid species name in Look Up Tables worksheet
    Dim species_r As New Collection
    Dim sr As Range
    For Each sr In Application.Worksheets("Look Up Tables").Range("lookupelem")
        species_r.Add Item:=sr.Row, Key:=sr.Value
    Next
    
    'inputColl_r matches item = row # with key = species name in OLI-Input worksheet
    'outflowDict matches item = 0 (for initialization) with key = species name in OLI-Input worksheet
    Dim inputColl_r As New Collection
    Dim outflowDict As New Scripting.Dictionary
    Dim pc2of1Dict As New Scripting.Dictionary
    Dim pc2of2Dict As New Scripting.Dictionary
    Dim c6ofDict As New Scripting.Dictionary
    Dim o As Range
    For Each o In InputOLI.Range("specrange")
        inputColl_r.Add Item:=o.Row, Key:=o.Value
        outflowDict.Add Item:=0, Key:=o.Value
        pc2of1Dict.Add Item:=0, Key:=o.Value
        pc2of2Dict.Add Item:=0, Key:=o.Value
        c6ofDict.Add Item:=0, Key:=o.Value
    Next
    
    Dim PC2toC3Split As New Collection
    PC2toC3Split.Add Item:=1, Key:="Apr"
    PC2toC3Split.Add Item:=0.6, Key:="May"
    PC2toC3Split.Add Item:=0.63, Key:="Jun"
    PC2toC3Split.Add Item:=0.71, Key:="Jul"
    PC2toC3Split.Add Item:=0.65, Key:="Aug"
    PC2toC3Split.Add Item:=0.67, Key:="Sep"
    PC2toC3Split.Add Item:=0.8, Key:="Oct"
    PC2toC3Split.Add Item:=0.65, Key:="Nov"
    PC2toC3Split.Add Item:=0.71, Key:="Dec"
    PC2toC3Split.Add Item:=0.71, Key:="Jan"
    PC2toC3Split.Add Item:=0.71, Key:="Feb"
    PC2toC3Split.Add Item:=0.71, Key:="Mar"
    
    mc_rand_row = MC.Range("mcRands").Row
    mcruns = MC.Range("mcruns").Value
    
    'Export MC output to text file to keep results in case Excel crashes
    'Open "C:\APC\MC" & Format(Now, "ddmmmyyyy-hhnn") & ".csv" For Output As #1
    
    For mc_run = 1 To mcruns
    
        mc_run_step = MC.Range("mcInputs").Row + mc_run
        MC.Rows(mc_rand_row).Calculate
        
        SeepFactor1 = MC.Range("mcSeepFactor1").Value
        SeepFactor2 = MC.Range("mcSeepFactor2").Value
        saltSSP = MC.Range("mcsaltSSP").Value
        SSP = MC.Range("mcslurrysolidspercent").Value
        mcMinEvapCF = MC.Range("mcMinEvapCF").Value
        'harvestedamt = MC.Range("mcharvestedamt").Value
        
        MC.Range("mcRands").Copy
        MC.Cells(mc_run_step, 2).PasteSpecial Paste:=xlPasteValues
        
        MC.Range("intakerandCF").Copy
        GlobalInputs.Cells(5, 11).PasteSpecial Paste:=xlPasteValues, Transpose:=True
        Range("MinEvapCF").Value = mcMinEvapCF
        
        With GlobalInputs
            .Range("SeepFactor1").Value = SeepFactor1
            .Range("SeepFactor2").Value = SeepFactor2
            .Range("saltSSP").Value = saltSSP
            .Range("slurrysolidspercent").Value = SSP
            '.Range("harvestedamt").Value = harvestedamt
        End With
        
        For cpv = 0 To 21
            MC.Range(MC.Cells(7, 19 + (cpv * 12)).Address(), MC.Cells(7, 30 + (cpv * 12)).Address()).Copy
            Lookup.Cells(74, 20 + cpv).PasteSpecial Paste:=xlPasteValues, Transpose:=True
        Next cpv
        
        'Iteration through timesteps (outer loop)
        '----------------------------------------
        TimeStep = StartPond.Range("liqrange").Row + 1
        For Step = 1 To totsteps
        ResetCount = 0
StepReset:
            ContinueRun = True           'stop condition reset
            Set inputpond = StartPond    'start at top pond
            pondint = 1
            
            'Iteration through ponds (inner loop)
            '------------------------------------
            Do While ContinueRun
                    
                CurrentWS = inputpond.name
                inputpond.Rows(TimeStep).Calculate
                currmonth = inputpond.Cells(TimeStep, Application.Worksheets("SP0A").Range("month").Column).Value
                StartVol = inputpond.Range("currvolume").Value
                MaxDepth = inputpond.Range("G8").Value
                
                If CurrentWS = "SP0B" Then
                    GlobalInputs.Rows(DeadSeaMon_r(currmonth)).Calculate
                    DeadSeaRevisedIntake = GlobalInputs.Cells(DeadSeaMon_r(currmonth), 10).Value
                    DeadSeaRevisedIntakePCT = GlobalInputs.Cells(DeadSeaMon_r(currmonth), 11).Value
                    inputpond.Cells(TimeStep, 1).Value = DeadSeaRevisedIntake
                End If
                
                targetLevel = inputpond.Cells(TimeStep, inputpond.Range("TVolpondOut").Column - 1).Value
                pondArea = inputpond.Range("G5").Value
                'MaxVol = inputpond.Range("MaxVolume").Value
                MaxVol = targetLevel * pondArea
                
                If ((currmonth = "Dec" Or currmonth = "Jan" Or currmonth = "Feb" Or currmonth = "Mar") And (Left(CurrentWS, 1) = "S") And WinterRise <> 0) Then
                    wintervol = MaxVol * (MaxDepth + WinterRise) / MaxDepth
                    'MaxDepth = MaxDepth + WinterRise
                    If currmonth = "Dec" Then
                        MaxVol = wintervol
                    ElseIf currmonth = "Jan" Then
                        MaxVol = MaxVol + (0.75 * (wintervol - MaxVol))
                    ElseIf currmonth = "Feb" Then
                        MaxVol = MaxVol + (0.5 * (wintervol - MaxVol))
                    Else
                        MaxVol = MaxVol + (0.25 * (wintervol - MaxVol))
                    End If
                End If
                
                Application.DisplayStatusBar = True
                elapsedtime = currTime - Now()
                trstep = ((Step - 1) * numponds) + pondint
                timeremain = (elapsedtime / (trstep / trtotsteps)) - elapsedtime
                compstat = Round((trstep / trtotsteps) * 100, 1)
                Application.StatusBar = "Calculating MC Run #" & mc_run & " Time Step " & Step & " of " & totsteps & " (" & compstat & _
                "%) - " & currmonth & " - " & CurrentWS & " - Elapsed Time: " & Format(elapsedtime, "HH:MM:SS") & _
                " (Estimated Time Remaining = " & Format(timeremain, "HH:MM:SS") & ")"
                
                
                'H2O balance and input
                '---------------------
                With inputpond
                    C_nim1 = .Cells(TimeStep, liquidsColl_c("H2O")).Value                       'current timestep moles
                    P_ni = .Cells(TimeStep, .Range("precipmol").Column).Value                   'precipitation moles
                    E_ni = .Cells(TimeStep, .Range("evapmol").Column).Value                     'evaporation moles
                    R_ni = .Cells(TimeStep, .Range("runoffvol").Column).Value / 18.02 * 1000000 'runoff moles
                End With
                
                totPrecip = totPrecip + (P_ni * 18.02)
                totEvap = totEvap + (E_ni * 18.02)
                
                If Step = 1 Then
                    initPondMass = initPondMass + (C_nim1 * 18.02)
                End If
                
                D_nm1i = outflowDict("H2O")
                D_nm1i_2 = 0
                
                Select Case CurrentWS
                    Case "SP0B"
                        D_nm1i = GlobalInputs.Cells(DeadSeaMon_r(currmonth), BrineDict_c("H2O")).Value / TimeFactor
                        totIntake = totIntake + (D_nm1i * 18.02)
                    Case "C4"
                        D_nm1i = GlobalInputs.Cells(ReturnMon_r(currmonth), BrineDict_c("H2O")).Value / TimeFactor
                        totRetTails = totRetTails + (D_nm1i * 18.02)
                    Case "C3"
                        D_nm1i_2 = pc2of1Dict("H2O")
                    Case "C7"
                        D_nm1i_2 = c6ofDict("H2O")
                    Case "C8"
                        D_nm1i = pc2of2Dict("H2O")
                End Select
                
                OLIBal = C_nim1 - E_ni + P_ni + D_nm1i + D_nm1i_2 + R_ni                'mole balance
                If OLIBal < 0 Then                                                      'exit sub if negative H2O input
                    'MsgBox "Input water balance is negative. End of run." & vbCrLf & _
                    '"Calculations stopped at pond " & CurrentWS & ". Timestep " & step & _
                    '" of " & totsteps & "."
                    GoTo endline
                End If
                InputOLI.Cells(16, 3).Value = OLIBal                                    'output to OLI-Input
        
                'Remaining species balance and input
                '-----------------------------------
                Dim Row As Range
                For Each Row In InputOLI.Range("oliinput")
                    species = InputOLI.Cells(Row.Row, 1).Value                          'get species name
                    C_nim1 = inputpond.Cells(TimeStep, liquidsColl_c(species)).Value    'get species current value
                    If solidsDict_c.Exists(species) Then
                        S_nim1 = inputpond.Cells(TimeStep, solidsDict_c(species)).Value 'get solid species if exists
                    Else
                        S_nim1 = 0
                    End If
                    
                    D_nm1i = outflowDict(species)
                    D_nm1i_2 = 0
                    
                    If Step = 1 Then
                        initPondMass = initPondMass + (C_nim1 * Lookup.Cells(species_r(species), 3).Value)
                    End If
                    
                    Select Case CurrentWS
                        Case "SP0B", "C4"
                            If CurrentWS = "SP0B" And BrineDict_c.Exists(species) Then
                                D_nm1i = GlobalInputs.Cells(DeadSeaMon_r(currmonth), BrineDict_c(species)).Value / TimeFactor
                                totIntake = totIntake + (D_nm1i * Lookup.Cells(species_r(species), 3).Value)
                            ElseIf CurrentWS = "C4" And BrineDict_c.Exists(species) Then
                                D_nm1i = GlobalInputs.Cells(ReturnMon_r(currmonth), BrineDict_c(species)).Value / TimeFactor
                                totRetTails = totRetTails + (D_nm1i * Lookup.Cells(species_r(species), 3).Value)
                            Else
                                D_nm1i = 0
                            End If
                        Case "C3"
                            D_nm1i_2 = pc2of1Dict(species)
                        Case "C7"
                            D_nm1i_2 = c6ofDict(species)
                        Case "C8"
                            D_nm1i = pc2of2Dict(species)
                    End Select
                                   
                    OLIBal = C_nim1 + S_nim1 + D_nm1i + D_nm1i_2                        'mole balance
                    InputOLI.Cells(Row.Row, 3).Value = OLIBal                           'output to OLI-Input
                Next
                
                'temperature transfer from SP0A to OLI-Input
                InputOLI.Cells(4, 3).Value = inputpond.Cells(TimeStep, inputpond.Range("temperature").Column).Value
                
                'run OLI Engine using Calculate() subroutine
                ThisWorkbook.Calculate
                
                'parameter transfer for current timestep
                '---------------------------------------
                With OutputOLI
                    SolidMass = .Range("solidmass").Value / 1000000                                          'solid mass in tonne
                    WaterMass = .Range("watermass").Value * 18.01528                                         'water mass for TDS calculation
                    LiquidMass = .Range("LiquidMass").Value                                                    'total mass for TDS calculation
                    SaltVol = .Range("saltvolume").Value                                                     'salt/solid volume
                    LiqPH = .Range("liqPH").Value                                                            'pH value
                    
                    'calculate TDS
                    If LiquidMass <> 0 Then                          'V_lnij <> 0 Then
                        TDS = (LiquidMass - WaterMass) / LiquidMass  'V_lnij
                    Else                                                                                        'exit sub if no liquid after OLI call
                        MsgBox "No liquid species present in OLI run. End of run." & vbCrLf & _
                        "Calculations stopped at pond " & CurrentWS & ". Timestep " & Step & _
                        " of " & totsteps & "."
                        GoTo endline
                    End If
                    
                    'LiqDensity = .Range("liqdensity").Value / 1000                                           'liquid density
                    LiqDensity = 1.2651846 * TDS + 0.888759231
                    
                    V_lnij = LiquidMass / LiqDensity / 1000000  ' .Range("liqvolume").Value
                    V_tnij = V_lnij + SaltVol                   ' .Range("totvolume").Value
                    If TimeStep <> (inputpond.Range("liqrange").Row + 1) Then
                        V_tnim1 = inputpond.Cells(TimeStep - 1, inputpond.Range("TVolpondOut").Column).Value 'previous volume
                    Else
                        V_tnim1 = StartVol                                                                   'starting volume for first timestep
                    End If
                End With
                
                With inputpond
                    HarvestRatio = .Cells(TimeStep, .Range("harvest").Column).Value                         'get harvest ratio
                    .Cells(TimeStep, .Range("TDSpondOut").Column).Value = TDS                               'output TDS
                    .Cells(TimeStep, .Range("DenpondOut").Column).Value = LiqDensity                        'output liquid density
                    .Cells(TimeStep, .Range("TPSpondOut").Column).Value = SolidMass                         'output total solid mass
                    .Cells(TimeStep, .Range("SaltvolpondOut").Column).Value = SaltVol * (1 - HarvestRatio)  'unharvested salt volume
                    .Cells(TimeStep, .Range("ph").Column).Value = LiqPH                                     'output pH
                    .Rows(TimeStep).Calculate                                                               'calculate row for Q
                    Q = .Cells(TimeStep, .Range("seepage").Column).Value                                    'get seepage flowrate
                End With
                           
                'calculate seepage and outflow ratios
                SeepageRatio = Q / V_lnij
                
                Select Case CurrentWS
                    Case "SP0A", "SP1-1", "SP1-2", "SP1-3", "SP1-4", "SP1-5", "SP1-6"
                        SeepageRatio = SeepageRatio * SeepFactor1
                    Case Else
                        SeepageRatio = SeepageRatio * SeepFactor2
                End Select
    
                If ((V_lnij - Q) < MaxVol) Then
                    DischargeVol = MinDisch
                Else
                    DischargeVol = V_lnij - Q - MaxVol
                    If DischargeVol > MaxDisch Then
                        DischargeVol = MaxDisch
                    End If
                End If
                
                'If (currmonth = "Dec" And (Left(CurrentWS, 1) = "S") And DischargeVol <> 0) Then
                    'DischargeVol = DischargeVol - ((wintervol - StartVol) / TimeFactor)
                'End If
    
                OutflowRatio = DischargeVol / V_lnij
                
                totalVolume = V_tnij - Q - DischargeVol
                
                With inputpond
                    .Cells(TimeStep, .Range("TVolpondOut").Column).Value = totalVolume                                'output total volume
                    .Cells(TimeStep, .Range("LVolpondOut").Column).Value = totalVolume - SaltVol * (1 - HarvestRatio) 'output liquid volume
                    .Cells(TimeStep, .Range("DispondOut").Column).Value = DischargeVol
                End With
                
                'solid species balance and output
                '--------------------------------
                Dim outputsol As Range
                For Each outputsol In inputpond.Range("solidrange")
                    species = inputpond.Cells(inputpond.Range("solidrange").Row, outputsol.Column).Value  'get solid species name
                    S_nij = OutputOLI.Cells(olioutsolColl_r(species), 2).Value                            'get solid species value
                    
                    If ResetCount < 5 And ResetBool And species = "KMGCL3.6H2O" And (Left(CurrentWS, 1) = "S" Or CurrentWS = "PC2") And Not (currmonth = "Apr") Then
                        If (Left(CurrentWS, 1) = "S" And S_nij > 0) Or (CurrentWS = "PC2" And S_nij > carn_max_limit) Then
                            GlobalInputs.Cells(DeadSeaMon_r(currmonth), 11).Value = DeadSeaRevisedIntakePCT * DeadSeaRevisedIntakeInc
                            ResetCount = ResetCount + 1
                            GoTo StepReset
                        ElseIf CurrentWS = "PC2" And S_nij < 1 Then
                            GlobalInputs.Cells(DeadSeaMon_r(currmonth), 11).Value = DeadSeaRevisedIntakePCT * DeadSeaRevisedIntakeDec
                            ResetCount = ResetCount + 1
                            GoTo StepReset
                        End If
                    End If
                    
                    If HarvestRatio > 0 Then
                        H_ni = S_nij * HarvestRatio                                                       'solid species harvested
                        totHarvest = totHarvest + (H_ni * Lookup.Cells(species_r(species), 3).Value)
                        S_ni = S_nij - H_ni                                                               'new solid species adjusted per harvest
                        inputpond.Cells(TimeStep, HarvDict_c(species)).Value = H_ni                       'Harvested solid species on same row as the result of this timestep
                    Else
                        S_ni = S_nij
                    End If
                    
                    If Step = totsteps Then
                        finPondMass = finPondMass + (S_ni * Lookup.Cells(species_r(species), 3).Value)
                    End If
    
                    inputpond.Cells(TimeStep + 1, outputsol.Column).Value = S_ni  'output solid species on next row as initial solids for next timestep
                Next
                
                'end the loop to populate final parameters for the last timestep
                'If step = totsteps Then
                    'GoTo LastLine
                'End If
                
                'liquid species balance and output
                '---------------------------------
                Dim outputliq As Range
                For Each outputliq In inputpond.Range("liqrange")
                    species = inputpond.Cells(inputpond.Range("liqrange").Row, outputliq.Column).Value   'get liquid species name
                    C_nij = OutputOLI.Cells(olioutliqColl_r(species), 2).Value                           'get liquid species value
                    F_ni = C_nij * SeepageRatio                                                          'liquid species seepage value
                    totSeepage = totSeepage + (F_ni * Lookup.Cells(species_r(species), 3).Value)
                    D_ni = C_nij * OutflowRatio                                                          'liquid species outflow value
                    outflowDict(species) = D_ni                                                          'add outflow to dictionary for next linked pond
                    
                    If CurrentWS = "PC2" Then
                        PC2toC3 = PC2toC3Split(inputpond.Cells(TimeStep, 3).Value)
                        pc2of1Dict(species) = D_ni * PC2toC3
                        pc2of2Dict(species) = D_ni * (1 - PC2toC3)
                    ElseIf CurrentWS = "C6" Then
                        c6ofDict(species) = D_ni
                    ElseIf CurrentWS = "C7" Then
                        totDischarge = totDischarge + (D_ni * Lookup.Cells(species_r(species), 3).Value)
                    End If
                    
                    If C_nij <> 0 Then
                        C_ni = C_nij - F_ni - D_ni                                                        'mole balance
                    Else
                        C_ni = 0
                    End If
                    
                    inputpond.Cells(TimeStep + 1, outputliq.Column).Value = C_ni                          'output liquid species
                    
                    If Step = totsteps Then
                        finPondMass = finPondMass + (C_ni * Lookup.Cells(species_r(species), 3).Value)
                    End If
                Next
                
    'LastLine:
            
                inputpond.Activate
                Call element1(HarvDict_c, liquidsColl_c, species_r, Step)
                
                Select Case CurrentWS
                    Case "C7"
                        ContinueRun = False          'stop condition for current timestep
                    Case "PC2"
                        Set inputpond = Sheets("C4")
                    Case "C4"
                        Set inputpond = Sheets("C3")
                    Case "C6"
                        Set inputpond = Sheets("C8")
                    Case "C11"
                        Set inputpond = Sheets("C7")
                    Case Else
                        If inputpond.Range("outflow").Value = "End" Then
                            ContinueRun = False
                        Else
                            Set inputpond = Sheets(inputpond.Range("outflow").Value)
                        End If
                End Select
                pondint = pondint + 1
            
            Loop
            
            'increment timestep
            TimeStep = TimeStep + 1
            
        Next Step
        
        GlobalInputs.Range("totIntake").Value = totIntake
        GlobalInputs.Range("totRetTails").Value = totRetTails
        GlobalInputs.Range("totEvap").Value = totEvap
        GlobalInputs.Range("totPrecip").Value = totPrecip
        GlobalInputs.Range("totSeepage").Value = totSeepage
        GlobalInputs.Range("totDischarge").Value = totDischarge
        GlobalInputs.Range("totHarvest").Value = totHarvest
        GlobalInputs.Range("initPondMass").Value = initPondMass
        GlobalInputs.Range("finPondMass").Value = finPondMass
        
        'For Each ws In ActiveWorkbook.Worksheets
        '    If ws.Range("A1").Value = "Evaporation Pond Design Spreadsheet" Then
        '        ws.Activate
        '        Call elementAll(HarvDict_c, liquidsColl_c, species_r, totsteps)
        '    End If
        'Next
        
        With MC
            ActiveWorkbook.Calculate
            mcCDBV = .Range("mcCDBV").Value
            mcCHS = .Range("mcCHS").Value
            mcCHSA = .Range("mcCHSA").Value
            mcCDBVA = .Range("mcCDBVA").Value
            mcPC2CDBVA = .Range("mcPC2CDBVA").Value
            mcAHSKCL = .Range("mcAHSKCL").Value
            mcSP22Carn = .Range("mcSP22Carn").Value
            mcPC2Carn = .Range("mcPC2Carn").Value
            mcMinPC2K = .Range("mcMinPC2K").Value
            .Cells(mc_run_step, .Range("mcCDBV").Column).Value = mcCDBV
            .Cells(mc_run_step, .Range("mcCHS").Column).Value = mcCHS
            .Cells(mc_run_step, .Range("mcCHSA").Column).Value = mcCHSA
            .Cells(mc_run_step, .Range("mcCDBVA").Column).Value = mcCDBVA
            .Cells(mc_run_step, .Range("mcPC2CDBVA").Column).Value = mcPC2CDBVA
            .Cells(mc_run_step, .Range("mcAHSKCL").Column).Value = mcAHSKCL
            .Cells(mc_run_step, .Range("mcSP22Carn").Column).Value = mcSP22Carn
            .Cells(mc_run_step, .Range("mcPC2Carn").Column).Value = mcPC2Carn
            .Cells(mc_run_step, .Range("mcMinPC2K").Column).Value = mcMinPC2K
            
            'Dim col As Integer
            'For col = 2 To .Range("mcMinPC2K").Column           'UPDATE THIS IF NEW CALCULATED PARAMETER IS ADDED
                'Print #1, .Cells(mc_run_step, col), ", ";
            'Next
            'Print #1,
            
        End With
    
endline:
    Next mc_run
    
    'Close '#1

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.Calculation = xlAutomatic

GlobalInputs.Activate

End Sub





