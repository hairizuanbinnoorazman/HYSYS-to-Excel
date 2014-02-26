Option Explicit

'This module contains only the code from ExtractValuesQuick subroutine
'This subroutine is a slightly modified version of the initial extract values subroutine in module 1
'Msgbox statement prompting for user to adjust feed tray is removed from this code.
'This is a less optimized case for hysys

Sub ExtractValuesQuick()
    'Part 1
    'Obtain required settings
    Worksheets("Settings").Activate
    Dim hysysFileName As String
    Dim columnStart As Integer
    Dim totalCases As Integer
    Dim fullHysysFileName As String
    With Worksheets("Settings")
        hysysFileName = .Cells(2, 2).Value
        totalCases = .Cells(3, 2).Value
    End With
    fullHysysFileName = Application.ActiveWorkbook.Path & "\" & hysysFileName & ".hsc"
    
    'Part 2
    'Start HYSYS Instance
    Dim hyApp As HYSYS.Application
    Dim hyCase As SimulationCase
    
    Set hyApp = CreateObject("HYSYS.Application")
    hyApp.SimulationCases.Open (fullHysysFileName)
    hyApp.Visible = True
    Set hyCase = GetObject(fullHysysFileName, "HYSYS.SimulationCase")
    hyCase.Visible = True
     
    'Part 3a
    'Defining the loop constraints variables for transfer
    Dim loopCounter As Integer
    loopCounter = 1
    Dim PFRTemp1 As Double
    Dim PFRTemp2 As Double
    Dim PFRVol1 As Double
    Dim PFRVol2 As Double
    PFRTemp1 = 0
    PFRTemp2 = 0
    PFRVol1 = 0
    PFRVol2 = 0
    Dim intLoop As Integer
    intLoop = 0
    
    'Defining the objects for Tolerance
    Dim InitialToleranceADJ1 As Double
    Dim InitialToleranceADJ2 As Double
    Dim FinalToleranceADJ1 As Double
    Dim FinalToleranceADJ2 As Double
    Dim hysysAdjust As HYSYS.AdjustOp
    
    With Worksheets("Settings")
        InitialToleranceADJ1 = .Cells(2, 6)
        InitialToleranceADJ2 = .Cells(3, 6)
        FinalToleranceADJ1 = .Cells(4, 6)
        FinalToleranceADJ2 = .Cells(5, 6)
    End With

    'Part 3b
    'Defining the objects for part 4
    Dim hySS As SpreadsheetOp
    Dim hyCellPFRTemp1 As SpreadsheetCell
    Dim hyCellPFRTemp2 As SpreadsheetCell
    Dim hyCellPFRVol1 As SpreadsheetCell
    Dim hyCellPFRVol2 As SpreadsheetCell
    
    Dim hyCellTransfer As SpreadsheetCell
     
    Dim OPCost(14) As Variant
    Dim PFR3Phase(3) As Variant
    Dim Furnace(5) As Variant
    Dim Tower1(8) As Variant
    Dim Tower2(8) As Variant
    Dim Pump(4) As Variant
    
    Dim Lights(10) As Variant
    Dim Vent(10) As Variant
    Dim R1In(10) As Variant
    Dim R2Out(10) As Variant
    

    
    'Pull values from Excel Spreadsheet and put into HYSYS
    'Loop begins at this point
    Do While loopCounter <= totalCases
        
        'Part 4a
        'Extract required values from Excel Spreadsheet
        Worksheets("Summary Sheet").Activate
        With Worksheets("Summary Sheet")
            PFRTemp1 = .Cells(loopCounter + 1, 2)
            PFRTemp2 = .Cells(loopCounter + 1, 3)
            PFRVol1 = .Cells(loopCounter + 1, 4)
            PFRVol2 = .Cells(loopCounter + 1, 5)
        End With
        
        'Part 4b
        'Transfer values into hysys
        
        Set hySS = hyCase.Flowsheet.Operations("spreadsheetop").Item("ADJUST SHEET")
        
        hyCase.Solver.CanSolve = False
        Set hyCellPFRTemp1 = hySS.Cell(1, 1)
        Set hyCellPFRTemp2 = hySS.Cell(1, 2)
        Set hyCellPFRVol1 = hySS.Cell(3, 1)
        Set hyCellPFRVol2 = hySS.Cell(3, 2)
        
        hyCellPFRTemp1.CellValue = PFRTemp1
        hyCellPFRTemp2.CellValue = PFRTemp2
        hyCellPFRVol1.CellValue = PFRVol1
        hyCellPFRVol2.CellValue = PFRVol2
        
        'Initial Adjustment to quickly obtain results from hysys
        'Adjustment can be put the settings page
        Set hysysAdjust = hyCase.Flowsheet.Operations.Item("Adj-1")
        hysysAdjust.ToleranceValue = InitialToleranceADJ1
        Set hysysAdjust = hyCase.Flowsheet.Operations.Item("Adj-2")
        hysysAdjust.ToleranceValue = InitialToleranceADJ2
        hyCase.Solver.CanSolve = True
        
        hyCase.Save
                
        'Final Adjustment to obtain final results from hysys
        'Adjustment can be put into the settings page
        hyCase.Solver.CanSolve = False
        Set hysysAdjust = hyCase.Flowsheet.Operations.Item("Adj-1")
        hysysAdjust.ToleranceValue = FinalToleranceADJ1
        Set hysysAdjust = hyCase.Flowsheet.Operations.Item("Adj-2")
        hysysAdjust.ToleranceValue = FinalToleranceADJ2
        hyCase.Solver.CanSolve = True
                
        hyCase.Save
        
        
        'Part 4c
        'Pushing for alert
        'This is a modified version of the code in module 1
        'This code allows for rapid data collection in order to collect data without waiting for user interaction
        'to adjust the feed tray.
        
        
        'Part 4d
        'Extracting values from hysys
        
        Set hySS = hyCase.Flowsheet.Operations("spreadsheetop").Item("COSTING SHEET")
        
        'For Operating Costs
        intLoop = 2
        Do While intLoop <= 16
            Set hyCellTransfer = hySS.Cell(intLoop, 2)
            OPCost(intLoop - 2) = hyCellTransfer.CellValue
            intLoop = intLoop + 1
        Loop
        
        'For PFR and 3 Phase
        intLoop = 1
        Do While intLoop <= 4
            Set hyCellTransfer = hySS.Cell(intLoop, 6)
            PFR3Phase(intLoop - 1) = hyCellTransfer.CellValue
            intLoop = intLoop + 1
        Loop
        
        'For Furnace
        intLoop = 1
        Do While intLoop <= 6
            Set hyCellTransfer = hySS.Cell(intLoop, 10)
            Furnace(intLoop - 1) = hyCellTransfer.CellValue
            intLoop = intLoop + 1
        Loop
        
        'For Tower1
        intLoop = 0
        Do While intLoop <= 8
            Set hyCellTransfer = hySS.Cell(intLoop, 14)
            Tower1(intLoop) = hyCellTransfer.CellValue
            intLoop = intLoop + 1
        Loop
        
        'For Tower2
        intLoop = 0
        Do While intLoop <= 8
            Set hyCellTransfer = hySS.Cell(intLoop, 18)
            Tower2(intLoop) = hyCellTransfer.CellValue
            intLoop = intLoop + 1
        Loop
        
        'For Pump
        intLoop = 1
        Do While intLoop <= 5
            Set hyCellTransfer = hySS.Cell(intLoop, 22)
            Pump(intLoop - 1) = hyCellTransfer.CellValue
            intLoop = intLoop + 1
        Loop
        
        'Part 4e
        'Extracting the values for flow rates
        'Need to reference new spreadsheet
        hyCase.Save
        Set hySS = hyCase.Flowsheet.Operations("spreadsheetop").Item("FLOWS SHEET")
        
        'For Lights
        intLoop = 1
        Do While intLoop <= 11
            Set hyCellTransfer = hySS.Cell(1, intLoop)
            Lights(intLoop - 1) = hyCellTransfer.CellValue
            intLoop = intLoop + 1
        Loop
        
        'For Vents
        intLoop = 1
        Do While intLoop <= 11
            Set hyCellTransfer = hySS.Cell(2, intLoop)
            Vent(intLoop - 1) = hyCellTransfer.CellValue
            intLoop = intLoop + 1
        Loop
        
        'For R1In
        intLoop = 1
        Do While intLoop <= 11
            Set hyCellTransfer = hySS.Cell(3, intLoop)
            R1In(intLoop - 1) = hyCellTransfer.CellValue
            intLoop = intLoop + 1
        Loop
        
        'For R2Out
        intLoop = 1
        Do While intLoop <= 11
            Set hyCellTransfer = hySS.Cell(4, intLoop)
            R2Out(intLoop - 1) = hyCellTransfer.CellValue
            intLoop = intLoop + 1
        Loop
        
        'Part 4f
        'Putting values into respective worksheets in hysys
        
        'For Operating Cost
        intLoop = 0
        Do While intLoop <= 14
            Worksheets("Operating Costs").Activate
            With Worksheets("Operating Costs")
                .Cells(loopCounter + 1, 2 + intLoop).Value = OPCost(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'For PFR 3 Phase
        intLoop = 0
        Do While intLoop <= 3
            Worksheets("PFR_3Phase_Capital").Activate
            With Worksheets("PFR_3Phase_Capital")
                .Cells(loopCounter + 1, 2 + intLoop).Value = PFR3Phase(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'For Furnace
        intLoop = 0
        Do While intLoop <= 5
            Worksheets("Furnace_HE_Capital").Activate
            With Worksheets("Furnace_HE_Capital")
                .Cells(loopCounter + 1, 2 + intLoop).Value = Furnace(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'For Tower1
        intLoop = 0
        Do While intLoop <= 8
            Worksheets("Tower1_Capital").Activate
            With Worksheets("Tower1_Capital")
                .Cells(loopCounter + 1, 2 + intLoop).Value = Tower1(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'For Tower2
        intLoop = 0
        Do While intLoop <= 8
            Worksheets("Tower2_Capital").Activate
            With Worksheets("Tower2_Capital")
                .Cells(loopCounter + 1, 2 + intLoop).Value = Tower2(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'Pump
        intLoop = 0
        Do While intLoop <= 4
            Worksheets("Pump_Compressor_Capital").Activate
            With Worksheets("Pump_Compressor_Capital")
                .Cells(loopCounter + 1, 2 + intLoop).Value = Pump(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'Part 4g
        'Putting values into respective worksheets in hysys
        
        'Lights
        intLoop = 0
        Do While intLoop <= 10
            Worksheets("Lights").Activate
            With Worksheets("Lights")
                .Cells(loopCounter + 1, 2 + intLoop).Value = Lights(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'Vent
        intLoop = 0
        Do While intLoop <= 10
            Worksheets("Vent").Activate
            With Worksheets("Vent")
                .Cells(loopCounter + 1, 2 + intLoop).Value = Vent(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'Lights
        intLoop = 0
        Do While intLoop <= 10
            Worksheets("R1In").Activate
            With Worksheets("R1In")
                .Cells(loopCounter + 1, 2 + intLoop).Value = R1In(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'Lights
        intLoop = 0
        Do While intLoop <= 10
            Worksheets("R2Out").Activate
            With Worksheets("R2Out")
                .Cells(loopCounter + 1, 2 + intLoop).Value = R2Out(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        
        'Increase loopCounter by 1 to go to the next row
        loopCounter = loopCounter + 1
    Loop
End Sub

