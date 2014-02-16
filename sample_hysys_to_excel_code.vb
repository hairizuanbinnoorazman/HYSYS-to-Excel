Sub ExtractValues()
    'Part 1
    'Obtain required settings
    Worksheets("Settings").Activate
    Dim fileLocation As String
    Dim columnStart As Integer
    Dim totalCases As Integer
    With Worksheets("Settings")
        fileLocation = .Cells(2, 2).Value
        totalCases = .Cells(3, 2).Value
    End With
    
    'Part 2
    'Start HYSYS Instance
    Dim hyApp As HYSYS.Application
    Dim hyCase As SimulationCase

    Set hyApp = CreateObject("HYSYS.Application")
    hyApp.SimulationCases.Open (fileLocation)
    hyApp.Visible = True
    Set hyCase = GetObject(fileLocation, "HYSYS.SimulationCase")
    hyCase.Visible = True
     
    'Part 3a
    'Defining the loop constraints variables for transfer
    Dim loopCounter As Integer
    loopCounter = 1
    Dim PFRTemp1 As Integer
    Dim PFRTemp2 As Integer
    Dim PFRVol1 As Integer
    Dim PFRVol2 As Integer
    PFRTemp1 = 0
    PFRTemp2 = 0
    PFRVol1 = 0
    PFRVol2 = 0
    Dim intLoop As Integer
    intLoop = 0

    'Part 3b
    'Defining the objects for part 4
    Dim hySS As SpreadsheetOp
    Dim hyCellPFRTemp1 As SpreadsheetCell
    Dim hyCellPFRTemp2 As SpreadsheetCell
    Dim hyCellPFRVol1 As SpreadsheetCell
    Dim hyCellPFRVol2 As SpreadsheetCell
    
    Dim hyCellTransfer As SpreadsheetCell
     
    Dim OPCost(15) As Variant
    Dim PFR3Phase(3) As Variant
    Dim Furnace(6) As Variant
    Dim Tower1(8) As Variant
    Dim Tower2(8) As Variant
    Dim Pump(4) As Variant
    
    
    'Pull values from Excel Spreadsheet and put into HYSYS
    'Loop begins at this point
    Do While loopCounter <= totalCases
        
        'Part 4a
        'Extract required values from Excel Spreadsheet
        Worksheets("Summary Sheet").Activate
        With Worksheets("Summary Sheet")
            PFRTemp1 = .Cells(loopCounter + 1, 1)
            PFRTemp2 = .Cells(loopCounter + 1, 2)
            PFRVol1 = .Cells(loopCounter + 1, 3)
            PFRVol2 = .Cells(loopCounter + 1, 4)
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
        hyCase.Solver.CanSolve = True
        
        hyCase.Save
                
                
        'Part 4c
        'Pushing for alert
        Beep
        MsgBox ("Adjust feed tray column in HYSYS. Click OK to continue")
        
        
        'Part 4d
        'Extracting values from hysys
        hyCase.Save
        Set hySS = hyCase.Flowsheet.Operations("spreadsheetop").Item("COSTING SHEET")
        
        'For Operating Costs
        intLoop = 1
        Do While intLoop <= 16
            Set hyCellTransfer = hySS.Cell(intLoop, 2)
            OPCost(intLoop - 1) = hyCellTransfer.CellValue
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
        Do While intLoop <= 7
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
        'Putting values into respective worksheets in hysys
        'Final part of the loop
        
        'For Operating Cost
        intLoop = 0
        Do While intLoop <= 15
            Worksheets("Operating Costs").Activate
            With Worksheets("Operating Costs")
                .Cells(loopCounter + 1, 5 + intLoop).Value = OPCost(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'For PFR 3 Phase
        intLoop = 0
        Do While intLoop <= 3
            Worksheets("PFR_3Phase_Capital").Activate
            With Worksheets("PFR_3Phase_Capital")
                .Cells(loopCounter + 1, 5 + intLoop).Value = PFR3Phase(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'For Furnace
        intLoop = 0
        Do While intLoop <= 6
            Worksheets("Furnace_HE_Capital").Activate
            With Worksheets("Furnace_HE_Capital")
                .Cells(loopCounter + 1, 5 + intLoop).Value = Furnace(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'For Tower1
        intLoop = 0
        Do While intLoop <= 8
            Worksheets("Tower1_Capital").Activate
            With Worksheets("Tower1_Capital")
                .Cells(loopCounter + 1, 5 + intLoop).Value = Tower1(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'For Tower2
        intLoop = 0
        Do While intLoop <= 8
            Worksheets("Tower2_Capital").Activate
            With Worksheets("Tower2_Capital")
                .Cells(loopCounter + 1, 5 + intLoop).Value = Tower2(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'Pump
        intLoop = 0
        Do While intLoop <= 4
            Worksheets("Pump_Compressor_Capital").Activate
            With Worksheets("Pump_Compressor_Capital")
                .Cells(loopCounter + 1, 5 + intLoop).Value = Pump(intLoop)
            End With
            intLoop = intLoop + 1
        Loop
        
        'Increase loopCounter by 1 to go to the next row
        loopCounter = loopCounter + 1
    Loop
End Sub
