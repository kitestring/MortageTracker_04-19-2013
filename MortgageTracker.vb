Attribute VB_Name = "MortgageTracker"
Option Explicit

Sub IsolateMortgateCategories()

Dim DataAnalysisWkbk As Workbook
Dim RawDataWkbk As Workbook
Dim textFile As Variant

Const StartingColumnLetter As String = "G"
Const StartingColumnNumber As Integer = 7
Const StartingRowNumber As Integer = 8
Dim Row As Integer
Dim Column As Integer
Dim CategoryName(2) As String
Dim NoOfTransactions(2) As Integer
Dim MaxTransactions As Integer
Dim Transactions() As String
Dim c As Integer
Dim t As Integer
Dim d As Integer
Dim Continue As Boolean
Dim HitCounter As Integer

Dim CategoryColumnNumbers(2) As Integer
Const CategoryStartingRowNumber As Integer = 5
Const InterestDeltaEqun As String = "=U4-U5"

    Application.ScreenUpdating = False

'Define analysis workbook
    Set DataAnalysisWkbk = ActiveWorkbook
    
'Prompt user for text file to be data mined
    'textFile = Application.GetOpenFilename(Title:="Select text File")
    textFile = Application.GetOpenFilename("Text Files (*.txt), *.txt", Title:="Select text File")
    If VarType(textFile) = vbBoolean Then Exit Sub
    
'Open and define raw data workbook
    Workbooks.OpenText Filename:=textFile, Origin _
        :=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote _
        , ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:= _
        False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1) _
        , Array(3, 1)), TrailingMinusNumbers:=True
    Set RawDataWkbk = ActiveWorkbook
    
'Delete old raw data set
    DataAnalysisWkbk.Activate
    Sheets("Raw Data").Select
    Cells.Select
    Selection.ClearContents
    
'Copy new raw data to clipboard
    RawDataWkbk.Activate
    Cells.Select
    Selection.Copy
    
'Paste new raw data into analysis workbook and close RawDataWkbk
    DataAnalysisWkbk.Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    RawDataWkbk.Close

'Mine Raw data
    Application.ScreenUpdating = False

    CategoryName(0) = "Mortgage:Normal Payment"
    CategoryName(1) = "Mortgage:Additional Principal"
    CategoryName(2) = "Mortgage:Interest"
    
    For c = 0 To 2
        NoOfTransactions(c) = CountCategoryTransactions(CategoryName(c), StartingColumnLetter, StartingColumnNumber, c)
    Next c
    
    Cells(4, StartingColumnNumber - 1).Value = "=MAX(F1:F3)-1"
    MaxTransactions = Cells(4, StartingColumnNumber - 1).Value
    Range(Cells(1, StartingColumnNumber - 1), Cells(4, StartingColumnNumber - 1)).ClearContents
    
    ReDim Transactions(2, MaxTransactions, 1) As String
    
    For c = 0 To 2
        HitCounter = 0
        Continue = True
        Row = StartingRowNumber
        Column = StartingColumnNumber
        Do While Continue = True
            If Cells(Row, Column).Value = CategoryName(c) Then
                Transactions(c, HitCounter, 0) = Cells(Row, 2).Value
                Transactions(c, HitCounter, 1) = Cells(Row, 10).Value
                HitCounter = HitCounter + 1
            End If
            Row = Row + 1
            If HitCounter = NoOfTransactions(c) Then Continue = False
        Loop
    Next c
    
'Dump parsed raw data into categories sheet
    Sheets("Categories").Select
    
    CategoryColumnNumbers(0) = 6
    CategoryColumnNumbers(1) = 13
    CategoryColumnNumbers(2) = 20
    
    For c = 0 To 2
        For t = 0 To NoOfTransactions(c) - 1
            Row = CategoryStartingRowNumber + t
            For d = 0 To 1
                Column = CategoryColumnNumbers(c) + d
                Cells(Row, Column).Value = Transactions(c, t, d)
            Next d
        Next t
    Next c
    
    Column = Column + 1
    Cells(CategoryStartingRowNumber, Column).Value = InterestDeltaEqun
    Cells(CategoryStartingRowNumber, Column).Select
    Selection.AutoFill Destination:=Range(Cells(CategoryStartingRowNumber, Column), Cells(Row, Column))
    
'Begin Year Statistics Section
Dim RowRange(1) As Integer
Dim YearColumnNumber(2) As Integer
Dim YearColumnLetter(2) As String
Const YearDataColumn As Integer = 39
Dim YearRow As Integer
Dim YearCheck As String
Dim CL As String
Dim Equation As String
Dim MonthNumber As Integer
Dim TotalMounthCount As Integer
Dim YearCount As Integer
    
    YearCount = CountNumberOfYears(NoOfTransactions(0))
    YearColumnNumber(0) = 40
    YearColumnNumber(1) = 42
    YearColumnNumber(2) = 44
    YearColumnLetter(0) = DetermineColumnLetter(YearColumnNumber(0))
    YearColumnLetter(1) = DetermineColumnLetter(YearColumnNumber(1))
    YearColumnLetter(2) = DetermineColumnLetter(YearColumnNumber(2))
    YearRow = CategoryStartingRowNumber
    
    For c = 0 To 2
        Row = CategoryStartingRowNumber
        RowRange(0) = Row
        Column = CategoryColumnNumbers(c)
        CL = DetermineColumnLetter(CategoryColumnNumbers(c) + 1)
        YearCheck = Year(Cells(Row, Column).Value)
        YearRow = CategoryStartingRowNumber
        
        Do While IsEmpty(Cells(Row, Column)) = False
            If Year(Cells(Row, Column).Value) <> YearCheck Then
                RowRange(1) = Row - 1
                Equation = Function_ContinuiousRange("SUM", CL, RowRange(0), RowRange(1))
                Cells(YearRow, YearColumnNumber(c)).Value = Equation
                Cells(YearRow, YearColumnNumber(c) + 1).Value = "=" & YearColumnLetter(c) & YearRow & "/12"
                Cells(YearRow, YearDataColumn).Value = YearCheck
                
                YearCheck = Year(Cells(Row, Column).Value)
                RowRange(0) = Row
                YearRow = YearRow + 1
                
            End If
            Row = Row + 1
        Loop
        
        Row = Row - 1
        MonthNumber = Month(Cells(Row, Column).Value)
        RowRange(1) = Row
        Equation = Function_ContinuiousRange("SUM", CL, RowRange(0), RowRange(1))
        Cells(YearRow, YearColumnNumber(c)).Value = Equation
        Cells(YearRow, YearColumnNumber(c) + 1).Value = "=" & YearColumnLetter(c) & YearRow & "/" & MonthNumber
        Cells(YearRow, YearDataColumn).Value = YearCheck
        
        Equation = Function_ContinuiousRange("SUM", YearColumnLetter(c), CategoryStartingRowNumber, YearRow)
        Cells(YearRow + 2, YearColumnNumber(c)).Value = Equation
        
        TotalMounthCount = ((YearCount - 1) * 12) + MonthNumber
        Cells(YearRow + 2, YearColumnNumber(c) + 1).Value = "=" & YearColumnLetter(c) & (YearRow + 2) & "/" & TotalMounthCount
        Cells(YearRow + 2, YearDataColumn).Value = "Total"
    
    Next c
    
'Create Net & Payment Distributation equations
    
    Row = CategoryStartingRowNumber + YearCount - 1
    Range("AT5").Value = "=SUM(AN5,AP5,AR5)"
    Range("AU5").Value = "=$BP$3-SUM($AT$5:AT5)"
    Range("AV5").Value = "=SUM(AN5,AP5)"
    Range("AW5").Value = "=SUM(AO5,AQ5)"
    Range("AX5").Value = "=(SUM($AO5,$AS5)/SUM($AO5,$AQ5))"
    Range("AY5").Value = "=($AQ5/SUM($AO5,$AQ5))"
    Range("AZ5").Value = "=(ABS($AS5)/SUM($AO5,$AQ5))"
    
    Range("AT5:AZ5").Copy
    
    Range("AT6:AZ" & Row + 2).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Cells(YearRow + 2, YearColumnNumber(c - 1) + 3) = "=AU" & Row
    
    Range("BP5").Value = "=BP4+AT" & (YearRow + 2)
    
    Range(Cells(YearRow + 1, YearDataColumn), Cells(YearRow + 1, YearColumnNumber(c - 1) + 8)).ClearContents
    
    Cells(2, 1).Select
    
'Delete raw data text file
    Call RemoveRawDataFile(textFile)

End Sub

Private Function CountCategoryTransactions(ByVal Category As String, ByVal ColumnLetter As String, ByVal ColumnNumber As Integer, ByVal Iteration As Byte) As Integer
Dim Equation As String
Dim Column As Integer
Dim Row As Integer

    Column = ColumnNumber - 1
    Row = 1 + Iteration
    Equation = "=COUNTIF(" & ColumnLetter & ":" & ColumnLetter & "," & Chr(34) & Category & Chr(34) & ")"
    Cells(Row, Column).Value = Equation
    CountCategoryTransactions = Cells(Row, Column).Value
    
End Function

Private Function Function_ContinuiousRange(ByVal FunctionName As String, ByVal ColumnLetter As String, ByVal TopRow As Integer, ByVal BottomRow As Integer) As String

    Function_ContinuiousRange = "=" & FunctionName & "(" & ColumnLetter & TopRow & ":" & ColumnLetter & BottomRow & ")"

End Function

Private Function CountNumberOfYears(ByVal NormalPaymentEntries As Integer) As Integer

Dim LastRow As Integer

    Application.DisplayAlerts = False
    Sheets("Raw Data").Select
    
    LastRow = NormalPaymentEntries + 4
    
    Range("Z5").Value = "=YEAR(Categories!F5)"
    Range("Z5").Select
    
    Selection.AutoFill Destination:=Range("Z5:Z" & LastRow), Type:=xlFillDefault
    
    Range("Z5:Z" & LastRow).Select
    
    ActiveSheet.Range("Z5:Z" & LastRow).RemoveDuplicates Columns:=1, Header:=xlNo
    
    Range("Y1").Value = "=COUNT(Z:Z)"
    CountNumberOfYears = Range("Y1").Value
    Range("Y1:Z" & LastRow).ClearContents
    Cells(1, 1).Select
    
    Sheets("Categories").Select
    Application.DisplayAlerts = True
    
End Function

Private Sub RemoveRawDataFile(ByVal FilePath As String)

Dim FSO

'Set Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
'Delete File
    FSO.DeleteFile FilePath, True

End Sub

Private Function DetermineColumnLetter(ByVal ColumnNumber As Integer) As String
Dim ColumnLetterOutput(156) As String

    ColumnLetterOutput(1) = "A"
    ColumnLetterOutput(2) = "B"
    ColumnLetterOutput(3) = "C"
    ColumnLetterOutput(4) = "D"
    ColumnLetterOutput(5) = "E"
    ColumnLetterOutput(6) = "F"
    ColumnLetterOutput(7) = "G"
    ColumnLetterOutput(8) = "H"
    ColumnLetterOutput(9) = "I"
    ColumnLetterOutput(10) = "J"
    ColumnLetterOutput(11) = "K"
    ColumnLetterOutput(12) = "L"
    ColumnLetterOutput(13) = "M"
    ColumnLetterOutput(14) = "N"
    ColumnLetterOutput(15) = "O"
    ColumnLetterOutput(16) = "P"
    ColumnLetterOutput(17) = "Q"
    ColumnLetterOutput(18) = "R"
    ColumnLetterOutput(19) = "S"
    ColumnLetterOutput(20) = "T"
    ColumnLetterOutput(21) = "U"
    ColumnLetterOutput(22) = "V"
    ColumnLetterOutput(23) = "W"
    ColumnLetterOutput(24) = "X"
    ColumnLetterOutput(25) = "Y"
    ColumnLetterOutput(26) = "Z"
    ColumnLetterOutput(27) = "AA"
    ColumnLetterOutput(28) = "AB"
    ColumnLetterOutput(29) = "AC"
    ColumnLetterOutput(30) = "AD"
    ColumnLetterOutput(31) = "AE"
    ColumnLetterOutput(32) = "AF"
    ColumnLetterOutput(33) = "AG"
    ColumnLetterOutput(34) = "AH"
    ColumnLetterOutput(35) = "AI"
    ColumnLetterOutput(36) = "AJ"
    ColumnLetterOutput(37) = "AK"
    ColumnLetterOutput(38) = "AL"
    ColumnLetterOutput(39) = "AM"
    ColumnLetterOutput(40) = "AN"
    ColumnLetterOutput(41) = "AO"
    ColumnLetterOutput(42) = "AP"
    ColumnLetterOutput(43) = "AQ"
    ColumnLetterOutput(44) = "AR"
    ColumnLetterOutput(45) = "AS"
    ColumnLetterOutput(46) = "AT"
    ColumnLetterOutput(47) = "AU"
    ColumnLetterOutput(48) = "AV"
    ColumnLetterOutput(49) = "AW"
    ColumnLetterOutput(50) = "AX"
    ColumnLetterOutput(51) = "AY"
    ColumnLetterOutput(52) = "AZ"
    ColumnLetterOutput(53) = "BA"
    ColumnLetterOutput(54) = "BB"
    ColumnLetterOutput(55) = "BC"
    ColumnLetterOutput(56) = "BD"
    ColumnLetterOutput(57) = "BE"
    ColumnLetterOutput(58) = "BF"
    ColumnLetterOutput(59) = "BG"
    ColumnLetterOutput(60) = "BH"
    ColumnLetterOutput(61) = "BI"
    ColumnLetterOutput(62) = "BJ"
    ColumnLetterOutput(63) = "BK"
    ColumnLetterOutput(64) = "BL"
    ColumnLetterOutput(65) = "BM"
    ColumnLetterOutput(66) = "BN"
    ColumnLetterOutput(67) = "BO"
    ColumnLetterOutput(68) = "BP"
    ColumnLetterOutput(69) = "BQ"
    ColumnLetterOutput(70) = "BR"
    ColumnLetterOutput(71) = "BS"
    ColumnLetterOutput(72) = "BT"
    ColumnLetterOutput(73) = "BU"
    ColumnLetterOutput(74) = "BV"
    ColumnLetterOutput(75) = "BW"
    ColumnLetterOutput(76) = "BX"
    ColumnLetterOutput(77) = "BY"
    ColumnLetterOutput(78) = "BZ"
    ColumnLetterOutput(79) = "CA"
    ColumnLetterOutput(80) = "CB"
    ColumnLetterOutput(81) = "CC"
    ColumnLetterOutput(82) = "CD"
    ColumnLetterOutput(83) = "CE"
    ColumnLetterOutput(84) = "CF"
    ColumnLetterOutput(85) = "CG"
    ColumnLetterOutput(86) = "CH"
    ColumnLetterOutput(87) = "CI"
    ColumnLetterOutput(88) = "CJ"
    ColumnLetterOutput(89) = "CK"
    ColumnLetterOutput(90) = "CL"
    ColumnLetterOutput(91) = "CM"
    ColumnLetterOutput(92) = "CN"
    ColumnLetterOutput(93) = "CO"
    ColumnLetterOutput(94) = "CP"
    ColumnLetterOutput(95) = "CQ"
    ColumnLetterOutput(96) = "CR"
    ColumnLetterOutput(97) = "CS"
    ColumnLetterOutput(98) = "CT"
    ColumnLetterOutput(99) = "CU"
    ColumnLetterOutput(100) = "CV"
    ColumnLetterOutput(101) = "CW"
    ColumnLetterOutput(102) = "CX"
    ColumnLetterOutput(103) = "CY"
    ColumnLetterOutput(104) = "CZ"
    ColumnLetterOutput(105) = "DA"
    ColumnLetterOutput(106) = "DB"
    ColumnLetterOutput(107) = "DC"
    ColumnLetterOutput(108) = "DD"
    ColumnLetterOutput(109) = "DE"
    ColumnLetterOutput(110) = "DF"
    ColumnLetterOutput(111) = "DG"
    ColumnLetterOutput(112) = "DH"
    ColumnLetterOutput(113) = "DI"
    ColumnLetterOutput(114) = "DJ"
    ColumnLetterOutput(115) = "DK"
    ColumnLetterOutput(116) = "DL"
    ColumnLetterOutput(117) = "DM"
    ColumnLetterOutput(118) = "DN"
    ColumnLetterOutput(119) = "DO"
    ColumnLetterOutput(120) = "DP"
    ColumnLetterOutput(121) = "DQ"
    ColumnLetterOutput(122) = "DR"
    ColumnLetterOutput(123) = "DS"
    ColumnLetterOutput(124) = "DT"
    ColumnLetterOutput(125) = "DU"
    ColumnLetterOutput(126) = "DV"
    ColumnLetterOutput(127) = "DW"
    ColumnLetterOutput(128) = "DX"
    ColumnLetterOutput(129) = "DY"
    ColumnLetterOutput(130) = "DZ"
    ColumnLetterOutput(131) = "EA"
    ColumnLetterOutput(132) = "EB"
    ColumnLetterOutput(133) = "EC"
    ColumnLetterOutput(134) = "ED"
    ColumnLetterOutput(135) = "EE"
    ColumnLetterOutput(136) = "EF"
    ColumnLetterOutput(137) = "EG"
    ColumnLetterOutput(138) = "EH"
    ColumnLetterOutput(139) = "EI"
    ColumnLetterOutput(140) = "EJ"
    ColumnLetterOutput(141) = "EK"
    ColumnLetterOutput(142) = "EL"
    ColumnLetterOutput(143) = "EM"
    ColumnLetterOutput(144) = "EN"
    ColumnLetterOutput(145) = "EO"
    ColumnLetterOutput(146) = "EP"
    ColumnLetterOutput(147) = "EQ"
    ColumnLetterOutput(148) = "ER"
    ColumnLetterOutput(149) = "ES"
    ColumnLetterOutput(150) = "ET"
    ColumnLetterOutput(151) = "EU"
    ColumnLetterOutput(152) = "EV"
    ColumnLetterOutput(153) = "EW"
    ColumnLetterOutput(154) = "EX"
    ColumnLetterOutput(155) = "EY"
    ColumnLetterOutput(156) = "EZ"

    DetermineColumnLetter = ColumnLetterOutput(ColumnNumber)
End Function
