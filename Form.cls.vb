VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cbo_AcctNumbers_Change()
    Range("D6").Value = "=VLOOKUP(" & cbo_AcctNumbers.Value & ",C74:G100,5,FALSE)"
End Sub

Private Sub cbo_Calculate_Click()
    'Calculate
    
    If Range("L4").Value <> "" Then
        
        '5% VAT
        Range("L6").Value = Format(Range("L4").Value / 1.12 * 0.05, "standard")
        Range("P6").Value = Range("L4").Value / 1.12 * 0.05
        
        '2% VAT
        Range("L7").Value = Format(Range("L4").Value / 1.12 * 0.02, "standard")
        Range("P7").Value = Range("L4").Value / 1.12 * 0.02
        
        'Total Tax
        Range("L9").Value = Range("L6").Value + Range("L7").Value
        
        'Balance
        Range("L11").Value = Range("L4").Value - Range("L9").Value
        
    Else
        MsgBox "Please enter the charge amount.", vbCritical
    End If
End Sub

Private Sub cbo_ItemAdd_Click()
    If chk_Includes.Value = True Then
        'Add Item
    If Range("O18").Value = 1 Then
        'Cellphone Number
        Range("B22").Value = Range("D6").Value
        'Account Number
        Range("E22").Value = Range("E70").Value
        'Current charges
        Range("H22").Value = Range("L4").Value
        '5% VAT
        Range("J22").Value = Range("L6").Value
        '2% VAT
        Range("L22").Value = Range("L7").Value
        'Total Tax
        Range("N22").Value = Range("L9").Value
        'Balance
        Range("P22").Value = Range("L11").Value
        'Bill Period
        Range("R22").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 2 Then
        'Cellphone Number
        Range("B23").Value = Range("D6").Value
        'Account Number
        Range("E23").Value = Range("E70").Value
        'Current charges
        Range("H23").Value = Range("L4").Value
        '5% VAT
        Range("J23").Value = Range("L6").Value
        '2% VAT
        Range("L23").Value = Range("L7").Value
        'Total Tax
        Range("N23").Value = Range("L9").Value
        'Balance
        Range("P23").Value = Range("L11").Value
        Range("R23").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 3 Then
        'Cellphone Number
        Range("B24").Value = Range("D6").Value
        'Account Number
        Range("E24").Value = Range("E70").Value
        'Current charges
        Range("H24").Value = Range("L4").Value
        '5% VAT
        Range("J24").Value = Range("L6").Value
        '2% VAT
        Range("L24").Value = Range("L7").Value
        'Total Tax
        Range("N24").Value = Range("L9").Value
        'Balance
        Range("P24").Value = Range("L11").Value
        Range("R24").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 4 Then
        'Cellphone Number
        Range("B25").Value = Range("D6").Value
        'Account Number
        Range("E25").Value = Range("E70").Value
        'Current charges
        Range("H25").Value = Range("L4").Value
        '5% VAT
        Range("J25").Value = Range("L6").Value
        '2% VAT
        Range("L25").Value = Range("L7").Value
        'Total Tax
        Range("N25").Value = Range("L9").Value
        'Balance
        Range("P25").Value = Range("L11").Value
        Range("R25").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 5 Then
        'Cellphone Number
        Range("B26").Value = Range("D6").Value
        'Account Number
        Range("E26").Value = Range("E70").Value
        'Current charges
        Range("H26").Value = Range("L4").Value
        '5% VAT
        Range("J26").Value = Range("L6").Value
        '2% VAT
        Range("L26").Value = Range("L7").Value
        'Total Tax
        Range("N26").Value = Range("L9").Value
        'Balance
        Range("P26").Value = Range("L11").Value
        Range("R26").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 6 Then
        'Cellphone Number
        Range("B27").Value = Range("D6").Value
        'Account Number
        Range("E27").Value = Range("E70").Value
        'Current charges
        Range("H27").Value = Range("L4").Value
        '5% VAT
        Range("J27").Value = Range("L6").Value
        '2% VAT
        Range("L27").Value = Range("L7").Value
        'Total Tax
        Range("N27").Value = Range("L9").Value
        'Balance
        Range("P27").Value = Range("L11").Value
        Range("R27").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 7 Then
        'Cellphone Number
        Range("B28").Value = Range("D6").Value
        'Account Number
        Range("E28").Value = Range("E70").Value
        'Current charges
        Range("H28").Value = Range("L4").Value
        '5% VAT
        Range("J28").Value = Range("L6").Value
        '2% VAT
        Range("L28").Value = Range("L7").Value
        'Total Tax
        Range("N28").Value = Range("L9").Value
        'Balance
        Range("P28").Value = Range("L11").Value
        Range("R28").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 8 Then
        'Cellphone Number
        Range("B29").Value = Range("D6").Value
        'Account Number
        Range("E29").Value = Range("E70").Value
        'Current charges
        Range("H29").Value = Range("L4").Value
        '5% VAT
        Range("J29").Value = Range("L6").Value
        '2% VAT
        Range("L29").Value = Range("L7").Value
        'Total Tax
        Range("N29").Value = Range("L9").Value
        'Balance
        Range("P29").Value = Range("L11").Value
        Range("R29").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    Else
        'Add Item
    If Range("O18").Value = 1 Then
        'Current charges
        Range("H22").Value = Range("L4").Value
        '5% VAT
        Range("J22").Value = Range("L6").Value
        '2% VAT
        Range("L22").Value = Range("L7").Value
        'Total Tax
        Range("N22").Value = Range("L9").Value
        'Balance
        Range("P22").Value = Range("L11").Value
        Range("R22").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 2 Then
        'Current charges
        Range("H23").Value = Range("L4").Value
        '5% VAT
        Range("J23").Value = Range("L6").Value
        '2% VAT
        Range("L23").Value = Range("L7").Value
        'Total Tax
        Range("N23").Value = Range("L9").Value
        'Balance
        Range("P23").Value = Range("L11").Value
        Range("R23").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 3 Then
        'Current charges
        Range("H24").Value = Range("L4").Value
        '5% VAT
        Range("J24").Value = Range("L6").Value
        '2% VAT
        Range("L24").Value = Range("L7").Value
        'Total Tax
        Range("N24").Value = Range("L9").Value
        'Balance
        Range("P24").Value = Range("L11").Value
        Range("R24").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 4 Then
        'Current charges
        Range("H25").Value = Range("L4").Value
        '5% VAT
        Range("J25").Value = Range("L6").Value
        '2% VAT
        Range("L25").Value = Range("L7").Value
        'Total Tax
        Range("N25").Value = Range("L9").Value
        'Balance
        Range("P25").Value = Range("L11").Value
        Range("R25").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 5 Then
        'Current charges
        Range("H26").Value = Range("L4").Value
        '5% VAT
        Range("J26").Value = Range("L6").Value
        '2% VAT
        Range("L26").Value = Range("L7").Value
        'Total Tax
        Range("N26").Value = Range("L9").Value
        'Balance
        Range("P26").Value = Range("L11").Value
        Range("R26").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 6 Then
        'Current charges
        Range("H27").Value = Range("L4").Value
        '5% VAT
        Range("J27").Value = Range("L6").Value
        '2% VAT
        Range("L27").Value = Range("L7").Value
        'Total Tax
        Range("N27").Value = Range("L9").Value
        'Balance
        Range("P27").Value = Range("L11").Value
        Range("R27").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 7 Then
        'Current charges
        Range("H28").Value = Range("L4").Value
        '5% VAT
        Range("J28").Value = Range("L6").Value
        '2% VAT
        Range("L28").Value = Range("L7").Value
        'Total Tax
        Range("N28").Value = Range("L9").Value
        'Balance
        Range("P28").Value = Range("L11").Value
        Range("R28").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    
    If Range("O18").Value = 8 Then
        'Current charges
        Range("H29").Value = Range("L4").Value
        '5% VAT
        Range("J29").Value = Range("L6").Value
        '2% VAT
        Range("L29").Value = Range("L7").Value
        'Total Tax
        Range("N29").Value = Range("L9").Value
        'Balance
        Range("P29").Value = Range("L11").Value
        Range("R29").Value = Format(Range("L13").Value, "dd mmm yyyy") & " - " & Format(Range("P13").Value, "dd mmm yyyy")
    End If
    End If
    
    'Increase item number for next item to be added
    If Range("O18").Value >= 8 Then
        Range("O18").Value = 1
    Else
        Range("O18").Value = Range("O18").Value + 1
    End If
End Sub

Private Sub cbo_Recalculate_Click()
    'Re Calculate
    'Total Tax
    Range("L9").Value = Range("L6").Value + Range("L7").Value
        
    'Balance
    Range("L11").Value = Range("L4").Value - Range("L9").Value
    
End Sub

Private Sub cmd_ItemsClear_Click()
    Range("B21:R29").ClearContents
    'Back to start (no. 1, first item)
    Range("O18").Value = 1
End Sub

Private Sub cmd_Log_Click()
    If Sheets("Settings").Range("E18").Value = "" Then
        MsgBox "Log Directory is not set on 'Settings' sheet.", vbOKCancel, "Log Directory not set"
    Else
         
        'On Error GoTo errmsg
         
        Dim file As String
        Dim textfile As Integer
        
        file = Sheets("Settings").Range("E18").Value & "\" & "GLOBE TELECOM [" & Sheets("FORM").Range("V31").Value & "] - " & Format(Sheets("FORM").Range("E36").Value, "standard") & ".txt"
        
        textfile = FreeFile
        
        Open file For Output As textfile
        
        ' Print GLOBE NAME, ADDRESS, AMOUNT
        Print #textfile, "Voucher Name: " & Sheets("_temp.Voucher").Range("C12").Value
        Print #textfile, "Payee Address: " & Sheets("_temp.Voucher").Range("C14").Value
        Print #textfile, "Gross Amount: " & Format(Range("E36").Value, "Standard") & vbNewLine
        
        ' Print Date, Time, CLERK
        Print #textfile, "Date: " & Format(Now, "mmmm-dd-yyyy")
        Print #textfile, "Time: " & Format(Now, "hh:mm am/pm")
        Print #textfile, "Clerk: " & Sheets("Settings").Range("E14").Value & vbNewLine
        
        ' Print GROSS AMOUNT, GROSS AMOUNT IN WORDS, NET AMOUNT
        Print #textfile, "Gross Amount: " & Format(Range("E36").Value, "Standard")
        Print #textfile, "In words amount: " & Range("E38").Value
        Print #textfile, "Net Amount: " & Format(Range("R38").Value & vbNewLine, "Standard")
        
        'Print:
            ' Total Charges
            ' Total 5% VAT
            ' Total 2% VAT
            ' Total TAX
        Print #textfile, "Total Charges: " & Format(Range("H31").Value, "Standard")
        Print #textfile, "Total 5% VAT: " & Format(Range("J31").Value, "Standard")
        Print #textfile, "Total 2% VAT: " & Format(Range("L31").Value, "Standard")
        Print #textfile, "Total TAX: " & Format(Range("N31").Value, "Standard") & vbNewLine & vbNewLine
        
        'Print per Phone Number Charges
        'print Header First
        Print #textfile, "CELL NO." & vbTab & vbTab & "ACC. NO." & vbTab & vbTab & "CUR. CHRGES." & vbTab & vbTab & "5% VAT" & vbTab & vbTab & "2% VAT" & vbTab & vbTab & "TAX" & vbTab & vbTab & "BAL." & vbTab & vbTab & "PERIOD" & vbNewLine
        
        'Print per item
        Print #textfile, Range("B22").Value & vbTab & Range("E22").Value & vbTab & Format(Range("H22").Value, "Standard") & vbTab & vbTab & vbTab & Format(Range("J22").Value, "standard") & vbTab & vbTab & Format(Range("L22").Value, "Standard") & vbTab & vbTab & Format(Range("N22").Value, "Standard") & vbTab & vbTab & Format(Range("P22").Value, "Standard") & vbTab & vbTab & Range("R22").Value
        Print #textfile, Range("B23").Value & vbTab & Range("E23").Value & vbTab & Format(Range("H23").Value, "Standard") & vbTab & vbTab & vbTab & Format(Range("J23").Value, "standard") & vbTab & vbTab & Format(Range("L23").Value, "Standard") & vbTab & vbTab & Format(Range("N23").Value, "Standard") & vbTab & vbTab & Format(Range("P23").Value, "Standard") & vbTab & vbTab & Range("R23").Value
        Print #textfile, Range("B24").Value & vbTab & Range("E24").Value & vbTab & Format(Range("H24").Value, "Standard") & vbTab & vbTab & vbTab & Format(Range("J24").Value, "standard") & vbTab & vbTab & Format(Range("L24").Value, "Standard") & vbTab & vbTab & Format(Range("N24").Value, "Standard") & vbTab & vbTab & Format(Range("P24").Value, "Standard") & vbTab & vbTab & Range("R24").Value
        Print #textfile, Range("B25").Value & vbTab & Range("E25").Value & vbTab & Format(Range("H25").Value, "Standard") & vbTab & vbTab & vbTab & Format(Range("J25").Value, "standard") & vbTab & vbTab & Format(Range("L25").Value, "Standard") & vbTab & vbTab & Format(Range("N25").Value, "Standard") & vbTab & vbTab & Format(Range("P25").Value, "Standard") & vbTab & vbTab & Range("R25").Value
        Print #textfile, Range("B26").Value & vbTab & Range("E26").Value & vbTab & Format(Range("H26").Value, "Standard") & vbTab & vbTab & vbTab & Format(Range("J26").Value, "standard") & vbTab & vbTab & Format(Range("L26").Value, "Standard") & vbTab & vbTab & Format(Range("N26").Value, "Standard") & vbTab & vbTab & Format(Range("P26").Value, "Standard") & vbTab & vbTab & Range("R26").Value
        Print #textfile, Range("B27").Value & vbTab & Range("E27").Value & vbTab & Format(Range("H27").Value, "Standard") & vbTab & vbTab & vbTab & Format(Range("J27").Value, "standard") & vbTab & vbTab & Format(Range("L27").Value, "Standard") & vbTab & vbTab & Format(Range("N27").Value, "Standard") & vbTab & vbTab & Format(Range("P27").Value, "Standard") & vbTab & vbTab & Range("R27").Value
        Print #textfile, Range("B28").Value & vbTab & Range("E28").Value & vbTab & Format(Range("H28").Value, "Standard") & vbTab & vbTab & vbTab & Format(Range("J28").Value, "standard") & vbTab & vbTab & Format(Range("L28").Value, "Standard") & vbTab & vbTab & Format(Range("N28").Value, "Standard") & vbTab & vbTab & Format(Range("P28").Value, "Standard") & vbTab & vbTab & Range("R28").Value
        Print #textfile, Range("B29").Value & vbTab & Range("E29").Value & vbTab & Format(Range("H29").Value, "Standard") & vbTab & vbTab & vbTab & Format(Range("J29").Value, "standard") & vbTab & vbTab & Format(Range("L29").Value, "Standard") & vbTab & vbTab & Format(Range("N29").Value, "Standard") & vbTab & vbTab & Format(Range("P29").Value, "Standard") & vbTab & vbTab & Range("R29").Value
        
        Close textfile
        
        'save to Bill Record sheets
        Dim i As Integer
        
        Sheets("BILL_RECORD").Range("B5").EntireRow.Insert
        Sheets("BILL_RECORD").Range("B5").EntireRow.Insert
        
        For i = 29 To 22 Step -1
            If Range("B" & i).Value <> "" Then
                Sheets("BILL_RECORD").Range("B5").EntireRow.Insert
                
                Sheets("BILL_RECORD").Range("B5").Value = Sheets("FORM").Range("E" & i).Value
                Sheets("BILL_RECORD").Range("E5").Value = Sheets("FORM").Range("B" & i).Value
                Sheets("BILL_RECORD").Range("H5").Value = Sheets("FORM").Range("W" & i).Value
                Sheets("BILL_RECORD").Range("K5").Value = Sheets("FORM").Range("R" & i).Value
                'Sheets("BILL_RECORD").Range("L5").Value = Sheets("FORM").Range("B" & i).Value
                'Sheets("BILL_RECORD").Range("M5").Value = Sheets("FORM").Range("B" & i).Value
                'Sheets("BILL_RECORD").Range("O5").Value = Sheets("FORM").Range("B" & i).Value
                'Sheets("BILL_RECORD").Range("R5").Value = Sheets("FORM").Range("B" & i).Value
                Sheets("BILL_RECORD").Range("U5").Value = Sheets("FORM").Range("H" & i).Value
            End If
        Next
        
    
        MsgBox "Voucher successfully logged!" & vbNewLine & vbNewLine & Sheets("_temp.Voucher").Range("C12").Value & vbNewLine & "Amount: " & Format(Range("E36").Value, "Standard") & vbNewLine, vbOKOnly
        
errmsg:
        If Err.Number > 0 Then
            MsgBox "An unexpectederror has occured!" & vbNewLine & Err.Description & "Please check the voucher log directory path in the 'Configuration' sheet", vbCritical, "Unexpected error occured"
        End If
    End If
End Sub

Private Sub cmd_Print_Click()
    '*** ACTUAL PRINT ***'
    
    'Copy total charges
    Sheets("_temp.Voucher").Range("D32").Value = Range("H31").Value
    'Copy total 5% VAT
    Sheets("_temp.Voucher").Range("E32").Value = Range("J31").Value
    'Copy total 2% VAT
    Sheets("_temp.Voucher").Range("F32").Value = Range("L31").Value
    'Copy total Tax
    Sheets("_temp.Voucher").Range("G32").Value = Range("N31").Value
    'Copy total Tax 2
    Sheets("_temp.Voucher").Range("J32").Value = Range("N31").Value
    'Copy Net Amount
    Sheets("_temp.Voucher").Range("H32").Value = Range("P31").Value
    'Copy Net Amount 2
    Sheets("_temp.Voucher").Range("J34").Value = Range("P31").Value

    'Copy total amount to voucher
    Sheets("_temp.Voucher").Range("J20").Value = Range("E36").Value
    'Copy total amount to alobs
    Sheets("_temp.Alobs").Range("G15").Value = Range("E36").Value
    Sheets("_temp.Alobs").Range("G18").Value = Range("E36").Value
    
    'Copy total amount in words
    Sheets("_temp.Voucher").Range("B19").Value = Range("E38").Value
    
    CopyCellNos
    CopyAccNos
    CopyCharges
    CopyVAT05
    CopyVAT02
    CopyTotalTax
    AmtBalance
    BillPeriod
    
    ' PRINT SEQUENCE
    If Sheets("Settings").Range("E6").Value = "Print Voucher first" Then
        Sheets("_temp.Voucher").PrintOut From:=1, To:=1, Copies:=Sheets("Settings").Range("E8").Value
        Sheets("_temp.Alobs").PrintOut From:=1, To:=1, Copies:=Sheets("Settings").Range("E10").Value
    Else
        Sheets("_temp.Alobs").PrintOut From:=1, To:=1, Copies:=Sheets("Settings").Range("E10").Value
        Sheets("_temp.Voucher").PrintOut From:=1, To:=1, Copies:=Sheets("Settings").Range("E8").Value
    End If
End Sub

Private Sub cmd_PrintAlternative_Click()
    '*** ALTERNATIVE PRINT ***'
    
    'Copy total charges
    Sheets("_temp.Voucher").Range("D32").Value = Range("H31").Value
    'Copy total 5% VAT
    Sheets("_temp.Voucher").Range("E32").Value = Range("J31").Value
    'Copy total 2% VAT
    Sheets("_temp.Voucher").Range("F32").Value = Range("L31").Value
    'Copy total Tax
    Sheets("_temp.Voucher").Range("G32").Value = Range("N31").Value
    'Copy total Tax 2
    Sheets("_temp.Voucher").Range("J32").Value = Range("N31").Value
    'Copy Net Amount
    Sheets("_temp.Voucher").Range("H32").Value = Range("P31").Value
    'Copy Net Amount 2
    Sheets("_temp.Voucher").Range("J34").Value = Range("P31").Value

    'Copy total amount to voucher
    Sheets("_temp.Voucher").Range("J20").Value = Range("E36").Value
    'Copy total amount to alobs
    Sheets("_temp.Alobs").Range("G15").Value = Range("E36").Value
    Sheets("_temp.Alobs").Range("G18").Value = Range("E36").Value
    
    'Copy total amount in words
    Sheets("_temp.Voucher").Range("B19").Value = Range("E38").Value
    
    CopyCellNos
    CopyAccNos
    CopyCharges
    CopyVAT05
    CopyVAT02
    CopyTotalTax
    AmtBalance
    BillPeriod
    
    ' PRINT SEQUENCE
    If Sheets("Settings").Range("E6").Value = "Print Voucher first" Then
        Sheets("_temp.Voucher").PrintOut From:=1, To:=1, Copies:=Range("O45").Value
        Sheets("_temp.Alobs").PrintOut From:=1, To:=1, Copies:=Range("O47").Value
    Else
        Sheets("_temp.Alobs").PrintOut From:=1, To:=1, Copies:=Range("O47").Value
        Sheets("_temp.Voucher").PrintOut From:=1, To:=1, Copies:=Range("O45").Value
    End If
End Sub

Private Sub cmd_PrintPreview_Click()
    'Copy total charges
    Sheets("_temp.Voucher").Range("D32").Value = Range("H31").Value
    'Copy total 5% VAT
    Sheets("_temp.Voucher").Range("E32").Value = Range("J31").Value
    'Copy total 2% VAT
    Sheets("_temp.Voucher").Range("F32").Value = Range("L31").Value
    'Copy total Tax
    Sheets("_temp.Voucher").Range("G32").Value = Range("N31").Value
    'Copy total Tax 2
    Sheets("_temp.Voucher").Range("J32").Value = Range("N31").Value
    'Copy Net Amount
    Sheets("_temp.Voucher").Range("H32").Value = Range("P31").Value
    'Copy Net Amount 2
    Sheets("_temp.Voucher").Range("J34").Value = Range("P31").Value

    'Copy total amount to voucher
    Sheets("_temp.Voucher").Range("J20").Value = Range("E36").Value
    'Copy total amount to alobs
    Sheets("_temp.Alobs").Range("G15").Value = Range("E36").Value
    Sheets("_temp.Alobs").Range("G18").Value = Range("E36").Value
    
    'Copy total amount in words
    Sheets("_temp.Voucher").Range("B19").Value = Range("E38").Value
    
    CopyCellNos
    CopyAccNos
    CopyCharges
    CopyVAT05
    CopyVAT02
    CopyTotalTax
    AmtBalance
    BillPeriod
    
    Sheets("_temp.Voucher").PrintPreview
End Sub
