Attribute VB_Name = "Module2"
Public Sub CopyCellNos()
    Sheets("_temp.Voucher").Range("B23").Value = Sheets("FORM").Range("B22").Value
    Sheets("_temp.Voucher").Range("B24").Value = Sheets("FORM").Range("B23").Value
    Sheets("_temp.Voucher").Range("B25").Value = Sheets("FORM").Range("B24").Value
    Sheets("_temp.Voucher").Range("B26").Value = Sheets("FORM").Range("B25").Value
    Sheets("_temp.Voucher").Range("B27").Value = Sheets("FORM").Range("B26").Value
    Sheets("_temp.Voucher").Range("B28").Value = Sheets("FORM").Range("B27").Value
    Sheets("_temp.Voucher").Range("B29").Value = Sheets("FORM").Range("B28").Value
    Sheets("_temp.Voucher").Range("B30").Value = Sheets("FORM").Range("B29").Value
End Sub

Public Sub CopyAccNos()
    Sheets("_temp.Voucher").Range("C23").Value = Sheets("FORM").Range("E22").Value
    Sheets("_temp.Voucher").Range("C24").Value = Sheets("FORM").Range("E23").Value
    Sheets("_temp.Voucher").Range("C25").Value = Sheets("FORM").Range("E24").Value
    Sheets("_temp.Voucher").Range("C26").Value = Sheets("FORM").Range("E25").Value
    Sheets("_temp.Voucher").Range("C27").Value = Sheets("FORM").Range("E26").Value
    Sheets("_temp.Voucher").Range("C28").Value = Sheets("FORM").Range("E27").Value
    Sheets("_temp.Voucher").Range("C29").Value = Sheets("FORM").Range("E28").Value
    Sheets("_temp.Voucher").Range("C30").Value = Sheets("FORM").Range("E29").Value
End Sub

Public Sub CopyCharges()
    Sheets("_temp.Voucher").Range("D23").Value = Sheets("FORM").Range("H22").Value
    Sheets("_temp.Voucher").Range("D24").Value = Sheets("FORM").Range("H23").Value
    Sheets("_temp.Voucher").Range("D25").Value = Sheets("FORM").Range("H24").Value
    Sheets("_temp.Voucher").Range("D26").Value = Sheets("FORM").Range("H25").Value
    Sheets("_temp.Voucher").Range("D27").Value = Sheets("FORM").Range("H26").Value
    Sheets("_temp.Voucher").Range("D28").Value = Sheets("FORM").Range("H27").Value
    Sheets("_temp.Voucher").Range("D29").Value = Sheets("FORM").Range("H28").Value
    Sheets("_temp.Voucher").Range("D30").Value = Sheets("FORM").Range("H29").Value
End Sub

Public Sub CopyVAT05()
    Sheets("_temp.Voucher").Range("E23").Value = Sheets("FORM").Range("J22").Value
    Sheets("_temp.Voucher").Range("E24").Value = Sheets("FORM").Range("J23").Value
    Sheets("_temp.Voucher").Range("E25").Value = Sheets("FORM").Range("J24").Value
    Sheets("_temp.Voucher").Range("E26").Value = Sheets("FORM").Range("J25").Value
    Sheets("_temp.Voucher").Range("E27").Value = Sheets("FORM").Range("J26").Value
    Sheets("_temp.Voucher").Range("E28").Value = Sheets("FORM").Range("J27").Value
    Sheets("_temp.Voucher").Range("E29").Value = Sheets("FORM").Range("J28").Value
    Sheets("_temp.Voucher").Range("E30").Value = Sheets("FORM").Range("J29").Value
End Sub

Public Sub CopyVAT02()
    Sheets("_temp.Voucher").Range("F23").Value = Sheets("FORM").Range("L22").Value
    Sheets("_temp.Voucher").Range("F24").Value = Sheets("FORM").Range("L23").Value
    Sheets("_temp.Voucher").Range("F25").Value = Sheets("FORM").Range("L24").Value
    Sheets("_temp.Voucher").Range("F26").Value = Sheets("FORM").Range("L25").Value
    Sheets("_temp.Voucher").Range("F27").Value = Sheets("FORM").Range("L26").Value
    Sheets("_temp.Voucher").Range("F28").Value = Sheets("FORM").Range("L27").Value
    Sheets("_temp.Voucher").Range("F29").Value = Sheets("FORM").Range("L28").Value
    Sheets("_temp.Voucher").Range("F30").Value = Sheets("FORM").Range("L29").Value
End Sub

Public Sub CopyTotalTax()
    Sheets("_temp.Voucher").Range("G23").Value = Sheets("FORM").Range("N22").Value
    Sheets("_temp.Voucher").Range("G24").Value = Sheets("FORM").Range("N23").Value
    Sheets("_temp.Voucher").Range("G25").Value = Sheets("FORM").Range("N24").Value
    Sheets("_temp.Voucher").Range("G26").Value = Sheets("FORM").Range("N25").Value
    Sheets("_temp.Voucher").Range("G27").Value = Sheets("FORM").Range("N26").Value
    Sheets("_temp.Voucher").Range("G28").Value = Sheets("FORM").Range("N27").Value
    Sheets("_temp.Voucher").Range("G29").Value = Sheets("FORM").Range("N28").Value
    Sheets("_temp.Voucher").Range("G30").Value = Sheets("FORM").Range("N29").Value
End Sub

Public Sub AmtBalance()
    Sheets("_temp.Voucher").Range("H23").Value = Sheets("FORM").Range("P22").Value
    Sheets("_temp.Voucher").Range("H24").Value = Sheets("FORM").Range("P23").Value
    Sheets("_temp.Voucher").Range("H25").Value = Sheets("FORM").Range("P24").Value
    Sheets("_temp.Voucher").Range("H26").Value = Sheets("FORM").Range("P25").Value
    Sheets("_temp.Voucher").Range("H27").Value = Sheets("FORM").Range("P26").Value
    Sheets("_temp.Voucher").Range("H28").Value = Sheets("FORM").Range("P27").Value
    Sheets("_temp.Voucher").Range("H29").Value = Sheets("FORM").Range("P28").Value
    Sheets("_temp.Voucher").Range("H30").Value = Sheets("FORM").Range("P29").Value
End Sub

Public Sub BillPeriod()
    Sheets("_temp.Voucher").Range("I23").Value = Sheets("FORM").Range("R22").Value
    Sheets("_temp.Voucher").Range("I24").Value = Sheets("FORM").Range("R23").Value
    Sheets("_temp.Voucher").Range("I25").Value = Sheets("FORM").Range("R24").Value
    Sheets("_temp.Voucher").Range("I26").Value = Sheets("FORM").Range("R25").Value
    Sheets("_temp.Voucher").Range("I27").Value = Sheets("FORM").Range("R26").Value
    Sheets("_temp.Voucher").Range("I28").Value = Sheets("FORM").Range("R27").Value
    Sheets("_temp.Voucher").Range("I29").Value = Sheets("FORM").Range("R28").Value
    Sheets("_temp.Voucher").Range("I30").Value = Sheets("FORM").Range("R29").Value
End Sub
