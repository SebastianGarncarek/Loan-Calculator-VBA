Attribute VB_Name = "LoanLogic"

Option Explicit

Function CalculateInstallment(principal As Currency, totalInstallments As Integer, interestRate As Double) As Currency

'Calculates loan installments for a given principal, number of installments, and interest rate

Dim installmentSum As Double 'variable storing the total sum of all installments
Dim Index As Integer

installmentSum = 0

For Index = 1 To totalInstallments
    installmentSum = installmentSum + (1 / (1 + (interestRate / 12))) ^ Index 'summing the installments in a loop
Next Index
        
CalculateInstallment = CCur(principal / installmentSum) 'calculating the single installment amount

End Function

Function CalculateInterestRate(margin As Double, targetDate As Double) As Double

'Calculates the total interest rate (margin + 60-day moving average of WIBOR) for selected date

Dim movingAvgRangeStr As String 'stores the cell range for the moving average as a string
Dim movingAvgRange As Range 'range object for calculating average
Dim averageRate As Double 'variable storing the computed average rate
Dim isDateFound As Boolean 'flag indicating if the target date (or closest match) was found in the "Wibor" sheet
Dim today As Date 'variable storing the current date
Dim latestWiborDate As Double 'the latest date in the Wibor sheet
Dim Index As Integer
Dim lastRow As Long

today = Date

lastRow = wksWibor.Cells(wksWibor.Rows.Count, 1).End(xlUp).Row
latestWiborDate = CDbl(wksWibor.Cells(lastRow, 1).Value) 'retrieve the latest date in the "Wibor" sheet

'If the target date exceeds out database, cap it at the latest available date
If targetDate > latestWiborDate Then
    targetDate = latestWiborDate
End If

If (targetDate >= CDbl(today) + 60) Then
    'For dates further in the future (over 60 days from today), use the last available WIBOR value
    averageRate = wksWibor.Cells(lastRow, 5) / 100
Else
    'Search for the exact closest preceding date available in the "Wibor" sheet
    Dim bestIndex As Integer
    bestIndex = 2 'default value in case the date is older than the ones in database
    
    For Index = 2 To wksWibor.UsedRange.Rows.Count 'iterate through the entire "Wibor" sheet
        If (CDbl(wksWibor.Cells(Index, 1).Value) <= targetDate) Then
            bestIndex = Index 'update until we find date after the target one
        Else
            Exit For 'If we get a higher date it means that bestIndex is the target date or is the closest to it and not greater
        End If
    Next Index
    
    'Safety net in case of a negative index
    Dim startRow As Integer
    startRow = bestIndex - 60
    If startRow < 2 Then startRow = 2 'we bound the start by the first date
    
    movingAvgRangeStr = "E" & startRow & ":E" & bestIndex 'define the range
    Set movingAvgRange = wksWibor.Range(movingAvgRangeStr)
    'averageRate = (Application.WorksheetFunction.Average(movingAvgRange)) / 100 'calculate the 60-day average and convert to percentage

    Dim cell As Range
    Dim sumWibor As Double
    Dim countWibor As Integer

    sumWibor = 0
    countWibor = 0

    For Each cell In movingAvgRange
        If Not IsError(cell.Value) Then
            ' Sprawdzamy, czy Excel traktuje wartoœæ w komórce jako prawdziw¹ liczbê, a nie ukryty tekst
            If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
                sumWibor = sumWibor + CDbl(cell.Value)
                countWibor = countWibor + 1
            End If
        End If
    Next cell

    ' Liczymy œredni¹ omijaj¹c komórki z tekstem i b³êdy
    If countWibor > 0 Then
        averageRate = (sumWibor / countWibor) / 100
    Else
        averageRate = 0 ' Zabezpieczenie przed dzieleniem przez zero
    End If
End If

CalculateInterestRate = margin + averageRate 'total interest rate is the sum of bank margin and the computed WIBOR average
        
End Function

Sub GenerateSchedule(principal As Currency, margin As Double, totalInstallments As Integer, dueDay As Integer, startMonth As Integer, startYear As Integer, isFixedRate As Boolean)

'Generates the credit repayment schedule table based on principal, bank margin, total installments, payment day, and rate type (fixed/floating)

Application.ScreenUpdating = False 'disable screen updating to boost execution performance

wksMain.Protect Password:="MAIN", UserInterfaceOnly:=True
wksSchedule.Protect Password:="SCHEDULE", UserInterfaceOnly:=True

wksSchedule.Cells.Clear 'clear the sheet before generating the table

Dim Index As Integer
Dim installment As Currency
Dim currentDate As Date
Dim interestPaid As Currency
Dim interestRate As Double
Dim latestWiborDate As Double
Dim lastWiborRow As Long

lastWiborRow = wksWibor.Cells(wksWibor.Rows.Count, 1).End(xlUp).Row
latestWiborDate = CDbl(wksWibor.Cells(lastWiborRow, 1).Value)

interestRate = margin
currentDate = DateSerial(startYear, startMonth, dueDay)

With wksMain
    .Range("DshStartDate").Value = ""
    .Range("DshPrincipal").Value = ""
    .Range("DshMargin").Value = ""
    .Range("DshTerm").Value = ""
    .Range("DshType").Value = ""
    
    .Range("DshStartDate").Value = DateSerial(startYear, startMonth, dueDay)
    .Range("DshStartDate").NumberFormat = "dd/mm/yyyy"
    
    .Range("DshPrincipal").Value = principal
    .Range("DshPrincipal").NumberFormat = "#,##0.00 ""PLN"""
    
    .Range("DshMargin").Value = margin
    .Range("DshMargin").NumberFormat = "0.00%"
    
    .Range("DshTerm").Value = totalInstallments / 12
    .Range("DshTerm").NumberFormat = "0"
    
    If isFixedRate Then
        .Range("DshType").Value = "Fixed (Margin only)"
    Else
        .Range("DshType").Value = "Floating (Margin + WIBOR)"
    End If
End With

'Generate table column heads
wksSchedule.Cells(1, 1).Value = "No."
wksSchedule.Cells(1, 2).Value = "Due Date"
wksSchedule.Cells(1, 3).Value = "Installment Amount"
wksSchedule.Cells(1, 4).Value = "Interest Paid"
wksSchedule.Cells(1, 5).Value = "Principal Repaid"
wksSchedule.Cells(1, 6).Value = "Remaining Balance"
wksSchedule.Cells(1, 7).Value = "Rate Type (Source)"

If (isFixedRate = True) Then installment = CalculateInstallment(principal, totalInstallments, margin) 'for fixed-rate loans, the installment is calculated only once

For Index = 1 To (totalInstallments)
    currentDate = DateAdd("m", 1, currentDate)
    If (Index Mod 3 = 1 And isFixedRate = False) Then 'for floating-rate loans, recalculate the installment every 3 months
        interestRate = CalculateInterestRate(margin, CDbl(currentDate))
        installment = CalculateInstallment(principal, totalInstallments - Index + 1, interestRate)
    End If
    interestPaid = CCur(interestRate / 12 * principal)
    If (Index = totalInstallments) Then installment = interestPaid + principal 'last installment adjusted to clear the exact remaining principal
    
    'Fill the schedule rows
    wksSchedule.Cells(Index + 1, 1).Value = "'" & Format(Index, "0") & "."
    wksSchedule.Cells(Index + 1, 2).Value = currentDate
    wksSchedule.Cells(Index + 1, 3).Value = installment
    wksSchedule.Cells(Index + 1, 4).Value = interestPaid
    wksSchedule.Cells(Index + 1, 5).Value = installment - interestPaid
    principal = principal - (installment - interestPaid)
    wksSchedule.Cells(Index + 1, 6).Value = principal
    
    If isFixedRate Then
        wksSchedule.Cells(Index + 1, 7).Value = "Fixed Rate"
    ElseIf CDbl(currentDate) > latestWiborDate Then
        wksSchedule.Cells(Index + 1, 7).Value = "Forecasted (Latest WIBOR)"
    Else
        wksSchedule.Cells(Index + 1, 7).Value = "Historical (60-day avg)"
    End If
    
    wksSchedule.Range(wksSchedule.Cells(Index + 1, 3), wksSchedule.Cells(Index + 1, 6)).NumberFormat = "#,##0.00 ""PLN"""
Next Index

wksSchedule.UsedRange.EntireColumn.AutoFit
Application.ScreenUpdating = True 're-enable screen updating after table generation completes

End Sub

Function UpcomingInstallment(principal As Currency, margin As Double, totalInstallments As Integer, dueDay As Integer, startMonth As Integer, startYear As Integer, isFixedRate As Boolean) As Currency

'Calculates the upcoming installment amount relative to today's date based on loan specifications
    
Dim installment As Currency
Dim currentDate As Date
Dim interestPaid As Currency
Dim interestRate As Double
Dim today As Date
Dim isPeriodFound As Boolean 'flag indicating if the active loan period matching today's date was found
Dim Index As Integer

today = Date
interestRate = margin
currentDate = DateSerial(startYear, startMonth, dueDay)

If (isFixedRate = True) Then installment = CalculateInstallment(principal, totalInstallments, margin)

For Index = 1 To (totalInstallments)
    currentDate = DateAdd("m", 1, currentDate)
    If (Index Mod 3 = 1 And isFixedRate = False) Then
        interestRate = CalculateInterestRate(margin, CDbl(currentDate))
        installment = CalculateInstallment(principal, totalInstallments - Index + 1, interestRate)
    End If
    interestPaid = CCur(interestRate / 12 * principal)
    If (Index = totalInstallments) Then installment = interestPaid + principal
    If (today <= currentDate And today > DateAdd("m", -1, currentDate)) Then 'check if today falls within the current monthly installment period
        isPeriodFound = True 'match found, abort search to return current period's rate
        Exit For
    End If
    principal = principal - (installment - interestPaid)
Next Index

If (isPeriodFound = False) Then installment = 0 'if no active period matches, the loan has already been fully repaid
UpcomingInstallment = installment

End Function

Function UpcomingDueDate(totalInstallments As Integer, dueDay As Integer, startMonth As Integer, startYear As Integer) As Date

'Returns the due date for the upcoming loan installment based on the total terms and payment day

Dim currentDate As Date
Dim today As Date
Dim isPeriodFound As Boolean 'flag indicating if the target payment window was identified
Dim Index As Integer

today = Date
currentDate = DateSerial(startYear, startMonth, dueDay)

For Index = 1 To totalInstallments
    currentDate = DateAdd("m", 1, currentDate)
    If (today <= currentDate And today > DateAdd("m", -1, currentDate)) Then 'check if today falls within this monthly cycle
        isPeriodFound = True 'target payment date identified, exit search
        Exit For
    End If
Next Index

If (isPeriodFound = False) Then currentDate = 0 'if no period matches, the loan is fully repaid
UpcomingDueDate = currentDate

End Function

Sub OpenCalculator()

'Macro to initialize and display the graphical calculator interface

Dim loanCalculator As CalculatorUI

Set loanCalculator = New CalculatorUI 'instantiate the Calculator UserForm object

loanCalculator.Show 'display the user interface

Unload loanCalculator

End Sub
