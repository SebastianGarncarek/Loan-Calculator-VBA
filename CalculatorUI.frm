VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalculatorUI 
   Caption         =   "Loan Installments Calculator"
   ClientHeight    =   8412.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5040
   OleObjectBlob   =   "CalculatorUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CalculatorUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    
Dim Index As Integer
Dim lastRow As Long
Dim minYear As Integer
Dim maxYear As Integer
Dim MonthsAsText As Variant

MonthsAsText = Array("January", "February", "March", "April", "May", "June", _
                     "July", "August", "September", "October", "November", "December")
    
'Fill the day combobox with accepted days
For Index = 1 To 28
    cmbDueDay.AddItem CStr(Index)
Next Index
cmbDueDay.Text = "10" 'default due day is the 10th

'Fill the month combobox with month names
For Index = 0 To 11
    cmbStartMonth.AddItem MonthsAsText(Index)
Next Index
cmbStartMonth.ListIndex = 0 'default start month is January
    
'Fill the year combobox with accepted years
lastRow = wksWibor.Cells(wksWibor.Rows.Count, 1).End(xlUp).Row
minYear = 2011
maxYear = Year(wksWibor.Cells(lastRow, 1).Value)

For Index = minYear To maxYear
    cmbStartYear.AddItem CStr(Index)
Next Index

cmbStartYear.Text = "2011"

End Sub

Private Sub tglLoanType_Click()

'Handles switching between fixed and floating interest rates

If (tglLoanType.Value = True) Then
    tglLoanType.Caption = "Set Floating Interest Rate"
    lblInterestType.Caption = "Fixed interest rate enabled (margin only)."
    
    'Clear old upcoming installment values after state change
    txtUpcomingInstallment.Value = ""
    txtInstallmentDate.Value = ""
Else
    tglLoanType.Caption = "Set Fixed Interest Rate"
    lblInterestType.Caption = "Floating interest rate enabled (margin + WIBOR)."
    
    'Clear old upcoming installment values after state change
    txtUpcomingInstallment.Value = ""
    txtInstallmentDate.Value = ""
End If
    
End Sub

Private Sub ValidateInputs()

Dim isPrincipalValid As Boolean
Dim isMarginValid As Boolean

isPrincipalValid = IsNumeric(txtPrincipal.Value) And Len(txtPrincipal.Value) > 0
isMarginValid = IsNumeric(txtMargin.Value) And Len(txtMargin.Value) > 0

If isPrincipalValid And isMarginValid Then
    If (CDbl(txtPrincipal.Value) >= 500 And CDbl(txtMargin.Value) > 0) Then
        btnGenerateSchedule.Locked = False
        btnGenerateSchedule.BackColor = &H8000000F
        btnUpcomingInstallment.Locked = False
        btnUpcomingInstallment.BackColor = &H8000000F
        
        Exit Sub
    End If
End If

btnGenerateSchedule.Locked = True
btnGenerateSchedule.BackColor = &H80000011
btnUpcomingInstallment.Locked = True
btnUpcomingInstallment.BackColor = &H80000011

End Sub

Private Sub txtPrincipal_AfterUpdate()

'Validates input fields after principal updated to enable/disable calculation buttons

ValidateInputs

'Clear old upcoming installment values after state change
txtUpcomingInstallment.Value = ""
txtInstallmentDate.Value = ""

End Sub

Private Sub txtMargin_AfterUpdate()

'Validates input fields after margin update to enable/disable calculation buttons

ValidateInputs

'Clear old upcoming installment values after state change
txtUpcomingInstallment.Value = ""
txtInstallmentDate.Value = ""

End Sub

Private Sub scrLoanDuration_Change()

'Updates the loan duration text box based on the slider value
    
txtLoanDuration.Value = scrLoanDuration.Value

'Clear old upcoming installment values after state change
txtUpcomingInstallment.Value = ""
txtInstallmentDate.Value = ""
    
End Sub

Private Sub btnUpcomingInstallment_Click()

'Calls functions from Module1 to calculate the upcoming installment amount and its due date
    
Dim installmentAmount As Currency
Dim installmentStr As String
Dim installmentDueDate As Date
Dim sDay As Integer
Dim sMonth As Integer
Dim sYear As Integer

sDay = CInt(cmbDueDay.Text)
sMonth = cmbStartMonth.ListIndex + 1
sYear = CInt(cmbStartYear.Text)

installmentAmount = LoanLogic.UpcomingInstallment(CCur(txtPrincipal.Value), CDbl(txtMargin.Value / 100), CInt(12 * txtLoanDuration.Value), sDay, sMonth, sYear, CBool(tglLoanType.Value))
installmentStr = Format(installmentAmount, "0.00") & " PLN" 'Format output to 2 decimal places with currency standard

txtUpcomingInstallment.Value = installmentStr

installmentDueDate = LoanLogic.UpcomingDueDate(CInt(12 * txtLoanDuration.Value), sDay, sMonth, sYear)

'Check if a valid date was returned, then format do standard DD/MM/YYYY
If (CDbl(installmentDueDate) > 1000) Then
    txtInstallmentDate.Value = Format(installmentDueDate, "dd/mm/yyyy")
End If

End Sub

Private Sub btnGenerateSchedule_Click()

'Calls the macro from Module1 to generate the dynamic repayment schedule table
    
Dim sDay As Integer
Dim sMonth As Integer
Dim sYear As Integer

sDay = CInt(cmbDueDay.Text)
sMonth = cmbStartMonth.ListIndex + 1
sYear = CInt(cmbStartYear.Text)
    
LoanLogic.GenerateSchedule CCur(txtPrincipal.Value), CDbl(txtMargin.Value / 100), CInt(12 * txtLoanDuration.Value), sDay, sMonth, sYear, CBool(tglLoanType.Value)

'Unload the calculator interface after generating the table
Unload Me
    
End Sub

Private Sub btnExit_Click()

'Closes the calculator UserForm interface

Unload Me
    
End Sub

Private Sub UserForm_Click()

End Sub
