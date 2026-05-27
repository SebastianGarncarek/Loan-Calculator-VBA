VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalculatorUI 
   Caption         =   "Loan Installments Calculator"
   ClientHeight    =   6468
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

Private Sub tglLoanType_Click()

'Handles switching between fixed and floating interest rates

If (tglLoanType.Value = True) Then
    lblInterestType.Caption = "Fixed interest rate enabled (margin only)."
    
    'Clear old upcoming installment values after state change
    txtUpcomingInstallment.Value = ""
    txtInstallmentDate.Value = ""
Else
    lblInterestType.Caption = "Floating interest rate enabled (margin + WIBOR)."
    
    'Clear old upcoming installment values after state change
    txtUpcomingInstallment.Value = ""
    txtInstallmentDate.Value = ""
End If
    
End Sub

Private Sub txtPrincipal_AfterUpdate()

'Validates input fields after principal updated to enable/disable calculation buttons

If (txtPrincipal.Value >= 500 And txtMargin.Value > 0 And Len(txtPrincipal.Value) > 0 And Len(txtMargin.Value) > 0 And IsNumeric(txtPrincipal.Value) And IsNumeric(txtMargin.Value)) Then
    btnGenerateSchedule.Locked = False 'unlock button
    btnGenerateSchedule.BackColor = &H8000000F 'change button color to active
    btnUpcomingInstallment.Locked = False
    btnUpcomingInstallment.BackColor = &H8000000F
    
    'Clear old upcoming installment values after state change
    txtUpcomingInstallment.Value = ""
    txtInstallmentDate.Value = ""
Else
    btnGenerateSchedule.Locked = True 'lock button
    btnGenerateSchedule.BackColor = &H80000011 'change button color to inactive
    btnUpcomingInstallment.Locked = True
    btnUpcomingInstallment.BackColor = &H80000011
    
    'Clear old upcoming installment values after state change
    txtUpcomingInstallment.Value = ""
    txtInstallmentDate.Value = ""
End If

End Sub

Private Sub txtMargin_AfterUpdate()

'Validates input fields after margin update to enable/disable calculation buttons

If (txtPrincipal.Value >= 500 And txtMargin.Value > 0 And Len(txtPrincipal.Value) > 0 And Len(txtMargin.Value) > 0 And IsNumeric(txtPrincipal.Value) And IsNumeric(txtMargin.Value)) Then
    btnGenerateSchedule.Locked = False 'unlock button
    btnGenerateSchedule.BackColor = &H8000000F 'change button color to active
    btnUpcomingInstallment.Locked = False
    btnUpcomingInstallment.BackColor = &H8000000F
    
    'Clear old upcoming installment values after state change
    txtUpcomingInstallment.Value = ""
    txtInstallmentDate.Value = ""
Else
    btnGenerateSchedule.Locked = True 'lock button
    btnGenerateSchedule.BackColor = &H80000011 'change button color to inactive
    btnUpcomingInstallment.Locked = True
    btnUpcomingInstallment.BackColor = &H80000011
    
    'Clear old upcoming installment values after state change
    txtUpcomingInstallment.Value = ""
    txtInstallmentDate.Value = ""
End If

End Sub

Private Sub scrLoanDuration_Change()

'Updates the loan duration text box based on the slider value
    
txtLoanDuration.Value = scrLoanDuration.Value

'Clear old upcoming installment values after state change
txtUpcomingInstallment.Value = ""
txtInstallmentDate.Value = ""
    
End Sub

Private Sub scrDueDay_Change()

'Updates the monthly due day text box based on the slider value

txtDueDay.Value = scrDueDay.Value

'Clear old upcoming installment values after state change
txtUpcomingInstallment.Value = ""
txtInstallmentDate.Value = ""

End Sub

Private Sub btnUpcomingInstallment_Click()

'Calls functions from Module1 to calculate the upcoming installment amount and its due date
    
Dim installmentAmount As Currency
Dim installmentStr As String
Dim installmentDueDate As Date

installmentAmount = LoanLogic.UpcomingInstallment(CCur(txtPrincipal.Value), CDbl(txtMargin.Value / 100), CInt(12 * txtLoanDuration.Value), CInt(txtDueDay.Value), CBool(tglLoanType.Value))
installmentStr = Format(installmentAmount, "0.00") & " PLN" 'Format output to 2 decimal places with currency standard

txtUpcomingInstallment.Value = installmentStr

installmentDueDate = LoanLogic.UpcomingDueDate(CInt(12 * txtLoanDuration.Value), CInt(txtDueDay.Value))

'Check if a valid date was returned, then format do standard DD/MM/YYYY
If (CDbl(installmentDueDate) > 1000) Then
    txtInstallmentDate.Value = Format(installmentDueDate, "dd/mm/yyyy")
End If

End Sub

Private Sub btnGenerateSchedule_Click()

'Calls the macro from Module1 to generate the dynamic repayment schedule table
    
LoanLogic.GenerateSchedule CCur(txtPrincipal.Value), CDbl(txtMargin.Value / 100), CInt(12 * txtLoanDuration.Value), CInt(txtDueDay.Value), CBool(tglLoanType.Value)

'Unload the calculator interface after generating the table
Unload Me
    
End Sub

Private Sub btnExit_Click()

'Closes the calculator UserForm interface

Unload Me
    
End Sub
