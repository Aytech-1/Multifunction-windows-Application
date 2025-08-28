Public Class Form1

    Private Sub close_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles close_btn.Click
        Dim answer As Integer
        answer = MessageBox.Show("Are you sure you want to Exit", "Exit Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If answer = vbYes Then
            Me.Close()
        End If
    End Sub

    Private Sub minimize_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles minimize_btn.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AfooTECH_home_page_btn.Show()
        calculator_panel.Hide()
        age_panel.Hide()
        boyles_law_panel.Hide()
        loan_panel.Hide()
        unit_converter_panel.Hide()
        afotech_gp_panel.Hide()
        home_page_panel.Show()

        ' Initialize current date values
        current_day = DateTime.Now.Day
        current_month = DateTime.Now.Month
        current_year = DateTime.Now.Year

        ' Display current date in textboxes
        current_day_txt.Text = current_day.ToString()
        current_month_txt.Text = current_month.ToString()
        current_year_txt.Text = current_year.ToString()

        ' Set up error provider settings
        errorProvider.BlinkStyle = ErrorBlinkStyle.NeverBlink ' Optional: Prevents blinking of error icon

    End Sub
    Private Sub home_page_panel_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles home_page_panel.Paint
        home_page_panel.Show()
        AfooTECH_home_page_btn.Show()
        calculator_panel.Hide()
        age_panel.Hide()
        boyles_law_panel.Hide()
        loan_panel.Show()
        unit_converter_panel.Hide()
        afotech_gp_panel.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        AfooTECH_home_page_btn.Show()
        home_page_panel.Hide()
        calculator_panel.Show()
        age_panel.Hide()
        boyles_law_panel.Hide()
        loan_panel.Hide()
        unit_converter_panel.Hide()
        afotech_gp_panel.Hide()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        home_page_panel.Hide()
        AfooTECH_home_page_btn.Show()
        calculator_panel.Hide()
        age_panel.Show()
        boyles_law_panel.Hide()
        loan_panel.Hide()
        unit_converter_panel.Hide()
        afotech_gp_panel.Hide()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        home_page_panel.Hide()
        AfooTECH_home_page_btn.Show()
        calculator_panel.Hide()
        age_panel.Hide()
        boyles_law_panel.Show()
        loan_panel.Hide()
        unit_converter_panel.Hide()
        afotech_gp_panel.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        home_page_panel.Hide()
        AfooTECH_home_page_btn.Show()
        calculator_panel.Hide()
        age_panel.Hide()
        boyles_law_panel.Hide()
        loan_panel.Show()
        unit_converter_panel.Hide()
        afotech_gp_panel.Hide()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        home_page_panel.Hide()
        AfooTECH_home_page_btn.Show()
        calculator_panel.Hide()
        age_panel.Hide()
        boyles_law_panel.Hide()
        loan_panel.Hide()
        unit_converter_panel.Hide()
        afotech_gp_panel.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        home_page_panel.Hide()
        AfooTECH_home_page_btn.Show()
        calculator_panel.Hide()
        age_panel.Hide()
        boyles_law_panel.Hide()
        loan_panel.Hide()
        unit_converter_panel.Show()
        afotech_gp_panel.Hide()
    End Sub

    Private Sub AfooTECH_home_page_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AfooTECH_home_page_btn.Click
        home_page_panel.Show()
        AfooTECH_home_page_btn.Show()
        calculator_panel.Show()
        age_panel.Show()
        boyles_law_panel.Show()
        loan_panel.Show()
        unit_converter_panel.Show()
        afotech_gp_panel.Show()
    End Sub

    Private Sub calculate_btn_txt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles calculate_btn_txt.Click
        Dim SDT101, MAD102, SDA103, WAD104, NS105, AGD106 As Integer
        Dim Gradepoint1, Gradepoint2, Gradepoint3, Gradepoint4, Gradepoint5, Gradepoint6 As Integer
        Dim GPA, TGP, TCU As Double

        AddHandler SDT101_txt.KeyPress, AddressOf TextBox_KeyPress
        AddHandler MAD102_txt.KeyPress, AddressOf TextBox_KeyPress
        AddHandler SDA103_txt.KeyPress, AddressOf TextBox_KeyPress
        AddHandler WAD104_txt.KeyPress, AddressOf TextBox_KeyPress
        AddHandler NS105_txt.KeyPress, AddressOf TextBox_KeyPress
        AddHandler AGD106_txt.KeyPress, AddressOf TextBox_KeyPress

        AddHandler SDT101_txt.TextChanged, AddressOf TextBox_TextChanged
        AddHandler MAD102_txt.TextChanged, AddressOf TextBox_TextChanged
        AddHandler SDA103_txt.TextChanged, AddressOf TextBox_TextChanged
        AddHandler WAD104_txt.TextChanged, AddressOf TextBox_TextChanged
        AddHandler NS105_txt.TextChanged, AddressOf TextBox_TextChanged
        AddHandler AGD106_txt.TextChanged, AddressOf TextBox_TextChanged

        Try
            SDT101 = Double.Parse(SDT101_txt.Text)
            MAD102 = Double.Parse(MAD102_txt.Text)
            SDA103 = Double.Parse(SDA103_txt.Text)
            WAD104 = Double.Parse(WAD104_txt.Text)
            NS105 = Double.Parse(NS105_txt.Text)
            AGD106 = Double.Parse(AGD106_txt.Text)

            If (SDT101 < 0 Or SDT101 > 100) Or (MAD102 < 0 Or MAD102 > 100) Or (SDA103 < 0 Or SDA103 > 100) Or (WAD104 < 0 Or WAD104 > 100) Or (NS105 < 0 Or NS105 > 100) Or (AGD106 < 0 Or AGD106 > 100) Then
                MessageBox.Show("Scores Must Not < 0 or > 100.", " Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Gradepoint1 = GetGradePoint(SDT101, 4)
            Gradepoint2 = GetGradePoint(MAD102, 4)
            Gradepoint3 = GetGradePoint(SDA103, 5)
            Gradepoint4 = GetGradePoint(WAD104, 4)
            Gradepoint5 = GetGradePoint(NS105, 4)
            Gradepoint6 = GetGradePoint(AGD106, 4)

            TGP = Gradepoint1 + Gradepoint2 + Gradepoint3 + Gradepoint4 + Gradepoint5 + Gradepoint6
            TCU = 25
            GPA = TGP / TCU
            result_tgp_txt.Text = "  " & TGP.ToString
            result_tcu_txt.Text = "  " & TCU.ToString
            result_cgpa_txt.Text = "  " & GPA.ToString("F2")

            If GPA >= 3.5 And GPA <= 4.0 Then
                result_grade_txt.Text = "  " & "DISTINCTION"
            ElseIf GPA >= 3.0 And GPA <= 3.49 Then
                result_grade_txt.Text = "  " & "UPPER CREDIT"
            ElseIf GPA >= 2.5 And GPA <= 2.99 Then
                result_grade_txt.Text = "  " & "LOWER CREDIT"
            ElseIf GPA >= 2.0 And GPA <= 2.49 Then
                result_grade_txt.Text = "  " & "PASS"
            Else
                result_grade_txt.Text = "  " & "FAIL"
            End If

        Catch ex As Exception
            MessageBox.Show("Please enter valid scores in ")
            Exit Sub
        End Try
    End Sub

    Private Function GetGradePoint(ByVal score As Integer, ByVal creditUnits As Integer) As Integer
        Select Case score
            Case 80 To 100
                Return creditUnits * 4
            Case 70 To 79
                Return creditUnits * 3.5
            Case 60 To 69
                Return creditUnits * 3
            Case 50 To 59
                Return creditUnits * 2.5
            Case 40 To 49
                Return creditUnits * 2
            Case Else
                Return creditUnits * 0
        End Select
    End Function

    Private Sub clear_bnt_txt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clear_bnt_txt.Click
        SDT101_txt.Text = ""
        MAD102_txt.Text = ""
        SDA103_txt.Text = ""
        WAD104_txt.Text = ""
        NS105_txt.Text = ""
        AGD106_txt.Text = ""
        result_tgp_txt.Text = ""
        result_tcu_txt.Text = ""
        result_cgpa_txt.Text = ""
        result_grade_txt.Text = ""
    End Sub

    Private Sub TextBox_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles v1_txt.KeyPress, v2_txt.KeyPress, p1_txt.KeyPress, p2_txt.KeyPress
        Dim textBox As TextBox = CType(sender, TextBox)

        ' Allow only digits, one decimal point, and control keys
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
            MessageBox.Show("Please enter only digits or one decimal point.")
            Return
        End If
    End Sub

    Private Sub TextBox_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim textBox As TextBox = CType(sender, TextBox)
        Dim text As String = textBox.Text

        ' Ensure the input does not start with a decimal point
        If text.StartsWith(".") Then
            textBox.Text = ""
            MessageBox.Show("Scores must not start with a decimal point.")
            Return
        End If

        ' Ensure no more than 3 digits before the decimal point
        If text.Length > 3 AndAlso Not text.Contains(".") Then
            textBox.Text = text.Substring(0, 3)
            MessageBox.Show("Scores must not exceed 3 digits.")
            Return
        End If

        ' Ensure only one decimal point is allowed
        If text.Count(Function(c) c = "."c) > 1 Then
            textBox.Text = text.Remove(text.LastIndexOf("."c), 1)
            MessageBox.Show("Only one decimal point is allowed.")
            Return
        End If

        ' Reset the cursor position
        textBox.SelectionStart = textBox.Text.Length
    End Sub

    '-------------------------------SOLUTION FOR UNIT CONVERTER ---------------------------------------------------------

    Private Sub ComboBoxFrom_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxFrom.SelectedIndexChanged
        Dim input, output As Double

        If Input_txt.Text = "" Then
            MessageBox.Show("This Field Cannot be empty", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            input = Input_txt.Text
            If ComboBoxFrom.Text = "Kilometer" Then
                output = input * 1000

            ElseIf ComboBoxFrom.Text = "Decimeter" Then
                output = input * 10

            ElseIf ComboBoxFrom.Text = "Centimeter" Then
                output = input / 100

            ElseIf ComboBoxFrom.Text = "Millimeter" Then
                output = input / 1000

            ElseIf ComboBoxFrom.Text = "Inch" Then
                output = input * 0.0254

            ElseIf ComboBoxFrom.Text = "Foot" Then
                output = input * 0.3048

            ElseIf ComboBoxFrom.Text = "Yard" Then
                output = input * 0.9144
            End If

            Output_txt.Text = output
        End If
    End Sub
    Private Sub Input_txt_TextChanged(ByVal sender As System.Object, ByVal e As KeyPressEventArgs) Handles Input_txt.KeyPress
        'Accept Control Keys, Digit, dot and negative
        If Not Char.IsControl(e.KeyChar) AndAlso
            Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
            MessageBox.Show("This field can accepts only numbers", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Sub clear_unit_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clear_unit_btn.Click
        Input_txt.Text = ""
        Output_txt.Text = ""
        ComboBoxFrom.SelectedIndex = 0
    End Sub

    '-------------------------------SOLUTION FOR BOYLES LAW---------------------------------------------------------

    Dim v1, v2, p1, p2, result As Double
    Dim calculationType As String

    Private Sub v1_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles v1_btn.Click
        calculationType = "v1"
        SetFormState(False, True, True, True)
        ClearInputFields()
        v1_txt.Text = "?"
    End Sub

    Private Sub v2_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles v2_btn.Click
        calculationType = "v2"
        SetFormState(True, False, True, True)
        ClearInputFields()
        v2_txt.Text = "?"
    End Sub

    Private Sub p1_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles p1_btn.Click
        calculationType = "p1"
        SetFormState(True, True, False, True)
        ClearInputFields()
        p1_txt.Text = "?"
    End Sub

    Private Sub p2_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles p2_btn.Click
        calculationType = "p2"
        SetFormState(True, True, True, False)
        ClearInputFields()
        p2_txt.Text = "?"
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If Not ValidateInputFields() Then
            MessageBox.Show("Kindly fill all the input fields", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        v1 = If(Double.TryParse(v1_txt.Text, v1), v1, 0)
        v2 = If(Double.TryParse(v2_txt.Text, v2), v2, 0)
        p1 = If(Double.TryParse(p1_txt.Text, p1), p1, 0)
        p2 = If(Double.TryParse(p2_txt.Text, p2), p2, 0)

        Select Case calculationType
            Case "v1"
                result = (p2 * v2) / p1
                result_txt.Text = "V1 = " & result.ToString("N") & " ml"
            Case "v2"
                result = (p1 * v1) / p2
                result_txt.Text = "V2 = " & result.ToString("N") & " ml³"
            Case "p1"
                result = (v2 * p2) / v1
                result_txt.Text = "P1 = " & result.ToString("N") & " atm"
            Case "p2"
                result = (v1 * p1) / v2
                result_txt.Text = "P2 = " & result.ToString("N") & " atm"
        End Select
    End Sub

    Private Sub SetFormState(ByVal v1Enabled As Boolean, ByVal v2Enabled As Boolean, ByVal p1Enabled As Boolean, ByVal p2Enabled As Boolean)
        v1_txt.Enabled = v1Enabled
        v2_txt.Enabled = v2Enabled
        p1_txt.Enabled = p1Enabled
        p2_txt.Enabled = p2Enabled
    End Sub

    Private Sub ClearInputFields()
        v1_txt.Text = ""
        v2_txt.Text = ""
        p1_txt.Text = ""
        p2_txt.Text = ""
        result_txt.Text = ""
    End Sub

    Private Sub ResetForm()
        SetFormState(True, True, True, True)
        ClearInputFields()
        calculationType = ""
    End Sub

    Private Function ValidateInputFields() As Boolean
        If calculationType = "v1" AndAlso (String.IsNullOrWhiteSpace(p1_txt.Text) Or String.IsNullOrWhiteSpace(p2_txt.Text) Or String.IsNullOrWhiteSpace(v2_txt.Text)) Then Return False
        If calculationType = "v2" AndAlso (String.IsNullOrWhiteSpace(p1_txt.Text) Or String.IsNullOrWhiteSpace(p2_txt.Text) Or String.IsNullOrWhiteSpace(v1_txt.Text)) Then Return False
        If calculationType = "p1" AndAlso (String.IsNullOrWhiteSpace(v1_txt.Text) Or String.IsNullOrWhiteSpace(p2_txt.Text) Or String.IsNullOrWhiteSpace(v2_txt.Text)) Then Return False
        If calculationType = "p2" AndAlso (String.IsNullOrWhiteSpace(v1_txt.Text) Or String.IsNullOrWhiteSpace(p1_txt.Text) Or String.IsNullOrWhiteSpace(v2_txt.Text)) Then Return False
        Return True
    End Function

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        v1_txt.Text = ""
        v2_txt.Text = ""
        p1_txt.Text = ""
        p2_txt.Text = ""
        result_txt.Text = ""
    End Sub

    '-------------------------------SOLUTION  FOR LOAN APPLICATION---------------------------------------------------------

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim loan_amount, loan_duration, repayment, monthly_repayment, interest_on_loan, total_interest, total_monthly_repayment As Double
        Dim errorProvider As New ErrorProvider()

        If Double.TryParse(loan_amount_txt.Text, loan_amount) AndAlso Double.TryParse(loan_duration_txt.Text, loan_duration) Then
            lstRepaymentDetails.Items.Clear()
            lstRepaymentDetails.Items.Add("Month | Repayment   | Interest       | Monthly Repayment")
            lstRepaymentDetails.Items.Add("__________________________________________________________")

            total_interest = 0
            total_monthly_repayment = 0

            For month As Integer = 1 To loan_duration
                repayment = loan_amount / loan_duration
                interest_on_loan = (1.5 / 100) * (loan_amount - ((month - 1) * repayment))
                monthly_repayment = repayment + interest_on_loan
                total_interest += interest_on_loan
                total_monthly_repayment += monthly_repayment

                lstRepaymentDetails.Items.Add(month.ToString().PadRight(9) & "|" & repayment.ToString("n2").PadRight(15) & "|" & interest_on_loan.ToString("n2").PadRight(15) & "|" & monthly_repayment.ToString("n2").PadRight(12))
            Next
            total_interest_txt.Text = ("#") & total_interest.ToString("n2")
            total_monthly_repayment_txt.Text = ("#") & total_monthly_repayment.ToString("n2")
        Else
            MessageBox.Show("Please enter valid numeric values for loan amount and duration.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        loan_amount_txt.Clear()
        loan_duration_txt.Clear()
        total_interest_txt.Clear()
        total_monthly_repayment_txt.Clear()
        lstRepaymentDetails.Items.Clear()

    End Sub

    Private Sub loan_amount_txt_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles loan_amount_txt.KeyPress
        ' Allow only digits and control characters (e.g., backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
            MessageBox.Show("This field can accepts only numbers", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        ' Disallow entering a decimal point as the first character
        If e.KeyChar = "." AndAlso loan_amount_txt.Text.Length = 0 Then
            e.Handled = True
            MessageBox.Show("This field can accepts only numbers", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        ' Disallow entering more than one decimal point
        If e.KeyChar = "." AndAlso loan_amount_txt.Text.Contains(".") Then
            e.Handled = True
            MessageBox.Show("This field can accepts only numbers", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub loan_duration_txt_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles loan_duration_txt.KeyPress
        ' Allow only digits and control characters (e.g., backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
            MessageBox.Show("This field can accepts only numbers", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub loan_amount_txt_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles loan_amount_txt.TextChanged
        ValidateInput(loan_amount_txt)
    End Sub

    Private Sub loan_duration_txt_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles loan_duration_txt.TextChanged
        ValidateInput(loan_duration_txt)
    End Sub

    Private Sub ValidateInput(ByVal textBox As TextBox)
        Dim value As Double
        If Not Double.TryParse(textBox.Text, value) Then
        End If
    End Sub

    '-------------------------------SOLUTION ---------------------------------------------------------

    ' Declare variables to hold date values
    Dim current_day, current_month, current_year As Integer
    Dim birth_day, birth_month, birth_year As Integer
    Dim age_day, age_month, age_year As Integer

    ' Error provider for displaying error messages
    Dim errorProvider As New ErrorProvider()

   

    Private Sub birth_day_txt_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles birth_day_txt.TextChanged
        ' Validate birth day input
        If Not Integer.TryParse(birth_day_txt.Text, birth_day) OrElse
           birth_day <= 0 OrElse birth_day > 31 Then
            errorProvider.SetError(birth_day_txt, "Invalid day. Enter a valid day (1-31).")
            birth_month_txt.Enabled = False ' Disable subsequent inputs
            birth_year_txt.Enabled = False
        Else
            errorProvider.SetError(birth_day_txt, "")
            birth_month_txt.Enabled = True ' Enable subsequent inputs
            birth_year_txt.Enabled = True
        End If
    End Sub

    Private Sub birth_month_txt_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles birth_month_txt.TextChanged
        ' Validate birth month input
        If Not Integer.TryParse(birth_month_txt.Text, birth_month) OrElse
           birth_month <= 0 OrElse birth_month > 12 Then
            errorProvider.SetError(birth_month_txt, "Invalid month. Enter a valid month (1-12).")
            birth_year_txt.Enabled = False ' Disable subsequent inputs
        Else
            errorProvider.SetError(birth_month_txt, "")
            birth_year_txt.Enabled = True ' Enable subsequent inputs
        End If
    End Sub
    Private Sub birth_year_txt_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles birth_year_txt.TextChanged
        ' Validate birth year input
        If Not Integer.TryParse(birth_year_txt.Text, birth_year) OrElse
           birth_year <= 0 OrElse birth_year > current_year Then
            errorProvider.SetError(birth_year_txt, "Invalid year. Enter a valid year.")
            Return
        End If
    End Sub

    Private Sub calculate_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles calculate_btn.Click
        ' Check if all inputs are valid
        If Not Integer.TryParse(birth_year_txt.Text, birth_year) Then
            MessageBox.Show("Date is incorrect, Enter the correct date.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        ElseIf birth_year < 0 Then
            MessageBox.Show("Invalid Date.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        ElseIf birth_day > DateTime.DaysInMonth(birth_year, birth_month) Then
            MessageBox.Show("Invalid Date.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        Else

            ' Calculate age
            age_year = current_year - birth_year

            If current_month >= birth_month Then
                age_month = current_month - birth_month
            Else
                age_month = 12 + current_month - birth_month
                age_year -= 1
            End If

            If current_day >= birth_day Then
                age_day = current_day - birth_day
            Else
                age_month -= 1
                age_day = DateTime.DaysInMonth(birth_year, birth_month) + current_day - birth_day
            End If
            If age_month < 0 Then
                age_month = 11
                age_year -= 1
            End If
            If age_year < 0 Then
                MessageBox.Show("Invalid Date.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Display age in textboxes
            age_day_txt.Text = age_day.ToString()
            age_month_txt.Text = age_month.ToString()
            age_year_txt.Text = age_year.ToString()
        End If
    End Sub

    Private Function ValidateInputs() As Boolean
        ' Validate all inputs
        If birth_day <= 0 Or birth_day > 31 Then
            errorProvider.SetError(birth_day_txt, "Invalid day. Enter a valid day (1-31).")
            Return False
        End If

        If birth_month <= 0 Or birth_month > 12 Then
            errorProvider.SetError(birth_month_txt, "Invalid month. Enter a valid month (1-12).")
            Return False
        End If

        If birth_year <= 0 Or birth_year > current_year Then
            errorProvider.SetError(birth_year_txt, "Invalid year. Enter a valid year.")
            Return False
        End If

        ' Validate leap year for birth year
        If Not DateTime.IsLeapYear(birth_year) Then
            errorProvider.SetError(birth_year_txt, "Not a leap year. Enter a valid leap year.")
            Return False
        End If

        ' Clear any existing error messages
        errorProvider.SetError(birth_day_txt, "")
        errorProvider.SetError(birth_month_txt, "")
        errorProvider.SetError(birth_year_txt, "")
        Return True
    End Function

    Private Sub clear_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles clear_btn.Click
        ' Clear all input and output textboxes
        birth_day_txt.Text = ""
        birth_month_txt.Text = ""
        birth_year_txt.Text = ""
        age_day_txt.Text = ""
        age_month_txt.Text = ""
        age_year_txt.Text = ""

        ' Clear error providers
        errorProvider.SetError(birth_day_txt, "")
        errorProvider.SetError(birth_month_txt, "")
        errorProvider.SetError(birth_year_txt, "")
    End Sub


    ' ----------------------------------------------- SOLUTION FOR CALCULATOR ---------------------------------------------------------------



    Private Sub scientific_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles scientific_btn.Click
        standard_panel.Hide()
        sci_panel.Show()

    End Sub

    Private Sub standard_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles standard_btn.Click
        sci_panel.Hide()
        standard_panel.Show()
    End Sub

    ' -------------------------------- SCIENTIFIC CALCULATOR SOLUTION --------------------------------------------


    Dim firstnum, secondnum, answer As Double
    Dim opera, ops As String
    Private Sub button_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn9.Click, btn8.Click, btn7.Click, btn6.Click, btn5.Click, btn4.Click, btn3.Click, btn2.Click, btn1.Click, btn0.Click
        Dim b As Button = sender
        If txtDisplay.Text = "0" Then
            txtDisplay.Text = b.Text
        Else
            txtDisplay.Text = txtDisplay.Text + b.Text
        End If
    End Sub
    Private Sub Arithmetic_Operator(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDivide.Click, btnMult.Click, btnSub.Click, btnadd.Click, btnMod.Click, btnExp.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim ops As Button = sender
            firstnum = txtDisplay.Text
            op_label.Text = txtDisplay.Text
            txtDisplay.Text = ""
            opera = ops.Text
            op_label.Text = op_label.Text + "    " + opera
        End If
    End Sub



    Private Sub btnEquals_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEquals.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        ElseIf opera = "" Then
            MessageBox.Show("Click on an Operator", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            secondnum = txtDisplay.Text
            If opera = "+" Then
                answer = firstnum + secondnum
                txtDisplay.Text = answer
                op_label.Text = ""
            ElseIf opera = "-" Then
                answer = firstnum - secondnum
                txtDisplay.Text = answer
                op_label.Text = ""

            ElseIf opera = "x" Then
                answer = firstnum * secondnum
                txtDisplay.Text = answer
                op_label.Text = ""

            ElseIf opera = "/" Then
                answer = firstnum / secondnum
                txtDisplay.Text = answer
                op_label.Text = ""
            ElseIf opera = "Mod" Then
                answer = firstnum Mod secondnum
                txtDisplay.Text = answer
                op_label.Text = ""
            ElseIf opera = "Exp" Then
                answer = firstnum ^ secondnum
                txtDisplay.Text = answer
                op_label.Text = ""

            End If
        End If
    End Sub

    Private Sub back_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles back_btn.Click
        If txtDisplay.Text.Length > 0 Then
            txtDisplay.Text = txtDisplay.Text.Remove(txtDisplay.Text.Length - 1, 1)
        End If
    End Sub


    Private Sub btnDot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles back_btn.Click
        If InStr(txtDisplay.Text, ".") = 0 Then
            txtDisplay.Text = txtDisplay.Text + "."
        End If
    End Sub
    Private Sub clear_sci_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clear_sci_btn.Click
        txtDisplay.Text = "0"
        op_label.Text = ""
    End Sub


    Private Sub btnPM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPM.Click
        If (txtDisplay.Text.Contains("-")) Then
            txtDisplay.Text = txtDisplay.Text.Remove(0, 1)
        Else
            txtDisplay.Text = "-" + txtDisplay.Text
        End If
    End Sub

    Private Sub clear_2_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clear_2_btn.Click
        txtDisplay.Text = "0"
        op_label.Text = ""
    End Sub

    Private Sub Button82_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPi.Click
        txtDisplay.Text = "3.141592653589976323"
    End Sub

    Private Sub btnLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLog.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim ilog As Double

            ilog = Double.Parse(txtDisplay.Text)
            op_label.Text = System.Convert.ToString("log" + "(" + (txtDisplay.Text) + ")")
            ilog = Math.Log10(ilog)
            txtDisplay.Text = System.Convert.ToString(ilog)
        End If
    End Sub

    Private Sub btnSqrt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSqrt.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim isqrt As Double

            isqrt = Double.Parse(txtDisplay.Text)
            op_label.Text = System.Convert.ToString("sqrt" + "(" + (txtDisplay.Text) + ")")
            isqrt = Math.Sqrt(isqrt)
            txtDisplay.Text = System.Convert.ToString(isqrt)
        End If
    End Sub

    Private Sub btnSinh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSinh.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim iSinh As Double

            iSinh = Double.Parse(txtDisplay.Text)
            op_label.Text = System.Convert.ToString("Sinh" + "(" + (txtDisplay.Text) + ")")
            iSinh = Math.Sinh(iSinh)
            txtDisplay.Text = System.Convert.ToString(iSinh)
        End If
    End Sub

    Private Sub btnCosh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCosh.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim iCosh As Double

            iCosh = Double.Parse(txtDisplay.Text)
            op_label.Text = System.Convert.ToString("Cosh" + "(" + (txtDisplay.Text) + ")")
            iCosh = Math.Cosh(iCosh)
            txtDisplay.Text = System.Convert.ToString(iCosh)
        End If
    End Sub

    Private Sub btnTanh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTanh.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim iTanh As Double

            iTanh = Double.Parse(txtDisplay.Text)
            op_label.Text = System.Convert.ToString("Tanh" + "(" + (txtDisplay.Text) + ")")
            iTanh = Math.Tanh(iTanh)
            txtDisplay.Text = System.Convert.ToString(iTanh)
        End If
    End Sub

    Private Sub btnSin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSin.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim iSin As Double

            iSin = Double.Parse(txtDisplay.Text)
            op_label.Text = System.Convert.ToString("Sin" + "(" + (txtDisplay.Text) + ")")
            iSin = Math.Sin(iSin)
            txtDisplay.Text = System.Convert.ToString(iSin)
        End If
    End Sub

    Private Sub btnCos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCos.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim iCos As Double

            iCos = Double.Parse(txtDisplay.Text)
            op_label.Text = System.Convert.ToString("Cos" + "(" + (txtDisplay.Text) + ")")
            iCos = Math.Cos(iCos)
            txtDisplay.Text = System.Convert.ToString(iCos)
        End If
    End Sub

    Private Sub btnTan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTan.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim iTan As Double

            iTan = Double.Parse(txtDisplay.Text)
            op_label.Text = System.Convert.ToString("Tan" + "(" + (txtDisplay.Text) + ")")
            iTan = Math.Tan(iTan)
            txtDisplay.Text = System.Convert.ToString(iTan)
        End If
    End Sub

    Private Sub btnHexadecimal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHexadecimal.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim a As Integer = Integer.Parse(txtDisplay.Text)
            txtDisplay.Text = System.Convert.ToString(a, 16)
        End If
    End Sub

    Private Sub btnBinary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBinary.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim a As Integer = Integer.Parse(txtDisplay.Text)
            txtDisplay.Text = System.Convert.ToString(a, 2)
        End If
    End Sub

    Private Sub btnDecimal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDecimal.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim a As Integer = Integer.Parse(txtDisplay.Text)
            txtDisplay.Text = System.Convert.ToString(a, 10)
        End If
    End Sub

    Private Sub btnOct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOct.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        Else
            Dim a As Integer = Integer.Parse(txtDisplay.Text)
            txtDisplay.Text = System.Convert.ToString(a, 8)
        End If
    End Sub

    Private Sub btnSqr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSqr.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim a As Double
            a = Convert.ToDouble(txtDisplay.Text) * Convert.ToDouble(txtDisplay.Text)
            txtDisplay.Text = System.Convert.ToString(a)
        End If
    End Sub

    Private Sub btnCube_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCube.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim a As Double
            a = Convert.ToDouble(txtDisplay.Text) * Convert.ToDouble(txtDisplay.Text) * Convert.ToDouble(txtDisplay.Text)
            txtDisplay.Text = System.Convert.ToString(a)
        End If
    End Sub

    Private Sub btnPercent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPercent.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim a As Double
            a = Convert.ToDouble(txtDisplay.Text) / Convert.ToDouble(100)
            txtDisplay.Text = System.Convert.ToString(a)
        End If
    End Sub

    Private Sub btnlnverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnlnverse.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim a As Double
            a = Convert.ToDouble(1.0 / Convert.ToDouble(txtDisplay.Text))
            txtDisplay.Text = System.Convert.ToString(a)
        End If
    End Sub

    Private Sub log_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles log.Click
        If txtDisplay.Text = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim ilog As Double

            ilog = Double.Parse(txtDisplay.Text)
            op_label.Text = System.Convert.ToString("log" + "(" + (txtDisplay.Text) + ")")
            ilog = Math.Log(ilog)
            txtDisplay.Text = System.Convert.ToString(ilog)
        End If
    End Sub

    '------------------------------------- STANDARD CALCULATOR SOLUTION -----------------------------------------------------------------------


    Private currentNumber As String = ""

    Private currentOperation As String = ""
    Private Sub btn_Digit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn00.Click, btn01.Click, btn02.Click, btn03.Click, btn04.Click, btn05.Click, btn06.Click, btn07.Click, btn08.Click, btn09.Click
        Dim btn As Button = CType(sender, Button)
        currentNumber &= btn.Text
        displayTxt.Text = currentNumber
    End Sub
    Private Sub btnOperation_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_add.Click, btn_sub.Click, btn_div.Click, btn_mult.Click
        If currentNumber = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim btn As Button = CType(sender, Button)
            currentOperation = btn.Text
            result = Double.Parse(currentNumber)
            currentNumber = ""
        End If

    End Sub
    Private Sub btn_answer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_answer.Click
        If currentNumber = "" Or currentOperation = "" Then
            MessageBox.Show("This Field cannot be empty", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else

            Dim secondNumber As Double = Double.Parse(currentNumber)

            If currentOperation = "+" Then
                result += secondNumber
            ElseIf currentOperation = "-" Then
                result -= secondNumber
            ElseIf currentOperation = "x" Then
                result *= secondNumber
            ElseIf currentOperation = "/" Then
                result /= secondNumber
            Else
                MessageBox.Show("Kindly click on an Operator", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            displayTxt.Text = result.ToString()
            currentNumber = ""

        End If
    End Sub

    Private Sub percentBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles percentBtn.Click
        If currentNumber = "" Then
            MessageBox.Show("Enter a number first", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return

        Else
            Dim number As Double = Double.Parse(currentNumber)
            result = number / 100
            displayTxt.Text = result.ToString
            currentNumber = ""
        End If
    End Sub

    Private Sub clear_stand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clear_stand.Click
        displayTxt.Text = ""
        currentNumber = ""
        result = 0
        currentOperation = ""
    End Sub

    Private Sub clear_ce_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clear_ce.Click
        displayTxt.Text = ""
        currentNumber = ""
        result = 0
        currentOperation = ""
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        If displayTxt.Text.Length > 0 Then
            displayTxt.Text = displayTxt.Text.Remove(displayTxt.Text.Length - 1, 1)
        End If
    End Sub
End Class