Imports System.IO
Dim strFilename As String

Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    'Get the desired Filename from the user
    strFilename = InputBox("Enter the Filename")
End Sub

Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
    'Constant for number of inputs
    Const intINPUT_SET As Integer = 1

    'Local Variables
    Dim strFirstName As String
    Dim strMiddleName As String
    Dim strLastName As String
    Dim strEmployeeNum As String
    Dim strDepartment As String
    Dim strPhone As String
    Dim strExtension As String
    Dim strEmail As String
    Dim friendFile As StreamWriter

    Try
        'Open the file and Append text
        friendFile = File.AppendText(strFilename)

        For intCount = 1 To intINPUT_SET
            'Get Inputs from User
            strFirstName = txtFirstName.Text
            strMiddleName = txtMiddleName.Text
            strLastName = txtLastName.Text
            strEmployeeNum = txtEmployeeNum.Text
            strDepartment = cboxDepartment.SelectedItem.ToString()
            strPhone = txtPhone.Text
            strExtension = txtExtension.Text
            strEmail = txtEmail.Text

            'Write data to the file
            friendFile.WriteLine(strFirstName)
            friendFile.WriteLine(strMiddleName)
            friendFile.WriteLine(strLastName)
            friendFile.WriteLine(strEmployeeNum)
            friendFile.WriteLine(strDepartment)
            friendFile.WriteLine(strPhone)
            friendFile.WriteLine(strExtension)
            friendFile.WriteLine(strEmail)
        Next

        'Auto-Close File to save data
        friendFile.Close()

        'Update Status Bar
        lblStatus.Text = "Your File " & strFilename & " has been created!"

    Catch ex As Exception
        'Error Message
        MessageBox.Show("Error: This File Cannot Be Created.")
    End Try

End Sub

Friend WithEvents lblStatus As Label

Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
    'Clear All Fields
    txtFirstName.Text = ""
    txtMiddleName.Text = ""
    txtLastName.Text = ""
    txtEmployeeNum.Text = ""
    cboxDepartment.SelectedIndex = -1
    txtPhone.Text = ""
    txtExtension.Text = ""
    txtEmail.Text = ""
    lblStatus.Text = ""
End Sub

Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
    Me.Close()
End Sub

Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

    'Shortcut exit key set to Alt + X
    If e.KeyCode = Keys.X AndAlso e.Modifiers = Keys.Alt Then
        e.Handled = True
        btnExit_Click(sender, e) 'or cmdExit.PerformClick()
    End If

    'Shortcut clear key set to Alt + C
    If e.KeyCode = Keys.C AndAlso e.Modifiers = Keys.Alt Then
        e.Handled = True
        btnClear_Click(sender, e)
    End If

    'Shortcut save key set to Alt + S
    If e.KeyCode = Keys.S AndAlso e.Modifiers = Keys.Alt Then
        e.Handled = True
        btnSave_Click(sender, e)
    End If
End Sub


