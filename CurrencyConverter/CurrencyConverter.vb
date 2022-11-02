Imports System.IO
Imports Newtonsoft.Json.Linq

Public Class CurrencyConverter
    Public currencies As New Dictionary(Of String, String)

    Private Sub ExitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButton.Click
        End
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim json As String = File.ReadAllText("currencies.json")
        Dim data = JToken.Parse(json)
        For Each c In data
            Try
                currencies.Add(c("name"), c("cc"))
            Catch ex As Exception

            End Try
            ComboBox1.Items.Add(c("name"))
            ComboBox2.Items.Add(c("name"))
        Next
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
    End Sub

    Private Sub ConvertButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConvertButton.Click
        Dim currency As String = currencies(ComboBox1.SelectedItem)
        Dim TargetCurrency As String = currencies(ComboBox2.SelectedItem)
        Dim amount As String = TextBox1.Text
        Dim json As String = New System.Net.WebClient().DownloadString("https://api.exchangerate.host/convert?from=" & currency & "&to=" & TargetCurrency & "&amount=" & amount)
        Dim parsejson As JObject = JObject.Parse(json)
        Dim result As Decimal = parsejson.SelectToken("result")
        TextBox2.Text = Math.Round(result, 2) & TargetCurrency
        TextBox2.Focus()
    End Sub

    Private Sub TextBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Enter
        TextBox1.SelectAll()
    End Sub


    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        Select Case e.KeyChar
            Case "0" To "9"
            Case "."
            Case Chr(Keys.Back)
            Case Else
                Beep()
                e.Handled = True
        End Select
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            ConvertButton.Enabled = False
        Else
            ConvertButton.Enabled = True
        End If
    End Sub
End Class
