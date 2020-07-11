Public Class InputValidation
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ErrorProvider1.Clear()




        Dim tbx As TextBox
        Dim cmb As ComboBox
        Dim ctrl As Control
        Dim TbxList As New List(Of Control)

        For Each ctrl In Me.TabPageBasicDetails.Controls
            If TypeOf ctrl Is TextBox Then
                tbx = CType(ctrl, TextBox)
                If String.IsNullOrEmpty(tbx.Text) AndAlso tbx.Tag <> "" Then
                    TbxList.Add(tbx)
                    tbx.BackColor = Color.MistyRose
                    ErrorProvider1.SetIconAlignment(tbx, ErrorIconAlignment.MiddleRight)
                    ErrorProvider1.SetIconPadding(tbx, -20)
                    ErrorProvider1.SetError(tbx, "Please enter " & tbx.Tag)
                Else
                    tbx.BackColor = Color.White
                End If
            ElseIf TypeOf ctrl Is ComboBox Then
                cmb = CType(ctrl, ComboBox)
                If String.IsNullOrEmpty(cmb.Text) AndAlso cmb.Tag <> "" Then
                    TbxList.Add(cmb)
                    cmb.BackColor = Color.MistyRose
                    ErrorProvider1.SetIconAlignment(cmb, ErrorIconAlignment.MiddleRight)
                    ErrorProvider1.SetIconPadding(cmb, -40)
                    ErrorProvider1.SetError(cmb, "Please enter " & cmb.Tag)
                Else
                    cmb.BackColor = Color.White
                End If
            End If
        Next ctrl

        Dim str As String = ""

        For i As Integer = 0 To TbxList.Count - 1
            str &= TbxList(i).Tag.ToString() & Environment.NewLine
        Next i

        If Not String.IsNullOrEmpty(str) Then
            Dim msg As String = "Please enter data for all the required fields below."
            Dim style = MsgBoxStyle.Information Or vbOKOnly

            MsgBox(msg & vbNewLine & StrDup(64, "-") & vbCrLf & str, style)
        Else
            MsgBox("Code A Minute says:" & Environment.NewLine & "Well done!")
        End If
    End Sub

    Private Sub TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles _
        TextBox14.KeyPress,
        TextBox27.KeyPress,
        TextBox12.KeyPress,
        TextBox15.KeyPress,
        TextBox25.KeyPress,
        TextBox22.KeyPress

        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

    End Sub
End Class
