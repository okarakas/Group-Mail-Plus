Imports System.Net.Mail
Imports System.Net
Imports System.Management
Imports System.Net.NetworkInformation


Public Class Form1
    Dim w As IO.StreamWriter
    Dim r As IO.StreamReader
    Dim intr As Integer = 11
    Dim mail As Integer = 0
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            Const CS_NOCLOSE As Integer = &H200
            cp.ClassStyle = cp.ClassStyle Or CS_NOCLOSE
            CS_NOCLOSE.ToString.Trim()
            Return cp
        End Get
    End Property

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        RichTextBox1.SelectedText = "<b></b>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        RichTextBox1.SelectedText = "<i></i>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        RichTextBox1.SelectedText = "<u></u>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        RichTextBox1.SelectedText = "<s></s>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        RichTextBox1.SelectedText = "#*NAME*#"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        RichTextBox1.SelectedText = "#*SURNAME*#"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        RichTextBox1.SelectedText = "#*APPEAL*#"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        MenuStrip1.Enabled = False
        GroupBox1.Visible = False
        GroupBox4.Visible = True
        mail = 0
        ProgressBar1.Maximum = ListBox2.Items.Count
        ProgressBar1.Value = 0
        mailgonder.Interval = 100
        mailgonder.Start()
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.SelectedText = "<font color=""" + ColorTranslator.ToHtml(ColorDialog1.Color) + """></font>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        ListBox1.Items.Add(OpenFileDialog1.FileName)
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        RichTextBox1.SelectedText = "<font face=""" + ComboBox2.Text + """></font>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        RichTextBox1.SelectedText = "<font size=""" + NumericUpDown1.Text + """></font>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        RichTextBox1.SelectedText = "<img src=""ENTER-URL"">"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        RichTextBox1.SelectedText = "<a href=""LİNK"">TEXT</a>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        RichTextBox1.SelectedText = "<p align=""left"">" + Chr(13) + Chr(13) + "</p>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        RichTextBox1.SelectedText = "<p align=""center"">" + Chr(13) + Chr(13) + "</p>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        RichTextBox1.SelectedText = "<p align=""right"">" + Chr(13) + Chr(13) + "</p>"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ListBox1.SelectedItem <> "" Then
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
        End If
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        RichTextBox1.SelectedText = "<br />"
        RichTextBox1.Focus()
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Try
            Dim onc As Integer
            onc = MessageBox.Show("Draft will be sent to your own e-mail address as a preview." + Chr(13) + "Do you confirm ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk)
            If onc = 6 Then
                Dim msg As New MailMessage()
                If TabControl1.SelectedIndex = 0 Then
                    msg.IsBodyHtml = False
                ElseIf TabControl1.SelectedIndex = 1 Then
                    msg.IsBodyHtml = True
                End If
                msg.[To].Add(TextBox2.Text) 'Gönderilen Kişi
                msg.From = New MailAddress(TextBox2.Text, TextBox4.Text, System.Text.Encoding.UTF8)
                msg.Subject = "RE: " + TextBox1.Text 'Mesaj Konusu
                If TabControl1.SelectedIndex = 0 Then
                    msg.Body = RichTextBox3.Text 'Mesaj
                ElseIf TabControl1.SelectedIndex = 1 Then

                    If CheckBox1.Checked = True Then
                        msg.Body = RichTextBox1.Text + "<br /><br />______________________________________________________<br/ >" + RichTextBox2.Text 'Mesaj
                    Else
                        msg.Body = RichTextBox1.Text 'Mesaj
                    End If
                End If
                Dim smp As New SmtpClient()
                '----------------------------------
                Dim atcs As Integer = 0
                If ListBox1.Items.Count > 0 Then
                    For atcs = 1 To ListBox1.Items.Count
                        Dim atc As New Mail.Attachment(ListBox1.Items(atcs - 1))
                        msg.Attachments.Add(atc)
                    Next
                End If
                '----------------------------------
                smp.Credentials = New NetworkCredential(TextBox2.Text, TextBox3.Text)
                If ComboBox1.Text = "Manually Entering (SSL Not Supported)" Then
                    smp.Port = TextBox8.Text
                    smp.Host = TextBox7.Text
                    smp.EnableSsl = False
                Else
                    smp.Port = 587
                    smp.Host = ComboBox1.Text
                    smp.EnableSsl = True
                End If
                smp.Send(msg)
                MessageBox.Show("Test Mail Sent to " + TextBox2.Text + ".", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Mail Preview Canceled...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("An Error Occurred While Sending Mail. Please Check Your Settings...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        TextBox1.Clear()
        ListBox1.Items.Clear()
        RichTextBox1.Clear()
        RichTextBox3.Clear()
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            CheckBox1.Enabled = False
            CheckBox1.Checked = False

        ElseIf TabControl1.SelectedIndex = 1 Then
            CheckBox1.Enabled = True
        End If
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Dim cvp As Integer
        cvp = MessageBox.Show("Do You Confirm Data Entry?" + Chr(13) + Chr(13) + "Name: " + TextBox5.Text + Chr(13) + "Gender: " + ComboBox3.Text + Chr(13) + "E-Mail: " + TextBox6.Text, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If cvp = 6 Then
            nm.Items.Add(TextBox5.Text)
            cst.Items.Add(ComboBox3.Text)
            ml.Items.Add(TextBox6.Text)
            ListBox2.Items.Add(TextBox6.Text)
            ListBox3.Items.Add(TextBox5.Text + Chr(9) + Chr(9) + Chr(9) + ComboBox3.Text + Chr(9) + Chr(9) + Chr(9) + TextBox6.Text)
            Dim i As Integer
            w = New IO.StreamWriter(Application.StartupPath + "/nm.dll")
            For i = 0 To nm.Items.Count - 1
                w.WriteLine(nm.Items.Item(i))
            Next
            w.Close()
            w = New IO.StreamWriter(Application.StartupPath + "/cst.dll")
            For i = 0 To cst.Items.Count - 1
                w.WriteLine(cst.Items.Item(i))
            Next
            w.Close()
            w = New IO.StreamWriter(Application.StartupPath + "/ml.dll")
            For i = 0 To ml.Items.Count - 1
                w.WriteLine(ml.Items.Item(i))
            Next
            w.Close()
            MessageBox.Show("The Person You Want To Add Has Been Added To The List...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox5.Clear()
            TextBox6.Clear()
        Else
            MessageBox.Show("Data Entry Canceled ...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        r = New IO.StreamReader(Application.StartupPath + "/mz.dll")
        While (r.Peek() > -1)
            RichTextBox2.Text = RichTextBox2.Text + (r.ReadLine) + Chr(13)
        End While
        r.Close()
        ComboBox4.SelectedIndex = 0
        r = New IO.StreamReader(Application.StartupPath + "/config.dll")
        While (r.Peek() > -1)
            settings.Items.Add(r.ReadLine)
        End While
        r.Close()
        TextBox2.Text = settings.Items(0)
        TextBox3.Text = settings.Items(1)
        TextBox4.Text = settings.Items(2)
        ComboBox1.Text = settings.Items(3)
        TextBox7.Text = settings.Items(4)
        TextBox8.Text = settings.Items(5)
        NumericUpDown3.Text = settings.Items(6)
        r = New IO.StreamReader(Application.StartupPath + "/nm.dll")
        While (r.Peek() > -1)
            nm.Items.Add(r.ReadLine)
        End While
        r.Close()
        r = New IO.StreamReader(Application.StartupPath + "/cst.dll")
        While (r.Peek() > -1)
            cst.Items.Add(r.ReadLine)
        End While
        r.Close()
        r = New IO.StreamReader(Application.StartupPath + "/ml.dll")
        While (r.Peek() > -1)
            ml.Items.Add(r.ReadLine)
        End While
        r.Close()

        Dim z As Integer
        For z = 0 To nm.Items.Count - 1
            ListBox3.Items.Add(nm.Items(z) + Chr(9) + Chr(9) + Chr(9) + cst.Items(z) + Chr(9) + Chr(9) + Chr(9) + ml.Items(z))
            ListBox2.Items.Add(ml.Items(z))
        Next
        ComboBox5.SelectedIndex = 0
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        If ListBox3.SelectedItem <> "" Then
            Dim ind As Integer
            ind = ListBox3.SelectedIndex
            Dim cvp As Integer
            cvp = MessageBox.Show("Do You Confirm Deletion?" + Chr(13) + Chr(13) + "Name: " + nm.Items(ind) + Chr(13) + "Gender: " + cst.Items(ind) + Chr(13) + "E-Mail: " + ml.Items(ind), "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If cvp = 6 Then

                ListBox3.Items.RemoveAt(ind)
                nm.Items.RemoveAt(ind)
                cst.Items.RemoveAt(ind)
                ml.Items.RemoveAt(ind)
                ListBox2.Items.RemoveAt(ind)
                Dim i As Integer
                w = New IO.StreamWriter(Application.StartupPath + "/nm.dll")
                For i = 0 To nm.Items.Count - 1
                    w.WriteLine(nm.Items.Item(i))
                Next
                w.Close()
                w = New IO.StreamWriter(Application.StartupPath + "/cst.dll")
                For i = 0 To cst.Items.Count - 1
                    w.WriteLine(cst.Items.Item(i))
                Next
                w.Close()
                w = New IO.StreamWriter(Application.StartupPath + "/ml.dll")
                For i = 0 To ml.Items.Count - 1
                    w.WriteLine(ml.Items.Item(i))
                Next
                w.Close()
                MessageBox.Show("Registration Successfully Deleted...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Deletion Canceled...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub MailGönderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SendMailToolStripMenuItem.Click
        GroupBox1.Visible = True
        GroupBox2.Visible = False
        GroupBox3.Visible = False
        GroupBox4.Visible = False
        GroupBox5.Visible = False
    End Sub

    Private Sub KişilerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContactsToolStripMenuItem.Click
        GroupBox1.Visible = False
        GroupBox2.Visible = True
        GroupBox3.Visible = False
        GroupBox4.Visible = False
        GroupBox5.Visible = False
    End Sub

    Private Sub AyarlarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettingsToolStripMenuItem.Click
        GroupBox1.Visible = False
        GroupBox2.Visible = False
        GroupBox3.Visible = True
        GroupBox4.Visible = False
        GroupBox5.Visible = False
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        TextBox5.Clear()
        TextBox6.Clear()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "Manually Entering (SSL Not Supported)" Then
            Panel6.Enabled = True
        Else
            Panel6.Enabled = False
            TextBox7.Clear()
            TextBox8.Clear()
        End If
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Panel6.Enabled = False
        Panel7.Enabled = False
        Button26.Enabled = False
        Button27.Enabled = True
        Panel9.Enabled = False
        settings.Items.Add(TextBox2.Text)
        settings.Items.Add(TextBox3.Text)
        settings.Items.Add(TextBox4.Text)
        settings.Items.Add(ComboBox1.Text)
        settings.Items.Add(TextBox7.Text)
        settings.Items.Add(TextBox8.Text)
        settings.Items.Add(NumericUpDown3.Text)
        Dim i As Integer
        w = New IO.StreamWriter(Application.StartupPath + "/config.dll")
        For i = 0 To settings.Items.Count - 1
            w.WriteLine(settings.Items.Item(i))
        Next
        w.Close()
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        settings.Items.Clear()
        If ComboBox1.Text = "Manually Entering (SSL Not Supported)" Then
            Panel6.Enabled = True
        Else
            Panel6.Enabled = False
        End If
        Panel7.Enabled = True
        Button26.Enabled = True
        Button27.Enabled = False
        Panel9.Enabled = True
    End Sub

    Private Sub TextBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.Click
        TextBox2.Clear()

    End Sub

    Private Sub TextBox3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.Click
        TextBox3.Clear()

    End Sub

    Private Sub TextBox4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.Click
        TextBox4.Clear()

    End Sub

    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button38.Click
        If Button38.Text = "Save Signature" Then
            Button38.Text = "Edit Signature"
            Dim i As Integer
            w = New IO.StreamWriter(Application.StartupPath + "/mz.dll")
            For i = 0 To RichTextBox2.Lines.Length - 1
                w.WriteLine(RichTextBox2.Lines(i))
            Next
            w.Close()
            RichTextBox2.Enabled = False
            Panel8.Enabled = False
        Else
            Button38.Text = "Save Signature"
            RichTextBox2.Enabled = True
            Panel8.Enabled = True
        End If
    End Sub

    Private Sub Button44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button44.Click
        RichTextBox2.SelectedText = "<b></b>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click
        RichTextBox2.SelectedText = "<i></i>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click
        RichTextBox2.SelectedText = "<u></u>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button41.Click
        RichTextBox2.SelectedText = "<s></s>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        ColorDialog1.ShowDialog()
        RichTextBox2.SelectedText = "<font color=""" + ColorTranslator.ToHtml(ColorDialog1.Color) + """></font>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        RichTextBox2.SelectedText = "<img src=""ENTER-URL"">"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        RichTextBox2.SelectedText = "<a href=""LİNK"">TEXT</a>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        RichTextBox2.SelectedText = "<font face=""" + ComboBox4.Text + """></font>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        RichTextBox2.SelectedText = "<font size=""" + NumericUpDown2.Text + """></font>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        RichTextBox2.SelectedText = "<p align=""left"">" + Chr(13) + Chr(13) + "</p>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        RichTextBox2.SelectedText = "<p align=""center"">" + Chr(13) + Chr(13) + "</p>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        RichTextBox2.SelectedText = "<p align=""right"">" + Chr(13) + Chr(13) + "</p>"
        RichTextBox2.Focus()
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        RichTextBox2.SelectedText = "<br />"
        RichTextBox2.Focus()
    End Sub

    Private Sub ÇıkışToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Dim cvp As Integer
        cvp = MessageBox.Show("Are You Sure You Want To Close The Program?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If cvp = 6 Then
            Me.Close()
        Else
            MessageBox.Show("Shutdown Process Canceled.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ToolStripLabel1.Text = TimeSerial(Now.Hour, Now.Minute, Now.Second) + " | " + Now.Date
        If ToolStripLabel2.Text = "No Internet Connection..." Then
            intr -= 1
            Me.Text = "No Internet Connection. The program will be shut down after " + intr.ToString + " Seconds..."
            If intr = 0 Then
                Me.Close()
            End If
        Else
            intr = 11
        End If

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            Dim internet = My.Computer.Network.Ping("208.67.222.222")

            If internet = True Then
                ToolStripLabel2.ForeColor = Color.Green
                ToolStripLabel2.Text = "Internet Connection Available..."
            Else
                ToolStripLabel2.ForeColor = Color.Red
                ToolStripLabel2.Text = "No Internet connection..."
            End If
        Catch ex As Exception
            ToolStripLabel2.ForeColor = Color.Red
            ToolStripLabel2.Text = "No Internet connection..."
        End Try
    End Sub

    Private Sub NumericUpDown3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown3.ValueChanged

    End Sub

    Private Sub mailgonder_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mailgonder.Tick
        'Try
        Dim msg As New MailMessage()
        '--HTML Kontrol -----------------------------------------
        If TabControl1.SelectedIndex = 0 Then
            msg.IsBodyHtml = False
        ElseIf TabControl1.SelectedIndex = 1 Then
            msg.IsBodyHtml = True
        End If
        '--HTML Kontrol ------------------------------------------
        msg.[To].Add(ListBox2.Items(mail)) 'Gönderilen Kişi
        msg.From = New MailAddress(TextBox2.Text, TextBox4.Text, System.Text.Encoding.UTF8)
        msg.Subject = TextBox1.Text 'Mesaj Konusu
        '--Özel Alan Değiştirme ------------------------------------------
        If TabControl1.SelectedIndex = 0 Then
            Dim mesajj As String = RichTextBox3.Text
            Dim i As Integer
            For i = 1 To Len(RichTextBox1.Text)
                mesajj = Replace(mesajj, "#*NAME*#", nm.Items(mail), 1, 1, CompareMethod.Text)
                mesajj = Replace(mesajj, "#*SURNAME*#", ml.Items(mail), 1, 1, CompareMethod.Text)
                Dim htp As String
                If cst.Items(mail) = "Man" Then
                    htp = "Mr. "
                Else
                    htp = "Mrs. "
                End If
                mesajj = Replace(mesajj, "#*APPEAL*#", htp, 1, 1, CompareMethod.Text)
            Next
            msg.Body = mesajj 'Mesaj
            '--Özel Alan Değiştirme ------------------------------------------
        ElseIf TabControl1.SelectedIndex = 1 Then
            '--Özel Alan Değiştirme ------------------------------------------
            Dim mesajj As String = RichTextBox1.Text
            Dim i As Integer
            For i = 1 To Len(RichTextBox1.Text)
                mesajj = Replace(mesajj, "#*NAME*#", nm.Items(mail), 1, 1, CompareMethod.Text)
                mesajj = Replace(mesajj, "#*SURNAME*#", ml.Items(mail), 1, 1, CompareMethod.Text)
                Dim htp As String
                If cst.Items(mail) = "Man" Then
                    htp = "Mr. "
                Else
                    htp = "Mrs. "
                End If
                mesajj = Replace(mesajj, "#*APPEAL*#", htp, 1, 1, CompareMethod.Text)
            Next
            '--Özel Alan Değiştirme ------------------------------------------

            '--İmza Ekleme ------------------------------------------
            If CheckBox1.Checked = True Then
                msg.Body = mesajj + "<br /><br />______________________________________________________<br/ >" + RichTextBox2.Text 'Mesaj
            Else
                msg.Body = mesajj 'Mesaj
            End If
            '--İmza Ekleme ------------------------------------------
        End If
        Dim smp As New SmtpClient()
        '--Ekler --------------------------------
        Dim atcs As Integer = 0
        If ListBox1.Items.Count > 0 Then
            For atcs = 1 To ListBox1.Items.Count
                Dim atc As New Mail.Attachment(ListBox1.Items(atcs - 1))
                msg.Attachments.Add(atc)
            Next
        End If
        '--Ekler --------------------------------
        smp.Credentials = New NetworkCredential(TextBox2.Text, TextBox3.Text)
        '--SSL Ayarı --------------------------------
        If ComboBox1.Text = "Manually Entering (SSL Not Supported)" Then
            smp.Port = TextBox8.Text
            smp.Host = TextBox7.Text
            smp.EnableSsl = False
        Else
            smp.Port = 587
            smp.Host = ComboBox1.Text
            smp.EnableSsl = True
        End If
        '--SSL Ayarı --------------------------------
        smp.Send(msg)
        ListBox4.Items.Add(ListBox2.Items(mail))
        'Catch ex As Exception
        'ListBox5.Items.Add(ListBox2.Items(mail))
        'End Try
        ProgressBar1.Value = mail + 1
        Label27.Text = (mail + 1).ToString + "/" + (ListBox2.Items.Count).ToString
        mail += 1

        If mail = ListBox2.Items.Count Then
            mailgonder.Stop()
            mail = 0
            MenuStrip1.Enabled = True
        End If
    End Sub

    Private Sub Button20_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        MenuStrip1.Enabled = False
        GroupBox1.Visible = False
        GroupBox4.Visible = True
        mail = 0
        ProgressBar1.Maximum = ListBox2.Items.Count
        mailgonder.Interval = 1000 * NumericUpDown3.Value
        mailgonder.Start()
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Dim filt As Integer = 0
        ListBox3.Items.Clear()
        ListBox2.Items.Clear()
        For filt = 0 To cst.Items.Count - 1
            If cst.Items(filt) = "Man" Then
                ListBox3.Items.Add(nm.Items(filt) + Chr(9) + Chr(9) + Chr(9) + cst.Items(filt) + Chr(9) + Chr(9) + Chr(9) + ml.Items(filt))
                ListBox2.Items.Add(ml.Items(filt))
            End If
        Next
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        Dim filt As Integer = 0
        ListBox3.Items.Clear()
        ListBox2.Items.Clear()
        For filt = 0 To cst.Items.Count - 1
            If cst.Items(filt) = "Woman" Then
                ListBox3.Items.Add(nm.Items(filt) + Chr(9) + Chr(9) + Chr(9) + cst.Items(filt) + Chr(9) + Chr(9) + Chr(9) + ml.Items(filt))
                ListBox2.Items.Add(ml.Items(filt))
            End If
        Next
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Dim filt As Integer = 0
        ListBox3.Items.Clear()
        ListBox2.Items.Clear()
        For filt = 0 To cst.Items.Count - 1
            ListBox3.Items.Add(nm.Items(filt) + Chr(9) + Chr(9) + Chr(9) + cst.Items(filt) + Chr(9) + Chr(9) + Chr(9) + ml.Items(filt))
            ListBox2.Items.Add(ml.Items(filt))
        Next
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        If ComboBox5.SelectedIndex = 0 Then
            Panel10.Visible = True
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
        ElseIf ComboBox5.SelectedIndex = 1 Then
            Panel10.Visible = False
            Panel11.Visible = True
            Panel12.Visible = False
            Panel13.Visible = False
        ElseIf ComboBox5.SelectedIndex = 2 Then
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = True
            Panel13.Visible = False
        ElseIf ComboBox5.SelectedIndex = 3 Then
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = True
        End If
    End Sub

    Private Sub YardımToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        GroupBox1.Visible = False
        GroupBox2.Visible = False
        GroupBox3.Visible = False
        GroupBox4.Visible = False
        GroupBox5.Visible = True
    End Sub

    Private Sub ToolStripLabel3_Click(sender As Object, e As EventArgs)

    End Sub
End Class
