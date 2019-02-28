Imports System.ComponentModel
Imports System.IO
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ofd As New OpenFileDialog
        If ofd.ShowDialog = DialogResult.OK Then
            TextBox1.Text = ofd.FileName
            Dim x As Byte = 0
            Dim Ex As Byte = 0

            Do
                Dim f As String = Path.GetDirectoryName(ofd.FileName) & "\" & Path.GetFileNameWithoutExtension(ofd.FileName) & "-" & x & Path.GetExtension(ofd.FileName)
                If Not File.Exists(f) Then
                    Ex = 1
                    TextBox3.Text = f
                End If
                x = x + 1
            Loop Until Ex = 1
        End If

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ofd As New OpenFileDialog With {.Filter = "SRT files (*.srt)|*.srt"}
        If ofd.ShowDialog = DialogResult.OK Then
            TextBox2.Text = ofd.FileName
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If FontDialog1.ShowDialog() = DialogResult.OK Then
            Button3.Font = FontDialog1.Font
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim colorDialog As New ColorDialog With {.FullOpen = True}
        If colorDialog.ShowDialog = DialogResult.OK Then
            Label4.BackColor = colorDialog.Color
        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim colorDialog As New ColorDialog With {.FullOpen = True}
        If colorDialog.ShowDialog = DialogResult.OK Then
            Label5.BackColor = colorDialog.Color
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim colorDialog As New ColorDialog With {.FullOpen = True}
        If colorDialog.ShowDialog = DialogResult.OK Then
            Label6.BackColor = colorDialog.Color
        End If
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim colorDialog As New ColorDialog With {.FullOpen = True}
        If colorDialog.ShowDialog = DialogResult.OK Then
            Label7.BackColor = colorDialog.Color
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GroupBox3.Hide()
        ProgressBar1.Hide()
        ProgressBar1.Location = New Point(12, 235)
        Label20.Location = New Point(625, 215)
        GroupBox11.Location = New Point(12, 151)
        GroupBox9.Location = New Point(332, 151)
        TextBox4.Location = New Point(12, 212)
        Button8.Location = New Point(835, 216)
        Me.Height = 285
        Button3.Font = FontDialog1.Font
        Label14.Text = TrackBar1.Value & "%"
        Label15.Text = TrackBar2.Value & "%"
        Label16.Text = TrackBar3.Value & "%"
        Label17.Text = TrackBar4.Value & "%"
        ComboBox1.SelectedIndex = 5

    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            NumericUpDown4.Enabled = False
            NumericUpDown5.Enabled = False
        Else
            NumericUpDown4.Enabled = True
            NumericUpDown5.Enabled = True
        End If
    End Sub
    Private Function GetHEXColor(ByVal A As Byte, ByVal color As Color)
        Dim aColor As Color = Color.FromArgb(((A) / 100) * 255, color.R, color.G, color.B)
        Dim Result As String = "&H00FFFFFF"
        Result = String.Format("&H{0:X2}{1:X2}{2:X2}{3:X2}", aColor.A, aColor.B, aColor.G, aColor.R)
        Return Result
    End Function
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        If Button8.Text = "START" Then

            If Not File.Exists(TextBox1.Text) Then
                MsgBox("Please Select Valid Input Video File")
                Exit Sub
            End If
            If Not File.Exists(TextBox2.Text) Then
                MsgBox("Please Select Valid SRT Subtitle File")
                Exit Sub
            End If
            If File.Exists(TextBox3.Text) Then

                Dim result1 As DialogResult = MessageBox.Show("The Selected Output File Already Exists!!" & vbNewLine &
                                                          "Do You want to replace it??",
                                                      "Replace File",
                                                      MessageBoxButtons.YesNo)
                If Not result1 = DialogResult.Yes Then
                    Exit Sub
                End If
            End If

            Button8.Text = "Cancel"

            TextBox4.Show()
            ProgressBar1.Show()
            Label20.Show()
            ProgressBar1.Value = 0
            GroupBox1.Enabled = False
            GroupBox2.Enabled = False
            GroupBox3.Enabled = False
            GroupBox9.Enabled = False
            GroupBox11.Enabled = False

            Dim InputFile As String = TextBox1.Text
            Dim OutputFile As String = TextBox3.Text
            Dim SRTFile As String = TextBox2.Text.ToLower.Replace("\", "\\\\")
            SRTFile = SRTFile.ToLower.Replace(":", "\\:")
            Dim Fontname As String = FontDialog1.Font.Name
            Dim Fontsize As Byte = FontDialog1.Font.Size
            Dim PrimaryColour As String = GetHEXColor(TrackBar1.Value, Label4.BackColor)
            Dim SecondaryColour As String = GetHEXColor(TrackBar2.Value, Label5.BackColor)
            Dim OutlineColour As String = GetHEXColor(TrackBar3.Value, Label6.BackColor)
            Dim BackColour As String = GetHEXColor(TrackBar4.Value, Label7.BackColor)
            Dim Bold As SByte = IIf(FontDialog1.Font.Bold, -1, 0)
            Dim Italic As SByte = IIf(FontDialog1.Font.Italic, -1, 0)
            Dim Underline As SByte = IIf(FontDialog1.Font.Underline, -1, 0)
            Dim StrikeOut As SByte = IIf(FontDialog1.Font.Strikeout, -1, 0)
            Dim ScaleX As Byte = NumericUpDown8.Value
            Dim ScaleY As Byte = NumericUpDown9.Value
            Dim Spacing As UShort = NumericUpDown6.Value
            Dim Angle As UShort = NumericUpDown7.Value
            Dim BorderStyle As UShort = IIf(RadioButton1.Checked, 1, 3)
            Dim Outline As Single = FormatNumber(NumericUpDown4.Value, 1)
            Dim Shadow As Single = FormatNumber(NumericUpDown5.Value, 1)
            Dim Alignment As Byte = 2
            Dim MarginL As UShort = NumericUpDown1.Value
            Dim MarginR As UShort = NumericUpDown2.Value
            Dim MarginV As UShort = NumericUpDown3.Value


            If Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE").ToString().ToLower.Contains("64") Then
                Arch = 64
            Else
                Arch = 32
            End If

            If TopLeft.Checked Then
                Alignment = 4
            ElseIf TopCenter.Checked Then
                Alignment = 6
            ElseIf TopRight.Checked Then
                Alignment = 7
            ElseIf MiddleLeft.Checked Then
                Alignment = 8
            ElseIf MiddleCenter.Checked Then
                Alignment = 10
            ElseIf MiddleRight.Checked Then
                Alignment = 11
            ElseIf BottomLeft.Checked Then
                Alignment = 1
            ElseIf BottomCenter.Checked Then
                Alignment = 2
            ElseIf BottomRight.Checked Then
                Alignment = 3
            End If

            If CheckBox1.Checked Then
                Command = $"-i ""{InputFile}"" -hide_banner -c:v libx264 -preset {ComboBox1.SelectedItem} -vf subtitles=""{SRTFile}"":force_style=" &
    $"'FontName=""{Fontname}""," &
    $"Fontsize={Fontsize}," &
    $"PrimaryColour=""{PrimaryColour}""," &
    $"SecondaryColour=""{SecondaryColour}""," &
    $"OutlineColour=""{OutlineColour}""," &
    $"BackColour=""{BackColour}""," &
    $"Bold={Bold}," &
    $"Italic={Italic}," &
    $"Underline={Underline}," &
    $"StrikeOut={StrikeOut}," &
    $"ScaleX={ScaleX}," &
    $"ScaleY={ScaleY}," &
    $"Spacing={Spacing}," &
    $"Angle={Angle}," &
    $"BorderStyle={BorderStyle}," &
    $"Outline={Outline}," &
    $"Shadow={Shadow}," &
    $"Alignment={Alignment}," &
    $"MarginL={MarginL}," &
    $"MarginR={MarginR}," &
    $"MarginV={MarginV}' -y ""{OutputFile}"""

            Else
                Command = $"-i ""{InputFile}"" -hide_banner -c:v libx264 -preset {ComboBox1.SelectedItem} -vf subtitles=""{SRTFile}"":force_style=" &
    $"'FontName=""{Fontname}""," &
    $"Fontsize={Fontsize}," &
    $"PrimaryColour=""{PrimaryColour}""," &
    $"SecondaryColour=""{SecondaryColour}""," &
    $"OutlineColour=""{OutlineColour}""," &
    $"BackColour=""{BackColour}""," &
    $"Bold={Bold}," &
    $"Italic={Italic}," &
    $"Underline={Underline}," &
    $"StrikeOut={StrikeOut}," &
    $"Spacing={Spacing}," &
    $"Angle={Angle}," &
    $"BorderStyle={BorderStyle}," &
    $"Outline={Outline}," &
    $"Shadow={Shadow}," &
    $"Alignment={Alignment}," &
    $"MarginL={MarginL}," &
    $"MarginR={MarginR}," &
    $"MarginV={MarginV}' -y ""{OutputFile}"""
            End If

            StartWorker()

        ElseIf Button8.Text = "Cancel" Then
            Try
                cmd.Kill()
            Catch ex As Exception
            End Try

        End If

    End Sub
    Private Sub StartWorker()
        Dim powershellWorker As New BackgroundWorker
        AddHandler powershellWorker.DoWork, AddressOf DoWork

        powershellWorker.RunWorkerAsync()
    End Sub
    Private Sub DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Try
            cmd.Kill()
        Catch ex As Exception
        End Try

        Dim psi As New ProcessStartInfo(Application.StartupPath & "\" & Arch & "\ffmpeg.exe", Command)
        Dim systemencoding As System.Text.Encoding
        System.Text.Encoding.GetEncoding(Globalization.CultureInfo.CurrentUICulture.TextInfo.OEMCodePage)
        With psi
            .UseShellExecute = False
            .RedirectStandardError = True
            .RedirectStandardOutput = True
            .RedirectStandardInput = True
            .CreateNoWindow = True
            .StandardOutputEncoding = systemencoding
            .StandardErrorEncoding = systemencoding
        End With

        cmd = New Process With {.StartInfo = psi, .EnableRaisingEvents = True}
        AddHandler cmd.ErrorDataReceived, AddressOf Async_Data_Received
        AddHandler cmd.OutputDataReceived, AddressOf Async_Data_Received
        cmd.Start()
        cmd.BeginOutputReadLine()
        cmd.BeginErrorReadLine()
        cmd.WaitForExit()

        Invoke(Sub()
                   TextBox4.Hide()
                   ProgressBar1.Hide()
                   Label20.Hide()
                   GroupBox1.Enabled = True
                   GroupBox2.Enabled = True
                   GroupBox3.Enabled = True
                   GroupBox9.Enabled = True
                   GroupBox11.Enabled = True
                   Button8.Text = "START"
               End Sub)

        MsgBox("DONE")
    End Sub

    Private Arch As Byte = 32
    Private Command As String
    Private psi As ProcessStartInfo
    Private cmd As Process
    Private Delegate Sub InvokeWithString(ByVal text As String)
    Private Sub Async_Data_Received(ByVal sender As Object, ByVal e As DataReceivedEventArgs)
        Try
            Me.Invoke(New InvokeWithString(AddressOf Sync_Output), e.Data)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Sync_Output(ByVal text As String)
        TextBox4.Text = text
        Clipboard.SetText(text)
        If text.Contains("Duration:") Then
            TextBox4.Tag = text.Split(",")(0)
            Dim strArr() As String = TextBox4.Tag.Split(" ")
            TextBox4.Tag = GetSeconds(strArr(strArr.Length - 1))
        End If

        Dim time As String
        time = text.Split("=")(5)
        time = time.Split(" ")(0)
        time = GetSeconds(time)

        Dim Speed As String
        Dim strArr1() As String = text.Split("=")
        Speed = strArr1(strArr1.Length - 1)
        Speed = Speed.Split(" ")(0)
        Speed = Speed.TrimEnd("x")
        Label20.Text = "Remaining Time : " & GetTime((TextBox4.Tag - time) / Speed)
        ProgressBar1.Value = (time / TextBox4.Tag) * 100
    End Sub
    Private Function GetTime(ByVal Seconds As Decimal)
        Dim iSecond As Double = Seconds 'Total number of seconds
        Dim iSpan As TimeSpan = TimeSpan.FromSeconds(iSecond)

        Dim txtFormattedTime = iSpan.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                iSpan.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                iSpan.Seconds.ToString.PadLeft(2, "0"c)
        Return txtFormattedTime
    End Function
    Private Function GetSeconds(ByVal txt As String)
        Dim result As Decimal
        result = (txt.Split(":")(0) * 60 * 60) + (txt.Split(":")(1) * 60) + txt.Split(":")(2)
        Return result
    End Function
    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Label14.Text = TrackBar1.Value & "%"
    End Sub
    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll
        Label15.Text = TrackBar2.Value & "%"
    End Sub
    Private Sub TrackBar3_Scroll(sender As Object, e As EventArgs) Handles TrackBar3.Scroll
        Label16.Text = TrackBar3.Value & "%"
    End Sub
    Private Sub TrackBar4_Scroll(sender As Object, e As EventArgs) Handles TrackBar4.Scroll
        Label17.Text = TrackBar4.Value & "%"
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim sfd As New SaveFileDialog With {
            .Filter = "AVI files (*.avi)|*.avi|" &
            "FLV files (*.flv)|*.flv|" &
            "M4V files (*.m4v)|*.m4v|" &
            "MKV files (*.mkv)|*.mkv|" &
            "MOV files (*.mov)|*.mov|" &
            "MP4 files (*.mp4)|*.mp4|" &
            "FLV files (*.flv)|*.flv|" &
            "MPEG files (*.mpeg)|*.mpeg|" &
            "MPG files (*.mpg)|*.mpg|" &
            "MTS files (*.mts)|*.mts|" &
            "VOB files (*.vob)|*.vob|" &
            "WMV files (*.wmv)|*.wmv",
            .CheckPathExists = True,
            .FileName = IO.Path.GetFileNameWithoutExtension(TextBox1.Text),
            .RestoreDirectory = True,
            .ValidateNames = True,
            .FilterIndex = 6
        }

        If sfd.ShowDialog = DialogResult.OK Then
            TextBox3.Text = sfd.FileName
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked Then
            MsgBox("BE CAREFULL ..." & vbNewLine &
                   "If You use Scaling the Burning Process will be Very Slow." & vbNewLine &
                   "Use ONLY if you really want it.")
            NumericUpDown8.Enabled = True
            NumericUpDown9.Enabled = True
        Else
            NumericUpDown8.Enabled = False
            NumericUpDown9.Enabled = False
        End If
    End Sub
    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click

        If GroupBox3.Visible = True Then
            Me.Height = 285
            GroupBox3.Hide()
            ProgressBar1.Location = New Point(12, 235)
            Label20.Location = New Point(625, 215)
            GroupBox11.Location = New Point(12, 151)
            GroupBox9.Location = New Point(332, 151)
            TextBox4.Location = New Point(12, 212)
            Button8.Location = New Point(835, 216)
            Label19.Text = "Show Options ------------------------------------------------------------------------------------------------↓↓"
        Else
            Me.Height = 465
            GroupBox3.Show()
            ProgressBar1.Location = New Point(12, 415)
            Label20.Location = New Point(625, 395)
            GroupBox11.Location = New Point(12, 331)
            GroupBox9.Location = New Point(332, 331)
            TextBox4.Location = New Point(12, 392)
            Button8.Location = New Point(835, 396)
            Label19.Text = "Hide Options ------------------------------------------------------------------------------------------------↑↑"
        End If

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("https://github.com/DrAliRagab/Subtitle-Burner")
    End Sub
End Class
