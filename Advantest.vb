'Option Strict On
Imports System.Windows.Forms.DataVisualization.Charting
Public Class Advantest
    Dim meter As NationalInstruments.NI4882.Device
    Dim tipeA As Integer = 1
    Dim l As ListViewItem
    Dim i As Integer
    Dim exc As New Excel.Application
    Dim wbk As Excel.Workbook
    Dim sht As Excel.Worksheet

    Dim InstRead(100) As Double 'tentukan jumlah pengambilan data (tipe A)
    Dim abort As Integer
    Dim reading As String
    Dim iterasi As Integer
    Dim baris As Integer

    Private Property SaveFileDialog1 As Object

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        meter.Write("F1") '-->menentukan chanel


    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        meter.Write("F2") '-->menentukan chanel

    End Sub


    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        meter.Write("F3") '-->menentukan chanel

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim kuadrat As Long
        Button1.Text = "Operate"
        kuadrat = 0
        'If Not Button2.Enabled = False Then
        ListView1.Items.Clear()
        'End If


        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        RadioButton3.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox4.Enabled = True
        TextBox4.Show()
        Label4.Text = "Jml. data terambil"
        Button1.Enabled = False
        Button2.Enabled = True
        Button6.Enabled = False
        'textbox1

        If TextBox1.Text = "" Then
            MsgBox("GPIB address harus di isi!")
            Button6.Enabled = True
            Button1.Enabled = True
            reset()

            Exit Sub
        End If

        meter = New NationalInstruments.NI4882.Device(0, TextBox1.Text, 0, TimeoutValue.T100s) 'masukan alamat GPIB untuk meter
        'combobox1
        If ComboBox1.Text = "60 MHz - 1.5 GHz" Then
            meter.Write("A2") '--->menentukan gate time
        ElseIf ComboBox1.Text = "1.5 GHz- 3 GHz" Then
            meter.Write("A3")
        Else
            MsgBox("Silahkan tentukan range pengukuran")
            Button1.Enabled = True
            Button6.Enabled = True
            reset()
            Exit Sub
        End If
        ' COmbobox 3
        Dim twait As Double
        If ComboBox3.Text = "0.1 ms" Then
            meter.Write("GT1") '--->menentukan gate time
            twait = 0.0001
        ElseIf ComboBox3.Text = "1.0 ms" Then
            meter.Write("GT2")
            twait = 0.001
        ElseIf ComboBox3.Text = "10.0 ms" Then
            meter.Write("GT3")
            twait = 0.01
        ElseIf ComboBox3.Text = "0.1 s" Then
            meter.Write("GT4")
            twait = 0.1
        ElseIf ComboBox3.Text = "1.0 s" Then
            meter.Write("GT5")
            twait = 1.0
        ElseIf ComboBox3.Text = "10.0 s" Then
            meter.Write("GT6")
            twait = 10
        Else
            MsgBox("Silahkan Pilih Nilai resolusi/Gate Time")
            Button6.Enabled = True
            'meter.Write("CONT0")
            Button1.Enabled = True
            reset()
            Exit Sub
        End If

        'textbox 2

        If TextBox2.Text = "" Then
            MsgBox("Silahkan cantumkan jumlah titik ukur")
            Button6.Enabled = True
            'meter.Write("CONT0")
            Button1.Enabled = True
            reset()
            Exit Sub
        End If

        'combobox4

        If ComboBox4.Text = "10 ms" Then
            meter.Write("SR1") '--->menentukan gate time
        ElseIf ComboBox4.Text = "80 ms" Then
            meter.Write("SR2")
        ElseIf ComboBox4.Text = "320 ms" Then
            meter.Write("SR3")
        ElseIf ComboBox4.Text = "2.5 s" Then
            meter.Write("SR4")
        Else
            MsgBox("Silahkan Pilih Nilai sample rate")
            Button6.Enabled = True
            'meter.Write("CONT0")
            Button1.Enabled = True
            reset()
            Exit Sub
        End If

        'meter.Write("SR1") '--->sample rate (10 ms)
        ' meter.Write("SQR ON")
        'meter.Write("CONT1")

        meter.Write("H0") '---> menhilangkan header, agar data terbaca sebagai number




        If ComboBox2.Text = "AC Connection" Then
            meter.Write("B3") '--->menentukan gate time
        ElseIf ComboBox2.Text = "DC Connection" Then
            meter.Write("B2")
        ElseIf RadioButton1.Checked Then
            'ComboBox2.Text = "Auto Connection"
            ComboBox2.Enabled = False

        Else
            MsgBox("Silahkan tentukan jenis koneksi (AC/DC)")
            Button1.Enabled = True
            Button6.Enabled = True
            reset()
            Exit Sub
        End If
        'iterasi

        For Me.baris = 1 To Me.TextBox2.Text
            l = Me.ListView1.Items.Add("")
            For j As Integer = 1 To Me.ListView1.Columns.Count
                l.SubItems.Add("")
            Next
            For Me.iterasi = 1 To tipeA
                If Button2.Enabled = False Then
                    Exit Sub
                End If


                Button4.Enabled = False
                Button3.Enabled = False

                ' meter.Write("Enter Cnt;A$")
                meter.Write("A$") '--->mendisplaykan angka
                ' meter.Write("GOTO 1040")
                'meter.Write("GT5")

                'meter.Write("GOTO 1060")                 
                'meter.Write("END")
                InstRead(0) = Val(meter.ReadString)
                'reading = Val(meter.ReadString)
                'reading = InstRead(0) 'Format(InstRead(0), "##0.000 000" + " Hz")
                reading = Format(InstRead(0), "")
                'reading = Format(reading, "000 000 000 000")
                'TextBox3.Text = reading
                'reading = reading.Replace("F   ", "")
                'reading = Val(meter.ReadString)
                'meter.Write("s5")
                wait(twait)

                ' If Button2.Enabled = True Then
                'Exit Sub
                'Else
                Me.ListView1.Items(baris - 1).SubItems(1).Text() = reading
                ' Me.ListView1.Items(baris - 1).SubItems(0).Text() = baris
                Me.ListView1.Items(baris - 1).SubItems(0).Text() = Date.Now.ToString("HH:mm:ss")
                ' Me.ListView1.Items(baris - 1).SubItems(2).Text() = Date.Now.ToString("dd/MM/yyy")
                'wait(ComboBox3.SelectedItem)
                ' meter.Write("SR5")
                'End If
                TextBox4.Text = baris
                'Dim i As Integer
                ' For i = 1 To TextBox2.Text
                ''    For Each i In ListView1.Items 'To TextBox2.Text


                'TextBox3.Text = Val(ListView1.Items.Item(i - 1).SubItems(1).Text.ToString) + Val(ListView1.Items.Item(i - 2).SubItems(1).Text.ToString)
                'wait(1)
                'TextBox3.Text = Val(TextBox3.Text.ToString) + Val(ListView1.Items.Item(baris - 1).SubItems(1).Text.ToString)
                'TextBox6.Text = Math.Sqrt((Val(TextBox6.Text.ToString) + (Val(ListView1.Items.Item(i - 1).SubItems(1).Text.ToString) - Val(ListView1.Items.Item(i - 2).SubItems(1).Text.ToString)) ^ 2) / (2 * Val(TextBox2.Text - 1)))
                ' TextBox6.Text = v(k)+
                'TextBox6.Text = Val(ListView1.Items.Item(baris - 1).SubItems(1).Text.ToString) + Val(ListView1.Items.Item(baris - 2).SubItems(1).Text.ToString)
                'TextBox6.Text = Val(TextBox6.Text.ToString) + (baris + 1) * 2 - baris

                
                ' Dim i As Long

                ' End If
                'Next i
                ' End With



            Next
        Next
        'meter.Write("A$")
        meter.Write("CONT0") '--> menghentikan pengukuran
        Button4.Enabled = True
        Button3.Enabled = True
        Button1.Enabled = True
        Button6.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        'Dim sigma As String
        TextBox6.Text = ""
        'TextBox5.Text = ""
        'TextBox7.Text = ""
        Dim i As Integer
        'Dim sigma As String
        For i = 2 To TextBox2.Text
          
            TextBox6.Text = (Val(TextBox6.Text.ToString) + (Val(ListView1.Items.Item(i - 1).SubItems(1).Text.ToString) - Val(ListView1.Items.Item(i - 2).SubItems(1).Text.ToString)) ^ 2)

           
        Next
        ' Next
        TextBox6.Text = Math.Sqrt(Val(TextBox6.Text.ToString) / (2 * Val(TextBox2.Text - 1)))

        'menentukan titik minimum
        ' Dim i As Integer

        TextBox5.Text = Val(ListView1.Items(0).SubItems(1).Text.ToString)
        For i = 2 To Val(TextBox2.Text)


            If Val(ListView1.Items(i - 1).SubItems(1).Text.ToString) > Val(ListView1.Items(i - 2).SubItems(1).Text.ToString) And TextBox5.Text > Val(ListView1.Items(i - 2).SubItems(1).Text.ToString) Then
                TextBox5.Text = Val(ListView1.Items(i - 2).SubItems(1).Text.ToString)
            Else
                TextBox5.Text = TextBox5.Text
            End If
        Next i
        'menentukan titik maksimum
        Dim n As Integer
        For n = 1 To TextBox2.Text

            Dim maxvalue As String

            maxvalue = Val(TextBox3.Text.ToString)

            If Val(ListView1.Items(n - 1).SubItems(1).Text.ToString) > maxvalue Then
                maxvalue = Val(ListView1.Items(n - 1).SubItems(1).Text.ToString)
                TextBox3.Text = maxvalue
            End If
        Next n


        MsgBox("Pengukuran Selesai!")
        'Next
        
       

        Exit Sub

    End Sub
    Public Sub wait(ByVal Dt As Double)
        Dim IDay As Double = Date.Now.DayOfYear
        Dim CDay As Double
        Dim ITime As Double = Date.Now.TimeOfDay.TotalSeconds
        Dim CTime As Double
        Dim DiffDay As Double

        Do
            Application.DoEvents()
            CDay = Date.Now.DayOfYear
            CTime = Date.Now.TimeOfDay.TotalSeconds
            DiffDay = CDay - IDay
            CTime = CTime + 86400 * DiffDay
            If CTime >= ITime + Dt Then Exit Do

        Loop
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ListView1.Items.Clear()
        Button2.Enabled = True
        Chart1.Series(0).Points.Clear()
        'Chart1.Series(1).Points.Clear()

        Exit Sub
    End Sub



    'Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
    '   meter.Write("F0") '-->menentukan chanel
    'End Sub






    Private Sub RadioButton1_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        ComboBox2.Text = "Auto Connection"
        ComboBox2.Enabled = False
        ComboBox1.Enabled = True
    End Sub

    Private Sub RadioButton2_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        ComboBox1.Text = "Auto Range"
        ComboBox1.Enabled = False
        ComboBox2.Enabled = True
        ComboBox2.Text = "-AC/DC-"
    End Sub
    Private Sub RadioButton3_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        ComboBox1.Text = "Auto Range"
        ComboBox1.Enabled = False
        ComboBox2.Enabled = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        meter.Write("SP")
        meter.Write("SR5")
        Button2.Enabled = False
        Button4.Enabled = True
        Button3.Enabled = True
        Button1.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True

        Button1.Text = "Return"
        Exit Sub

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView1.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView1.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView1.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = ListView1.Items.Item(i).Text.ToString
            xlWorkSheet.Cells(i + 2, 2) = ListView1.Items.Item(i).SubItems(1).Text
            ' xlWorkSheet.Cells(i + 2, 3) = ListView1.Items.Item(i).SubItems(2).Text
            ' xlWorkSheet.Cells(i + 2, 4) = ListView1.Items.Item(i).SubItems(3).Text

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub


    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Button2.Enabled = False
        Button4.Enabled = True
        Button3.Enabled = True
        Button1.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        'TextBox4.Enabled = False

        RadioButton1.PerformClick()
        TextBox1.Text = 1
        TextBox2.Text = 10
        ComboBox1.Text = "60 MHz - 1.5 GHz"
        ComboBox2.Text = "Auto Connection"
        ComboBox3.Text = "1.0 s"
        ComboBox4.Text = "10 ms"


    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim intv As Double

        TextBox4.Hide()
        Label4.Text = ""
        Label5.Text = Date.Now.ToString("dd/MM/yyyy")
        Me.Location = New Point(CInt((Screen.PrimaryScreen.WorkingArea.Width / 2) - (Me.Width / 2)), CInt((Screen.PrimaryScreen.WorkingArea.Height / 2) - (Me.Height / 2)))

        Chart1.ChartAreas(0).AxisX.Minimum = 0
        Chart1.ChartAreas(0).AxisX.LabelStyle.Angle = 90
        Chart1.ChartAreas(0).AxisY.Interval = 1
        'Chart1.ChartAreas(0).AxisY.Minimum = Val(reading - 1)

        Chart1.ChartAreas(0).AxisX.ScrollBar.Size = 10
        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
        Chart1.Series(0).IsVisibleInLegend = False
        Chart1.ChartAreas(0).AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
        Chart1.ChartAreas(0).AxisX.ScrollBar.IsPositionedInside = True
        Chart1.ChartAreas(0).AxisX.ScrollBar.BackColor = Color.LightGray
        Chart1.ChartAreas(0).AxisX.ScrollBar.ButtonColor = Color.Gray
        'Timer1.Interval = 500
    End Sub

    Private Sub reset()
        Button2.Enabled = False
        Button4.Enabled = True
        Button3.Enabled = True
        Button1.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox4.Hide()
        Label4.Text = ""
    End Sub



    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

   

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        'menentukan titik minimum
        Dim i As Integer

        TextBox5.Text = Val(ListView1.Items(0).SubItems(1).Text.ToString)
        For i = 2 To Val(TextBox2.Text)


            If Val(ListView1.Items(i - 1).SubItems(1).Text.ToString) > Val(ListView1.Items(i - 2).SubItems(1).Text.ToString) And TextBox5.Text > Val(ListView1.Items(i - 2).SubItems(1).Text.ToString) Then
                TextBox5.Text = Val(ListView1.Items(i - 2).SubItems(1).Text.ToString)
            Else
                TextBox5.Text = TextBox5.Text
            End If
        Next i
        'menentukan titik maksimum
        Dim j As Integer
        For j = 1 To TextBox2.Text

            Dim maxvalue As String

            maxvalue = Val(TextBox3.Text.ToString)

            If Val(ListView1.Items(j - 1).SubItems(1).Text.ToString) > maxvalue Then
                maxvalue = Val(ListView1.Items(j - 1).SubItems(1).Text.ToString)
                TextBox3.Text = maxvalue
            End If
        Next j

        'membuat grafik
        Dim k As Integer

        For k = 1 To TextBox2.Text
            Chart1.Series(0).Points.AddXY(ListView1.Items(k - 1).SubItems(0).Text.ToString, ListView1.Items(k - 1).SubItems(1).Text.ToString)
            Chart1.ChartAreas(0).AxisY.Minimum = Val(TextBox5.Text.ToString) - 1
            Chart1.ChartAreas(0).AxisY.Maximum = Val(TextBox3.Text + 1)
            Chart1.ChartAreas("ChartArea1").AxisX.Title = "Waktu Pengambilan Data"
            Chart1.ChartAreas("ChartArea1").AxisY.Title = "frekuensi (Hz)"
            Chart1.ChartAreas(0).AxisX.ScaleView.Size = Val(TextBox2.Text.ToString)
            'Chart1.AutoScaling = True
        Next k
    End Sub
    Dim Rand As New Random
    Dim Counter As Double = 1



    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            .AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            .AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Chart1.ChartAreas(0).AxisX.Interval = Val(TextBox2.Text.ToString) * 10
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
            .AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub
End Class
