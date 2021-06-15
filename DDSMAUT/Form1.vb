Public Class Form1

    Private Sub btnKriteriaSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKriteriaSave.Click
        Call disabletxt()
        MsgBox("Data Kriteria Telah di Set", vbInformation, "SPK - MAUT")
    End Sub
    Public Sub disabletxt()
        txtc1.Enabled = False
        txtc2.Enabled = False
        txtc3.Enabled = False
        btnKriteriaSave.Visible = False
        btnKriteriaEdit.Visible = True
    End Sub

    Public Sub truetxt()
        txtc1.Enabled = True
        txtc2.Enabled = True
        txtc3.Enabled = True
        btnKriteriaSave.Visible = True
        btnKriteriaEdit.Visible = False
    End Sub

    Public Sub loadapp()
        Call disabletxt()

        'Bobot Kriteria Secara Default'
        txtc1.Text = "0.30"
        txtc2.Text = "0.45"
        txtc3.Text = "0.65"

        txtmxc1.Enabled = False
        txtmxc1.Clear()
        txtmxc2.Enabled = False
        txtmxc2.Clear()
        txtmxc3.Enabled = False
        txtmxc3.Clear()
        txtmnc1.Enabled = False
        txtmnc1.Clear()
        txtmnc2.Enabled = False
        txtmnc2.Clear()
        txtmnc3.Enabled = False
        txtmnc3.Clear()

        txta1c1.Clear()
        txta1c2.Clear()
        txta1c3.Clear()

        txta2c1.Clear()
        txta2c2.Clear()
        txta2c3.Clear()

        txta3c1.Clear()
        txta3c2.Clear()
        txta3c3.Clear()

        txtba1c1.Enabled = False
        txtba1c1.Clear()
        txtba1c2.Enabled = False
        txtba1c2.Clear()
        txtba1c3.Enabled = False
        txtba1c3.Clear()

        txtba2c1.Enabled = False
        txtba2c1.Clear()
        txtba2c2.Enabled = False
        txtba2c2.Clear()
        txtba2c3.Enabled = False
        txtba2c3.Clear()

        txtba3c1.Enabled = False
        txtba3c1.Clear()
        txtba3c2.Enabled = False
        txtba3c2.Clear()
        txtba3c3.Enabled = False
        txtba3c3.Clear()

        txtq1.Enabled = False
        txtq1.Clear()
        txtq2.Enabled = False
        txtq2.Clear()
        txtq3.Enabled = False
        txtq3.Clear()

        txtp1.Enabled = False
        txtp1.Clear()
        txtp2.Enabled = False
        txtp2.Clear()
        txtp3.Enabled = False
        txtp3.Clear()
    End Sub

    Private Sub btnKriteriaEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKriteriaEdit.Click
        Call truetxt()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call loadapp()
        btnreset.Visible = False
    End Sub

    Private Sub btnhitung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnhitung.Click
        Dim maxc1 As Integer
        Dim maxc2 As Integer
        Dim maxc3 As Integer
        Dim minc1 As Integer
        Dim minc2 As Integer
        Dim minc3 As Integer


        'Menggambil nilai terbesar dari baris C1'
        If Val(txta1c1.Text) >= Val(txta2c1.Text) And Val(txta1c1.Text) >= Val(txta3c1.Text) Then
            maxc1 = Val(txta1c1.Text)
            txtmxc1.Text = maxc1
        ElseIf Val(txta2c1.Text) >= Val(txta1c1.Text) And Val(txta2c1.Text) >= Val(txta3c1.Text) Then
            maxc1 = Val(txta2c1.Text)
            txtmxc1.Text = maxc1
        ElseIf Val(txta3c1.Text) >= Val(txta1c1.Text) And Val(txta3c1.Text) >= Val(txta2c1.Text) Then
            maxc1 = Val(txta3c1.Text)
            txtmxc1.Text = maxc1
        End If

        'Menggambil nilai terkecil dari baris C1'
        If Val(txta1c1.Text) <= Val(txta2c1.Text) And Val(txta1c1.Text) <= Val(txta3c1.Text) Then
            minc1 = Val(txta1c1.Text)
            txtmnc1.Text = minc1
        ElseIf Val(txta2c1.Text) <= Val(txta1c1.Text) And Val(txta2c1.Text) <= Val(txta3c1.Text) Then
            minc1 = Val(txta2c1.Text)
            txtmnc1.Text = minc1
        ElseIf Val(txta3c1.Text) <= Val(txta1c1.Text) And Val(txta3c1.Text) <= Val(txta2c1.Text) Then
            minc1 = Val(txta3c1.Text)
            txtmnc1.Text = minc1
        End If

        'Menggambil nilai terbesar dari baris C2'
        If Val(txta1c2.Text) >= Val(txta2c2.Text) And Val(txta1c2.Text) >= Val(txta3c2.Text) Then
            maxc2 = Val(txta1c2.Text)
            txtmxc2.Text = maxc2
        ElseIf Val(txta2c2.Text) >= Val(txta1c2.Text) And Val(txta2c2.Text) >= Val(txta3c2.Text) Then
            maxc2 = Val(txta2c2.Text)
            txtmxc2.Text = maxc2
        ElseIf Val(txta3c2.Text) >= Val(txta1c2.Text) And Val(txta3c2.Text) >= Val(txta2c2.Text) Then
            maxc2 = Val(txta3c2.Text)
            txtmxc2.Text = maxc2
        End If

        'Menggambil nilai terkecil dari baris C2'
        If Val(txta1c2.Text) <= Val(txta2c2.Text) And Val(txta1c2.Text) <= Val(txta3c2.Text) Then
            minc2 = Val(txta1c2.Text)
            txtmnc2.Text = minc2
        ElseIf Val(txta2c2.Text) <= Val(txta1c2.Text) And Val(txta2c2.Text) <= Val(txta3c2.Text) Then
            minc2 = Val(txta2c2.Text)
            txtmnc2.Text = minc2
        ElseIf Val(txta3c2.Text) <= Val(txta1c2.Text) And Val(txta3c2.Text) <= Val(txta2c2.Text) Then
            minc2 = Val(txta3c2.Text)
            txtmnc2.Text = minc2
        End If

        'Menggambil nilai terbesar dari baris C3'
        If Val(txta1c3.Text) >= Val(txta2c3.Text) And Val(txta1c3.Text) >= Val(txta3c3.Text) Then
            maxc3 = Val(txta1c3.Text)
            txtmxc3.Text = maxc3
        ElseIf Val(txta2c3.Text) >= Val(txta1c3.Text) And Val(txta2c3.Text) >= Val(txta3c3.Text) Then
            maxc3 = Val(txta2c3.Text)
            txtmxc3.Text = maxc3
        ElseIf Val(txta3c3.Text) >= Val(txta1c3.Text) And Val(txta3c3.Text) >= Val(txta2c3.Text) Then
            maxc3 = Val(txta3c3.Text)
            txtmxc3.Text = maxc3
        End If

        'Menggambil nilai terkecil dari baris C3'
        If Val(txta1c3.Text) <= Val(txta2c3.Text) And Val(txta1c3.Text) <= Val(txta3c3.Text) Then
            minc3 = Val(txta1c3.Text)
            txtmnc3.Text = minc3
        ElseIf Val(txta2c3.Text) <= Val(txta1c3.Text) And Val(txta2c3.Text) <= Val(txta3c3.Text) Then
            minc3 = Val(txta2c3.Text)
            txtmnc3.Text = minc3
        ElseIf Val(txta3c3.Text) <= Val(txta1c3.Text) And Val(txta3c3.Text) <= Val(txta2c3.Text) Then
            minc3 = Val(txta3c3.Text)
            txtmnc3.Text = minc3
        End If

        Dim ben1c1 As Double
        ben1c1 = Val(txta1c1.Text) / maxc1
        txtba1c1.Text = ben1c1

        Dim ben2c1 As Double
        ben2c1 = Val(txta2c1.Text) / maxc1
        txtba2c1.Text = ben2c1

        Dim ben3c1 As Double
        ben3c1 = Val(txta3c1.Text) / maxc1
        txtba3c1.Text = ben3c1

        Dim ben1c2 As Double
        ben1c2 = Val(txta1c2.Text) / maxc2
        txtba1c2.Text = ben1c2

        Dim ben2c2 As Double
        ben2c2 = Val(txta2c2.Text) / maxc2
        txtba2c2.Text = ben2c2

        Dim ben3c2 As Double
        ben3c2 = Val(txta3c2.Text) / maxc2
        txtba3c2.Text = ben3c2

        Dim ben1c3 As Double
        ben1c3 = Val(txta1c3.Text) / maxc3
        txtba1c3.Text = ben1c3

        Dim ben2c3 As Double
        ben2c3 = Val(txta2c3.Text) / maxc3
        txtba2c3.Text = ben2c3

        Dim ben3c3 As Double
        ben3c3 = Val(txta3c3.Text) / maxc3
        txtba3c3.Text = ben3c3


        Dim a1x As Double
        Dim a1y As Double
        a1x = (0.5 * (txtba1c1.Text * txtc1.Text) + (txtba1c2.Text * txtc2.Text) + (txtba1c3.Text * txtc3.Text))
        a1y = (0.5 * (txtba1c1.Text ^ txtc1.Text) + (txtba1c2.Text ^ txtc2.Text) + (txtba1c3.Text ^ txtc3.Text))
        txtq1.Text = a1x + a1y

        Dim a2x As Double
        Dim a2y As Double
        a2x = (0.5 * (txtba2c1.Text * txtc1.Text) + (txtba2c2.Text * txtc2.Text) + (txtba2c3.Text * txtc3.Text))
        a2y = (0.5 * (txtba2c1.Text ^ txtc1.Text) + (txtba2c2.Text ^ txtc2.Text) + (txtba2c3.Text ^ txtc3.Text))
        txtq2.Text = a2x + a2y

        Dim a3x As Double
        Dim a3y As Double
        a3x = (0.5 * (txtba3c1.Text * txtc1.Text) + (txtba3c2.Text * txtc2.Text) + (txtba3c3.Text * txtc3.Text))
        a3y = (0.5 * (txtba3c1.Text ^ txtc1.Text) + (txtba3c2.Text ^ txtc2.Text) + (txtba3c3.Text ^ txtc3.Text))
        txtq3.Text = a3x + a3y

        'Penentuan peringkat di hasil'
        If txtq1.Text >= txtq2.Text And txtq1.Text >= txtq3.Text Then
            txtp1.Text = 1
        ElseIf txtq1.Text <= txtq2.Text And txtq1.Text >= txtq3.Text Then
            txtp1.Text = 2
        ElseIf txtq1.Text >= txtq2.Text And txtq1.Text <= txtq3.Text Then
            txtp1.Text = 2
        Else
            txtp1.Text = 3
        End If

        If txtq2.Text >= txtq1.Text And txtq2.Text >= txtq3.Text Then
            txtp2.Text = 1
        ElseIf txtq2.Text <= txtq1.Text And txtq2.Text >= txtq3.Text Then
            txtp2.Text = 2
        ElseIf txtq2.Text >= txtq1.Text And txtq2.Text <= txtq3.Text Then
            txtp2.Text = 2
        Else
            txtp2.Text = 3
        End If

        If txtq3.Text >= txtq1.Text And txtq3.Text >= txtq2.Text Then
            txtp3.Text = 1
        ElseIf txtq3.Text <= txtq1.Text And txtq3.Text >= txtq2.Text Then
            txtp3.Text = 2
        ElseIf txtq3.Text >= txtq1.Text And txtq3.Text <= txtq2.Text Then
            txtp3.Text = 2
        Else
            txtp3.Text = 3
        End If

        btnhitung.Visible = False
        btnreset.Visible = True

    End Sub

    Private Sub btnreset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnreset.Click
        Call loadapp()
        MsgBox("Data telah di reset seperti semula", vbInformation, "SPK - MAUT")
        btnhitung.Visible = True
        btnreset.Visible = False
    End Sub
End Class
