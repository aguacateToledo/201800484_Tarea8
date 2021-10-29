Public Class Form1
    Public TIPO As Double, PG As Double, REC As Double
    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub MostrarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MostrarToolStripMenuItem.Click
        Dim MOSTRAR As Byte
        Dim CALCU As Double
        DataGridView1.Rows.Clear()
        If TextBox1.Text = "Básico" Then
            CALCU = 350
        ElseIf TextBox1.Text = "Diversificado" Then
            CALCU = 450
        End If

        For MOSTRAR = 0 To 6
            If Nombre(MOSTRAR) <> Nothing Then
                DataGridView1.Rows.Add(Nombre(MOSTRAR), Carné(MOSTRAR), Nivel(MOSTRAR), Carrera(MOSTRAR), Pago(MOSTRAR), Pparcial(MOSTRAR), Recargo(MOSTRAR), Pfinal(MOSTRAR))
            Else
                Exit For
            End If
        Next MOSTRAR
    End Sub

    Private Sub LimpiarVectoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarVectoresToolStripMenuItem.Click
        Dim LIMPIAR As Byte
        For LIMPIAR = 0 To 6
            Nombre(LIMPIAR) = Nothing
            Carné(LIMPIAR) = Nothing
            Nivel(LIMPIAR) = Nothing
            Carrera(LIMPIAR) = Nothing
            Pago(LIMPIAR) = Nothing
        Next LIMPIAR
        MsgBox("Vectores borrados correctamente", vbInformation)
        DataGridView1.Rows.Clear()
        INDICE = 0
        INICIAR()
    End Sub

    Private Sub LimpiarMatrizToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarMatrizToolStripMenuItem.Click
        LIMPIARMATRIZ()
    End Sub

    Private Sub MostrarResultadosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MostrarResultadosToolStripMenuItem.Click
        Dim R As Byte
        Dim R1 As Byte
        Dim R2 As Byte
        Dim R3 As Byte
        Dim R4 As Byte
        Dim R5 As Byte
        Dim R6 As Byte
        Dim R7 As Byte
        Dim R8 As Byte
        Dim R9 As Byte
        Dim R10 As Byte
        Dim R11 As Byte
        Dim R12 As Byte
        Dim R13 As Byte
        Dim R14 As Byte

        For R = 0 To 6
            If (Pago(R) = "Efectivo" And Nivel(R) = "Básico") Then
                R1 = R1 + 1
            End If
            If (Pago(R) = "Efectivo" And Nivel(R) = "Diversificado") Then
                R2 = R2 + 1
            End If
            If (Pago(R) = "Tarjeta de Crédito" And Nivel(R) = "Básico") Then
                R3 = R3 + 1
            End If
            If (Pago(R) = "Tarjeta de Crédito" And Nivel(R) = "Diversificado") Then
                R4 = R4 + 1
            End If
            If (Pago(R) = "ACH" And Nivel(R) = "Básico") Then
                R5 = R5 + 1
            End If
            If (Pago(R) = "ACH" And Nivel(R) = "Diversificado") Then
                R6 = R6 + 1
            End If
            If (Pago(R) = "Deposito Bancario" And Nivel(R) = "Básico") Then
                R7 = R7 + 1
            End If
            If (Pago(R) = "Deposito Bancario" And Nivel(R) = "Diversificado") Then
                R8 = R8 + 1
            End If
            If (Nivel(R) = "Diversificado") Then
                R9 = R9 + 1
            End If
            If (Nivel(R) = "Básico") Then
                R10 = R10 + 1
            End If
            If (Carrera(R) = "Perito Contador") Then
                R11 = R11 + 1
            End If
            If (Carrera(R) = "Bachillerato") Then
                R12 = R12 + 1
            End If
            If (Carrera(R) = "Electrónica") Then
                R13 = R13 + 1
            End If
            If (Carrera(R) = "Diseño Gráfico") Then
                R14 = R14 + 1
            End If
        Next R
        TextBox6.Text = Val(TextBox3.Text)
        TextBox7.Text = (Str(R1) * 600) + (Str(R2) * 800)
        TextBox8.Text = (Str(R3) * 600) + (Str(R4) * 800) + (0.1 * ((Str(R3) * 600) + (Str(R4) * 800)))
        TextBox9.Text = (Str(R5) * 600) + (Str(R6) * 800)
        TextBox10.Text = (Str(R7) * 600) + (Str(R8) * 800)
        TextBox15.Text = Str(R9) * 800
        TextBox16.Text = Str(R10) * 600
        TextBox11.Text = Str(R14)
        TextBox12.Text = Str(R13)
        TextBox13.Text = Str(R12)
        TextBox14.Text = Str(R11)
    End Sub

    Private Sub LimpiarTotalToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarTotalToolStripMenuItem.Click
        LIMPIARESTADISTICAS()
    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        If MsgBox("¿DESEA SALIR?", vbQuestion + vbYesNo, "salir") = vbYes Then
            Me.Close()
        End If
    End Sub

    Private Sub CalcularToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalcularToolStripMenuItem.Click
        If ComboBox1.Text = "Básico" Then
            TIPO = 350
        ElseIf ComboBox1.Text = "Diversificado" Then
            TIPO = 450
        End If
        If ComboBox1.Text = "Básico" Then
            PG = 250
        ElseIf ComboBox1.Text = "Diversificado" Then
            PG = 350
        End If
        If ComboBox3.Text = "Tarjeta de Crédito" Then
            REC = 0.1 * (PG + TIPO)
        Else
            REC = 0
        End If
        If ComboBox1.Text = "Básico" Then
            ComboBox2.Text = ""
            MsgBox("Se elimino la selección de carrera, No Aplica", vbInformation)
        End If
        If ComboBox1.Text = "" Or ComboBox3.Text = "" Then
            MsgBox("Debe llenar todos los campos", vbExclamation)
            TextBox1.Focus()
        ElseIf TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Debe llenar todos los campos", vbExclamation)
            TextBox1.Focus()

        ElseIf (INDICE <= 6) Then
            Nombre(INDICE) = TextBox1.Text
            Carné(INDICE) = TextBox2.Text
            Nivel(INDICE) = ComboBox1.Text
            Carrera(INDICE) = ComboBox2.Text
            Pago(INDICE) = ComboBox3.Text
            Pparcial(INDICE) = TIPO + PG
            Recargo(INDICE) = REC
            Pfinal(INDICE) = Pparcial(INDICE) + Recargo(INDICE)
            INDICE = INDICE + 1
            INICIAR()
        End If
        If INDICE >= 7 Then
            MsgBox("Registro de Estudiantes Lleno", vbExclamation)
        End If
    End Sub
End Class
