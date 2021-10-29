Module Module1
    Public INSCRIPCIONES(7, 8) As String
    Public INDICE As Byte = 0


    Public Nombre(6)
    Public Carné(6)
    Public Nivel(6)
    Public Carrera(6)
    Public Cuota(6)
    Public Pago(6)
    Public Pparcial(6)
    Public Recargo(6)
    Public Pfinal(6)



    Sub INICIAR()
        Form1.TextBox1.Clear()
        Form1.TextBox2.Clear()
        Form1.ComboBox1.Text = ""
        Form1.ComboBox2.Text = ""
        Form1.ComboBox3.Text = ""
    End Sub
    Sub LIMPIARESTADISTICAS()

        Form1.TextBox6.Clear()
        Form1.TextBox7.Clear()
        Form1.TextBox8.Clear()
        Form1.TextBox9.Clear()
        Form1.TextBox10.Clear()
        Form1.TextBox11.Clear()
        Form1.TextBox12.Clear()
        Form1.TextBox13.Clear()
        Form1.TextBox14.Clear()
        Form1.TextBox15.Clear()
        Form1.TextBox16.Clear()

    End Sub
    Sub SALIR()
        If MsgBox("¿Seguro que Desea Salir?", vbQuestion + vbYesNo, "Salir") = vbYes Then
            End
        Else
            INICIAR()
            LIMPIARESTADISTICAS()
        End If
    End Sub
    Sub LIMPIARMATRIZ()
        Form1.DataGridView1.Rows.Clear()

        INDICE = 0

    End Sub
End Module
