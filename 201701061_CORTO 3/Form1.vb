Public Class Form1
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub EliminarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EliminarToolStripMenuItem.Click
        nombre(vector) = Nothing
        nit(vector) = Nothing
        numerodias(vector) = Nothing
        tipohabitacion(vector) = Nothing
        sencilla(vector) = Nothing
        doble(vector) = Nothing
        cabaña(vector) = Nothing
        tipopago(vector) = Nothing
        efectivo(vector) = Nothing
        tajeta(vector) = Nothing
        trasferencia(vector) = Nothing
        deposito(vector) = Nothing
        pagoparcial(vector) = Nothing
        pagototal(vector) = Nothing

        For I = vector To 6

            nombre(I) = nombre(I + 1)
            nit(I) = nit(I + 1)
            numerodias(I) = numerodias(I + 1)
            tipohabitacion(I) = tipohabitacion(I + 1)
            sencilla(I) = sencilla(vector)
            doble(I) = doble(I + 1)
            cabaña(I) = cabaña(I + 1)
            tipopago(I) = tipopago(I + 1)
            efectivo(I) = efectivo(I + 1)
            tajeta(I) = tajeta(I + 1)
            trasferencia(I) = trasferencia(I + 1)
            deposito(I) = deposito(I + 1)
            pagoparcial(I) = pagoparcial(I + 1)
            pagototal(I) = pagototal(I + 1)
        Next I
        MsgBox("Registro Eliminado exitosamente")

        nombre(vector) = Nothing
        nit(vector) = Nothing
        numerodias(vector) = Nothing
        tipohabitacion(vector) = Nothing
        sencilla(vector) = Nothing
        doble(vector) = Nothing
        cabaña(vector) = Nothing
        tipopago(vector) = Nothing
        efectivo(vector) = Nothing
        tajeta(vector) = Nothing
        trasferencia(vector) = Nothing
        deposito(vector) = Nothing
        pagoparcial(vector) = Nothing
        pagototal(vector) = Nothing
        vector = I
        limpiar_entradas()
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub CálculoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CálculoToolStripMenuItem.Click
        If vector <= 6 Then
            If Val(TextBox1.Text <> "") Then
                nombre(vector) = Val(TextBox1.Text)
            Else
                MsgBox("Error, ingresar nombre")
                TextBox1.Focus() : Exit Sub
            End If
            If (IsNumeric(TextBox2.Text) And Val(TextBox2.Text <> "")) Then
                nit(vector) = Val(TextBox2.Text)
            Else
                MsgBox("Error, ingresar NIT")
                TextBox2.Focus() : Exit Sub
            End If
            If (IsNumeric(TextBox3.Text) And Val(TextBox3.Text <> "")) Then
                numerodias(vector) = Val(TextBox3.Text)
            Else
                MsgBox("Error, ingresar el numero de dias")
                TextBox2.Focus() : Exit Sub
            End If


            numerodias(vector) = TextBox2.Text


            Select Case ComboBox1.SelectedIndex
                Case 0
                    sencilla(vector) = Psencilla
                Case 1
                    doble(vector) = Pdoble
                Case 2
                    cabaña(vector) = Pcabaña

                Case Else
                    MsgBox("Error, no se seleciono ningun tipo")
                    TextBox1.Focus()
            End Select

            tipohabitacion(vector) = ComboBox1.Text


            Select Case ComboBox1.SelectedIndex
                Case 0
                    habitacion(vector) = sencilla
                Case 1
                    habitacion(vector) = doble
                Case 2
                    habitacion(vector) = cabaña
                Case Else
                    MsgBox("Error, no se seleciono ningun tipo")
                    TextBox1.Focus()
            End Select



            Select Case ComboBox2.SelectedIndex
                Case 0
                    pagoparcial(vector) = Val((tipohabitacion(vector)) - (0.15 * Val(tipohabitacion(vector)))) * Val(numerodias(vector))
                Case 1
                    pagoparcial(vector) = Val((tipohabitacion(vector)) + (0.03 * Val(tipohabitacion(vector)))) * Val(numerodias(vector))
                Case 2
                    pagoparcial(vector) = Val((tipohabitacion(vector))) * Val(numerodias(vector))
                Case 3
                    pagoparcial(vector) = Val((tipohabitacion(vector))) * Val(numerodias(vector))
                Case Else
                    MsgBox("Error, no se seleciono ningun tipo")
                    TextBox1.Focus()
            End Select



            Select Case ComboBox2.SelectedIndex
                Case 0
                    tipopago(vector) = efectivo
                Case 1
                    tipopago(vector) = tajeta
                Case 2
                    tipopago(vector) = trasferencia
                Case 3
                    tipopago(vector) = deposito
                Case Else
                    MsgBox("Error, no se seleciono ningun tipo")
                    TextBox1.Focus()
            End Select

            nit(vector) = TextBox3.Text



        End If
    End Sub

    Private Sub ConsultarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultarToolStripMenuItem.Click

        Dim existe As Boolean = False


        I = 0
        While (I <= 6) And Not (existe)

            If (nombre(I) = Val(TextBox4.Text)) Then
                existe = True
            Else
                I = I + 1
            End If
        End While


        If (existe) Then
            MsgBox("Registro Encontrado exitosamente")

            TextBox1.Text = nombre(I)
            TextBox2.Text = Str(nit(I))
            TextBox3.Text = Str(numerodias(I))
            ComboBox1.Text = tipohabitacion(I)
            ComboBox2.Text = tipopago(I)

            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add(Str(nombre(I), Str(nit(I)), Str(numerodias(I)), tipohabitacion(I), tipopago(I)))

            vector = I
        Else
            MsgBox("Carnet no encontrado")
            TextBox4.Focus()
        End If
    End Sub
End Class
