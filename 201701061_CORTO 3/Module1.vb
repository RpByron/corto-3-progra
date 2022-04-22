Module Module1


    Public nombre(7)
    Public nit(7)
    Public numerodias(7)
    Public tipohabitacion(7)
    Public sencilla(7)
    Public doble(7)
    Public cabaña(7)
    Public tipopago(7)
    Public efectivo(7)
    Public tajeta(7)
    Public trasferencia(7)
    Public deposito(7)
    Public pagoparcial(7)
    Public pagototal(7)
    Public habitacion(7)

    Public Const Psencilla = 250
    Public Const Pdoble = 400
    Public Const Pcabaña = 650

    Public vector As Byte = 0


    Public I As Byte

    Sub limpiar_entradas()
        Form1.TextBox1.Clear()
        Form1.TextBox2.Clear()
        Form1.TextBox3.Clear()
        Form1.ComboBox1.Text = ""
        Form1.ComboBox2.Text = ""
        Form1.TextBox1.Focus()
    End Sub

    Sub limpiar_vectores()

        Form1.DataGridView1.Rows.Clear()

        vector = 0

        For I = 0 To 6

            nombre(I) = Nothing
            nit(I) = Nothing
            numerodias(I) = Nothing
            tipohabitacion(I) = Nothing
            sencilla(I) = Nothing
            doble(I) = Nothing
            cabaña(I) = Nothing
            tipopago(I) = Nothing
            efectivo(I) = Nothing
            tajeta(I) = Nothing
            trasferencia(I) = Nothing
            deposito(I) = Nothing
            pagoparcial(I) = Nothing
            pagototal(I) = Nothing
        Next I

    End Sub
End Module
