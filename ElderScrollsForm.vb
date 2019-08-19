Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Salud'
        SaludBase.Text = 100 + ((Fuerza.Value) * 5) + ((Resistencia.Value) * 10)
        SaludTotal.Text = SaludBase.Text * ((SaludExtra.Text / 100) + 1)
        'Magia
        MagiaBase.Text = 100 + ((Voluntad.Value) * 5) + ((Inteligencia.Value) * 10)
        MagiaTotal.Text = MagiaBase.Text * ((MagiaExtra.Text / 100) + 1)
        'Aguante
        AguanteBase.Text = 100 + ((Destreza.Value) * 5) + ((Velocidad.Value) * 5) + ((Agilidad.Value) * 5)
        AguanteTotal.Text = AguanteBase.Text * ((AguanteExtra.Text / 100) + 1)
        'Nivel
        NivelAnterior = Nivel.Text
        'Estadisticas
        SBA = SaludBase.Text
        ABA = AguanteBase.Text
        MBA = MagiaBase.Text
        'Atributos
        AA(0) = Fuerza.Value
        AA(1) = Resistencia.Value
        AA(2) = Voluntad.Value
        AA(3) = Destreza.Value
        AA(4) = Agilidad.Value
        AA(5) = Personalidad.Value
        AA(6) = Inteligencia.Value
        AA(7) = Velocidad.Value
        AA(8) = Suerte.Value
        'Puntos
        PuntosAtributo.Text = 10
        PuntosEstadistica.Text = 100
        PuntosHabilidades.Text = 50

    End Sub

    Private Sub Nivel_ValueChanged(sender As Object, e As EventArgs) Handles Nivel.ValueChanged
        If (NivelAnterior >= Nivel.Value) Then
            Nivel.Value = NivelAnterior
        ElseIf (NivelAnterior < Nivel.Value) Then
            PuntosEstadistica.Text = Val(PuntosEstadistica.Text) + (50 * (Nivel.Value - NivelAnterior))
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (20 * (Nivel.Value - NivelAnterior))
            PuntosAtributo.Text = Val(PuntosAtributo.Text) + (2 * (Nivel.Value - NivelAnterior))
        End If
        NivelAnterior = Nivel.Value
    End Sub

    Private Sub Misc_SaludBase_Click(sender As Object, e As EventArgs) Handles Misc_SaludBase.Click
        SaludBase.Text = InputBox("Introduzca la nueva Salud.", "Salud Base", 1)
        If (Val(SaludBase.Text) < (100 + ((Fuerza.Value) * 5) + ((Resistencia.Value) * 10))) Then
            SaludBase.Text = SBA
        ElseIf (Val(SaludBase.Text) >= (100 + ((Fuerza.Value) * 5) + ((Resistencia.Value) * 10))) Then
            PuntosEstadistica.Text = Val(PuntosEstadistica.Text) - (Val(SaludBase.Text) - SBA)
            If (PuntosEstadistica.Text < 0) Then
                PuntosEstadistica.Text = Val(PuntosEstadistica.Text) + (Val(SaludBase.Text) - SBA)
                SaludBase.Text = SBA
            End If
        End If
        SBA = Val(SaludBase.Text)
        SaludTotal.Text = Val(SaludBase.Text) * ((Val(SaludExtra.Text) / 100) + 1)
    End Sub

    Private Sub Misc_MagiaBase_Click(sender As Object, e As EventArgs) Handles Misc_MagiaBase.Click
        MagiaBase.Text = InputBox("Introduzca la nueva Magia.", "Magia Base", 1)
        If (Val(MagiaBase.Text) < (100 + ((Voluntad.Value) * 5) + ((Inteligencia.Value) * 10))) Then
            MagiaBase.Text = MBA
        ElseIf (Val(MagiaBase.Text) >= (100 + ((Voluntad.Value) * 5) + ((Inteligencia.Value) * 10))) Then
            PuntosEstadistica.Text = Val(PuntosEstadistica.Text) - (Val(MagiaBase.Text) - MBA)
            If (PuntosEstadistica.Text < 0) Then
                PuntosEstadistica.Text = Val(PuntosEstadistica.Text) + (Val(MagiaBase.Text) - MBA)
                MagiaBase.Text = MBA
            End If
        End If
        MBA = Val(MagiaBase.Text)
        MagiaTotal.Text = Val(MagiaBase.Text) * ((Val(MagiaExtra.Text) / 100) + 1)
    End Sub

    Private Sub Misc_AguanteBase_Click(sender As Object, e As EventArgs) Handles Misc_AguanteBase.Click
        AguanteBase.Text = InputBox("Introduzca el nuevo Aguante.", "Aguante Base", 1)
        If (Val(AguanteBase.Text) < (100 + ((Destreza.Value) * 5) + ((Velocidad.Value) * 5) + ((Agilidad.Value) * 5))) Then
            AguanteBase.Text = ABA
        ElseIf (Val(AguanteBase.Text) >= (100 + ((Destreza.Value) * 5) + ((Velocidad.Value) * 5) + ((Agilidad.Value) * 5))) Then
            PuntosEstadistica.Text = Val(PuntosEstadistica.Text) - (Val(AguanteBase.Text) - ABA)
            If (PuntosEstadistica.Text < 0) Then
                PuntosEstadistica.Text = Val(PuntosEstadistica.Text) + (Val(AguanteBase.Text) - ABA)
                AguanteBase.Text = ABA
            End If
        End If
        ABA = Val(AguanteBase.Text)
        AguanteTotal.Text = Val(AguanteBase.Text) * ((Val(AguanteExtra.Text) / 100) + 1)
    End Sub

    Private Sub Fuerza_ValueChanged(sender As Object, e As EventArgs) Handles Fuerza.ValueChanged
        If (AA(0) > Fuerza.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
            SaludBase.Text = Val(SaludBase.Text) - 5
        ElseIf (AA(0) < Fuerza.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) - 1
            SaludBase.Text = Val(SaludBase.Text) + 5
            If (PuntosAtributo.Text < 0) Then
                Fuerza.Value = AA(0)
                PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
                SaludBase.Text = Val(SaludBase.Text) - 5
            End If
        End If
        AA(0) = Fuerza.Value
        SBA = SaludBase.Text
        SaludTotal.Text = Val(SaludBase.Text) * ((Val(SaludExtra.Text) / 100) + 1)
    End Sub

    Private Sub Resistencia_ValueChanged(sender As Object, e As EventArgs) Handles Resistencia.ValueChanged
        If (AA(1) > Resistencia.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
            SaludBase.Text = Val(SaludBase.Text) - 10
        ElseIf (AA(1) < Resistencia.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) - 1
            SaludBase.Text = Val(SaludBase.Text) + 10
            If (PuntosAtributo.Text < 0) Then
                Resistencia.Value = AA(1)
                PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
                SaludBase.Text = Val(SaludBase.Text) - 10
            End If
        End If
        AA(1) = Resistencia.Value
        SBA = SaludBase.Text
        SaludTotal.Text = Val(SaludBase.Text) * ((Val(SaludExtra.Text) / 100) + 1)
    End Sub

    Private Sub Voluntad_ValueChanged(sender As Object, e As EventArgs) Handles Voluntad.ValueChanged
        If (AA(2) > Voluntad.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
            MagiaBase.Text = Val(MagiaBase.Text) - 5
        ElseIf (AA(2) < Voluntad.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) - 1
            MagiaBase.Text = Val(MagiaBase.Text) + 5
            If (PuntosAtributo.Text < 0) Then
                Voluntad.Value = AA(2)
                PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
                MagiaBase.Text = Val(MagiaBase.Text) - 5
            End If
        End If
        AA(2) = Voluntad.Value
        MBA = MagiaBase.Text
        MagiaTotal.Text = Val(MagiaBase.Text) * ((Val(MagiaExtra.Text) / 100) + 1)
    End Sub

    Private Sub Destreza_ValueChanged(sender As Object, e As EventArgs) Handles Destreza.ValueChanged
        If (AA(3) > Destreza.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
            AguanteBase.Text = Val(AguanteBase.Text) - 5
        ElseIf (AA(3) < Destreza.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) - 1
            AguanteBase.Text = Val(AguanteBase.Text) + 5
            If (PuntosAtributo.Text < 0) Then
                Destreza.Value = AA(3)
                PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
                AguanteBase.Text = Val(AguanteBase.Text) - 5
            End If
        End If
        AA(3) = Destreza.Value
        ABA = AguanteBase.Text
        AguanteTotal.Text = Val(AguanteBase.Text) * ((Val(AguanteExtra.Text) / 100) + 1)
    End Sub

    Private Sub Agilidad_ValueChanged(sender As Object, e As EventArgs) Handles Agilidad.ValueChanged
        If (AA(4) > Agilidad.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
            AguanteBase.Text = Val(AguanteBase.Text) - 5
        ElseIf (AA(4) < Agilidad.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) - 1
            AguanteBase.Text = Val(AguanteBase.Text) + 5
            If (PuntosAtributo.Text < 0) Then
                Agilidad.Value = AA(4)
                PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
                AguanteBase.Text = Val(AguanteBase.Text) - 5
            End If
        End If
        AA(4) = Agilidad.Value
        ABA = AguanteBase.Text
        AguanteTotal.Text = Val(AguanteBase.Text) * ((Val(AguanteExtra.Text) / 100) + 1)
    End Sub

    Private Sub Velocidad_ValueChanged(sender As Object, e As EventArgs) Handles Velocidad.ValueChanged
        If (AA(7) > Velocidad.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
            AguanteBase.Text = Val(AguanteBase.Text) - 5
        ElseIf (AA(7) < Velocidad.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) - 1
            AguanteBase.Text = Val(AguanteBase.Text) + 5
            If (PuntosAtributo.Text < 0) Then
                Velocidad.Value = AA(7)
                PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
                AguanteBase.Text = Val(AguanteBase.Text) - 5
            End If
        End If
        AA(7) = Velocidad.Value
        ABA = AguanteBase.Text
        AguanteTotal.Text = Val(AguanteBase.Text) * ((Val(AguanteExtra.Text) / 100) + 1)
    End Sub

    Private Sub Inteligencia_ValueChanged(sender As Object, e As EventArgs) Handles Inteligencia.ValueChanged
        If (AA(6) > Inteligencia.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
            MagiaBase.Text = Val(MagiaBase.Text) - 10
        ElseIf (AA(6) < Inteligencia.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) - 1
            MagiaBase.Text = Val(MagiaBase.Text) + 10
            If (PuntosAtributo.Text < 0) Then
                Inteligencia.Value = AA(6)
                PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
                MagiaBase.Text = Val(MagiaBase.Text) - 10
            End If
        End If
        AA(6) = Inteligencia.Value
        MBA = MagiaBase.Text
        MagiaTotal.Text = Val(MagiaBase.Text) * ((Val(MagiaExtra.Text) / 100) + 1)
    End Sub

    Private Sub Personalidad_ValueChanged(sender As Object, e As EventArgs) Handles Personalidad.ValueChanged
        If (AA(5) > Personalidad.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
            BElocuencia.Text = Val(BElocuencia.Text) - 1
            TElocuencia.Text = Val(TElocuencia.Text) - 1
        ElseIf (AA(5) < Personalidad.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) - 1
            BElocuencia.Text = Val(BElocuencia.Text) + 1
            TElocuencia.Text = Val(TElocuencia.Text) + 1
            If (PuntosAtributo.Text < 0) Then
                Personalidad.Value = AA(5)
                PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
                BElocuencia.Text = Val(BElocuencia.Text) - 1
                TElocuencia.Text = Val(TElocuencia.Text) - 1
            End If
        End If
        AA(5) = Personalidad.Value
    End Sub

    Private Sub Suerte_ValueChanged(sender As Object, e As EventArgs) Handles Suerte.ValueChanged
        If (AA(8) > Suerte.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
            BHerreria.Text = Val(BHerreria.Text) - 1
            THerreria.Text = Val(THerreria.Text) - 1
            BEncantamiento.Text = Val(BEncantamiento.Text) - 1
            TEncantamiento.Text = Val(TEncantamiento.Text) - 1
            BAlquimia.Text = Val(BAlquimia.Text) - 1
            TAlquimia.Text = Val(TAlquimia.Text) - 1
            BMedicina.Text = Val(BMedicina.Text) - 1
            TMedicina.Text = Val(TMedicina.Text) - 1
            BACerraduras.Text = Val(BACerraduras.Text) - 1
            TACerraduras.Text = Val(TACerraduras.Text) - 1

        ElseIf (AA(8) < Suerte.Value) Then
            PuntosAtributo.Text = Val(PuntosAtributo.Text) - 1
            BHerreria.Text = Val(BHerreria.Text) + 1
            THerreria.Text = Val(THerreria.Text) + 1
            BEncantamiento.Text = Val(BEncantamiento.Text) + 1
            TEncantamiento.Text = Val(TEncantamiento.Text) + 1
            BAlquimia.Text = Val(BAlquimia.Text) + 1
            TAlquimia.Text = Val(TAlquimia.Text) + 1
            BMedicina.Text = Val(BMedicina.Text) + 1
            TMedicina.Text = Val(TMedicina.Text) + 1
            BACerraduras.Text = Val(BACerraduras.Text) + 1
            TACerraduras.Text = Val(TACerraduras.Text) + 1

            If (PuntosAtributo.Text < 0) Then
                Suerte.Value = AA(8)
                PuntosAtributo.Text = Val(PuntosAtributo.Text) + 1
                BHerreria.Text = Val(BHerreria.Text) - 1
                THerreria.Text = Val(THerreria.Text) - 1
                BEncantamiento.Text = Val(BEncantamiento.Text) - 1
                TEncantamiento.Text = Val(TEncantamiento.Text) - 1
                BAlquimia.Text = Val(BAlquimia.Text) - 1
                TAlquimia.Text = Val(TAlquimia.Text) - 1
                BMedicina.Text = Val(BMedicina.Text) - 1
                TMedicina.Text = Val(TMedicina.Text) - 1
                BACerraduras.Text = Val(BACerraduras.Text) - 1
                TACerraduras.Text = Val(TACerraduras.Text) - 1

            End If
        End If
        AA(8) = Suerte.Value
    End Sub

    Private Sub Misc_SaludExtra_Click(sender As Object, e As EventArgs) Handles Misc_SaludExtra.Click
        SaludExtra.Text = InputBox("Introduzca el PORCENTAGE de Salud extra. También se puede aplicar porcentages negativos.", "Salud Extra", 1)
        SaludTotal.Text = Val(SaludBase.Text) * ((Val(SaludExtra.Text) / 100) + 1)
    End Sub

    Private Sub Misc_MagiaExtra_Click(sender As Object, e As EventArgs) Handles Misc_MagiaExtra.Click
        MagiaExtra.Text = InputBox("Introduzca el PORCENTAGE de Magia extra. También se puede aplicar porcentages negativos.", "Magia Extra", 1)
        MagiaTotal.Text = Val(MagiaBase.Text) * ((Val(MagiaExtra.Text) / 100) + 1)
    End Sub

    Private Sub Misc_AguanteExtra_Click(sender As Object, e As EventArgs) Handles Misc_AguanteExtra.Click
        AguanteExtra.Text = InputBox("Introduza el PORCENTAGE de aguante extra. También se puede aplicar porcentages negativos.", "Aguante Extra", 1)
        AguanteTotal.Text = Val(AguanteBase.Text) * ((Val(AguanteExtra.Text) / 100) + 1)

    End Sub

    Private Sub Misc_SaludTotal_Click(sender As Object, e As EventArgs) Handles Misc_SaludTotal.Click
        Mensaje = InputBox("Introduzca el valor a SUMAR de salud junto al total ya calculado. Si desea volver a calcular el total incicial ponga 0.", "Salud Total Extra", 1)
        If (Mensaje = 0) Then
            SaludTotal.Text = Val(SaludBase.Text) * ((Val(SaludExtra.Text) / 100) + 1)
        ElseIf (Mensaje < 0 Or Mensaje > 0) Then
            SaludTotal.Text = Val(SaludTotal.Text) + Mensaje
        End If
    End Sub

    Private Sub Misc_MagiaTotal_Click(sender As Object, e As EventArgs) Handles Misc_MagiaTotal.Click
        Mensaje = InputBox("Introduzca el valor a SUMAR de magia junto al total ya calculado. Si desea volver a calcular el total incicial ponga 0.", "Magia Total Extra", 1)
        If (Mensaje = 0) Then
            MagiaTotal.Text = Val(MagiaBase.Text) * ((Val(MagiaExtra.Text) / 100) + 1)
        ElseIf (Mensaje < 0 Or Mensaje > 0) Then
            MagiaTotal.Text = Val(MagiaTotal.Text) + Mensaje
        End If
    End Sub

    Private Sub Misc_AguanteTotal_Click(sender As Object, e As EventArgs) Handles Misc_AguanteTotal.Click
        Mensaje = InputBox("Introduzca el valor a SUMAR de aguante junto al total ya calculado. Si desea volver a calcular el total incicial ponga 0.", "Aguante Total Extra", 1)
        If (Mensaje = 0) Then
            AguanteTotal.Text = Val(AguanteBase.Text) * ((Val(AguanteExtra.Text) / 100) + 1)
        ElseIf (Mensaje < 0 Or Mensaje > 0) Then
            AguanteTotal.Text = Val(AguanteTotal.Text) + Mensaje
        End If
    End Sub


    Private Sub NUnaMano_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NUnaMano.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If


    End Sub

    Private Sub NUnaMano_TextChanged(sender As Object, e As EventArgs) Handles NUnaMano.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NUnaMano.Text) - HA(0, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NUnaMano.Text) - HA(0, 0))
            NUnaMano.Text = HA(0, 0)
        End If
        HA(0, 0) = Val(NUnaMano.Text)
        TUnaMano.Text = Val(BUnaMano.Text) + Val(NUnaMano.Text) + Val(MUnaMano.Text)
    End Sub

    Private Sub NAArrojadizas_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NAArrojadizas.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NACerraduras_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NACerraduras.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NALigera_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NALigera.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NAlquimia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NAlquimia.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NAlteracion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NAlteracion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NAMedia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NAMedia.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NAnimales_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NAnimales.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NAPesada_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NAPesada.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NArqueria_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NArqueria.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NBArcano_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NBArcano.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NBloqueo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NBloqueo.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NBMagico_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NBMagico.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NChaman_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NChaman.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NCocina_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NCocina.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NConjuracion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NConjuracion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NDesarmado_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NDesarmado.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NDestruccion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NDestruccion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NDosManos_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NDosManos.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NElocuencia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NElocuencia.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NEncantamiento_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NEncantamiento.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NEsquivar_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NEsquivar.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NFortuna_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NFortuna.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NFrialdad_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NFrialdad.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NHerreria_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NHerreria.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NIlusion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NIlusion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NKiai_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NKiai.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NMedicina_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NMedicina.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NMisticismo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NMisticismo.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NPercepcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NPercepcion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NRestauracion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NRestauracion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NRobo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NRobo.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NSigilo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NSigilo.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NThu_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NThu.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NTMagica_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NTMagica.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NTramperia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NTramperia.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MAArrojadizas_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MAArrojadizas.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MACerraduras_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MACerraduras.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MALigera_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MALigera.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MAlquimia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MAlquimia.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MAlteracion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MAlteracion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MAMedia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MAMedia.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MAnimales_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MAnimales.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MAPesada_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MAPesada.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MArqueria_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MArqueria.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MBArcano_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MBArcano.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MBloqueo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MBloqueo.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MBMagico_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MBMagico.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MChaman_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MChaman.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MCocina_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MCocina.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MConjuracion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MConjuracion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MDesarmado_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MDesarmado.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MDestruccion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MDestruccion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MDosManos_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MDosManos.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MElocuencia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MElocuencia.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MEncantamiento_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MEncantamiento.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MEsquivar_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MEsquivar.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MFortuna_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MFortuna.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MFrialdad_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MFrialdad.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MHerreria_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MHerreria.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MIlusion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MIlusion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MKiai_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MKiai.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MMedicina_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MMedicina.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MMisticismo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MMisticismo.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MPercepcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MPercepcion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MRestauracion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MRestauracion.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MRobo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MRobo.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MSigilo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MSigilo.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MThu_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MThu.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MTMagica_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MTMagica.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MTramperia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MTramperia.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub MUnaMano_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MUnaMano.KeyPress
        If Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 13 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub NDosManos_TextChanged(sender As Object, e As EventArgs) Handles NDosManos.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NDosManos.Text) - HA(1, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NDosManos.Text) - HA(1, 0))
            NDosManos.Text = HA(1, 0)
        End If
        HA(1, 0) = Val(NDosManos.Text)
        TDosManos.Text = Val(BDosManos.Text) + Val(NDosManos.Text) + Val(MDosManos.Text)
    End Sub

    Private Sub NDesarmado_TextChanged(sender As Object, e As EventArgs) Handles NDesarmado.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NDesarmado.Text) - HA(2, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NDesarmado.Text) - HA(2, 0))
            NDesarmado.Text = HA(2, 0)
        End If
        HA(2, 0) = Val(NDesarmado.Text)
        TDesarmado.Text = Val(BDesarmado.Text) + Val(NDesarmado.Text) + Val(MDesarmado.Text)
    End Sub

    Private Sub NArqueria_TextChanged(sender As Object, e As EventArgs) Handles NArqueria.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NArqueria.Text) - HA(3, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NArqueria.Text) - HA(3, 0))
            NArqueria.Text = HA(3, 0)
        End If
        HA(3, 0) = Val(NArqueria.Text)
        TArqueria.Text = Val(NArqueria.Text) + Val(BArqueria.Text) + Val(MArqueria.Text)
    End Sub

    Private Sub NBloqueo_TextChanged(sender As Object, e As EventArgs) Handles NBloqueo.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NBloqueo.Text) - HA(4, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NBloqueo.Text) - HA(4, 0))
            NBloqueo.Text = HA(4, 0)
        End If
        HA(4, 0) = Val(NBloqueo.Text)
        TBloqueo.Text = Val(NBloqueo.Text) + Val(BBloqueo.Text) + Val(MBloqueo.Text)
    End Sub

    Private Sub NAPesada_TextChanged(sender As Object, e As EventArgs) Handles NAPesada.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NAPesada.Text) - HA(5, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NAPesada.Text) - HA(5, 0))
            NAPesada.Text = HA(5, 0)
        End If
        HA(5, 0) = Val(NAPesada.Text)
        TAPesada.Text = Val(NAPesada.Text) + Val(BAPesada.Text) + Val(MAPesada.Text)
    End Sub

    Private Sub NHerreria_TextChanged(sender As Object, e As EventArgs) Handles NHerreria.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NHerreria.Text) - HA(6, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NHerreria.Text) - HA(6, 0))
            NHerreria.Text = HA(6, 0)
        End If
        HA(6, 0) = Val(NHerreria.Text)
        THerreria.Text = Val(NHerreria.Text) + Val(BHerreria.Text) + Val(MHerreria.Text)
    End Sub

    Private Sub NFrialdad_TextChanged(sender As Object, e As EventArgs) Handles NFrialdad.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NFrialdad.Text) - HA(7, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NFrialdad.Text) - HA(7, 0))
            NFrialdad.Text = HA(7, 0)
        End If
        HA(7, 0) = Val(NFrialdad.Text)
        TFrialdad.Text = Val(NFrialdad.Text) + Val(BFrialdad.Text) + Val(MFrialdad.Text)
    End Sub

    Private Sub NAArrojadizas_TextChanged(sender As Object, e As EventArgs) Handles NAArrojadizas.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NAArrojadizas.Text) - HA(8, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NAArrojadizas.Text) - HA(8, 0))
            NAArrojadizas.Text = HA(8, 0)
        End If
        HA(8, 0) = Val(NAArrojadizas.Text)
        TAArrojadizas.Text = Val(NAArrojadizas.Text) + Val(BAArrojadizas.Text) + Val(MAArrojadizas.Text)
    End Sub

    Private Sub NAnimales_TextChanged(sender As Object, e As EventArgs) Handles NAnimales.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NAnimales.Text) - HA(9, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NAnimales.Text) - HA(9, 0))
            NAnimales.Text = HA(9, 0)
        End If
        HA(9, 0) = Val(NAnimales.Text)
        TAnimales.Text = Val(NAnimales.Text) + Val(BAnimales.Text) + Val(MAnimales.Text)
    End Sub

    Private Sub NThu_TextChanged(sender As Object, e As EventArgs) Handles NThu.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NThu.Text) - HA(10, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NThu.Text) - HA(10, 0))
            NThu.Text = HA(10, 0)
        End If
        HA(10, 0) = Val(NThu.Text)
        TThu.Text = Val(NThu.Text) + Val(BThu.Text) + Val(MThu.Text)
    End Sub

    Private Sub NKiai_TextChanged(sender As Object, e As EventArgs) Handles NKiai.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NKiai.Text) - HA(11, 0))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NKiai.Text) - HA(11, 0))
            NKiai.Text = HA(11, 0)
        End If
        HA(11, 0) = Val(NKiai.Text)
        TKiai.Text = Val(NKiai.Text) + Val(BKiai.Text) + Val(MKiai.Text)
    End Sub

    Private Sub NDestruccion_TextChanged(sender As Object, e As EventArgs) Handles NDestruccion.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NDestruccion.Text) - HA(0, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NDestruccion.Text) - HA(0, 1))
            NDestruccion.Text = HA(0, 1)
        End If
        HA(0, 1) = Val(NDestruccion.Text)
        TDestruccion.Text = Val(NDestruccion.Text) + Val(BDestruccion.Text) + Val(MDestruccion.Text)
    End Sub

    Private Sub NConjuracion_TextChanged(sender As Object, e As EventArgs) Handles NConjuracion.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NConjuracion.Text) - HA(1, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NConjuracion.Text) - HA(1, 1))
            NConjuracion.Text = HA(1, 1)
        End If
        HA(1, 1) = Val(NConjuracion.Text)
        TConjuracion.Text = Val(NConjuracion.Text) + Val(BConjuracion.Text) + Val(MConjuracion.Text)
    End Sub

    Private Sub NAlteracion_TextChanged(sender As Object, e As EventArgs) Handles NAlteracion.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NAlteracion.Text) - HA(2, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NAlteracion.Text) - HA(2, 1))
            NAlteracion.Text = HA(2, 1)
        End If
        HA(2, 1) = Val(NAlteracion.Text)
        TAlteracion.Text = Val(NAlteracion.Text) + Val(BAlteracion.Text) + Val(MAlteracion.Text)
    End Sub

    Private Sub NIlusion_TextChanged(sender As Object, e As EventArgs) Handles NIlusion.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NIlusion.Text) - HA(3, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NIlusion.Text) - HA(3, 1))
            NIlusion.Text = HA(3, 1)
        End If
        HA(3, 1) = Val(NIlusion.Text)
        TIlusion.Text = Val(NIlusion.Text) + Val(BIlusion.Text) + Val(MIlusion.Text)
    End Sub

    Private Sub NRestauracion_TextChanged(sender As Object, e As EventArgs) Handles NRestauracion.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NRestauracion.Text) - HA(4, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NRestauracion.Text) - HA(4, 1))
            NRestauracion.Text = HA(4, 1)
        End If
        HA(4, 1) = Val(NRestauracion.Text)
        TRestauracion.Text = Val(BRestauracion.Text) + Val(NRestauracion.Text) + Val(MRestauracion.Text)
    End Sub

    Private Sub NMisticismo_TextChanged(sender As Object, e As EventArgs) Handles NMisticismo.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NMisticismo.Text) - HA(5, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NMisticismo.Text) - HA(5, 1))
            NMisticismo.Text = HA(5, 1)
        End If
        HA(5, 1) = Val(NMisticismo.Text)
        TMisticismo.Text = Val(BMisticismo.Text) + Val(NMisticismo.Text) + Val(MMisticismo.Text)
    End Sub

    Private Sub NBMagico_TextChanged(sender As Object, e As EventArgs) Handles NBMagico.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NBMagico.Text) - HA(6, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NBMagico.Text) - HA(6, 1))
            NBMagico.Text = HA(6, 1)
        End If
        HA(6, 1) = Val(NBMagico.Text)
        TBMagico.Text = Val(BBMagico.Text) + Val(NBMagico.Text) + Val(MBMagico.Text)
    End Sub

    Private Sub NALigera_TextChanged(sender As Object, e As EventArgs) Handles NALigera.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NALigera.Text) - HA(7, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NALigera.Text) - HA(7, 1))
            NALigera.Text = HA(7, 1)
        End If
        HA(7, 1) = Val(NALigera.Text)
        TALigera.Text = Val(BALigera.Text) + Val(NALigera.Text) + Val(MALigera.Text)
    End Sub

    Private Sub NTMagica_TextChanged(sender As Object, e As EventArgs) Handles NTMagica.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NTMagica.Text) - HA(8, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NTMagica.Text) - HA(8, 1))
            NTMagica.Text = HA(8, 1)
        End If
        HA(8, 1) = Val(NTMagica.Text)
        TTMagica.Text = Val(BTMagica.Text) + Val(NTMagica.Text) + Val(MTMagica.Text)
    End Sub

    Private Sub NEncantamiento_TextChanged(sender As Object, e As EventArgs) Handles NEncantamiento.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NEncantamiento.Text) - HA(9, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NEncantamiento.Text) - HA(9, 1))
            NEncantamiento.Text = HA(9, 1)
        End If
        HA(9, 1) = Val(NEncantamiento.Text)
        TEncantamiento.Text = Val(BEncantamiento.Text) + Val(NEncantamiento.Text) + Val(MEncantamiento.Text)
    End Sub

    Private Sub NBArcano_TextChanged(sender As Object, e As EventArgs) Handles NBArcano.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NBArcano.Text) - HA(10, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NBArcano.Text) - HA(10, 1))
            NBArcano.Text = HA(10, 1)
        End If
        HA(10, 1) = Val(NBArcano.Text)
        TBArcano.Text = Val(BBArcano.Text) + Val(NBArcano.Text) + Val(MBArcano.Text)
    End Sub

    Private Sub NChaman_TextChanged(sender As Object, e As EventArgs) Handles NChaman.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NChaman.Text) - HA(11, 1))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NChaman.Text) - HA(11, 1))
            NChaman.Text = HA(11, 1)
        End If
        HA(11, 1) = Val(NChaman.Text)
        TChaman.Text = Val(BChaman.Text) + Val(NChaman.Text) + Val(MChaman.Text)
    End Sub

    Private Sub NAMedia_TextChanged(sender As Object, e As EventArgs) Handles NAMedia.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NAMedia.Text) - HA(0, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NAMedia.Text) - HA(0, 2))
            NAMedia.Text = HA(0, 2)
        End If
        HA(0, 2) = Val(NAMedia.Text)
        TAMedia.Text = Val(BAMedia.Text) + Val(NAMedia.Text) + Val(MAMedia.Text)
    End Sub

    Private Sub NRobo_TextChanged(sender As Object, e As EventArgs) Handles NRobo.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NRobo.Text) - HA(1, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NRobo.Text) - HA(1, 2))
            NRobo.Text = HA(1, 2)
        End If
        HA(1, 2) = Val(NRobo.Text)
        TRobo.Text = Val(BRobo.Text) + Val(NRobo.Text) + Val(MRobo.Text)
    End Sub

    Private Sub NSigilo_TextChanged(sender As Object, e As EventArgs) Handles NSigilo.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NSigilo.Text) - HA(2, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NSigilo.Text) - HA(2, 2))
            NSigilo.Text = HA(2, 2)
        End If
        HA(2, 2) = Val(NSigilo.Text)
        TSigilo.Text = Val(BSigilo.Text) + Val(NSigilo.Text) + Val(MSigilo.Text)
    End Sub

    Private Sub NEsquivar_TextChanged(sender As Object, e As EventArgs) Handles NEsquivar.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NEsquivar.Text) - HA(3, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NEsquivar.Text) - HA(3, 2))
            NEsquivar.Text = HA(3, 2)
        End If
        HA(3, 2) = Val(NEsquivar.Text)
        TEsquivar.Text = Val(BEsquivar.Text) + Val(NEsquivar.Text) + Val(MEsquivar.Text)
    End Sub

    Private Sub NACerraduras_TextChanged(sender As Object, e As EventArgs) Handles NACerraduras.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NACerraduras.Text) - HA(4, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NACerraduras.Text) - HA(4, 2))
            NACerraduras.Text = HA(4, 2)
        End If
        HA(4, 2) = Val(NACerraduras.Text)
        TACerraduras.Text = Val(BACerraduras.Text) + Val(NACerraduras.Text) + Val(MACerraduras.Text)
    End Sub

    Private Sub NElocuencia_TextChanged(sender As Object, e As EventArgs) Handles NElocuencia.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NElocuencia.Text) - HA(5, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NElocuencia.Text) - HA(5, 2))
            NElocuencia.Text = HA(5, 2)
        End If
        HA(5, 2) = Val(NElocuencia.Text)
        TElocuencia.Text = Val(BElocuencia.Text) + Val(NElocuencia.Text) + Val(MElocuencia.Text)
    End Sub

    Private Sub NAlquimia_TextChanged(sender As Object, e As EventArgs) Handles NAlquimia.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NAlquimia.Text) - HA(6, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NAlquimia.Text) - HA(6, 2))
            NAlquimia.Text = HA(6, 2)
        End If
        HA(6, 2) = Val(NAlquimia.Text)
        TAlquimia.Text = Val(BAlquimia.Text) + Val(NAlquimia.Text) + Val(MAlquimia.Text)
    End Sub

    Private Sub NPercepcion_TextChanged(sender As Object, e As EventArgs) Handles NPercepcion.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NPercepcion.Text) - HA(7, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NPercepcion.Text) - HA(7, 2))
            NPercepcion.Text = HA(7, 2)
        End If
        HA(7, 2) = Val(NAlquimia.Text)
        TPercepcion.Text = Val(BPercepcion.Text) + Val(NPercepcion.Text) + Val(MPercepcion.Text)
    End Sub

    Private Sub NMedicina_TextChanged(sender As Object, e As EventArgs) Handles NMedicina.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NMedicina.Text) - HA(8, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NMedicina.Text) - HA(8, 2))
            NMedicina.Text = HA(8, 2)
        End If
        HA(8, 2) = Val(NMedicina.Text)
        TMedicina.Text = Val(BMedicina.Text) + Val(NMedicina.Text) + Val(MMedicina.Text)
    End Sub

    Private Sub NTramperia_TextChanged(sender As Object, e As EventArgs) Handles NTramperia.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NTramperia.Text) - HA(9, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NTramperia.Text) - HA(9, 2))
            NTramperia.Text = HA(9, 2)
        End If
        HA(9, 2) = Val(NTramperia.Text)
        TTramperia.Text = Val(BTramperia.Text) + Val(NTramperia.Text) + Val(MTramperia.Text)
    End Sub

    Private Sub NCocina_TextChanged(sender As Object, e As EventArgs) Handles NCocina.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NCocina.Text) - HA(10, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NCocina.Text) - HA(10, 2))
            NCocina.Text = HA(10, 2)
        End If
        HA(10, 2) = Val(NCocina.Text)
        TCocina.Text = Val(BCocina.Text) + Val(NCocina.Text) + Val(MCocina.Text)
    End Sub

    Private Sub NFortuna_TextChanged(sender As Object, e As EventArgs) Handles NFortuna.TextChanged
        PuntosHabilidades.Text = Val(PuntosHabilidades.Text) - (Val(NFortuna.Text) - HA(11, 2))
        If (PuntosHabilidades.Text < 0) Then
            PuntosHabilidades.Text = Val(PuntosHabilidades.Text) + (Val(NFortuna.Text) - HA(11, 2))
            NFortuna.Text = HA(11, 2)
        End If
        HA(11, 2) = Val(NFortuna.Text)
        TFortuna.Text = Val(BFortuna.Text) + Val(NFortuna.Text) + Val(MFortuna.Text)
    End Sub

    Private Sub MUnaMano_TextChanged(sender As Object, e As EventArgs) Handles MUnaMano.TextChanged
        TUnaMano.Text = Val(BUnaMano.Text) + Val(NUnaMano.Text) + Val(MUnaMano.Text)

    End Sub

    Private Sub MDosManos_TextChanged(sender As Object, e As EventArgs) Handles MDosManos.TextChanged
        TDosManos.Text = Val(BDosManos.Text) + Val(NDosManos.Text) + Val(MDosManos.Text)

    End Sub

    Private Sub MDesarmado_TextChanged(sender As Object, e As EventArgs) Handles MDesarmado.TextChanged
        TDesarmado.Text = Val(BDesarmado.Text) + Val(NDesarmado.Text) + Val(MDesarmado.Text)

    End Sub

    Private Sub MArqueria_TextChanged(sender As Object, e As EventArgs) Handles MArqueria.TextChanged
        TArqueria.Text = Val(MArqueria.Text) + Val(BArqueria.Text) + Val(NArqueria.Text)

    End Sub

    Private Sub MBloqueo_TextChanged(sender As Object, e As EventArgs) Handles MBloqueo.TextChanged
        TBloqueo.Text = Val(BBloqueo.Text) + Val(MBloqueo.Text) + Val(NBloqueo.Text)

    End Sub

    Private Sub MAPesada_TextChanged(sender As Object, e As EventArgs) Handles MAPesada.TextChanged
        TAPesada.Text = Val(BAPesada.Text) + Val(MAPesada.Text) + Val(NAPesada.Text)

    End Sub

    Private Sub MHerreria_TextChanged(sender As Object, e As EventArgs) Handles MHerreria.TextChanged
        THerreria.Text = Val(BHerreria.Text) + Val(NHerreria.Text) + Val(MHerreria.Text)
    End Sub

    Private Sub MFrialdad_TextChanged(sender As Object, e As EventArgs) Handles MFrialdad.TextChanged
        TFrialdad.Text = Val(NFrialdad.Text) + Val(MFrialdad.Text) + Val(BFrialdad.Text)
    End Sub

    Private Sub MAArrojadizas_TextChanged(sender As Object, e As EventArgs) Handles MAArrojadizas.TextChanged
        TAArrojadizas.Text = Val(BAArrojadizas.Text) + Val(NAArrojadizas.Text) + Val(MAArrojadizas.Text)
    End Sub

    Private Sub MAnimales_TextChanged(sender As Object, e As EventArgs) Handles MAnimales.TextChanged
        TAnimales.Text = Val(BAnimales.Text) + Val(MAnimales.Text) + Val(NAnimales.Text)

    End Sub

    Private Sub MThu_TextChanged(sender As Object, e As EventArgs) Handles MThu.TextChanged
        TThu.Text = Val(MThu.Text) + Val(BThu.Text) + Val(NThu.Text)

    End Sub

    Private Sub MKiai_TextChanged(sender As Object, e As EventArgs) Handles MKiai.TextChanged
        TKiai.Text = Val(BKiai.Text) + Val(NKiai.Text) + Val(MKiai.Text)

    End Sub

    Private Sub MDestruccion_TextChanged(sender As Object, e As EventArgs) Handles MDestruccion.TextChanged
        TDestruccion.Text = Val(MDestruccion.Text) + Val(NDestruccion.Text) + Val(BDestruccion.Text)

    End Sub

    Private Sub MConjuracion_TextChanged(sender As Object, e As EventArgs) Handles MConjuracion.TextChanged
        TConjuracion.Text = Val(MConjuracion.Text) + Val(BConjuracion.Text) + Val(NConjuracion.Text)

    End Sub

    Private Sub MAlteracion_TextChanged(sender As Object, e As EventArgs) Handles MAlteracion.TextChanged
        TAlteracion.Text = Val(MAlteracion.Text) + Val(NAlteracion.Text) + Val(BAlteracion.Text)

    End Sub

    Private Sub MIlusion_TextChanged(sender As Object, e As EventArgs) Handles MIlusion.TextChanged
        TIlusion.Text = Val(BIlusion.Text) + Val(NIlusion.Text) + Val(MIlusion.Text)

    End Sub

    Private Sub MRestauracion_TextChanged(sender As Object, e As EventArgs) Handles MRestauracion.TextChanged
        TRestauracion.Text = Val(BRestauracion.Text) + Val(NRestauracion.Text) + Val(MRestauracion.Text)

    End Sub

    Private Sub MMisticismo_TextChanged(sender As Object, e As EventArgs) Handles MMisticismo.TextChanged
        TMisticismo.Text = Val(MMisticismo.Text) + Val(NMisticismo.Text) + Val(BMisticismo.Text)

    End Sub

    Private Sub MBMagico_TextChanged(sender As Object, e As EventArgs) Handles MBMagico.TextChanged
        TBMagico.Text = Val(BBMagico.Text) + Val(NBMagico.Text) + Val(MBMagico.Text)

    End Sub

    Private Sub MALigera_TextChanged(sender As Object, e As EventArgs) Handles MALigera.TextChanged
        TALigera.Text = Val(NALigera.Text) + Val(BALigera.Text) + Val(MALigera.Text)

    End Sub

    Private Sub MTMagica_TextChanged(sender As Object, e As EventArgs) Handles MTMagica.TextChanged
        TTMagica.Text = Val(BTMagica.Text) + Val(NTMagica.Text) + Val(MTMagica.Text)

    End Sub

    Private Sub MEncantamiento_TextChanged(sender As Object, e As EventArgs) Handles MEncantamiento.TextChanged
        TEncantamiento.Text = Val(BEncantamiento.Text) + Val(NEncantamiento.Text) + Val(MEncantamiento.Text)

    End Sub

    Private Sub MBArcano_TextChanged(sender As Object, e As EventArgs) Handles MBArcano.TextChanged
        TBArcano.Text = Val(BBArcano.Text) + Val(NBArcano.Text) + Val(MBArcano.Text)

    End Sub

    Private Sub MChaman_TextChanged(sender As Object, e As EventArgs) Handles MChaman.TextChanged
        TChaman.Text = Val(BChaman.Text) + Val(NChaman.Text) + Val(MChaman.Text)

    End Sub

    Private Sub MAMedia_TextChanged(sender As Object, e As EventArgs) Handles MAMedia.TextChanged
        TAMedia.Text = Val(BAMedia.Text) + Val(NAMedia.Text) + Val(MAMedia.Text)

    End Sub

    Private Sub MRobo_TextChanged(sender As Object, e As EventArgs) Handles MRobo.TextChanged
        TRobo.Text = Val(BRobo.Text) + Val(NRobo.Text) + Val(MRobo.Text)

    End Sub

    Private Sub MSigilo_TextChanged(sender As Object, e As EventArgs) Handles MSigilo.TextChanged
        TSigilo.Text = Val(BSigilo.Text) + Val(NSigilo.Text) + Val(MSigilo.Text)

    End Sub

    Private Sub MEsquivar_TextChanged(sender As Object, e As EventArgs) Handles MEsquivar.TextChanged
        TEsquivar.Text = Val(BEsquivar.Text) + Val(NEsquivar.Text) + Val(MEsquivar.Text)

    End Sub

    Private Sub MACerraduras_TextChanged(sender As Object, e As EventArgs) Handles MACerraduras.TextChanged
        TACerraduras.Text = Val(BACerraduras.Text) + Val(NACerraduras.Text) + Val(MACerraduras.Text)

    End Sub

    Private Sub MElocuencia_TextChanged(sender As Object, e As EventArgs) Handles MElocuencia.TextChanged
        TElocuencia.Text = Val(BElocuencia.Text) + Val(NElocuencia.Text) + Val(MElocuencia.Text)

    End Sub

    Private Sub MAlquimia_TextChanged(sender As Object, e As EventArgs) Handles MAlquimia.TextChanged
        TAlquimia.Text = Val(NAlquimia.Text) + Val(BAlquimia.Text) + Val(MAlquimia.Text)

    End Sub

    Private Sub MPercepcion_TextChanged(sender As Object, e As EventArgs) Handles MPercepcion.TextChanged
        TPercepcion.Text = Val(BPercepcion.Text) + Val(NPercepcion.Text) + Val(MPercepcion.Text)

    End Sub

    Private Sub MMedicina_TextChanged(sender As Object, e As EventArgs) Handles MMedicina.TextChanged
        TMedicina.Text = Val(BMedicina.Text) + Val(NMedicina.Text) + Val(MMedicina.Text)

    End Sub

    Private Sub MTramperia_TextChanged(sender As Object, e As EventArgs) Handles MTramperia.TextChanged
        TTramperia.Text = Val(BTramperia.Text) + Val(NTramperia.Text) + Val(MTramperia.Text)

    End Sub

    Private Sub MCocina_TextChanged(sender As Object, e As EventArgs) Handles MCocina.TextChanged
        TCocina.Text = Val(BCocina.Text) + Val(NCocina.Text) + Val(MCocina.Text)

    End Sub

    Private Sub MFortuna_TextChanged(sender As Object, e As EventArgs) Handles MFortuna.TextChanged
        TFortuna.Text = Val(BFortuna.Text) + Val(NFortuna.Text) + Val(MFortuna.Text)

    End Sub


    Private Sub Raza_TextChanged(sender As Object, e As EventArgs) Handles Raza.TextChanged

        SaludExtra.Text = CSalud
        MagiaExtra.Text = CMagia
        AguanteExtra.Text = CAguante
        RMagia = 0
        RSalud = 0
        RAguante = 0
        SaludTotal.Text = Val(SaludBase.Text) * ((Val(SaludExtra.Text) / 100) + 1)
        MagiaTotal.Text = Val(MagiaBase.Text) * ((Val(MagiaExtra.Text) / 100) + 1)
        AguanteTotal.Text = Val(AguanteBase.Text) * ((Val(AguanteExtra.Text) / 100) + 1)

        'Altmer

        If (Raza.Text = "Altmer") Then
            'Raciales
            TVR1.Text = "Capacidad Altmer"
            VR1.Text = "Tienes 80% de Resistencia a emfermedades y 40% de Magia Extra"
            TVR2.Text = "Paz Mental"
            VR2.Text = "La magia aumenta en 60% y los hechizos de destrucción te hacen un 20% más de daño (Temporal, Escena)"
            TVR3.Text = "Don Arcano"
            VR3.Text = "Los hechizos de destucción, ilusión y misticismo cuestan un 25% menos"
            TVR4.Text = "Maestro Destructor"
            VR4.Text = "Los hechizos de destrucción ocasionan un 25% más de daño"

            'Calculo de magia'
            RMagia = 40
            MagiaExtra.Text = RMagia + CMagia
            MagiaTotal.Text = Val(MagiaBase.Text) * ((Val(MagiaExtra.Text) / 100) + 1)


            'Meritos
            Merito1.Text = "Intelectual"
            Merito1.ReadOnly = True

            'Habilidades base
            BHerreria.Text = 15 + Suerte.Value
            BAPesada.Text = 15
            BBloqueo.Text = 15
            BDosManos.Text = 15
            BUnaMano.Text = 15
            BDesarmado.Text = 15
            BArqueria.Text = 15
            BFrialdad.Text = 20
            BAArrojadizas.Text = 15
            BAnimales.Text = 15
            BThu.Text = 15
            BKiai.Text = 15
            BAMedia.Text = 15
            BEsquivar.Text = 15
            BSigilo.Text = 15
            BACerraduras.Text = 15 + Suerte.Value
            BRobo.Text = 15
            BElocuencia.Text = 15 + Personalidad.Value
            BAlquimia.Text = 15 + Suerte.Value
            BPercepcion.Text = 15
            BMedicina.Text = 15 + Suerte.Value
            BTramperia.Text = 15
            BCocina.Text = 15
            BFortuna.Text = 15
            BALigera.Text = 20
            BIlusion.Text = 20
            BConjuracion.Text = 20
            BDestruccion.Text = 25
            BRestauracion.Text = 20
            BAlteracion.Text = 20
            BEncantamiento.Text = 20 + Suerte.Value
            BTMagica.Text = 20
            BMisticismo.Text = 20
            BBMagico.Text = 20
            BBArcano.Text = 15
            BChaman.Text = 20

            'Habilidades totales
            TFortuna.Text = Val(BFortuna.Text) + Val(NFortuna.Text) + Val(MFortuna.Text)
            TCocina.Text = Val(BCocina.Text) + Val(NCocina.Text) + Val(MCocina.Text)
            TTramperia.Text = Val(BTramperia.Text) + Val(NTramperia.Text) + Val(MTramperia.Text)
            TMedicina.Text = Val(BMedicina.Text) + Val(NMedicina.Text) + Val(MMedicina.Text)
            TPercepcion.Text = Val(BPercepcion.Text) + Val(NPercepcion.Text) + Val(MPercepcion.Text)
            TAlquimia.Text = Val(NAlquimia.Text) + Val(BAlquimia.Text) + Val(MAlquimia.Text)
            TElocuencia.Text = Val(BElocuencia.Text) + Val(NElocuencia.Text) + Val(MElocuencia.Text)
            TACerraduras.Text = Val(BACerraduras.Text) + Val(NACerraduras.Text) + Val(MACerraduras.Text)
            TEsquivar.Text = Val(BEsquivar.Text) + Val(NEsquivar.Text) + Val(MEsquivar.Text)
            TSigilo.Text = Val(BSigilo.Text) + Val(NSigilo.Text) + Val(MSigilo.Text)
            TRobo.Text = Val(BRobo.Text) + Val(NRobo.Text) + Val(MRobo.Text)
            TAMedia.Text = Val(BAMedia.Text) + Val(NAMedia.Text) + Val(MAMedia.Text)
            TChaman.Text = Val(BChaman.Text) + Val(NChaman.Text) + Val(MChaman.Text)
            TBArcano.Text = Val(BBArcano.Text) + Val(NBArcano.Text) + Val(MBArcano.Text)
            TEncantamiento.Text = Val(BEncantamiento.Text) + Val(NEncantamiento.Text) + Val(MEncantamiento.Text)
            TTMagica.Text = Val(BTMagica.Text) + Val(NTMagica.Text) + Val(MTMagica.Text)
            TALigera.Text = Val(NALigera.Text) + Val(BALigera.Text) + Val(MALigera.Text)
            TBMagico.Text = Val(BBMagico.Text) + Val(NBMagico.Text) + Val(MBMagico.Text)
            TMisticismo.Text = Val(MMisticismo.Text) + Val(NMisticismo.Text) + Val(BMisticismo.Text)
            TRestauracion.Text = Val(BRestauracion.Text) + Val(NRestauracion.Text) + Val(MRestauracion.Text)
            TIlusion.Text = Val(BIlusion.Text) + Val(NIlusion.Text) + Val(MIlusion.Text)
            TAlteracion.Text = Val(MAlteracion.Text) + Val(NAlteracion.Text) + Val(BAlteracion.Text)
            TConjuracion.Text = Val(MConjuracion.Text) + Val(BConjuracion.Text) + Val(NConjuracion.Text)
            TDestruccion.Text = Val(MDestruccion.Text) + Val(NDestruccion.Text) + Val(BDestruccion.Text)
            TKiai.Text = Val(BKiai.Text) + Val(NKiai.Text) + Val(MKiai.Text)
            TThu.Text = Val(MThu.Text) + Val(BThu.Text) + Val(NThu.Text)
            TAnimales.Text = Val(BAnimales.Text) + Val(MAnimales.Text) + Val(NAnimales.Text)
            TAArrojadizas.Text = Val(BAArrojadizas.Text) + Val(NAArrojadizas.Text) + Val(MAArrojadizas.Text)
            TFrialdad.Text = Val(NFrialdad.Text) + Val(MFrialdad.Text) + Val(BFrialdad.Text)
            THerreria.Text = Val(BHerreria.Text) + Val(NHerreria.Text) + Val(MHerreria.Text)
            TAPesada.Text = Val(BAPesada.Text) + Val(MAPesada.Text) + Val(NAPesada.Text)
            TBloqueo.Text = Val(BBloqueo.Text) + Val(MBloqueo.Text) + Val(NBloqueo.Text)
            TBloqueo.Text = Val(BBloqueo.Text) + Val(MBloqueo.Text) + Val(NBloqueo.Text)
            TArqueria.Text = Val(MArqueria.Text) + Val(BArqueria.Text) + Val(NArqueria.Text)
            TDesarmado.Text = Val(BDesarmado.Text) + Val(NDesarmado.Text) + Val(MDesarmado.Text)
            TDosManos.Text = Val(BDosManos.Text) + Val(NDosManos.Text) + Val(MDosManos.Text)
            TUnaMano.Text = Val(BUnaMano.Text) + Val(NUnaMano.Text) + Val(MUnaMano.Text)

        End If
        If (Raza.Text = "Subnormal") Then
            MsgBox("Subnormal")
        End If
    End Sub

    Private Sub Costelacion_TextChanged(sender As Object, e As EventArgs) Handles Costelacion.TextChanged

        CSalud = 0
        CAguante = 0
        CMagia = 0
        SaludExtra.Text = RSalud
        MagiaExtra.Text = RMagia
        AguanteExtra.Text = RAguante
        SaludTotal.Text = Val(SaludBase.Text) * ((Val(SaludExtra.Text) / 100) + 1)
        MagiaTotal.Text = Val(MagiaBase.Text) * ((Val(MagiaExtra.Text) / 100) + 1)
        AguanteTotal.Text = Val(AguanteBase.Text) * ((Val(AguanteExtra.Text) / 100) + 1)

        'El Amante

        If (Costelacion.Text = "El Amante") Then
            CSalud = 25
            CAguante = 25
            CMagia = 25
            AguanteExtra.Text = RAguante + CAguante
            SaludExtra.Text = RSalud + CSalud
            MagiaExtra.Text = RMagia + CMagia
            SaludTotal.Text = Val(SaludBase.Text) * ((Val(SaludExtra.Text) / 100) + 1)
            MagiaTotal.Text = Val(MagiaBase.Text) * ((Val(MagiaExtra.Text) / 100) + 1)
            AguanteTotal.Text = Val(AguanteBase.Text) * ((Val(AguanteExtra.Text) / 100) + 1)
        End If

        'El Señor

        If Costelacion.Text = "El Señor" Then
            CSalud = 40
            SaludExtra.Text = RSalud + CSalud
            SaludTotal.Text = Val(SaludBase.Text) * ((Val(SaludExtra.Text) / 100) + 1)
        End If

        'El Corcel

        If Costelacion.Text = "El Corcel" Then
            CAguante = 40
            AguanteExtra.Text = RAguante + CAguante
            AguanteTotal.Text = Val(AguanteBase.Text) * ((Val(AguanteExtra.Text) / 100) + 1)
        End If

        'El Aprendiz

        If Costelacion.Text = "El Aprendiz" Then
            CMagia = 40
            MagiaExtra.Text = RMagia + CMagia
            MagiaTotal.Text = Val(MagiaBase.Text) * ((Val(MagiaExtra.Text) / 100) + 1)

        End If
    End Sub

    Private Sub Guardar_Click(sender As Object, e As EventArgs) Handles Guardar.Click
        Dim SaveFile As New SaveFileDialog
        SaveFile.FileName = NombreJugador.Text
        SaveFile.Filter = "Text Files (*.txt|*.txt"
        SaveFile.Title = NombreJugador.Text
        SaveFile.ShowDialog()
        Try
            Dim Write As New System.IO.StreamWriter(SaveFile.FileName)

            'Variables
            Write.WriteLine(NivelAnterior)
            Write.WriteLine(SBA)
            Write.WriteLine(MBA)
            Write.WriteLine(ABA)

            For Contador As Integer = 0 To 8
                Write.WriteLine(AA(Contador))
            Next

            Write.WriteLine(Mensaje)

            For Contador1 As Integer = 0 To 2
                For Contador2 As Integer = 0 To 11
                    Write.WriteLine(HA(Contador2, Contador1))
                Next
            Next

            Write.WriteLine(CMagia)
            Write.WriteLine(CSalud)
            Write.WriteLine(CAguante)

            Write.WriteLine(RMagia)
            Write.WriteLine(RSalud)
            Write.WriteLine(RAguante)

            'DatosJugador
            Write.WriteLine(NombreJugador.Text)
            Write.WriteLine(Nivel.Value)
            Write.WriteLine(Costelacion.Text)
            Write.WriteLine(Raza.Text)

            'Puntos Estadistica
            Write.WriteLine(PuntosEstadistica.Text)

            'Salud
            Write.WriteLine(SaludBase.Text)
            Write.WriteLine(SaludExtra.Text)
            Write.WriteLine(SaludTotal.Text)
            Write.WriteLine(SaludActual.Text)

            'Magia
            Write.WriteLine(MagiaBase.Text)
            Write.WriteLine(MagiaExtra.Text)
            Write.WriteLine(MagiaTotal.Text)
            Write.WriteLine(MagiaActual.Text)

            'Aguante
            Write.WriteLine(AguanteBase.Text)
            Write.WriteLine(AguanteExtra.Text)
            Write.WriteLine(AguanteTotal.Text)
            Write.WriteLine(AguanteActual.Text)

            'Meritos
            Write.WriteLine(Merito1.Text)
            Write.WriteLine(Merito2.Text)
            Write.WriteLine(Merito3.Text)
            Write.WriteLine(Merito4.Text)
            Write.WriteLine(Merito5.Text)
            Write.WriteLine(Merito6.Text)
            Write.WriteLine(Merito7.Text)

            'Defectos
            Write.WriteLine(Defecto1.Text)
            Write.WriteLine(Defecto2.Text)
            Write.WriteLine(Defecto3.Text)
            Write.WriteLine(Defecto4.Text)
            Write.WriteLine(Defecto5.Text)
            Write.WriteLine(Defecto6.Text)
            Write.WriteLine(Defecto7.Text)

            'PuntosAtributo
            Write.WriteLine(PuntosAtributo.Text)

            'Atributos
            Write.WriteLine(Fuerza.Value)
            Write.WriteLine(Resistencia.Value)
            Write.WriteLine(Destreza.Value)
            Write.WriteLine(Agilidad.Value)
            Write.WriteLine(Velocidad.Value)
            Write.WriteLine(Voluntad.Value)
            Write.WriteLine(Inteligencia.Value)
            Write.WriteLine(Personalidad.Value)
            Write.WriteLine(Suerte.Value)

            'PuntosHabilidades
            Write.WriteLine(PuntosHabilidades.Text)

            'HabilidadesGuerreroN
            Write.WriteLine(NUnaMano.Text)
            Write.WriteLine(NDosManos.Text)
            Write.WriteLine(NDesarmado.Text)
            Write.WriteLine(NArqueria.Text)
            Write.WriteLine(NBloqueo.Text)
            Write.WriteLine(NAPesada.Text)
            Write.WriteLine(NHerreria.Text)
            Write.WriteLine(NFrialdad.Text)
            Write.WriteLine(NAArrojadizas.Text)
            Write.WriteLine(NAnimales.Text)
            Write.WriteLine(NThu.Text)
            Write.WriteLine(NKiai.Text)

            'HabilidadesMagoN
            Write.WriteLine(NDestruccion.Text)
            Write.WriteLine(NConjuracion.Text)
            Write.WriteLine(NAlteracion.Text)
            Write.WriteLine(NIlusion.Text)
            Write.WriteLine(NRestauracion.Text)
            Write.WriteLine(NMisticismo.Text)
            Write.WriteLine(NBMagico.Text)
            Write.WriteLine(NALigera.Text)
            Write.WriteLine(NTMagica.Text)
            Write.WriteLine(NEncantamiento.Text)
            Write.WriteLine(NBArcano.Text)
            Write.WriteLine(NChaman.Text)

            'HabilidadesLadronN
            Write.WriteLine(NAMedia.Text)
            Write.WriteLine(NRobo.Text)
            Write.WriteLine(NSigilo.Text)
            Write.WriteLine(NEsquivar.Text)
            Write.WriteLine(NACerraduras.Text)
            Write.WriteLine(NElocuencia.Text)
            Write.WriteLine(NAlquimia.Text)
            Write.WriteLine(NPercepcion.Text)
            Write.WriteLine(NMedicina.Text)
            Write.WriteLine(NTramperia.Text)
            Write.WriteLine(NCocina.Text)
            Write.WriteLine(NFortuna.Text)

            'HabilidadesGuerreroM
            Write.WriteLine(MUnaMano.Text)
            Write.WriteLine(MDosManos.Text)
            Write.WriteLine(MDesarmado.Text)
            Write.WriteLine(MArqueria.Text)
            Write.WriteLine(MBloqueo.Text)
            Write.WriteLine(MAPesada.Text)
            Write.WriteLine(MHerreria.Text)
            Write.WriteLine(MFrialdad.Text)
            Write.WriteLine(MAArrojadizas.Text)
            Write.WriteLine(MAnimales.Text)
            Write.WriteLine(MThu.Text)
            Write.WriteLine(MKiai.Text)

            'HabilidadesMagoM
            Write.WriteLine(MDestruccion.Text)
            Write.WriteLine(MConjuracion.Text)
            Write.WriteLine(MAlteracion.Text)
            Write.WriteLine(MIlusion.Text)
            Write.WriteLine(MRestauracion.Text)
            Write.WriteLine(MMisticismo.Text)
            Write.WriteLine(MBMagico.Text)
            Write.WriteLine(MALigera.Text)
            Write.WriteLine(MTMagica.Text)
            Write.WriteLine(MEncantamiento.Text)
            Write.WriteLine(MBArcano.Text)
            Write.WriteLine(MChaman.Text)

            'HabilidadesLadronM
            Write.WriteLine(MAMedia.Text)
            Write.WriteLine(MRobo.Text)
            Write.WriteLine(MSigilo.Text)
            Write.WriteLine(MEsquivar.Text)
            Write.WriteLine(MACerraduras.Text)
            Write.WriteLine(MElocuencia.Text)
            Write.WriteLine(MAlquimia.Text)
            Write.WriteLine(MPercepcion.Text)
            Write.WriteLine(MMedicina.Text)
            Write.WriteLine(MTramperia.Text)
            Write.WriteLine(MCocina.Text)
            Write.WriteLine(MFortuna.Text)

            'HabilidadesAtributos
            Write.WriteLine(BHerreria.Text)
            Write.WriteLine(BACerraduras.Text)
            Write.WriteLine(BElocuencia.Text)
            Write.WriteLine(BAlquimia.Text)
            Write.WriteLine(BMedicina.Text)
            Write.WriteLine(BEncantamiento.Text)

            Write.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Cargar_Click(sender As Object, e As EventArgs) Handles Cargar.Click
        Dim OpenFile As New OpenFileDialog
        Dim NameFile As String
        OpenFile.Filter = "Text Files (*.txt|*.txt"
        OpenFile.ShowDialog()
        NameFile = OpenFile.FileName
        Try
            Dim fileReader As New System.IO.StreamReader(NameFile)

            'Variables
            NivelAnterior = fileReader.ReadLine()
            SBA = fileReader.ReadLine()
            MBA = fileReader.ReadLine()
            ABA = fileReader.ReadLine()

            For Contador As Integer = 0 To 8
                AA(Contador) = fileReader.ReadLine()
            Next

            Mensaje = fileReader.ReadLine()

            For Contador1 As Integer = 0 To 2
                For Contador2 As Integer = 0 To 11
                    HA(Contador2, Contador1) = fileReader.ReadLine()
                Next
            Next

            CMagia = fileReader.ReadLine()
            CSalud = fileReader.ReadLine()
            CAguante = fileReader.ReadLine()

            RMagia = fileReader.ReadLine()
            RSalud = fileReader.ReadLine()
            RAguante = fileReader.ReadLine()

            'DatosJugador
            NombreJugador.Text = fileReader.ReadLine()
            Nivel.Value = fileReader.ReadLine()
            Costelacion.Text = fileReader.ReadLine()
            Raza.Text = fileReader.ReadLine()

            'Puntos Estadistica
            PuntosEstadistica.Text = fileReader.ReadLine()

            'Salud
            SaludBase.Text = fileReader.ReadLine()
            SaludExtra.Text = fileReader.ReadLine()
            SaludTotal.Text = fileReader.ReadLine()
            SaludActual.Text = fileReader.ReadLine()

            'Magia
            MagiaBase.Text = fileReader.ReadLine()
            MagiaExtra.Text = fileReader.ReadLine()
            MagiaTotal.Text = fileReader.ReadLine()
            MagiaActual.Text = fileReader.ReadLine()

            'Aguante
            AguanteBase.Text = fileReader.ReadLine()
            AguanteExtra.Text = fileReader.ReadLine()
            AguanteTotal.Text = fileReader.ReadLine()
            AguanteActual.Text = fileReader.ReadLine()

            'Meritos
            Merito1.Text = fileReader.ReadLine()
            Merito2.Text = fileReader.ReadLine()
            Merito3.Text = fileReader.ReadLine()
            Merito4.Text = fileReader.ReadLine()
            Merito5.Text = fileReader.ReadLine()
            Merito6.Text = fileReader.ReadLine()
            Merito7.Text = fileReader.ReadLine()

            'Defectos
            Defecto1.Text = fileReader.ReadLine()
            Defecto2.Text = fileReader.ReadLine()
            Defecto3.Text = fileReader.ReadLine()
            Defecto4.Text = fileReader.ReadLine()
            Defecto5.Text = fileReader.ReadLine()
            Defecto6.Text = fileReader.ReadLine()
            Defecto7.Text = fileReader.ReadLine()

            'PuntosDeAtributo
            PuntosAtributo.Text = fileReader.ReadLine()

            'Atributos
            Fuerza.Value = fileReader.ReadLine()
            Resistencia.Value = fileReader.ReadLine()
            Destreza.Value = fileReader.ReadLine()
            Agilidad.Value = fileReader.ReadLine()
            Velocidad.Value = fileReader.ReadLine()
            Voluntad.Value = fileReader.ReadLine()
            Inteligencia.Value = fileReader.ReadLine()
            Personalidad.Value = fileReader.ReadLine()
            Suerte.Value = fileReader.ReadLine()

            'PuntosHabilidades
            PuntosHabilidades.Text = fileReader.ReadLine()

            'HabilidadesGuerreroN
            NUnaMano.Text = fileReader.ReadLine()
            NDosManos.Text = fileReader.ReadLine()
            NDesarmado.Text = fileReader.ReadLine()
            NArqueria.Text = fileReader.ReadLine()
            NBloqueo.Text = fileReader.ReadLine()
            NAPesada.Text = fileReader.ReadLine()
            NHerreria.Text = fileReader.ReadLine()
            NFrialdad.Text = fileReader.ReadLine()
            NAArrojadizas.Text = fileReader.ReadLine()
            NAnimales.Text = fileReader.ReadLine()
            NThu.Text = fileReader.ReadLine()
            NKiai.Text = fileReader.ReadLine()

            'HabilidadesMagoN
            NDestruccion.Text = fileReader.ReadLine()
            NConjuracion.Text = fileReader.ReadLine()
            NAlteracion.Text = fileReader.ReadLine()
            NIlusion.Text = fileReader.ReadLine()
            NRestauracion.Text = fileReader.ReadLine()
            NMisticismo.Text = fileReader.ReadLine()
            NBMagico.Text = fileReader.ReadLine()
            NALigera.Text = fileReader.ReadLine()
            NTMagica.Text = fileReader.ReadLine()
            NEncantamiento.Text = fileReader.ReadLine()
            NBArcano.Text = fileReader.ReadLine()
            NChaman.Text = fileReader.ReadLine()

            'HabilidadesLadronN
            NAMedia.Text = fileReader.ReadLine()
            NRobo.Text = fileReader.ReadLine()
            NSigilo.Text = fileReader.ReadLine()
            NEsquivar.Text = fileReader.ReadLine()
            NACerraduras.Text = fileReader.ReadLine()
            NElocuencia.Text = fileReader.ReadLine()
            NAlquimia.Text = fileReader.ReadLine()
            NPercepcion.Text = fileReader.ReadLine()
            NMedicina.Text = fileReader.ReadLine()
            NTramperia.Text = fileReader.ReadLine()
            NCocina.Text = fileReader.ReadLine()
            NFortuna.Text = fileReader.ReadLine()

            'HabilidadesGuerreroM
            MUnaMano.Text = fileReader.ReadLine()
            MDosManos.Text = fileReader.ReadLine()
            MDesarmado.Text = fileReader.ReadLine()
            MArqueria.Text = fileReader.ReadLine()
            MBloqueo.Text = fileReader.ReadLine()
            MAPesada.Text = fileReader.ReadLine()
            MHerreria.Text = fileReader.ReadLine()
            MFrialdad.Text = fileReader.ReadLine()
            MAArrojadizas.Text = fileReader.ReadLine()
            MAnimales.Text = fileReader.ReadLine()
            MThu.Text = fileReader.ReadLine()
            MKiai.Text = fileReader.ReadLine()

            'HabilidadesMagoM
            MDestruccion.Text = fileReader.ReadLine()
            MConjuracion.Text = fileReader.ReadLine()
            MAlteracion.Text = fileReader.ReadLine()
            MIlusion.Text = fileReader.ReadLine()
            MRestauracion.Text = fileReader.ReadLine()
            MMisticismo.Text = fileReader.ReadLine()
            MBMagico.Text = fileReader.ReadLine()
            MALigera.Text = fileReader.ReadLine()
            MTMagica.Text = fileReader.ReadLine()
            MEncantamiento.Text = fileReader.ReadLine()
            MBArcano.Text = fileReader.ReadLine()
            MChaman.Text = fileReader.ReadLine()

            'HabilidadesLadronM
            MAMedia.Text = fileReader.ReadLine()
            MRobo.Text = fileReader.ReadLine()
            MSigilo.Text = fileReader.ReadLine()
            MEsquivar.Text = fileReader.ReadLine()
            MACerraduras.Text = fileReader.ReadLine()
            MElocuencia.Text = fileReader.ReadLine()
            MAlquimia.Text = fileReader.ReadLine()
            MPercepcion.Text = fileReader.ReadLine()
            MMedicina.Text = fileReader.ReadLine()
            MTramperia.Text = fileReader.ReadLine()
            MCocina.Text = fileReader.ReadLine()
            MFortuna.Text = fileReader.ReadLine()

            'HabilidadesEspeciales
            BHerreria.Text = fileReader.ReadLine()
            BACerraduras.Text = fileReader.ReadLine()
            BElocuencia.Text = fileReader.ReadLine()
            BAlquimia.Text = fileReader.ReadLine()
            BMedicina.Text = fileReader.ReadLine()
            BEncantamiento.Text = fileReader.ReadLine()

            THerreria.Text = Val(BHerreria.Text) + Val(NHerreria.Text) + Val(MHerreria.Text)
            TEncantamiento.Text = Val(BEncantamiento.Text) + Val(NEncantamiento.Text) + Val(MEncantamiento.Text)
            TACerraduras.Text = Val(BACerraduras.Text) + Val(NACerraduras.Text) + Val(MACerraduras.Text)
            TAlquimia.Text = Val(NAlquimia.Text) + Val(BAlquimia.Text) + Val(MAlquimia.Text)
            TMedicina.Text = Val(BMedicina.Text) + Val(NMedicina.Text) + Val(MMedicina.Text)
            TElocuencia.Text = Val(BElocuencia.Text) + Val(NElocuencia.Text) + Val(MElocuencia.Text)

            fileReader.Close()
        Catch ex As Exception
        End Try
    End Sub
End Class



