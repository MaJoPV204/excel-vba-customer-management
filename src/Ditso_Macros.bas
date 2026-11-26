Attribute VB_Name = "ClientesDitso"
Option Explicit

Sub Guardar()
    Dim Celda As Range
    Dim Fila As Long
    If Hoja1.Range("E9").Value = "" Then
        MsgBox "El código esta vacio"
    Else
        Set Celda = Hoja2.Range("A:A").Find(What:=Hoja1.Range("E9").Value, _
                                         after:=Hoja2.Range("A1"))
        If Celda Is Nothing Then
            Fila = Hoja2.Cells(1048576, 1).End(xlUp).Row + 1
            Hoja2.Cells(Fila, 1).Value = Hoja1.Range("E9").Value
            Hoja2.Cells(Fila, 2).Value = Hoja1.Range("B5").Value
            Hoja2.Cells(Fila, 3).Value = Hoja1.Range("E5").Value
            Hoja2.Cells(Fila, 4).Value = Hoja1.Range("B7").Value
            Hoja2.Cells(Fila, 5).Value = Round((Date - Hoja1.Range("B7").Value) / 365, 0)
            Hoja2.Cells(Fila, 6).Value = Hoja1.Range("E7").Value
            Hoja2.Cells(Fila, 7).Value = Hoja1.Range("B9").Value
            MsgBox "Cliente guardado satisfactoriamente"
        Else
            MsgBox "El cliente ya existe, por favor ingrese otro código"
        End If
    End If
    Hoja1.Range("B5").Value = ""
    Hoja1.Range("E5").Value = ""
    Hoja1.Range("B7").Value = ""
    Hoja1.Range("E7").Value = ""
    Hoja1.Range("B9").Value = ""
    Hoja1.Range("E9").Value = ""
End Sub

Sub Buscar()
    Dim Celda As Range
    If Hoja1.Range("E9").Value = "" Then
        MsgBox "Ingrese un código para buscar"
    Else
        Set Celda = Hoja2.Range("A:A").Find(What:=Hoja1.Range("E9").Value, _
                                         after:=Hoja2.Range("A1"))
        If Celda Is Nothing Then
            MsgBox "El código ingresado no existe"
        Else
            Hoja1.Range("B5").Value = Celda.Offset(0, 1).Value
            Hoja1.Range("E5").Value = Celda.Offset(0, 2).Value
            Hoja1.Range("B7").Value = Celda.Offset(0, 3).Value
            Hoja1.Range("E7").Value = Celda.Offset(0, 5).Value
            Hoja1.Range("B9").Value = Celda.Offset(0, 6).Value
        End If
    End If
End Sub

Sub Limpiar()
    Hoja1.Range("B5").Value = ""
    Hoja1.Range("E5").Value = ""
    Hoja1.Range("B7").Value = ""
    Hoja1.Range("E7").Value = ""
    Hoja1.Range("B9").Value = ""
    Hoja1.Range("E9").Value = ""
End Sub

Sub Eliminar()
    Dim Celda As Range
    If Hoja1.Range("E9").Value = "" Then
        MsgBox "Ingrese un código"
    Else
        Set Celda = Hoja2.Range("A:A").Find(What:=Hoja1.Range("E9").Value, _
                                         after:=Hoja2.Range("A1"))
        If Celda Is Nothing Then
            Hoja1.Range("E9").Value = ""
            MsgBox "El código no existe"
        Else
            Celda.EntireRow.Delete
            Call Limpiar
            MsgBox "El cliente ha sido eliminado satisfactoriamente"
        End If
    End If
End Sub

' Procedimiento adicional de instalacion de botones.
Sub InstalarBotones()
    Dim nombres As Variant, zonas As Variant, i As Long
    Dim r As Range, b As Shape
    nombres = Array("Guardar", "Buscar", "Limpiar", "Eliminar")
    zonas = Array("A11:B11", "E11:F11", "A13:B13", "E13:F13")
    For i = 0 To 3
        Set r = Hoja1.Range(zonas(i))
        On Error Resume Next
        Hoja1.Shapes("Ditso_" & nombres(i)).Delete
        On Error GoTo 0
        Set b = Hoja1.Shapes.AddShape(msoShapeRoundedRectangle, r.Left, r.Top, r.Width, r.Height)
        b.Name = "Ditso_" & nombres(i)
        b.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & nombres(i)
        b.Fill.ForeColor.RGB = RGB(57, 125, 120)
        If i = 3 Then b.Fill.ForeColor.RGB = RGB(81, 56, 43)
        b.Line.Visible = msoFalse
        b.TextFrame.Characters.Text = UCase(nombres(i))
        b.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        b.TextFrame.Characters.Font.Bold = True
        b.TextFrame.HorizontalAlignment = xlHAlignCenter
        b.TextFrame.VerticalAlignment = xlVAlignCenter
    Next i
    Hoja1.Range("A16").Value = "Botones configurados."
    MsgBox "Botones instalados correctamente"
End Sub
