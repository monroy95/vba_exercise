Sub validar()
If Range("C7").Value = Empty Or Range("C9").Value = Empty Or Range("C11").Value = Empty Or Range("C13").Value = Empty Then
    MsgBox ("Debe ingresar todos los datos"), vbCritical, "AVISO"
    Exit Sub
End If
End Sub

Sub insertar()
On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim uf As Long
Dim fila As Integer
Dim fecha As Date
Dim correo As String
Dim estado_correo As Boolean

If Range("C7").Value = Empty Or Range("C9").Value = Empty Then
    MsgBox ("El campo para Nombre y Fecha de Nacimiento son obligatorios"), vbCritical, "AVISO"
    Exit Sub
End If

' En la hoja DB los registros se agregan al final
uf = Sheets("DB").Range("A" & Rows.Count).End(xlUp).Row

' nombre
Sheets("DB").Cells(uf + 1, "B") = Sheets("REGISTRO").Cells(7, "C")

' fecha nacimiento
If IsDate(Range("C9")) = True Then
    fecha = Format(Sheets("REGISTRO").Cells(9, "C"), "mm/dd/yyyy")
    Sheets("DB").Cells(uf + 1, "C") = fecha ' Sheets("REGISTRO").Cells(9, "C")
Else
    MsgBox ("El valor del campo Fecha de nacimiento debe ser una fecha valida"), vbCritical, "AVISO"
    Exit Sub
End If

If Range("C9").Value = Date Then
    MsgBox ("La fecha no puede ser la misma que hoy, por favor corregirla"), vbCritical, "AVISO"
    Exit Sub
End If

' correo electronico
correo = Sheets("REGISTRO").Cells(11, "C")
estado_correo = validarCorreo(correo)

If IsEmpty(Range("C11").Value) = False And estado_correo = True Then
    Sheets("DB").Cells(uf + 1, "D") = correo
End If
If IsEmpty(Range("C11").Value) = False And estado_correo = False Then
    MsgBox ("El correo ingresado no es valido"), vbCritical, "AVISO"
    Exit Sub
End If

' direccion
Sheets("DB").Cells(uf + 1, "E") = Sheets("REGISTRO").Cells(13, "C")

' codigo
Sheets("DB").Cells(uf + 1, "A") = Sheets("REGISTRO").Cells(5, "C")

Sheets("REGISTRO").Range("C5:C13").ClearContents
Sheets("REGISTRO").Cells(5, "C") = Application.WorksheetFunction.Max(Sheets("DB").Range("A1" & ":A" & uf + 1)) + 1
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub


' WIP
Sub actualizar()
MsgBox ("Work in progress")
Exit Sub

On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim uf As Long
Dim fila As Integer
Dim fecha As Date
Dim correo As String
Dim estado_correo As Boolean

If Range("C7").Value = Empty Or Range("C9").Value = Empty Then
    MsgBox ("El campo para Nombre y Fecha de Nacimiento son obligatorios"), vbCritical, "AVISO"
    Exit Sub
End If

' En la hoja DB los registros se agregan al final
uf = Sheets("DB").Range("A" & Rows.Count).End(xlUp).Row

' nombre
Sheets("DB").Cells(uf + 1, "B") = Sheets("REGISTRO").Cells(7, "C")

' fecha nacimiento
If IsDate(Range("C9")) = True Then
    fecha = Format(Sheets("REGISTRO").Cells(9, "C"), "mm/dd/yyyy")
    Sheets("DB").Cells(uf + 1, "C") = fecha ' Sheets("REGISTRO").Cells(9, "C")
Else
    MsgBox ("El valor del campo Fecha de nacimiento debe ser una fecha valida"), vbCritical, "AVISO"
    Exit Sub
End If

If Range("C9").Value = Date Then
    MsgBox ("La fecha no puede ser la misma que hoy, por favor corregirla"), vbCritical, "AVISO"
    Exit Sub
End If

' correo electronico
correo = Sheets("REGISTRO").Cells(11, "C")
estado_correo = validarCorreo(correo)

If IsEmpty(Range("C11").Value) = False And estado_correo = True Then
    Sheets("DB").Cells(uf + 1, "D") = correo
End If
If IsEmpty(Range("C11").Value) = False And estado_correo = False Then
    MsgBox ("El correo ingresado no es valido"), vbCritical, "AVISO"
    Exit Sub
End If

' direccion
Sheets("DB").Cells(uf + 1, "E") = Sheets("REGISTRO").Cells(13, "C")

' codigo
Sheets("DB").Cells(uf + 1, "A") = Sheets("REGISTRO").Cells(5, "C")

Sheets("REGISTRO").Range("C5:C13").ClearContents
Sheets("REGISTRO").Cells(5, "C") = Sheets("REGISTRO").Cells(5, "I")
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub


Sub limpiar()
On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim uf As Long

Range("C5").Value = Empty
Range("C7").Value = Empty
Range("C9").Value = Empty
Range("C11").Value = Empty
Range("C13").Value = Empty

' En la hoja DB los registros se agregan al final
uf = Sheets("DB").Range("A" & Rows.Count).End(xlUp).Row
Sheets("REGISTRO").Range("C5:C13").ClearContents
Sheets("REGISTRO").Cells(5, "C") = Application.WorksheetFunction.Max(Sheets("DB").Range("A1" & ":A" & uf + 1)) + 1
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub limpiarBusqueda()
Range("I5").Value = Empty
End Sub

' WIP
Sub eliminar()
MsgBox ("Work in progress")
End Sub

Sub buscar()
Dim valor_buscado As Integer
Dim numero_fila As Integer
Dim codigo As Integer
Dim nombre As String
Dim fecha As Date
Dim correo As String
Dim direccion As String
Dim vacio As Boolean

vacio = IsEmpty(Range("I5").Value)
If IsEmpty(Range("I5").Value) = True Then
    MsgBox ("Por favor ingrese un valor para ser buscado")
    Exit Sub
End If

valor_buscado = Range("I5").Value

If Not IsError(Application.VLookup(valor_buscado, Sheets("DB").Range("A1:E1000"), 1, 0)) Then
    MsgBox ("Registro Encontrado")
    codigo = Application.VLookup(valor_buscado, Sheets("DB").Range("A1:E1000"), 1, 0)
    Range("C5") = codigo

    nombre = Application.VLookup(valor_buscado, Sheets("DB").Range("A1:E1000"), 2, 0)
    Range("C7") = nombre

    fecha = Application.VLookup(valor_buscado, Sheets("DB").Range("A1:E1000"), 3, 0)
    Range("C9") = fecha

    correo = Application.VLookup(valor_buscado, Sheets("DB").Range("A1:E1000"), 4, 0)
    Range("C11") = correo

    direccion = Application.VLookup(valor_buscado, Sheets("DB").Range("A1:E1000"), 5, 0)
    Range("C13") = direccion

Else
    MsgBox ("Registro No se encuentra registrado")
    Exit Sub
End If

End Sub


Sub exportar()
Dim titulo As String
Dim continuar As String
Dim rango_datos As Range
Dim nueva_fila As Integer
Dim limpiar As String
Dim hoja_destino
Dim ruta As String

Dim destino As New Excel.Application
Dim archivo_destino As New Excel.Workbook

titulo = "Exportar Datos"

continuar = MsgBox("Desea exportar los datos", vbYesNo + vbExclamation, titulo)
If continuar = vbNo Then
Exit Sub
End If

ruta = ActiveWorkbook.Path

Set archivo_destino = destino.Workbooks.Open(ruta & "\Datos.xlsx")
Set hoja_destino = archivo_destino.Worksheets("Hoja1")

Set rango_datos = hoja_destino.Cells(1, 10).CurrentRegion

nueva_fila = rango_datos.Rows.Count + 1

With hoja_destino
    .Cells(nueva_fila, 1).Value = ThisWorkbook.Sheets(1).Range("A1")
    .Cells(nueva_fila, 2).Value = ThisWorkbook.Sheets(1).Range("B1")
    .Cells(nueva_fila, 3).Value = ThisWorkbook.Sheets(1).Range("C1")
    .Cells(nueva_fila, 4).Value = ThisWorkbook.Sheets(1).Range("D1")
    .Cells(nueva_fila, 5).Value = ThisWorkbook.Sheets(1).Range("E1")
End With

archivo_destino.Save
archivo_destino.Close

Set destino = Nothing
Set archivo_destino = Nothing

End Sub
