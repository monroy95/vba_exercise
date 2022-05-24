Function validarCorreo(correo_ingresado As String) As Boolean
Dim RegEx As Object
Dim validaEmail As Boolean
Set RegEx = CreateObject("vbscript.regexp")

With RegEx
    .Global = True
    .Pattern = "^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$"
End With

validaEmail = RegEx.test(correo_ingresado)

Set RegEx = Nothing
validarCorreo = validaEmail
End Function
