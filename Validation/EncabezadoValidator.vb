' Validation/EncabezadoValidator.vb
Public Class EncabezadoValidator
    Inherits Validator

    Private ReadOnly Ingenieros As String() = {
        "ESPINOZA HERNANDEZ DIEGO", "SEPULVEDA BRAVO JUAN", "HONORATO LOPEZ GONZALO",
        "OLAVARRIA ARANCIBIA TOMAS", "ASTUDILLO DIAZ RICARDO", "HERRERA MENDOZA DYLAN",
        "VEGA MOLINA BRIAM", "FERNANDEZ CARRASCO PEDRO", "ROMAN MUÑOZ CLAUDIO"
    }

    Private ReadOnly Funcionarios As String() = {"RODRIGUEZ GREY OSCAR", "PEREIRA SOTO FIDEL"}

    Public Function Validar(model As Object) As List(Of String)
        Dim encabezado As Encabezado = CType(model, Encabezado)
        Dim errores As New List(Of String)

        If String.IsNullOrEmpty(encabezado.Y2) Then
            errores.Add($"Archivo: {encabezado.Archivo}, Hoja: ENCABEZADO, Celda: Y2, Error: campo en blanco")
        End If
        If String.IsNullOrEmpty(encabezado.L7) Then
            errores.Add($"Archivo: {encabezado.Archivo}, Hoja: ENCABEZADO, Celda: L7, Error: campo en blanco")
        End If

        If Funcionarios.Contains(encabezado.L7) AndAlso encabezado.Y2 <> "1.- FONDO FIJO  (FF)" Then
            errores.Add($"Archivo: {encabezado.Archivo}, Hoja: ENCABEZADO, Celda: Y2, Error: debería ser '1.- FONDO FIJO  (FF)'")
        ElseIf Ingenieros.Contains(encabezado.L7) AndAlso encabezado.Y2 <> "2.- FONDO X RENDIR (FXR)" Then
            errores.Add($"Archivo: {encabezado.Archivo}, Hoja: ENCABEZADO, Celda: Y2, Error: debería ser '2.- FONDO X RENDIR (FXR)'")
        ElseIf Not Funcionarios.Contains(encabezado.L7) AndAlso Not Ingenieros.Contains(encabezado.L7) AndAlso encabezado.Y2 <> "3.- OTRAS FUENTES (RXG)" Then
            errores.Add($"Archivo: {encabezado.Archivo}, Hoja: ENCABEZADO, Celda: Y2, Error: debería ser '3.- OTRAS FUENTES (RXG)'")
        End If

        If String.IsNullOrEmpty(encabezado.V6) Then
            errores.Add($"Archivo: {encabezado.Archivo}, Hoja: ENCABEZADO, Celda: V6, Error: campo en blanco")
        End If

        If Not String.IsNullOrEmpty(encabezado.X15) AndAlso String.IsNullOrEmpty(encabezado.D17) Then
            errores.Add($"Archivo: {encabezado.Archivo}, Hoja: ENCABEZADO, Celda: D17, Error: campo en blanco")
        End If

        If (encabezado.Y2 = "1.- FONDO FIJO  (FF)" OrElse encabezado.Y2 = "2.- FONDO X RENDIR (FXR)") AndAlso String.IsNullOrEmpty(encabezado.S5) Then
            errores.Add($"Archivo: {encabezado.Archivo}, Hoja: ENCABEZADO, Celda: S5, Error: campo en blanco")
        End If

        Return errores
    End Function

    Public Function ValidarS3NoRepetido(encabezados As List(Of Encabezado)) As List(Of String)
        Dim duplicados As New List(Of String)
        Dim valores As New Dictionary(Of String, List(Of String))

        For Each encabezado In encabezados
            If valores.ContainsKey(encabezado.S3) Then
                valores(encabezado.S3).Add(encabezado.Archivo)
            Else
                valores(encabezado.S3) = New List(Of String) From {encabezado.Archivo}
            End If
        Next

        For Each kvp In valores
            If kvp.Value.Count > 1 Then
                duplicados.Add($"El valor '{kvp.Key}' en la celda S3 se repite en los archivos: {String.Join(", ", kvp.Value)}")
            End If
        Next

        Return duplicados
    End Function
End Class
