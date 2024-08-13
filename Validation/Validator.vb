' Validation/Validator.vb
Imports System.Text
Imports System.Globalization
Public Class Validator

    Protected Function QuitarTildes(texto As String) As String
        Dim normalizedString As String = texto.Normalize(NormalizationForm.FormD)
        Dim sb As New StringBuilder()

        For Each c As Char In normalizedString
            If CharUnicodeInfo.GetUnicodeCategory(c) <> UnicodeCategory.NonSpacingMark Then
                sb.Append(c)
            End If
        Next

        Return sb.ToString().Normalize(NormalizationForm.FormC)
    End Function

    Protected Function ValidarFechaOrden(fechaAnterior As Date, fechaActual As Date, hoja As String, columna As String, filaIndex As Integer, archivo As String) As String
        If fechaActual < fechaAnterior Then
            Return $"Archivo: {archivo}, Hoja: {hoja}, Fila: {filaIndex}, Columna: {columna}, Error: fecha no está en orden cronológico"
        End If
        Return String.Empty
    End Function

    Protected Function ValidarCampoNoVacio(valor As String, hoja As String, columna As String, filaIndex As Integer, archivo As String) As String
        If String.IsNullOrEmpty(valor) Then
            Return $"Archivo: {archivo}, Hoja: {hoja}, Fila: {filaIndex}, Columna: {columna}, Error: campo en blanco"
        End If
        Return String.Empty
    End Function

    Protected Function ValidarValorNumerico(valor As String, hoja As String, columna As String, filaIndex As Integer, archivo As String) As String
        If String.IsNullOrEmpty(valor) OrElse Not IsNumeric(valor) Then
            Return $"Archivo: {archivo}, Hoja: {hoja}, Fila: {filaIndex}, Columna: {columna}, Error: campo en blanco o no es un valor numérico"
        End If
        Return String.Empty
    End Function

    Protected Function ValidarNombreEnLista(nombre As String, lista As String(), hoja As String, columna As String, filaIndex As Integer, archivo As String) As String
        If String.IsNullOrEmpty(nombre) OrElse Not lista.Any(Function(item) nombre.Contains(item)) Then
            Return $"Archivo: {archivo}, Hoja: {hoja}, Fila: {filaIndex}, Columna: {columna}, Error: campo en blanco o valor inválido"
        End If
        Return String.Empty
    End Function

End Class
