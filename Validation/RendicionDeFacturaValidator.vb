' Validation/RendicionDeFacturaValidator.vb
Imports System.Text
Imports System.Globalization
Public Class RendicionDeFacturaValidator
    Inherits Validator

    Private ReadOnly Ingenieros As String() = {
        "DIEGO ESPINOZA", "JUAN SEPULVEDA", "GONZALO HONORATO",
        "TOMAS OLAVARRIA", "RICARDO ASTUDILLO", "DYLAN HERRERA",
        "BRIAM VEGA", "PEDRO FERNANDEZ", "CLAUDIO ROMAN"
    }

    Private ReadOnly Funcionarios As String() = {"OSCAR RODRIGUEZ", "FIDEL PEREIRA"}

    Public Function Validar(model As Object) As List(Of String)
        Dim rendicion As RendicionDeFactura = CType(model, RendicionDeFactura)
        Dim errores As New List(Of String)

        For i As Integer = 0 To rendicion.Data.Count - 1
            Dim fila = rendicion.Data(i)
            Dim filaIndex = i + 15 ' Ajustar según el índice de inicio de la tabla

            If String.IsNullOrEmpty(fila(0)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: B, Error: campo en blanco")
            End If

            If String.IsNullOrEmpty(fila(10)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: L, Error: campo en blanco")
            End If

            If String.IsNullOrEmpty(fila(12)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: N, Error: campo en blanco")
            End If

            Dim contenidoCeldaP As String = QuitarTildes(fila(14).ToUpper())
            If String.IsNullOrEmpty(contenidoCeldaP) OrElse Not (Ingenieros.Any(Function(ingeniero) contenidoCeldaP.Contains(ingeniero)) OrElse Funcionarios.Any(Function(funcionario) contenidoCeldaP.Contains(funcionario))) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: P, Error: campo en blanco o valor inválido")
            End If

            'If String.IsNullOrEmpty(fila(14)) OrElse Not (Ingenieros.Any(Function(ingeniero) fila(14).ToUpper().Contains(ingeniero)) OrElse Funcionarios.Any(Function(funcionario) fila(14).ToUpper().Contains(funcionario))) Then
            '    'Ingenieros.Any(Function(ingeniero) fila(14).ToUpper().Contains(ingeniero)) OrElse Not Funcionarios.Any(Function(funcionario) fila(14).ToUpper().Contains(funcionario)) Then '(Ingenieros.Contains(fila(14).ToUpper) OrElse Funcionarios.Contains(fila(14).ToUpper)) Then
            '    errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE BOLETAS, Fila: {filaIndex}, Columna: P, Error: campo en blanco o valor inválido")
            'End If

            If String.IsNullOrEmpty(fila(18)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: T, Error: campo en blanco")
            End If

            If String.IsNullOrEmpty(fila(19)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: U, Error: campo en blanco")
            End If

            If String.IsNullOrEmpty(fila(20)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: V, Error: campo en blanco")
            End If

            ' 8. Celda W - Validar orden cronológico de fechas
            If i > 0 Then
                Dim fechaAnterior As Date
                Dim fechaActual As Date
                If Date.TryParse(rendicion.Data(i - 1)(21), fechaAnterior) AndAlso Date.TryParse(fila(21), fechaActual) Then
                    If fechaActual < fechaAnterior Then
                        errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: W, Error: fecha no está en orden cronológico")
                    End If
                End If
            End If

            If String.IsNullOrEmpty(fila(22)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: X, Error: campo en blanco")
            End If

            If String.IsNullOrEmpty(fila(23)) OrElse Not IsNumeric(fila(23)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: Y, Error: campo en blanco o no es un valor numérico")
            Else
                ' Validar duplicados en la misma planilla
                For j As Integer = 0 To rendicion.Data.Count - 1
                    If j <> i AndAlso fila(23) = rendicion.Data(j)(23) Then
                        errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: Y, Error: valor duplicado en la misma planilla")
                        Exit For
                    End If
                Next
            End If

            If String.IsNullOrEmpty(fila(24)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: Z, Error: campo en blanco")
            End If

            If String.IsNullOrEmpty(fila(25)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE FACTURA, Fila: {filaIndex}, Columna: AA, Error: campo en blanco")
            End If
        Next

        Return errores
    End Function

    Public Function ValidarYNoRepetido(rendiciones As List(Of RendicionDeFactura)) As List(Of String)
        Dim duplicados As New List(Of String)
        Dim valores As New Dictionary(Of String, List(Of String))

        For Each rendicion In rendiciones
            For i As Integer = 0 To rendicion.Data.Count - 1
                Dim fila = rendicion.Data(i)
                Dim valorY = fila(23)
                Dim hojaNombre = "RENDICION DE FACTURA" ' Cambia este valor según la hoja específica
                If valores.ContainsKey(valorY) Then
                    valores(valorY).Add($"{rendicion.Archivo}, Hoja: {hojaNombre}, Fila: {i + 15}, Columna: Y")
                Else
                    valores(valorY) = New List(Of String) From {$"{rendicion.Archivo}, Hoja: {hojaNombre}, Fila: {i + 15}, Columna: Y"}
                End If
                'If valores.ContainsKey(valorY) Then
                '    valores(valorY).Add($"{rendicion.Archivo}, Fila: {i + 15}, Columna: Y")
                'Else
                '    valores(valorY) = New List(Of String) From {$"{rendicion.Archivo}, Fila: {i + 15}, Columna: Y"}
                'End If
            Next
        Next

        For Each kvp In valores
            If kvp.Value.Count > 1 Then
                duplicados.Add($"El valor '{kvp.Key}' en la columna Y se repite en: {String.Join(", ", kvp.Value)}")
            End If
        Next

        Return duplicados
    End Function
End Class
