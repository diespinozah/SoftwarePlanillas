' Validation/RendicionDeViaticosValidator.vb
Imports System.Text
Imports System.Globalization

Public Class RendicionDeViaticosValidator
    Inherits Validator

    Private ReadOnly Ingenieros As String() = {
        "ESPINOZA HERNANDEZ DIEGO", "SEPULVEDA BRAVO JUAN", "HONORATO LOPEZ GONZALO",
        "OLAVARRIA ARANCIBIA TOMAS", "ASTUDILLO DIAZ RICARDO", "HERRERA MENDOZA DYLAN",
        "VEGA MOLINA BRIAM", "FERNANDEZ CARRASCO PEDRO", "ROMAN MUÑOZ CLAUDIO"
    }

    Private ReadOnly Funcionarios As String() = {"RODRIGUEZ GREY OSCAR", "PEREIRA SOTO FIDEL"}

    Public Function Validar(model As Object) As List(Of String)
        Dim rendicion As RendicionDeViaticos = CType(model, RendicionDeViaticos)
        Dim errores As New List(Of String)

        For i As Integer = 0 To rendicion.Data.Count - 1
            Dim fila = rendicion.Data(i)
            Dim filaIndex = i + 15 ' Ajustar según el índice de inicio de la tabla

            ' Validar Celda L
            If String.IsNullOrEmpty(fila(10)) OrElse Not (Ingenieros.Any(Function(ingeniero) fila(10).Contains(ingeniero)) OrElse Funcionarios.Any(Function(funcionario) fila(10).Contains(funcionario))) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE VIATICOS, Fila: {filaIndex}, Columna: L, Error: campo en blanco o valor inválido")
            End If

            ' Validar Celda V y W - Orden cronológico y V <= W
            Dim fechaV As Date
            Dim fechaW As Date
            If Not Date.TryParse(fila(20), fechaV) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE VIATICOS, Fila: {filaIndex}, Columna: V, Error: fecha no válida")
            End If
            If Not Date.TryParse(fila(21), fechaW) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE VIATICOS, Fila: {filaIndex}, Columna: W, Error: fecha no válida")
            End If
            If i > 0 Then
                Dim fechaAnteriorV As Date
                If Date.TryParse(rendicion.Data(i - 1)(20), fechaAnteriorV) AndAlso Date.TryParse(fila(20), fechaV) Then
                    If fechaV < fechaAnteriorV Then
                        errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE VIATICOS, Fila: {filaIndex}, Columna: V, Error: fecha no está en orden cronológico")
                    End If
                End If
            End If
            If i > 0 Then
                Dim fechaAnteriorW As Date
                If Date.TryParse(rendicion.Data(i - 1)(21), fechaAnteriorW) AndAlso Date.TryParse(fila(21), fechaW) Then
                    If fechaW < fechaAnteriorW Then
                        errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE VIATICOS, Fila: {filaIndex}, Columna: W, Error: fecha no está en orden cronológico")
                    End If
                End If
            End If
            If fechaV > fechaW Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE VIATICOS, Fila: {filaIndex}, Columna: V, Error: fecha en Celda V es mayor que en Celda W")
            End If
            If fechaW < fechaV Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE VIATICOS, Fila: {filaIndex}, Columna: W, Error: fecha en Celda W es menor que en Celda V")
            End If

            ' Validar Celda Y - Valor numérico
            If String.IsNullOrEmpty(fila(23)) OrElse Not IsNumeric(fila(23)) Then
                errores.Add($"Archivo: {rendicion.Archivo}, Hoja: RENDICION DE VIATICOS, Fila: {filaIndex}, Columna: Y, Error: campo en blanco o no es un valor numérico")
            End If
        Next

        Return errores
    End Function

End Class
