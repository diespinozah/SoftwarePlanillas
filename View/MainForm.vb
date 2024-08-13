' View/MainForm.vb
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Public Class MainForm
    Private Sub BtnSeleccionarArchivos_Click(sender As Object, e As EventArgs) Handles BtnSeleccionarArchivos.Click
        Using ofd As New OpenFileDialog()
            ofd.Multiselect = True
            ofd.Filter = "Archivos de Excel|*.xls;*.xlsx;*.xlsm"
            If ofd.ShowDialog() = DialogResult.OK Then
                TxtArchivos.Text = String.Join(Environment.NewLine, ofd.FileNames)
            End If
        End Using
    End Sub

    Private Sub BtnComparar_Click(sender As Object, e As EventArgs) Handles BtnComparar.Click
        Dim rutas As String() = TxtArchivos.Text.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
        Dim encabezadoValidator As New EncabezadoValidator()
        Dim rendicionDeBoletasValidator As New RendicionDeBoletasValidator()
        Dim rendicionDeFacturaValidator As New RendicionDeFacturaValidator()
        Dim rendicionDeViaticosValidator As New RendicionDeViaticosValidator()
        Dim controller As New ComparisonController(encabezadoValidator, rendicionDeBoletasValidator, rendicionDeFacturaValidator, rendicionDeViaticosValidator)
        Dim archivosCargados = controller.CargarArchivos(rutas)
        Dim mensajes As New List(Of String)()

        ' Restaurar colores iniciales
        For Each ruta In rutas
            RestaurarColoresIniciales(ruta)
        Next

        ' Validar duplicados primero
        Dim duplicados = controller.ValidarS3(archivosCargados.Encabezados)
        mensajes.AddRange(duplicados)
        If duplicados.Count > 0 Then
            For Each encabezado In archivosCargados.Encabezados
                PintarCeldasConErrores(encabezado, duplicados, "ENCABEZADO")
            Next
        End If

        ' Validar duplicados en la celda Y
        Dim duplicadosY = controller.ValidarYNoRepetido(archivosCargados.RendicionesDeBoletas)
        mensajes.AddRange(duplicadosY)
        If duplicadosY.Count > 0 Then
            For Each rendicion In archivosCargados.RendicionesDeBoletas
                PintarCeldasConErrores(rendicion, duplicadosY, "RENDICION DE BOLETAS")
            Next
        End If

        ' Validar duplicados en la celda Y de la hoja RENDICION DE FACTURA
        Dim duplicadosFacturaY = controller.ValidarYNoRepetidoFactura(archivosCargados.RendicionesDeFactura)
        mensajes.AddRange(duplicadosFacturaY)
        If duplicadosFacturaY.Count > 0 Then
            For Each rendicion In archivosCargados.RendicionesDeFactura
                PintarCeldasConErrores(rendicion, duplicadosFacturaY, "RENDICION DE FACTURA")
            Next
        End If

        For Each encabezado In archivosCargados.Encabezados
            Dim errores = controller.CompararEncabezado(encabezado)
            If errores.Count > 0 Then
                ' Pinta las celdas con errores en el archivo
                PintarCeldasConErrores(encabezado, errores, "ENCABEZADO")
                mensajes.AddRange(errores)
            End If
        Next

        For Each rendicion In archivosCargados.RendicionesDeBoletas
            Dim errores = controller.CompararRendicionDeBoletas(rendicion)
            If errores.Count > 0 Then
                PintarCeldasConErrores(rendicion, errores, "RENDICION DE BOLETAS")
                mensajes.AddRange(errores)
            End If
        Next

        ' Validar y pintar errores en la hoja RENDICION DE FACTURA
        For Each rendicion In archivosCargados.RendicionesDeFactura
            Dim errores = controller.CompararRendicionDeFactura(rendicion)
            If errores.Count > 0 Then
                PintarCeldasConErrores(rendicion, errores, "RENDICION DE FACTURA")
                mensajes.AddRange(errores)
            End If
        Next

        ' Validar y pintar errores en la hoja RENDICION DE VIATICOS
        For Each rendicion In archivosCargados.RendicionesDeViaticos
            Dim errores = controller.CompararRendicionDeViaticos(rendicion)
            If errores.Count > 0 Then
                PintarCeldasConErrores(rendicion, errores, "RENDICION DE VIATICOS")
                mensajes.AddRange(errores)
            End If
        Next

        If mensajes.Count = 0 Then
            MessageBox.Show("Todas las planillas son válidas.")
        End If

        TxtArchivos.Clear()
        TxtArchivos.Text = String.Join(Environment.NewLine, rutas) & Environment.NewLine
        For Each mensaje In mensajes
            TxtArchivos.SelectionStart = TxtArchivos.TextLength
            TxtArchivos.SelectionLength = 0
            TxtArchivos.SelectionFont = New Font(TxtArchivos.Font, FontStyle.Bold)
            TxtArchivos.AppendText(mensaje & Environment.NewLine)
            TxtArchivos.SelectionFont = TxtArchivos.Font ' Restaura la fuente original
        Next
    End Sub

    Private Sub PintarCeldasConErrores(model As Object, errores As List(Of String), hojaNombre As String)
        Try
            Dim archivo As String = If(TypeOf model Is Encabezado, CType(model, Encabezado).Archivo,
                           If(TypeOf model Is RendicionDeBoletas, CType(model, RendicionDeBoletas).Archivo,
                           If(TypeOf model Is RendicionDeFactura, CType(model, RendicionDeFactura).Archivo,
                           CType(model, RendicionDeViaticos).Archivo)))

            If Not File.Exists(archivo) Then
                errores.Add($"Archivo: {archivo}, Error: El archivo no existe.")
                Return
            End If

            Using package As New ExcelPackage(New FileInfo(archivo))
                Dim hoja = package.Workbook.Worksheets.FirstOrDefault(Function(ws) ws.Name.Equals(hojaNombre, StringComparison.OrdinalIgnoreCase))
                If hoja Is Nothing Then
                    errores.Add($"Archivo: {archivo}, Error: Hoja '{hojaNombre}' no encontrada.")
                    Return
                End If

                If hojaNombre.Equals("ENCABEZADO", StringComparison.OrdinalIgnoreCase) Then
                    ' Pintar celdas específicas basadas en errores en la hoja ENCABEZADO
                    For Each errorMensaje In errores
                        If errorMensaje.Contains("Y2") Then
                            PintarCelda(hoja.Cells("Y2"))
                        ElseIf errorMensaje.Contains("L7") Then
                            PintarCelda(hoja.Cells("L7"))
                        ElseIf errorMensaje.Contains("V6") Then
                            PintarCelda(hoja.Cells("V6"))
                        ElseIf errorMensaje.Contains("D17") Then
                            PintarCelda(hoja.Cells("D17"))
                        ElseIf errorMensaje.Contains("X15") Then
                            PintarCelda(hoja.Cells("X15"))
                        ElseIf errorMensaje.Contains("S5") Then
                            PintarCelda(hoja.Cells("S5"))
                        ElseIf errorMensaje.Contains("S3") Then
                            PintarCelda(hoja.Cells("S3"))
                        End If
                    Next
                ElseIf hojaNombre.Equals("RENDICION DE BOLETAS", StringComparison.OrdinalIgnoreCase) OrElse
                       hojaNombre.Equals("RENDICION DE FACTURA", StringComparison.OrdinalIgnoreCase) OrElse
                       hojaNombre.Equals("RENDICION DE VIATICOS", StringComparison.OrdinalIgnoreCase) Then
                    ' Pintar celdas dinámicas basadas en errores en la hoja RENDICION DE BOLETAS o RENDICION DE FACTURA
                    For Each errorMensaje In errores
                        If errorMensaje.Contains("Fila: ") AndAlso errorMensaje.Contains("Columna: ") Then
                            Dim partes = errorMensaje.Split(New String() {", "}, StringSplitOptions.None)
                            Dim esRepetido = errorMensaje.Contains("se repite en")

                            If esRepetido Then
                                ' Recorrer todas las partes del mensaje que contienen información de archivo, hoja, fila y columna
                                For i As Integer = 0 To partes.Length - 1 Step 4
                                    If i + 2 < partes.Length Then
                                        Dim archivoInfo = partes(i).Trim()
                                        If archivoInfo.StartsWith("El valor") Then
                                            'archivoInfo = archivoInfo.Substring(":")
                                            archivoInfo = archivoInfo.Substring(archivoInfo.IndexOf(":") + 1).Trim()
                                        End If
                                        Dim hojaInfo = partes(i + 1).Trim()
                                        Dim filaInfo = partes(i + 2).Trim().Replace("Fila: ", String.Empty)
                                        Dim columnaInfo = "Y"

                                        ' Asegurarse de que el archivo y la hoja coinciden
                                        If archivo.Contains(archivoInfo) AndAlso hojaNombre.Equals(hojaInfo.Replace("Hoja: ", String.Empty), StringComparison.OrdinalIgnoreCase) Then
                                            PintarCelda(hoja.Cells($"{columnaInfo}{filaInfo}"))
                                        End If
                                    End If
                                Next
                            Else
                                ' Mensaje de error general
                                Dim filaIndex = partes(2).Replace("Fila: ", String.Empty).Trim()
                                Dim columna = partes(3).Replace("Columna: ", String.Empty).Trim()
                                PintarCelda(hoja.Cells($"{columna}{filaIndex}"))
                            End If
                        End If
                    Next
                End If

                package.Save()
            End Using
        Catch ex As Exception
            errores.Add($"Error: {ex.Message}")
        End Try
    End Sub
    Private Sub PintarCelda(celda As ExcelRange)
        If celda IsNot Nothing Then
            celda.Style.Fill.PatternType = ExcelFillStyle.Solid
            celda.Style.Fill.BackgroundColor.SetColor(Color.Red)
        End If
    End Sub
    Private Sub RestaurarColorCelda(celda As ExcelRange)
        If celda IsNot Nothing Then
            celda.Style.Fill.PatternType = ExcelFillStyle.Solid
            celda.Style.Fill.BackgroundColor.SetColor(Color.White)
        End If
    End Sub

    Private Sub RestaurarColoresIniciales(ruta As String)
        Try
            If Not File.Exists(ruta) Then
                Return
            End If

            Using package As New ExcelPackage(New FileInfo(ruta))
                ' Restaurar colores de la hoja ENCABEZADO
                Dim hojaEncabezado = package.Workbook.Worksheets.FirstOrDefault(Function(ws) ws.Name.Equals("ENCABEZADO", StringComparison.OrdinalIgnoreCase))
                If hojaEncabezado IsNot Nothing Then
                    RestaurarColorCelda(hojaEncabezado.Cells("Y2"))
                    RestaurarColorCelda(hojaEncabezado.Cells("L7"))
                    RestaurarColorCelda(hojaEncabezado.Cells("V6"))
                    RestaurarColorCelda(hojaEncabezado.Cells("D17"))
                    RestaurarColorCelda(hojaEncabezado.Cells("X15"))
                    RestaurarColorCelda(hojaEncabezado.Cells("S5"))
                    RestaurarColorCelda(hojaEncabezado.Cells("S3"))
                End If

                ' Restaurar colores de la hoja RENDICION DE BOLETAS
                Dim hojaRendicionDeBoletas = package.Workbook.Worksheets.FirstOrDefault(Function(ws) ws.Name.Equals("RENDICION DE BOLETAS", StringComparison.OrdinalIgnoreCase))
                If hojaRendicionDeBoletas IsNot Nothing Then
                    For row = 15 To 49
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"B{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"L{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"N{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"P{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"T{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"U{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"V{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"W{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"X{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"Y{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"Z{row}"))
                        RestaurarColorCelda(hojaRendicionDeBoletas.Cells($"AA{row}"))
                    Next
                End If

                Dim hojaRendicionDeFactura = package.Workbook.Worksheets.FirstOrDefault(Function(ws) ws.Name.Equals("RENDICION DE FACTURA", StringComparison.OrdinalIgnoreCase))
                If hojaRendicionDeFactura IsNot Nothing Then
                    For row = 15 To 49
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"B{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"L{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"N{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"P{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"T{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"U{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"V{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"W{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"X{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"Y{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"Z{row}"))
                        RestaurarColorCelda(hojaRendicionDeFactura.Cells($"AA{row}"))
                    Next
                End If

                ' Restaurar colores de la hoja RENDICION DE BOLETAS
                Dim hojaRendicionDeViaticos = package.Workbook.Worksheets.FirstOrDefault(Function(ws) ws.Name.Equals("RENDICION DE VIATICOS", StringComparison.OrdinalIgnoreCase))
                If hojaRendicionDeViaticos IsNot Nothing Then
                    For row = 15 To 49
                        RestaurarColorCelda(hojaRendicionDeViaticos.Cells($"B{row}"))
                        RestaurarColorCelda(hojaRendicionDeViaticos.Cells($"L{row}"))
                        RestaurarColorCelda(hojaRendicionDeViaticos.Cells($"N{row}"))
                        RestaurarColorCelda(hojaRendicionDeViaticos.Cells($"P{row}"))
                        RestaurarColorCelda(hojaRendicionDeViaticos.Cells($"T{row}"))
                        RestaurarColorCelda(hojaRendicionDeViaticos.Cells($"U{row}"))
                        RestaurarColorCelda(hojaRendicionDeViaticos.Cells($"V{row}"))
                        RestaurarColorCelda(hojaRendicionDeViaticos.Cells($"W{row}"))
                        RestaurarColorCelda(hojaRendicionDeViaticos.Cells($"X{row}"))
                        RestaurarColorCelda(hojaRendicionDeViaticos.Cells($"Y{row}"))
                    Next
                End If

                package.Save()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class