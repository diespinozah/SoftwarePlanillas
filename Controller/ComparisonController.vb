' Controller/ComparisonController.vb
Imports System.IO
Imports OfficeOpenXml

Public Class ComparisonController
    Private ReadOnly _encabezadoValidator As EncabezadoValidator
    Private ReadOnly _rendicionDeBoletasValidator As RendicionDeBoletasValidator
    Private ReadOnly _rendicionDeFacturaValidator As RendicionDeFacturaValidator

    Public Sub New(encabezadoValidator As EncabezadoValidator, rendicionDeBoletasValidator As RendicionDeBoletasValidator, rendicionDeFacturaValidator As RendicionDeFacturaValidator)
        _encabezadoValidator = encabezadoValidator
        _rendicionDeBoletasValidator = rendicionDeBoletasValidator
        _rendicionDeFacturaValidator = rendicionDeFacturaValidator
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial ' Configura la licencia
    End Sub

    Public Function CompararEncabezado(encabezado As Encabezado) As List(Of String)
        Return _encabezadoValidator.Validar(encabezado)
    End Function
    Public Function CompararRendicionDeBoletas(rendicion As RendicionDeBoletas) As List(Of String)
        Return _rendicionDeBoletasValidator.Validar(rendicion)
    End Function
    Public Function CompararRendicionDeFactura(rendicion As RendicionDeFactura) As List(Of String)
        Return _rendicionDeFacturaValidator.Validar(rendicion)
    End Function

    Public Function CargarArchivos(rutas As String()) As ArchivoCargado
        Dim resultado As New ArchivoCargado With {
            .Encabezados = New List(Of Encabezado)(),
            .RendicionesDeBoletas = New List(Of RendicionDeBoletas)(),
            .RendicionesDeFactura = New List(Of RendicionDeFactura)()
        }
        'Debug.WriteLine("Inicializando la carga de archivos.")

        For Each ruta In rutas
            If Not File.Exists(ruta) Then
                'Debug.WriteLine($"Archivo no encontrado: {ruta}")
                Continue For
            End If

            'Debug.WriteLine($"Cargando archivo: {ruta}")

            Using package As New ExcelPackage(New FileInfo(ruta))

                ' Listar todas las hojas para depuración
                'For Each hoja In package.Workbook.Worksheets
                '    Debug.WriteLine($"Hoja encontrada en el archivo: {hoja.Name}")
                'Next

                Dim hojaEncabezado = package.Workbook.Worksheets.FirstOrDefault(Function(ws) ws.Name.Equals("ENCABEZADO", StringComparison.OrdinalIgnoreCase))
                If hojaEncabezado IsNot Nothing Then

                    Dim encabezado As New Encabezado With {
                        .Y2 = hojaEncabezado.Cells("Y2").Text,
                        .L7 = hojaEncabezado.Cells("L7").Text,
                        .V6 = hojaEncabezado.Cells("V6").Text,
                        .D17 = hojaEncabezado.Cells("D17").Text,
                        .X15 = hojaEncabezado.Cells("X15").Text,
                        .S5 = hojaEncabezado.Cells("S5").Text,
                        .S3 = hojaEncabezado.Cells("S3").Text,
                        .Archivo = ruta
                }
                    resultado.Encabezados.Add(encabezado)
                    'Debug.WriteLine($"Hoja 'ENCABEZADO' encontrada y cargada en el archivo: {ruta}")
                Else
                    Debug.WriteLine($"Hoja 'ENCABEZADO' no encontrada en el archivo: {ruta}")
                End If

                ' Cargar RENDICION DE BOLETAS
                Dim hojaRendicionDeBoletas = package.Workbook.Worksheets.FirstOrDefault(Function(ws) ws.Name.Equals("RENDICION DE BOLETAS", StringComparison.OrdinalIgnoreCase))
                If hojaRendicionDeBoletas IsNot Nothing Then
                    ''Dim rendicion As New RendicionDeBoletas With {
                    ''    .Archivo = ruta,
                    ''    .Data = New List(Of List(Of String))()
                    ''}

                    ''For i As Integer = 15 To hojaRendicionDeBoletas.Dimension.End.Row
                    ''    Dim fila As New List(Of String)()
                    ''    For j As Integer = 2 To 27
                    ''        fila.Add(hojaRendicionDeBoletas.Cells(i, j).Text)
                    ''    Next
                    ''    If fila.Any(Function(c) Not String.IsNullOrEmpty(c)) Then
                    ''        rendicion.Data.Add(fila)
                    ''    End If
                    ''Next
                    ''resultado.RendicionesDeBoletas.Add(rendicion)
                    Dim data As New List(Of List(Of String))()
                    For row = 15 To hojaRendicionDeBoletas.Dimension.End.Row
                        If IsRowEmpty(hojaRendicionDeBoletas, row) Then Exit For
                        Dim filaData As New List(Of String)()
                        For col = 2 To 27 ' Columnas B to AA (2 to 27)
                            filaData.Add(hojaRendicionDeBoletas.Cells(row, col).Text)
                        Next
                        data.Add(filaData)
                    Next

                    Dim rendicion As New RendicionDeBoletas With {
                        .Data = data,
                        .Archivo = ruta
                    }
                    resultado.RendicionesDeBoletas.Add(rendicion)
                End If

                Dim hojaRendicionDeFactura = package.Workbook.Worksheets.FirstOrDefault(Function(ws) ws.Name.Equals("RENDICION DE FACTURA", StringComparison.OrdinalIgnoreCase))
                If hojaRendicionDeFactura IsNot Nothing Then
                    Dim data As New List(Of List(Of String))()
                    For row = 15 To hojaRendicionDeFactura.Dimension.End.Row
                        If IsRowEmpty(hojaRendicionDeFactura, row) Then Exit For
                        Dim filaData As New List(Of String)()
                        For col = 2 To 27 ' Columnas B to AA (2 to 27)
                            filaData.Add(hojaRendicionDeFactura.Cells(row, col).Text)
                        Next
                        data.Add(filaData)
                    Next

                    Dim rendicion As New RendicionDeFactura With {
                        .Data = data,
                        .Archivo = ruta
                    }
                    resultado.RendicionesDeFactura.Add(rendicion)
                End If
            End Using
        Next

        'Debug.WriteLine($"Total de encabezados cargados: {resultado.Encabezados.Count}")
        'Debug.WriteLine($"Total de rendiciones de boletas cargadas: {resultado.RendicionesDeBoletas.Count}")

        Return resultado
    End Function

    Private Function IsRowEmpty(hoja As ExcelWorksheet, row As Integer) As Boolean
        For col = 2 To 27 ' Columnas B to AA (2 to 27)
            If Not String.IsNullOrEmpty(hoja.Cells(row, col).Text) Then
                Return False
            End If
        Next
        Return True
    End Function

    Public Function ValidarS3(encabezados As List(Of Encabezado)) As List(Of String)
        Return _encabezadoValidator.ValidarS3NoRepetido(encabezados)
    End Function

    Public Function ValidarYNoRepetido(rendiciones As List(Of RendicionDeBoletas)) As List(Of String)
        Return _rendicionDeBoletasValidator.ValidarYNoRepetido(rendiciones)
    End Function

    Public Function ValidarYNoRepetidoFactura(rendiciones As List(Of RendicionDeFactura)) As List(Of String)
        Return _rendicionDeFacturaValidator.ValidarYNoRepetido(rendiciones)
    End Function

End Class
