Attribute VB_Name = "solicitudstockbodega"
Sub ExportarAWord_SolicitudStock()
    Dim wdApp As Object
    Dim wdDoc As Object

    ' Objetos y variables para la descarga de la plantilla
    Dim objHTTP As Object
    Dim objStream As Object
    Dim rutaDescargaTemporal As String
    Dim plantillaID As String
    Dim plantillaRuta As String
    
    Dim guardarRuta As Variant
    Dim ws As Worksheet
    Dim wsProductos As Worksheet
    Dim wsBase As Worksheet
    
    Dim lugar As String
    Dim siglas As String
    Dim responsableDeCompras As String
    Dim cargoCompras As String
    Dim objetoDeContratacion As String
    Dim firmaTecnico As String
    Dim cargoTecnico As String
    Dim fecha As String
    Dim siglaEntidad As String
    Dim periodo As String

    ' Clave para desproteger la hoja "SECUENCIAS"
    Const claveSecuencias As String = "Admin1991"
    ' Clave general para la estructura y otras hojas
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral
    
    ' Asignar la hoja "BBDD" y desprotegerla para leer el ID
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral
    
    ' Leer el ID de la plantilla desde la celda B134
    plantillaID = wsBase.range("B134").Value
    If plantillaID = "" Then
        MsgBox "No se encontr� el ID de la plantilla en la celda B134 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If
    
    ' Construir la URL de la plantilla (p.e. Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID
    
    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral
    
    ' Mostrar cuadro de di�logo para seleccionar la ubicaci�n donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename("DocumentoTerminado.docx", _
        "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
    If guardarRuta = False Or guardarRuta = "" Then
        MsgBox "Operaci�n cancelada por el usuario.", vbInformation
        Exit Sub
    End If
    
    ' Asignar la hoja de trabajo "SECUENCIAS"
    Set ws = ThisWorkbook.Sheets("SECUENCIAS")
    If ws.Visible = xlSheetVeryHidden Or ws.Visible = xlSheetHidden Then
        ws.Visible = xlSheetVisible
    End If
    ws.Unprotect password:=claveSecuencias

    ' Leer datos de Excel
    lugar = CStr(ws.range("FQ2").Value)
    siglas = CStr(ws.range("DB2").Value)
    responsableDeCompras = CStr(ws.range("CD2").Value)
    cargoCompras = CStr(ws.range("CE2").Value)
    objetoDeContratacion = CStr(ws.range("Q2").Value)
    firmaTecnico = CStr(ws.range("G2").Value)
    cargoTecnico = CStr(ws.range("H2").Value)
    fecha = CStr(ws.range("GZ2").Value)
    siglaEntidad = CStr(ws.range("HA2").Value)
    periodo = CStr(ws.range("HB2").Value)

    ' Proteger y ocultar la hoja nuevamente
    ws.Protect password:=claveSecuencias
    ws.Visible = xlSheetHidden

    ' Construir la ruta temporal para descargar la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_SolicitudStock_Temp.docx"

    ' Descargar la plantilla con MSXML2.ServerXMLHTTP
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.Open "GET", plantillaRuta, False
    objHTTP.send

    If objHTTP.Status = 200 Then
        ' Guardar el archivo descargado en la ubicaci�n temporal
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Tipo binario
        objStream.Open
        objStream.Write objHTTP.ResponseBody
        objStream.SaveToFile rutaDescargaTemporal, 2 ' Sobrescribe si existe
        objStream.Close
    Else
        MsgBox "Error al descargar la plantilla. Verifique la conexi�n o el enlace." & vbCrLf & _
               "C�digo de estado: " & objHTTP.Status & " - " & objHTTP.statusText, vbExclamation
        Exit Sub
    End If
    
    ' Iniciar Word y abrir la plantilla descargada
    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    If wdApp Is Nothing Then
        MsgBox "No se pudo iniciar Microsoft Word.", vbCritical
        Exit Sub
    End If

    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(rutaDescargaTemporal)

    If wdDoc Is Nothing Then
        MsgBox "No se pudo abrir el documento de Word.", vbCritical
        wdApp.Quit
        Exit Sub
    End If

    ' Insertar datos en los marcadores de la plantilla
    With wdDoc
        If .Bookmarks.Exists("Lugar") Then .Bookmarks("Lugar").range.Text = lugar
        If .Bookmarks.Exists("Siglas") Then .Bookmarks("Siglas").range.Text = siglas
        If .Bookmarks.Exists("Responsable_de_Compras") Then .Bookmarks("Responsable_de_Compras").range.Text = responsableDeCompras
        If .Bookmarks.Exists("Cargo_Compras") Then .Bookmarks("Cargo_Compras").range.Text = cargoCompras
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Firma_Tecnico") Then .Bookmarks("Firma_Tecnico").range.Text = firmaTecnico
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
        If .Bookmarks.Exists("Sigla_entidad") Then .Bookmarks("Sigla_entidad").range.Text = siglaEntidad
        If .Bookmarks.Exists("Periodo") Then .Bookmarks("Periodo").range.Text = periodo
    End With

    ' A�adir datos de productos desde el rango visible
    Set wsProductos = ThisWorkbook.Sheets("PRODUCTOS")
    wsProductos.Unprotect password:=claveGeneral

    Dim rangoVisible As range

    On Error Resume Next
    Set rangoVisible = wsProductos.range("Productosdt").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not rangoVisible Is Nothing Then
        rangoVisible.Copy
        If wdDoc.Bookmarks.Exists("Productos") Then
            With wdDoc.Bookmarks("Productos").range
                .PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
                .Tables(1).AutoFitBehavior wdAutoFitWindow
            End With
        Else
            MsgBox "El marcador 'Productos' no existe en la plantilla.", vbExclamation
        End If
    Else
        MsgBox "No hay datos visibles para copiar en la hoja PRODUCTOS.", vbExclamation
    End If

    wsProductos.Protect password:=claveGeneral, Scenarios:=True, AllowFormattingRows:=True

    ' Guardar y cerrar documento
    wdDoc.SaveAs2 Filename:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Ubicarse en la hoja "REQUERIMIENTO"
    ThisWorkbook.Sheets("REQUERIMIENTO").Activate

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True

    ' Eliminar el archivo temporal
    On Error Resume Next
    Kill rutaDescargaTemporal
    On Error GoTo 0

    ' Liberar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing
    Set wsProductos = Nothing
    Set wsBase = Nothing

    MsgBox "El documento se ha generado correctamente.", vbInformation
End Sub

