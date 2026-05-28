Attribute VB_Name = "M�dulo1"
public sub comp_medidores()
    dim wsOrigen as worksheet
    dim wsDestino as worksheet
    set wsOrigen = ThisWorkbook.Worksheets("MD")
    set wsDestino = ThisWorkbook.Worksheets("COMP_Medidores")

    'tomar valores de celdas específicas para cada seccion
    '=========
    ' SECCION1
    '=========
    Dim dato1 as string, dato2 as string, dato3 as string
    dato1 = wsOrigen.Range("CF53").Value '+N
    dato2 = wsOrigen.Range("CN53").Value '+C
    dato3 = wsOrigen.Range("CB57").Value ' Tag del medidor
    '=========
    ' SECCION2
    '=========
    Dim dato4 as string, dato5 as string, dato6 as string, dato7 as string, dato8 as string, dato9 as string
    dato4 = wsOrigen.Range("EF149").Value 'origen ubicacion columna x7-p2 (1-3)
    dato5 = wsOrigen.Range("EF158").Value 'origen ubicacion cubiculo x7-p2 (1-3)
    dato6 = wsOrigen.Range("EN149").Value 'origen ubicacion cubiculo x7-p2 (4)
    dato7 = wsOrigen.Range("EN158").Value 'origen alimentacion Tag x7-p2 (1-3)
    dato8 = wsOrigen.Range("EV149").Value ' conductor  AWG x7-p2 (1-4) 
    dato9 = wsOrigen.Range("FL149").Value ' Punto x7-p2 (1-4)
    '=========
    ' SECCION3
    '=========
    Dim dato10 as string, dato11 as string, dato12 as string, dato13 as string, dato14 as string
    dato10 = wsOrigen.Range("EF129").Value 'origen ubicacion columna x7-p1 (1)
    dato11 = wsOrigen.Range("EN129").Value 'origen ubicacion cUBICULO x7-p1 (2)
    dato12 = wsOrigen.Range("EV129").Value 'origen alimentacion Tag x7-p1 (1-2)
    dato13 = wsOrigen.Range("FD129").Value ' Punto x7-p1 (1)
    dato14 = wsOrigen.Range("FD133").Value ' Punto x7-p1 (2)
    '=========
    ' SECCION4
    '=========
    Dim dato15 as string, dato16 as string, dato17 as string, dato18 as string, dato19 as string
    dim dato20 as string, dato21 as string, dato22 as string, dato23 as string
    dato15 = wsOrigen.Range("FX129").Value 'Datos del interruptor columna 
    dato16 = wsOrigen.Range("GF129").Value 'Datos del interruptor cubiculo
    dato17 = wsOrigen.Range("GN129").Value 'Datos del interruptor tag
    dato18 = wsOrigen.Range("GV129").Value 'Datos CONEXION alimentacion columna
    dato19 = wsOrigen.Range("HD129").Value 'Datos CONEXION alimentacion cubiculo
    dato20 = wsOrigen.Range("HL129").Value 'Datos CONEXION alimentacion viene de
    dato21 = wsOrigen.Range("HT129").Value 'Datos CONEXION alimentacion columna viene de
    dato22 = wsOrigen.Range("IB129").Value 'Datos CONEXION alimentacion cubiculo viene de
    dato23 = wsOrigen.Range("IJ129").Value 'Datos CONEXION alimentacion tag viene va
    '=========
    ' SECCION5
    '=========
    Dim dato24 as string, dato25 as string, dato26 as string, dato27 as string, dato28 as string, dato29 as string
    Dim dato30 as string
    dato24 = wsOrigen.Range("GF149").Value 'Puntos de conexion serv. corriente
    dato25 = wsOrigen.Range("GF153").Value 'Puntos de conexion serv. corriente
    dato26 = wsOrigen.Range("GF157").Value 'Puntos de conexion serv. corriente
    dato27 = wsOrigen.Range("GN149").Value 'origen ubicacion columna
    dato28 = wsOrigen.Range("GV149").Value 'origen ubicacion cubiculo
    dato29 = wsOrigen.Range("HD149").Value 'origen ubicacion tag
    dato30 = wsOrigen.Range("HT149").Value 'Conductor AWG

    'CORROBORAR OBTENCION DE INFORMACION
    'MsgBox "Datos obtenidos: " & vbNewLine & _
       "Seccion 1: " & dato1 & ", " & dato2 & ", " & dato3 & vbNewLine & _
       "Seccion 2: " & dato4 & ", " & dato5 & ", " & dato6 & ", " & dato7 & ", " & dato8 & ", " & dato9 & ", " & vbNewLine & _
       "Seccion 3: " & dato10 & ", " & dato11 & ", " & dato12 & ", " & dato13 & ", " & dato14 & ", " & vbNewLine & _
       "Seccion 4: " &  dato15 & ", " & dato16 & ", " & dato17 & ", " & dato18 & ", " & dato19 & ", " & dato20 & ", " & dato21 & ", " & dato22 & ", " & dato23 & ", " & vbNewLine & _
       "Seccion 5: " &  dato24 & ", " & dato25 & ", " & dato26 & ", " & dato27 & ", " &dato28 & ", " &dato29 & ", " &dato30 & ", "
    '========== 
    ' REMPLAZAR EN EL DESTINO
    '==========
    ' SECCION1
    wsDestino.Range("C43").Value = dato1 
    wsDestino.Range("C44").Value = dato2
    ' DATO3 SE REPITE EN LAS CELDAS C7,8,14,35
    wsDestino.Range("C7").Value = dato3 
    WsDestino.Range("C8").Value = dato3 
    wsDestino.Range("C14").Value = dato3 
    wsDestino.Range("C35").Value = dato3
    'SECCION2
    wsDestino.Range("C30").Value = dato4
    wsDestino.Range("C33").Value = dato5
    wsDestino.Range("C31").Value = dato6
    wsDestino.Range("C34").Value = dato7
    wsDestino.Range("C32").Value = dato8
    wsDestino.Range("C41").Value = dato9
    'SECCION3
    wsDestino.Range("C39").Value = dato10
    wsDestino.Range("C40").Value = dato11
    wsDestino.Range("C36").Value = dato12
    wsDestino.Range("C37").Value = dato13
    wsDestino.Range("C38").Value = dato14
    'SECCION4
    wsDestino.Range("C25").Value = dato15
    wsDestino.Range("C26").Value = dato16
    ' DATO17 SE RETIPE EN C6,15
    wsDestino.Range("C6").Value = dato17
    wsDestino.Range("C15").Value = dato17
    wsDestino.Range("C21").Value = dato18
    wsDestino.Range("C22").Value = dato19
    wsDestino.Range("C19").Value = dato20
    wsDestino.Range("C23").Value = dato21
    wsDestino.Range("C24").Value = dato22
    wsDestino.Range("C20").Value = dato23
    'SECCION5
    'DATO24 SE REPITE EN C2,16
    wsDestino.Range("C2").Value = dato24
    wsDestino.Range("C16").Value = dato24
    'DATO25 SE REPITE EN C3,17
    wsDestino.Range("C3").Value = dato25
    wsDestino.Range("C17").Value = dato25
    'DATO26 SE REPITE EN C4,18
    wsDestino.Range("C4").Value = dato26
    wsDestino.Range("C18").Value = dato26
    wsDestino.Range("C27").Value = dato27
    wsDestino.Range("C28").Value = dato28
    wsDestino.Range("C29").Value = dato29
    wsDestino.Range("C42").Value = dato30

    '==========
    ' CREA NUEVO LIBRO
    '==========

    Dim wbNuevo As Workbook
    Dim wsNuevaMD As Worksheet
    Dim wsCompNuevo As Worksheet
    Dim rngAreaImpresion As Range
    Dim strAreaImpresion As String
    Dim rutaGuardado As String
    Dim rangoValores As Range
    Dim filaOffset As Long, colOffset As Long
    Dim celdaInicioDestino As Range

    ' Verificar área de impresión
    strAreaImpresion = wsOrigen.PageSetup.PrintArea

    If strAreaImpresion = "" Then
        MsgBox "La hoja MD no tiene área de impresión definida.", vbExclamation
        Exit Sub
    End If

    Set rngAreaImpresion = wsOrigen.Range(strAreaImpresion)

    ' Crear nuevo libro con 1 sola hoja
    Set wbNuevo = Workbooks.Add(xlWBATWorksheet)
    Set wsNuevaMD = wbNuevo.Sheets(1)
    wsNuevaMD.Name = "MD"

    ' --- Paso 1: Copiar área de impresión completa (formato + valores + todo) ---
    rngAreaImpresion.Copy
    wsNuevaMD.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    ' --- Paso 2: Convertir HL105 a valor neto ---
    Dim filaHL As Long, colHL As Long
    filaHL = wsOrigen.Range("HL105").Row - rngAreaImpresion.Row
    colHL = wsOrigen.Range("HL105").Column - rngAreaImpresion.Column
    wsNuevaMD.Cells(1 + filaHL, 1 + colHL).Value = wsOrigen.Range("HL105").Value

    ' --- Paso 3: Configurar tamaño de celdas MD ---
    wsNuevaMD.Cells.ColumnWidth = 0.38
    wsNuevaMD.Cells.RowHeight = 3.6

    ' --- Paso 4: Configurar página MD ---
    With wsNuevaMD.PageSetup
        .PrintArea = wsNuevaMD.Range("A1").Resize( _
            rngAreaImpresion.Rows.Count, rngAreaImpresion.Columns.Count).Address
        .Orientation = xlLandscape
        .LeftMargin = Application.CentimetersToPoints(0)
        .RightMargin = Application.CentimetersToPoints(0)
        .TopMargin = Application.CentimetersToPoints(0)
        .BottomMargin = Application.CentimetersToPoints(0)
        .CenterHorizontally = True
        .CenterVertically = True
        .Zoom = 100
    End With

    ' --- HOJA 2: COMP (copia de COMP_Medidores) ---
    wsDestino.Copy After:=wbNuevo.Sheets(wbNuevo.Sheets.Count)
    Set wsCompNuevo = wbNuevo.Sheets(wbNuevo.Sheets.Count)
    wsCompNuevo.Name = "COMP"

    ' Convertir fórmulas a valores en COMP
    Dim celdaComp As Range
    For Each celdaComp In wsCompNuevo.UsedRange
        On Error Resume Next
        If celdaComp.HasFormula Then
            celdaComp.Value = celdaComp.Value
        End If
        On Error GoTo 0
    Next celdaComp

    ' Eliminar hojas extra
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In wbNuevo.Sheets
        If ws.Name <> "MD" And ws.Name <> "COMP" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True

   ' --- Guardar con diálogo ---
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    With fd
        .Title = "Guardar libro de medidores"
        .InitialFileName = "COMP_Medidores_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xls"
        
        ' Forzar filtro a Excel 97-2003
        .FilterIndex = 2  ' Índice del filtro .xls en la lista
        
        If .Show = -1 Then
            rutaGuardado = .SelectedItems(1)
            
            ' Asegurar extensión .xls
            If Right(rutaGuardado, 4) <> ".xls" Then
                rutaGuardado = Replace(rutaGuardado, ".xlsx", ".xls")
                If Right(rutaGuardado, 4) <> ".xls" Then
                    rutaGuardado = rutaGuardado & ".xls"
                End If
            End If
            
            wbNuevo.CheckCompatibility = False
            Application.DisplayAlerts = False
            wbNuevo.SaveAs rutaGuardado, FileFormat:=xlExcel8
            Application.DisplayAlerts = True
            
            MsgBox "Proceso finalizado." & vbNewLine & _
                   "Archivo guardado en:" & vbNewLine & rutaGuardado
        Else
            MsgBox "Guardado cancelado. El libro sigue abierto sin guardar.", vbInformation
        End If
    End With

End Sub

           


