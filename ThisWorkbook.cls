' ============================================================================
' VARIABLES GLOBALES
' ============================================================================
Dim hojaAnterior As String

' ============================================================================
' EVENTO: ABRIR LIBRO
' ============================================================================
Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    OcultarTodasLasHojas
    MostrarHojaInicio
    ConfigurarInterfaz
    MostrarFormularioInicio
    
End Sub

' ============================================================================
' EVENTO: SEGUIR HIPERVÍNCULO
' ============================================================================
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    Dim nombreHoja As String
    
    On Error GoTo ManejadorError
    
    nombreHoja = Replace(Split(Target.SubAddress, "!")(0), "'", "")
    
    If Not SheetExists(nombreHoja) Then
        MsgBox "La hoja '" & nombreHoja & "' no existe.", vbExclamation, "Error"
        Exit Sub
    End If
    
    With Sheets(nombreHoja)
        .Visible = xlSheetVisible
        .Activate
    End With
    
    ' Ocultar hoja anterior SOLO si no es INICIO ni COMP_Medidores
    If hojaAnterior <> nombreHoja And hojaAnterior <> "" Then
        If hojaAnterior <> "INICIO" And hojaAnterior <> "COMP_Medidores" Then
            On Error Resume Next
            Sheets(hojaAnterior).Visible = xlSheetVeryHidden
            On Error GoTo 0
        End If
    End If
    
    hojaAnterior = nombreHoja
    
ManejadorError:
        If Err.Number <> 0 Then
            MsgBox "Error al navegar: " & Err.Description, vbCritical, "Error"
        End If
End Sub

' ============================================================================
' EVENTO: CAMBIO DE HOJA - Asegura que COMP_Medidores siempre esté visible
' ============================================================================
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If SheetExists("COMP_Medidores") Then
        Sheets("COMP_Medidores").Visible = xlSheetVisible
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' EVENTO: ANTES DE CERRAR
' ============================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' Asegurar que COMP_Medidores quede visible al cerrar
    If SheetExists("COMP_Medidores") Then
        Sheets("COMP_Medidores").Visible = xlSheetVisible
    End If
    
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    
    With ActiveWindow
        .DisplayHeadings = True
        .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    
    On Error GoTo 0
End Sub

' ============================================================================
' PROCEDIMIENTOS PRIVADOS
' ============================================================================

Private Sub OcultarTodasLasHojas()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ' NO ocultar INICIO ni COMP_Medidores
        If ws.Name <> "INICIO" And ws.Name <> "COMP_Medidores" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
End Sub

Private Sub MostrarHojaInicio()
    With Sheets("INICIO")
        .Visible = xlSheetVisible
        .Activate
    End With
    
    ' Asegurar COMP_Medidores visible al iniciar
    If SheetExists("COMP_Medidores") Then
        Sheets("COMP_Medidores").Visible = xlSheetVisible
    End If
    
    hojaAnterior = "INICIO"
End Sub

Private Sub ConfigurarInterfaz()
    On Error Resume Next
    
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With
    
    On Error GoTo 0
End Sub

Private Sub MostrarFormularioInicio()
    On Error Resume Next
    INICIO_APP.Show
    On Error GoTo 0
End Sub

' ============================================================================
' FUNCIONES AUXILIARES
' ============================================================================

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function