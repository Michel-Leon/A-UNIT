VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} INICIO_APP 
   Caption         =   "LOGIN"
   ClientHeight    =   5440
   ClientLeft      =   80
   ClientTop       =   310
   ClientWidth     =   3310
   OleObjectBlob   =   "INICIO_APP.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "INICIO_APP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    Me.Width  = 260
    Me.Height = 300
    Me.StartUpPosition = 2
    On Error Resume Next
    Me.PictureSizeMode = 1  ' 1 = Stretch
    On Error GoTo 0
End Sub
Private Sub cmdLogin_Click()
    Dim usuario As String, clave As String
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim loginCorrecto As Boolean
    Dim esSuperAdmin As Boolean
    Dim wsLog As Worksheet
    Dim filaLog As Long

    usuario = Trim(txtUsuario.Text)
    clave = Trim(txtPassword.Text)
    
    esSuperAdmin = False

    ' Super Admin codificado manualmente
    If usuario = "ADMIN" And clave = "Admin@GT78" Then
        esSuperAdmin = True
        loginCorrecto = True
        Sheets("INICIO").Range("HD4").Value = "Administrador Sara Michel"
    Else
        Set ws = ThisWorkbook.Sheets("Usuarios")
        lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
        
        For i = 2 To lastRow
            If ws.Cells(i, 5).Value = usuario And ws.Cells(i, 6).Value = clave Then
                loginCorrecto = True
                Exit For
            End If
        Next i
    End If

    If loginCorrecto Then
       Application.Visible = True
        Me.Hide

        '   Asegurar visibilidad solo de la hoja INICIO
        Dim hoja As Worksheet
        For Each hoja In ThisWorkbook.Worksheets
            If hoja.Name = "INICIO" Then
                hoja.Visible = xlSheetVisible
            Else
                hoja.Visible = xlSheetVeryHidden
            End If
        Next hoja

        Sheets("INICIO").Activate

        ' Registrar acceso
        Set wsLog = ThisWorkbook.Sheets("Accesos")
        filaLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
        wsLog.Cells(filaLog, 1).Value = Now
        wsLog.Cells(filaLog, 2).Value = usuario
        wsLog.Cells(filaLog, 3).Value = Environ("Username")

        ThisWorkbook.Save

        If Not esSuperAdmin Then
            Sheets("INICIO").Range("HD4").Value = ws.Cells(i, 4).Value
        Else
            MsgBox "Acceso concedido como Super Administrador.", vbInformation
        End If
    Else
        MsgBox "El usuario o la clave son incorrectos.", vbExclamation
    End If


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "Debe iniciar sesi�n para acceder a la aplicaci�n.", vbExclamation
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub
