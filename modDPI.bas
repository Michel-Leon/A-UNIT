Attribute VB_Name = "modDPI"
' MÓDULO: modDPI (CORREGIDO)
'=============================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
#End If

Private Const LOGPIXELSX As Long = 88

Public Function ObtenerDPI() As Long
    #If VBA7 Then
        Dim hDC As LongPtr
    #Else
        Dim hDC As Long
    #End If
    
    hDC = GetDC(0)
    ObtenerDPI = GetDeviceCaps(hDC, LOGPIXELSX)
    ReleaseDC 0, hDC
End Function

Public Function FactorEscala() As Double
    FactorEscala = 144 / ObtenerDPI()
End Function

' *** CAMBIO CLAVE: Se usa Object en vez de MSForms.UserForm ***
Public Sub EscalarFormulario(frm As Object)
    Dim factor As Double
    Dim ctrl As Object
    
    factor = FactorEscala()
    
    If Abs(factor - 1) < 0.01 Then Exit Sub
    
    Dim originalWidth As Double
    Dim originalHeight As Double
    originalWidth = frm.Width
    originalHeight = frm.Height
    
    frm.Width = originalWidth * factor
    frm.Height = originalHeight * factor
    
    For Each ctrl In frm.Controls
        On Error Resume Next  ' Protege contra controles sin estas propiedades
        ctrl.Left = ctrl.Left * factor
        ctrl.Top = ctrl.Top * factor
        ctrl.Width = ctrl.Width * factor
        ctrl.Height = ctrl.Height * factor
        ctrl.Font.Size = ctrl.Font.Size * factor
        On Error GoTo 0
    Next ctrl
    
    frm.StartUpPosition = 0
    frm.Left = Application.Left + (Application.Width - frm.Width) / 2
    frm.Top = Application.Top + (Application.Height - frm.Height) / 2
End Sub
