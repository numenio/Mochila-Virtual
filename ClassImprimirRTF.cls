VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImpresoraRTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'apis
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal _
        hDC As Long, ByVal nIndex As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As _
        Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function CreateDC Lib "gdi32" Alias _
        "CreateDCA" (ByVal lpDriverName As String, ByVal _
        lpDeviceName As String, ByVal lpOutput As Long, _
        ByVal lpInitData As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC _
        As Long) As Long
        
'constantes
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113
Private Const WM_USER As Long = &H400
Private Const EM_Rango As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72

'UDT
Private Type Rect
      Left As Long
      Top As Long
      Right As Long
      Bottom As Long
End Type

Private Type CharRange
  cpMin As Long
  cpMax As Long
End Type

Private Type Rango
  hDC As Long
  hdcTarget As Long
  rc As Rect
  rcp As Rect
  chrg As CharRange
End Type

Dim HDCPrinter As Long
Dim MargenTop As Long
Dim MargenLeft As Long
Dim MargenRight As Long
Dim MargenBottom As Long
Dim Encabezado As String
Dim Pie As String

Private anchoLinea As Long


Private Function ImpresoraInstalada() As Boolean
    On Error Resume Next
    Dim Impresora As String
    Impresora = Printer.DeviceName
    If Err.Number Then
        ImpresoraInstalada = False
    Else
        ImpresoraInstalada = True
     End If
     On Error GoTo 0
End Function

Private Function IniciarPrinter(UnRichText As RichTextBox) As Boolean
    Dim BW As Long, BH As Long, r As Long, tOff As Long, lOff As Long
    Dim Lm As Long, Rm As Long
    Dim swHayImpresora As Boolean 'para saber si hay o no instalada una impresora

    swHayImpresora = ImpresoraInstalada 'verificar si hay una impresora instalada
    
    If swHayImpresora = True Then
        lOff = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
           
        Lm = MargenLeft - lOff: Rm = (Printer.Width - MargenRight) - lOff
        anchoLinea = Rm - Lm
        
        HDCPrinter = CreateDC(Printer.DriverName, Printer.DeviceName, 0, 0)
        r = SendMessage(UnRichText.hwnd, EM_SETTARGETDEVICE, HDCPrinter, ByVal anchoLinea)
        IniciarPrinter = True
    Else
        IniciarPrinter = False
    End If
End Function

Public Sub Imprimir(UnRichText As RichTextBox)

Dim lOff As Long, tOff As Long
Dim Lm As Long, Tm As Long, Rm As Long, Bm As Long
Dim uFR As Rango, Rdt As Rect, rcp As Rect
Dim lentexto As Long, Pos As Long
Dim r As Long, v As Integer
Dim swHayImpresora As Boolean

'IniciarPrinter UnRichText
    
    swHayImpresora = IniciarPrinter(UnRichText)
    
    If swHayImpresora = True Then
           Printer.Print ""
           
           With Printer
               
               v = 1
               
               .ScaleMode = vbTwips
               lOff = .ScaleX(GetDeviceCaps(.hDC, _
                                       PHYSICALOFFSETX), vbPixels, _
                                       vbTwips)
                                       
               tOff = .ScaleY(GetDeviceCaps(.hDC, _
                                      PHYSICALOFFSETY), vbPixels, _
                                      vbTwips)
               Lm = MargenLeft - lOff
               Tm = MargenTop - tOff
               Rm = (.ScaleWidth - MargenRight) + lOff
                          
               Bm = (.ScaleHeight - MargenBottom) + tOff
           
               rcp.Left = 0
               rcp.Top = 0
               rcp.Right = .ScaleWidth
               rcp.Bottom = .ScaleHeight
               Rdt.Left = Lm
               Rdt.Top = Tm
               Rdt.Right = Rm
               Rdt.Bottom = Bm
           
               uFR.hDC = .hDC
               uFR.hdcTarget = .hDC
               uFR.rc = Rdt
               uFR.rcp = rcp
               uFR.chrg.cpMin = 0
               uFR.chrg.cpMax = -1
           
               
               lentexto = Len(UnRichText.Text)
               .Font = "Courier New": .FontSize = 11
              
           Do
               
               Pos = SendMessage(UnRichText.hwnd, EM_Rango, 1, uFR)
               
               .CurrentX = (.ScaleWidth - .TextWidth(Encabezado)) / 2
               .CurrentY = (Tm - .TextHeight("x")) / 2
                                   
               Printer.Print Encabezado
               
               .CurrentX = (.ScaleWidth - .TextWidth(Pie)) / 2
               .CurrentY = Bm + (.ScaleHeight - Bm - .TextHeight("x")) / 2
               
               Printer.Print Pie
               
               .CurrentX = .ScaleWidth - .TextWidth("Página " & v)
               .CurrentY = Bm + (.ScaleHeight - Bm - .TextHeight("x")) / 2
               Printer.Print "Página " & v
               
               If Pos >= lentexto Then Exit Do
               
               uFR.chrg.cpMin = Pos
               
               .NewPage
               Printer.Print ""
               uFR.hDC = .hDC
               uFR.hdcTarget = .hDC
               v = v + 1
           Loop
           
           'Imprime
           .EndDoc
            
        End With
           r = SendMessage(UnRichText.hwnd, EM_Rango, 0, ByVal CLng(0))
           
           DeleteDC HDCPrinter
    Else 'si no hay una impresora instalada
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "No Hay una impresora instalada. ¿Desea instalar una impresora en este momento?"
        frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado = True Then
            Shell "rundll32 shell32.dll,SHHelpShortcuts_RunDLL AddPrinter", vbNormalFocus
        End If

        
'        If MsgBox("No Hay una impresora instalada. ¿Desea abrir el cuadro de diálogo de windows para instalar una impresora ? ", vbInformation + vbYesNo) = vbYes Then
'            ' cuadro de diálogo de windows para instalar una nueva imp.
'            Shell "rundll32 shell32.dll,SHHelpShortcuts_RunDLL AddPrinter", vbNormalFocus
'        End If
    End If
End Sub




'Propiedades (margenes, encabezado y pie de pagina)

Public Property Get MarginTop() As Long
MarginTop = MargenTop
End Property

Public Property Let MarginTop(ByVal vNewValue As Long)
MargenTop = vNewValue
End Property

Public Property Get MarginLeft() As Long
MarginLeft = MargenLeft
End Property

Public Property Let MarginLeft(ByVal vNewValue As Long)
MargenLeft = vNewValue
End Property

Public Property Get MarginRight() As Long
MarginRight = MargenRight
End Property

Public Property Let MarginRight(ByVal vNewValue As Long)
MargenRight = vNewValue
End Property

Public Property Get MarginBottom() As Long
MargenBottom = MarginBottom
End Property

Public Property Let MarginBottom(ByVal vNewValue As Long)
MargenBottom = vNewValue
End Property

Public Property Get Header() As String
Header = Encabezado
End Property

Public Property Let Header(ByVal vNewValue As String)
Encabezado = vNewValue
End Property

Public Property Get Footer() As String
Footer = Pie
End Property

Public Property Let Footer(ByVal vNewValue As String)
Pie = vNewValue
End Property


Private Sub Class_Terminate()
DeleteDC HDCPrinter
End Sub


