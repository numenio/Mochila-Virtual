VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmCalculadora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculadora"
   ClientHeight    =   3810
   ClientLeft      =   2580
   ClientTop       =   1485
   ClientWidth     =   3270
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CALC.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "CALC.frx":08CA
   ScaleHeight     =   3810
   ScaleWidth      =   3270
   WhatsThisHelp   =   -1  'True
   Begin TransparentButton.ButtonTransparent Number1 
      Height          =   480
      Index           =   0
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   847
      Caption         =   "0"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Number1 
      Height          =   480
      Index           =   7
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "7"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Number1 
      Height          =   480
      Index           =   1
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "1"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Number1 
      Height          =   480
      Index           =   2
      Left            =   720
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "2"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Number1 
      Height          =   480
      Index           =   3
      Left            =   1320
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "3"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Number1 
      Height          =   480
      Index           =   4
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "4"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Number1 
      Height          =   480
      Index           =   5
      Left            =   720
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "5"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Number1 
      Height          =   480
      Index           =   6
      Left            =   1320
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "6"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Number1 
      Height          =   480
      Index           =   8
      Left            =   720
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "8"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Number1 
      Height          =   480
      Index           =   9
      Left            =   1320
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "9"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Decimal1 
      Height          =   480
      Left            =   1320
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2400
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "."
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Cancel1 
      Height          =   480
      Left            =   2040
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "C"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent CancelEntry1 
      Height          =   480
      Left            =   2640
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "CE"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Percent1 
      Height          =   480
      Left            =   2640
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2400
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "%"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Operator1 
      Height          =   480
      Index           =   1
      Left            =   2040
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1200
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "+"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Operator1 
      Height          =   480
      Index           =   0
      Left            =   2640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "/"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Operator1 
      Height          =   480
      Index           =   2
      Left            =   2040
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "X"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Operator1 
      Height          =   480
      Index           =   3
      Left            =   2640
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1200
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "-"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent Operator1 
      Height          =   480
      Index           =   4
      Left            =   2040
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2400
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "="
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "cancelar (C): suprimir; cancelar entrada (CE): borrar; leer la pantalla: f1, leer el número por cifras: f2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "frmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Op1, Op2                ' Operando introducido previamente.
Dim DecimalFlag As Integer  ' ¿Está presente la coma decimal?
Dim NumOps As Integer       ' Número de operandos.
Dim LastInput               ' Indica el tipo del último evento Keypress.
Dim OpFlag                  ' Indica la operación pendiente.
Dim TempReadout
Public swEstoyAbierto As Boolean

' Procedimiento de evento Click para la tecla C (cancelar).
' Restablece la presentación e inicializa las variables.
Private Sub Cancel1_Click()
    Readout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    Decir "borraste toda la cuenta"
    Operator1(4).SetFocus
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")
End Sub

' Procedimiento de evento Click para la tecla CE (cancelar entrada).
Private Sub CancelEntry1_Click()
    Dim aux As String
    If Readout.Caption <> "0," Then
        aux = Mid(Right(Readout, 2), 1, 1) ' Left(Readout, InStr(Readout, Format(0, ".")) - 2) + Format(0, ".")
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 2) + Format(0, ".")
        If Readout.Caption = "," Then
            Readout = Format(0, "0.")
            DecimalFlag = False
            LastInput = "CE"
            Decir "borraste todo el número"
        Else
            Decir "borrando el número " + Str(aux) + ". Queda escrito " + Readout
        End If
    Else
        Readout = Format(0, "0.")
        DecimalFlag = False
        LastInput = "CE"
        Decir "borraste todo el número"
    End If
    Operator1(4).SetFocus
End Sub

' Procedimiento de evento Click para la tecla de la coma decimal (,).
' Si la última tecla presionada fue un operador, inicializa
' la lectura a "0". Si no, agrega la coma decimal
' a la presentación.
Private Sub Decimal1_Click()
    If LastInput = "NEG" Then
        Readout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        Readout = Format(0, "0.")
    End If
    Decir "coma"
    DecimalFlag = True
    LastInput = "NUMS"
    Operator1(4).SetFocus
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")
    Decir "Abriendo la calculadora. ya se pueden ingresar números"
    swEstoyAbierto = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim aux As Integer
    Dim shiftkey As Integer
    
    shiftkey = Shift And 7 'por si se usa el más - menos - dividido, etc que no están en el teclado numérico
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
'        If swCuadernoAbierto = True Then
'            Decir "cerrando la calculadora, volviendo a tu carpeta"
'        ElseIf frmLectorEvaluaciones.swEstoyAbierto = True Then
'            Decir "cerrando la calculadora, volviendo a la evaluación"
'        Else
'            frmAccesorios.Show
'        End If
        Unload Me
        Exit Sub
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF12 Then frmControl.Show
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en la calculadora"
    
    If KeyCode = vbKeyF8 Then
        If swCuadernoAbierto = True Then 'F8 pasa al cuaderno
            frmCuaderno.Show
            Decir "Pasando a tu carpeta, para volver a la calculadora, apretá F8"
            Exit Sub
        End If
        
        If frmLectorEvaluaciones.swEstoyAbierto = True Then 'pasa a la evaluación
            frmLectorEvaluaciones.Show
            Decir "Pasando la evaluación, para volver a la calculadora, apretá F8"
            Exit Sub
        End If
    End If
    
    If KeyCode = vbKeyBack Then CancelEntry1_Click 'tecla borrar borra el último número
    If KeyCode = vbKeyDelete Then Cancel1_Click 'tecla suprimir borra todo el cálculo
    
    Dim i As Integer, auxstr As String
    If KeyCode = vbKeyF1 Then Decir "Está escrito el número " + Readout 'f1 relee lo que está escrito en la calc
    If KeyCode = vbKeyF2 Then 'f2 lee el número por cifras
        auxstr = Left(Readout, Len(Readout.Caption) - 1)
        Decir "está escrito el número ", True, True
        For i = 1 To Len(Readout.Caption) - 1
            If auxstr = "" Then Exit For
            Decir " " + Left(auxstr, 1) + " ", True, True
            auxstr = Right(auxstr, Len(auxstr) - 1)
        Next
    End If
    
    Operator1(4).SetFocus 'se pone el foco en el botón igual para que cuando el usuario apriete el enter se dé el igual
    If shiftkey = 0 And (KeyCode >= 48 And KeyCode <= 57) Then 'los números superiores a las letras
        If LastInput <> "NUMS" Then
            Readout = Format(0, ".")
            DecimalFlag = False
        End If
        If DecimalFlag Then
            Readout = Readout + Chr(KeyCode)
        Else
            Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Chr(KeyCode) + Format(0, ".")
        End If
        If LastInput = "NEG" Then Readout = "-" & Readout
        LastInput = "NUMS"
        Decir Readout 'Chr(KeyCode)
'        Number_Click (Chr(KeyCode))
    End If
    If shiftkey = 0 And (KeyCode >= 96 And KeyCode <= 105) Then 'los números del teclado numérico
        If LastInput <> "NUMS" Then
            Readout = Format(0, ".")
            DecimalFlag = False
        End If
        aux = KeyCode - 48
        If DecimalFlag Then
            Readout = Readout + Chr(aux)
        Else
            Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Chr(aux) + Format(0, ".")
        End If
        If LastInput = "NEG" Then Readout = "-" & Readout
        LastInput = "NUMS"
        Decir Readout 'Chr(aux)
        'aux = KeyCode - 48
        'Number_Click (Chr(aux)) 'Readout = Readout + Chr(KeyCode)
    End If
    
    'se les pasa las funcioens para +, -, x y / del teclado numérico
    Select Case KeyCode
        Case 106
            Operator1_Click (2) 'multiplicación
        Case Asc("d")
            Operator1_Click (2) 'multiplicación
        Case Asc("D")
            Operator1_Click (2) 'multiplicación
        Case 107
            Operator1_Click (1) 'suma
        Case Asc("a")
            Operator1_Click (1) 'suma
        Case Asc("A")
            Operator1_Click (1) 'suma
        Case 108
            Operator1_Click (4) 'igual
        Case 109
            Operator1_Click (3) 'resta
        Case Asc("s")
            Operator1_Click (3) 'resta
        Case Asc("S")
            Operator1_Click (3) 'resta
        Case 110
            Decimal1_Click 'punto
        Case Asc(",")
            Decimal1_Click 'punto
        Case 111
            Operator1_Click (0) 'dividido
        Case Asc("f")
            Operator1_Click (0) 'dividido
        Case Asc("F")
            Operator1_Click (0) 'dividido
    End Select
       
    If shiftkey = vbShiftMask And KeyCode = 16 Then Operator1_Click (1) 'suma
    If shiftkey = 0 And KeyCode = 187 Then Operator1_Click (2) 'multiplicación
    If shiftkey = 0 And KeyCode = 190 Then Decimal1_Click 'punto
    If shiftkey = 0 And KeyCode = 189 Then Operator1_Click (3) 'resta
    If shiftkey = vbShiftMask And KeyCode = 55 Then Operator1_Click (0) 'dividido
    If (shiftkey = vbShiftMask And KeyCode = 53) Or KeyCode = vbKeyG Then Percent1_Click 'porciento
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'control + c copia la pantalla al portapapeles
        Clipboard.Clear
        If Asc(Right(Readout, 1)) = Asc(",") Or Asc(Right(Readout, 1)) = Asc(".") Then
            Clipboard.SetText (Left(Readout, Len(Readout) - 1))
            Decir "texto copiado." + Left(Readout, Len(Readout) - 1)
        Else
            Clipboard.SetText (Readout)
            Decir "texto copiado." + Readout
        End If
    End If
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.calculadora
         frmAyuda.Show
         Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    swEstoyAbierto = False
    If swCuadernoAbierto = True Then
        Decir "cerrando la calculadora, volviendo a tu carpeta"
    ElseIf frmLectorEvaluaciones.swEstoyAbierto = True Then
        Decir "cerrando la calculadora, volviendo a la evaluación"
    Else
        frmAccesorios.Show
    End If
    'Call contarFormularios(False)
End Sub

' Procedimiento de evento para las teclas de los números (0-9).
' Agrega el número nuevo a la cantidad mostrada.
Private Sub Number1_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number1(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Number1(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
    Decir Readout 'Number(Index).Caption
    Operator1(4).SetFocus
End Sub

' Procedimiento de evento para las teclas de operación (+, -, x, /, =).
' Si la última tecla presionada formaba parte de un número,
' se incrementa en NumOps. Si ya hubiera presente un operando,
' establece Op1. Si hubiera dos, establece Op1 al resultado
' de la operación sobre Op1 y la cadena de entrada actual,
' y muestra el resultado.
Private Sub Operator1_Click(Index As Integer)
    Dim cadena As String
    
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
            If Operator1(Index).Caption = "-" And LastInput <> "NEG" Then
                Readout = "-" & Readout
                LastInput = "NEG"
                Decir "negativo"
            Else
                Decir "No se pueden apretar las teclas de operaciones sin escribir primero un número"
            End If
        Case 1
            Op1 = Readout
            If Operator1(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
                Readout = "-"
                LastInput = "NEG"
                Decir "negativo"
            End If
        Case 2
            Op2 = TempReadout
            Select Case OpFlag
                Case "+"
                    Op1 = CDbl(Op1) + CDbl(Op2)
                Case "-"
                    Op1 = CDbl(Op1) - CDbl(Op2)
                Case "X"
                    Op1 = CDbl(Op1) * CDbl(Op2)
                Case "/"
                    If Op2 = 0 Then
                       frmMsgBox.cadenaAMostrar = "No se puede dividir por cero"
                        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
                        frmMsgBox.Show 1
                    Else
                       Op1 = CDbl(Op1) / CDbl(Op2)
                    End If
                Case "="
                    Op1 = CDbl(Op2)
                Case "%"
                    Op1 = CDbl(Op1) * CDbl(Op2)
                End Select
            Readout = Op1 & Format(0, ".")
            NumOps = 1
            'Decir "El resultado es " + Readout
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator1(Index).Caption
        Select Case OpFlag
            Case "+"
                cadena = "más"
            Case "-"
                cadena = "menos"
            Case "X"
                cadena = "multiplicado por"
            Case "/"
                If Op2 = 0 Then
                   cadena = "No se puede dividir por cero"
                Else
                   cadena = "dividido por"
                End If
            Case "="
                cadena = "igual " + Readout
            Case "%"
                cadena = "porciento"
            End Select
    End If
    Decir cadena
    Operator1(4).SetFocus
End Sub

' Procedimiento de evento Click para la tecla de porcentaje (%).
' Calcula y muestra el porcentaje del primer operando.
Private Sub Percent1_Click()
    Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
    Decir "porciento. Queda escrito " + Readout
    Operator1(4).SetFocus
End Sub

