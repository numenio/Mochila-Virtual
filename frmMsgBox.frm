VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmMsgBox 
   Caption         =   "Hay que decidir"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3510
   Icon            =   "frmMsgBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMsgBox.frx":08CA
   ScaleHeight     =   3810
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   615
      Index           =   0
      Left            =   795
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      Caption         =   "Sí"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   615
      Index           =   1
      Left            =   795
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      Caption         =   "No"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   615
      Index           =   2
      Left            =   788
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      Caption         =   "Cancelar"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   2520
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cadenaAMostrar As String
Public swResultadoMostrado As Boolean
Public swCancelar As Boolean
Public swSíNoóAceptar As Boolean 'true sí o no, false acptar
Public swMostrarCancelar As Boolean
Public swEstoyAbierto As Boolean
Dim swComienzo As Boolean
Dim aux As String

Private Sub ButtonTransparent1_Click(Index As Integer)
    If Index = 0 Then
        swResultadoMostrado = True
    ElseIf Index = 1 Then
        swResultadoMostrado = False
    Else
        swCancelar = True
    End If
    Unload Me
End Sub

Private Sub ButtonTransparent1_GotFocus(Index As Integer)
    If Index <> 2 Then
        If swComienzo = True Then
            Decir cadenaAMostrar + aux + ButtonTransparent1(Index).Caption ', False, True
            swComienzo = False
        Else
            Decir ButtonTransparent1(Index).Caption ', True, False
        End If
    Else
        Decir "cancelar el salir de la carpeta" ', True, False
    End If
    
    If Index = 0 Then
        swResultadoMostrado = True
    Else
        swResultadoMostrado = False
    End If
End Sub

Private Sub ButtonTransparent1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift And 7 = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    'If KeyCode <> vbKeyReturn And KeyCode <> vbKeyDown Then Decir Label1.Caption
    If KeyCode = vbKeyF1 Then Decir Label1.Caption
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
    If KeyCode = vbKeyReturn Then Unload Me
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
'    If Shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
End Sub

Private Sub Form_Load()
    swEstoyAbierto = True
    swCancelar = False
    swComienzo = True
    Label1 = cadenaAMostrar
    swResultadoMostrado = False
    If swSíNoóAceptar = False Then 'el form puede ser de mensaje sí-no, o aceptar
        ButtonTransparent1(1).Visible = False
        ButtonTransparent1(0).Caption = "Aceptar"
        arreglarForm
        Me.Caption = "Información"
        aux = ". apretá espacio para aceptar este mensaje. "
    Else
        ButtonTransparent1(1).Visible = True
        ButtonTransparent1(1).Caption = "No"
        aux = ". Elegí con las flechas el botón sí o el botón no y aceptá con enter. "
    End If
    
    If swMostrarCancelar = True Then ButtonTransparent1(2).Visible = True
    
    Call centrarFormulario(frmMsgBox)
End Sub

Sub arreglarForm()
    Dim Centro As Single
      
    Centro = (Me.ScaleWidth - Label1.Width) / 2
    Label1.Move Centro, Label1.Top
    ButtonTransparent1(0).Top = 1800 'se baja el botón aceptar porque no se muestra el botón "no"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    swMostrarCancelar = False
    swEstoyAbierto = False
    cadenaAMostrar = ""
    Decir ""
End Sub
