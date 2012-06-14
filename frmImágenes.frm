VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmImágenes 
   Caption         =   "Elegí la imagen a insertar"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4095
   Icon            =   "frmImágenes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmImágenes.frx":08CA
   ScaleHeight     =   4635
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   1215
      Left            =   2880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Caption         =   "Insertar imagen"
      EstiloDelBoton  =   4
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
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      Height          =   1335
      Left            =   1260
      Top             =   3180
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vista previa de la imagen:"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "frmImágenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public swImagenDevuelta As String

Private Sub ButtonTransparent1_Click()
    Call Form_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Byte
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
        swImagenDevuelta = "ninguna"
        Unload Me
    End If
    
    If KeyCode = vbKeyReturn Then
        swImagenDevuelta = App.path + "\imagen\" + List1.List(List1.ListIndex)
        Unload Me
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.imágenes
         frmAyuda.Show
         Exit Sub
    End If
    
    If shiftkey = vbCtrlMask Then Decir ""
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Byte
    shiftkey = Shift And 7
    If shiftkey = 0 And (KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight) Then
        Decir List1.List(List1.ListIndex)
        Image1.Picture = LoadPicture(App.path + "\imagen\" + List1.List(List1.ListIndex))
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Call cargarImágenes
    Call centrarFormulario(Me)
    Me.swImagenDevuelta = "ninguna"
    Decir "Para poner una imagen en tu hoja, usá las flechas para seleccionarla y enter para insertarla"
End Sub

Sub cargarImágenes()
    Dim i As Integer, extensión As String
    File1.path = App.path + "\imagen"
    For i = 0 To File1.ListCount - 1
        extensión = LCase(Right(File1.List(i), 4))
        If extensión = ".jpg" Or extensión = ".gif" Or extensión = ".bmp" Or extensión = ".wmf" Or extensión = ".emf" Or extensión = ".ico" Then List1.AddItem File1.List(i)
    Next
End Sub

Private Sub List1_DblClick()
    Call Form_KeyDown(vbKeyReturn, 0)
End Sub
