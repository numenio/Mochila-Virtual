VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmDiccionarios 
   Caption         =   "Diccionarios inlcuídos en su Mochila"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmDiccionario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmDiccionario.frx":6852
   ScaleHeight     =   4170
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      Caption         =   "    Mostrar los capítulos del libro seleccionado"
      EstiloDelBoton  =   1
      Picture         =   "frmDiccionario.frx":88AA
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
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diccionarios disponibles:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   2265
   End
End
Attribute VB_Name = "frmDiccionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim swHizoClickParaAvanzar As Boolean

Private Sub Command1_Click()
    If List1.ListIndex <> -1 Then
        swHizoClickParaAvanzar = True
        Unload Me
    Else
        Decir "No se ha elegido ningún diccionario de la lista, por favor elija uno e intente de nuevo"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then Unload Me
    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en los libros"
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda
         frmAyuda.formulario = formularios.libros
         frmAyuda.Show
         Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Call centrarFormulario(Me)
    File1.path = App.path + "\Diccionarios\"
    'Dir1.Refresh
    For i = 0 To File1.ListCount - 1
        List1.AddItem Left(File1.List(i), InStrRev(File1.List(i), ".") - 1)
    Next
    swHizoClickParaAvanzar = False
    Decir "Entrando en los diccionarios. Elegí qué diccionario querés abrir y aceptá con enter"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If swHizoClickParaAvanzar = False Then
        frmAccesorios.Show
        Decir "cerrando los diccionarios, volviendo a los accesorios"
    Else
        frmDiccionarioElegido.diccionarioElegido = List1.List(List1.ListIndex) + ".txt"
        frmDiccionarioElegido.Show
    End If
End Sub

Private Sub List1_DblClick()
    Command1_Click
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Command1_Click
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        Decir List1.List(List1.ListIndex)
        sonido = sndPlaySound(App.path + "\sonidos\td.wav", SND_ASYNC)
    End If
End Sub

Private Sub List1_GotFocus()
    Decir List1.List(List1.ListIndex), True, True
End Sub

Private Sub Command1_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

