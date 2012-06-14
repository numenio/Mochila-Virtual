VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmMateriasEvaluaciones 
   Caption         =   "Elegir una materia"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   Icon            =   "frmMateriasEvaluaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMateriasEvaluaciones.frx":08CA
   ScaleHeight     =   5520
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      ItemData        =   "frmMateriasEvaluaciones.frx":2922
      Left            =   150
      List            =   "frmMateriasEvaluaciones.frx":2924
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   630
      TabIndex        =   1
      Top             =   4680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      Caption         =   " Evaluaciones de la materia seleccionada"
      EstiloDelBoton  =   1
      Picture         =   "frmMateriasEvaluaciones.frx":2926
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
End
Attribute VB_Name = "frmMateriasEvaluaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim swPulsóEnterParaAvanzar As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
'        frmPrincipal.Show
        Unload Me
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en las evaluaciones"
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.materiasEvaluaciones
         frmAyuda.Show
         Exit Sub
    End If
End Sub


Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swPulsóEnterParaAvanzar = False
    Decir "Elegí con las flechas la materia de que querés hacer una evaluación o ver las que ya hiciste, y aceptá con enter"
    
    Dim archivolibre As Byte, cadena As String
    archivolibre = FreeFile 'se abren las materias
    Open App.path + "\datos\materias.txt" For Input As archivolibre
    While Not EOF(archivolibre)
        Line Input #archivolibre, cadena
        List1.AddItem Trim(cadena) 'se añaden las materias al combo
    Wend
    Close #archivolibre
End Sub

Private Sub Form_Paint()
    List1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If swSalir = True Then
        If SalirDelPrograma = True Then
            chauPrograma
        Else
            Cancel = 1
            swSalir = False
        End If
        Exit Sub
    End If
    
    If swPulsóEnterParaAvanzar = False Then frmPrincipal.Show
    
    'Call contarFormularios(False)
End Sub

Private Sub List1_DblClick()
    If List1.ListIndex <> -1 Then
        frmEvaluaciones.swMateria = List1.Text
        frmEvaluaciones.Show
        swPulsóEnterParaAvanzar = True
        Unload Me
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then List1_DblClick
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

Private Sub Command1_GotFocus()
    Decir Command1.Caption
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Command1_Click()
    List1_DblClick
End Sub
