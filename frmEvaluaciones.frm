VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmEvaluaciones 
   Caption         =   "Evaluaciones"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4065
   Icon            =   "frmEvaluaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmEvaluaciones.frx":08CA
   ScaleHeight     =   4545
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   885
      TabIndex        =   0
      Top             =   1215
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Caption         =   "Hacer una evaluaci�n"
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
   Begin TransparentButton.ButtonTransparent Command2 
      Height          =   615
      Left            =   885
      TabIndex        =   1
      Top             =   2415
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Caption         =   "Ver las evaluaciones ya hechas"
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
End
Attribute VB_Name = "frmEvaluaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public swMateria As String
Dim swReci�nEmpiezo As Boolean
Dim swPuls�EnterParaAvanzar As Boolean

Private Sub Command1_Click()
    Dim d�a As Byte, a�o As Integer
    d�a = Day(Date)
    a�o = Year(Date)
    frmLectorEvaluaciones.swArchivoParaLeer = Trim(Str(d�a)) + "-" + Trim(Str(a�o))
    frmLectorEvaluaciones.swD�aParaAbrir = d�a
    frmLectorEvaluaciones.swMateriaParaAbrir = swMateria
    frmLectorEvaluaciones.swNumMesParaAbrir = Month(Date)
    frmLectorEvaluaciones.swS�loLeer = False
    frmLectorEvaluaciones.Show
    swPuls�EnterParaAvanzar = True
    Unload Me
End Sub

Private Sub Command1_GotFocus()
    If swReci�nEmpiezo = True Then
        Decir "Entrando en las evaluaciones. Eleg� con las flechas qu� quer�s hacer y acept� con enter. Est�s en" + Command1.Caption
        swReci�nEmpiezo = False
    Else
        Decir Command1.Caption
    End If
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Command2_Click()
    If usuario.mostrarA�oEnEvaluaciones = True Then  'si se muestran evaluaciones de todos los a�os
        frmA�oEvaluaciones.swMateria = Me.swMateria
        frmA�oEvaluaciones.Show
    Else
        frmMesEvaluaciones.swMateria = Me.swMateria
        frmMesEvaluaciones.a�oParaVerMeses = Year(Date)
        frmMesEvaluaciones.Show
    End If
    swPuls�EnterParaAvanzar = True
    Unload Me
End Sub

Private Sub Command2_GotFocus()
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
    Decir Command2.Caption
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el men� de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
'        Decir "volviendo al men� principal"
'        frmPrincipal.Show
        Unload Me
        Exit Sub
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de m�sica, ten�s que estar en el men� principal o en una carpeta. ahora est�s en las evaluaciones"
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al men� de la aplicaci�n. Para leer los �tems de este men� necesit�s jaws u otro lector de pantallas. Para volver a la mochila, apret� escape"
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.evaluaciones
         frmAyuda.Show
         Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    Call reproducirForm(formularios.actividades)
    swReci�nEmpiezo = True
    swPuls�EnterParaAvanzar = False
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
    
    If swPuls�EnterParaAvanzar = False Then
        Decir "volviendo al men� principal"
        frmPrincipal.Show
    End If
    
    'Call contarFormularios(False)
End Sub
