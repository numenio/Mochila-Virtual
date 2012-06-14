VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmRecordatorios 
   Caption         =   "Recordatorios"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   Icon            =   "frmRecordatorios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmRecordatorios.frx":08CA
   ScaleHeight     =   4455
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   540
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      Caption         =   "Añadir un recordatorio"
      EstiloDelBoton  =   1
      Picture         =   "frmRecordatorios.frx":14F2
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
      ForeColor       =   14737632
   End
   Begin TransparentButton.ButtonTransparent command2 
      Height          =   615
      Left            =   540
      TabIndex        =   1
      Top             =   2280
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      Caption         =   "Ver recordatorios ya añadidos"
      EstiloDelBoton  =   1
      Picture         =   "frmRecordatorios.frx":1DCC
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
      ForeColor       =   14737632
   End
End
Attribute VB_Name = "frmRecordatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim swEmpezando As Boolean 'para que la voz no hable pisando la introducción del load
Dim cadena As String
Dim swPulsóEnterParaAvanzar As Boolean

Private Sub Command1_Click()
    frmAñadirRecordatorio.Show
    swPulsóEnterParaAvanzar = True
    Unload Me
End Sub

Private Sub Command1_GotFocus()
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
    If swEmpezando = False Then
        Decir Command1.Caption
    Else
        Decir cadena + Command1.Caption
        swEmpezando = False
    End If
End Sub

Private Sub Command2_Click()
    frmFechaVerRec.Show
    swPulsóEnterParaAvanzar = True
    Unload Me
End Sub

Private Sub Command2_GotFocus()
    swEmpezando = False
    Decir Command2.Caption
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
'        frmAccesorios.Show 'si el cuaderno está cerrado se abre el principal
        Unload Me
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en los recordatorios"
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
        
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.recordatorios
         frmAyuda.Show
         Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    cadena = "entrando en los recordatorios, usá las flechas para moverte y enter para seleccionar"
    swEmpezando = True
    swPulsóEnterParaAvanzar = False
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
    
    If swPulsóEnterParaAvanzar = False Then frmAccesorios.Show
    
    'Call contarFormularios(False)
End Sub

