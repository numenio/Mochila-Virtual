VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmAccesorios 
   Caption         =   "Accesorios"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4200
   Icon            =   "frmAccesorios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAccesorios.frx":08CA
   ScaleHeight     =   5490
   ScaleWidth      =   4200
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   "Calculadora"
      EstiloDelBoton  =   1
      Picture         =   "frmAccesorios.frx":29F9
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
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   "Día y Hora"
      EstiloDelBoton  =   1
      Picture         =   "frmAccesorios.frx":32D3
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
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   "Recordatorios"
      EstiloDelBoton  =   1
      Picture         =   "frmAccesorios.frx":3BAD
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
   Begin TransparentButton.ButtonTransparent ButtonTransparent2 
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   2760
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   "Reproductor de música"
      EstiloDelBoton  =   1
      Picture         =   "frmAccesorios.frx":4487
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
   Begin TransparentButton.ButtonTransparent ButtonTransparent3 
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   4440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   "  Pegar una imagen"
      EstiloDelBoton  =   1
      Picture         =   "frmAccesorios.frx":4D61
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
   Begin TransparentButton.ButtonTransparent btnDiccionarios 
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   3600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   "  Diccionarios"
      EstiloDelBoton  =   1
      Picture         =   "frmAccesorios.frx":563B
      OriginalPicSizeW=   16
      OriginalPicSizeH=   16
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
Attribute VB_Name = "frmAccesorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim swEmpezando As Boolean 'para que la voz no hable pisando la introducción del load
Dim cadena As String
Dim swPulsóEnterParaAvanzar As Boolean

Private Sub btnDiccionarios_Click()
    frmDiccionarios.Show
    swPulsóEnterParaAvanzar = True
    Unload Me
End Sub

Private Sub btnDiccionarios_GotFocus()
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
    Decir btnDiccionarios.Caption
End Sub

Private Sub btnDiccionarios_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub ButtonTransparent1_Click()
    frmRecordatorios.Show
    swPulsóEnterParaAvanzar = True
    Unload Me
End Sub

Private Sub ButtonTransparent1_GotFocus()
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
    Decir ButtonTransparent1.Caption
End Sub

Private Sub ButtonTransparent1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub ButtonTransparent2_Click()
    frmReproductorMúsica.Show
    swPulsóEnterParaAvanzar = True
    Unload Me
End Sub

Private Sub ButtonTransparent2_GotFocus()
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
    Decir ButtonTransparent2.Caption
End Sub

Private Sub ButtonTransparent2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub ButtonTransparent3_Click()
    frmImágenes.Show 1
    
    If swCuadernoAbierto = True Then
        If frmImágenes.swImagenDevuelta <> "ninguna" Then
            frmCuaderno.Picture1.Picture = LoadPicture(frmImágenes.swImagenDevuelta)
            frmCuaderno.pegarImagen
        End If
    End If
    
    If frmLectorEvaluaciones.swEstoyAbierto = True Then
        frmLectorEvaluaciones.Picture1.Picture = LoadPicture(frmImágenes.swImagenDevuelta)
        frmLectorEvaluaciones.pegarImagen
    End If
    
    Unload Me
End Sub

Private Sub ButtonTransparent3_GotFocus()
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
    Decir ButtonTransparent3.Caption
End Sub

Private Sub ButtonTransparent3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Command1_Click()
    frmCalculadora.Show
    swPulsóEnterParaAvanzar = True
    Unload Me
End Sub

Private Sub Command1_GotFocus()
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
    If swEmpezando = False Then
        Decir Command1.Caption
    Else
        Decir cadena + ". estás en: " + Command1.Caption '@ 1 (estás en)
    End If
End Sub

Private Sub Command2_Click()
    Dim cadenaFecha As String, cadenaTiempo As String
    cadenaFecha = Format(Date, "Long Date")
    cadenaTiempo = Format(Time, "HH:mm") 'Time
    '@ 2 hoy es, 3 es la hora, 4 recordá que...
    Decir "Hoy es " + cadenaFecha + ". Es la hora " + cadenaTiempo + ". recordá que seguís en los accesorios"
    swEmpezando = False
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
   
    If KeyCode = vbKeyEscape Then Unload Me
    
    '@ 5 para abrir...
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir Trim(nombreUsuario) + "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en los accesorios"
    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
        
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.accesorios
         frmAyuda.Show
         Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    Call reproducirForm(formularios.accesorios)
    '@ 6 entrando en los...
    cadena = "entrando en los accesorios, usá las flechas para moverte y enter para seleccionar"
    swEmpezando = True
    swPulsóEnterParaAvanzar = False
    If swCuadernoAbierto = True Or frmLectorEvaluaciones.swEstoyAbierto = True Then
        ButtonTransparent3.Visible = True
    Else
        ButtonTransparent3.Visible = False
    End If
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
    
    If swPulsóEnterParaAvanzar = False Then
        If swCuadernoAbierto = False And frmLectorEvaluaciones.swEstoyAbierto = False Then
            frmPrincipal.Show 'si el cuaderno está cerrado se abre el principal
        Else
            Decir "volviendo a tu carpeta" '@ 7 volviendo...
        End If
    End If
    
    'Call contarFormularios(False)
End Sub
