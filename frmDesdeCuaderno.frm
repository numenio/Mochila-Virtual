VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmDesdeCuaderno 
   Caption         =   "Materia X"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4125
   Icon            =   "frmDesdeCuaderno.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmDesdeCuaderno.frx":08CA
   ScaleHeight     =   4605
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin TransparentButton.ButtonTransparent Command3 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Caption         =   "Abrir una hoja anterior de la carpeta"
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
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Caption         =   "Actividades"
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
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Caption         =   "Libros"
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
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Caption         =   "Abrir un archivo externo (tipo txt ó rtf)"
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
Attribute VB_Name = "frmDesdeCuaderno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim swReciénEmpiezo As Boolean 'para ver si la voz espera en el botón actividades

Private Sub ButtonTransparent1_Click()
    frmDiálogoAbrir.quéArchivosFiltrar = "*.rtf;*.txt"
    frmDiálogoAbrir.Show 1
    If frmDiálogoAbrir.archivoDevuelto <> "" Then
        frmCuaderno.swArchivoExterno = True
        frmCuaderno.RichTextBox1.LoadFile frmDiálogoAbrir.archivoDevuelto
    End If
    Unload Me
    Exit Sub
End Sub

Private Sub ButtonTransparent1_GotFocus()
    Decir ButtonTransparent1.Caption
End Sub

Private Sub ButtonTransparent1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Command1_Click()
    frmActividades.Show
    Unload Me
End Sub

Private Sub Command1_GotFocus()
    If swReciénEmpiezo = True Then
        Decir "Entrando a las actividades, libros y hojas ya escritas de " + miMateria + ".elegí con las flechas qué querés abrir y aceptá con enter. Estás en" + Command1.Caption
        swReciénEmpiezo = False
    Else
        Decir Command1.Caption
    End If
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Command2_Click()
    Unload Me
    frmLibros.Show
End Sub

Private Sub Command2_GotFocus()
    Decir Command2.Caption
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Command3_Click()
    Call reproducirForm(formularios.tareasAnt)
    If swMostrarAñoEnTareas Then 'si se muestran las tareas de todos los años, se muestra el form de los años
        frmAñoTareas.Show
    Else
        frmTareasAnt.añoParaVerMeses = Year(Date) 'si se ve sólo el año actual
        frmTareasAnt.Show
    End If
    frmCuaderno.swArchivoExterno = False
    Unload Me
End Sub

Private Sub Command3_GotFocus()
    Decir Command3.Caption
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
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
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en abrir una actividad, libro u hoja ya escrita"
    
    If KeyCode = vbKeyEscape Then
        Decir "volviendo a tu carpeta"
        Unload Me
    End If
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.desdeCuaderno
         frmAyuda.Show
         Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swReciénEmpiezo = True
    Me.Caption = "Actividades, libros y hojas antiguas de " + miMateria
    If swPermitirAbrirArchivos = True Then
        ButtonTransparent1.Visible = True
    Else
        ButtonTransparent1.Visible = False
    End If
    'Decir "Entrando a las actividades, libros y hojas ya escritas de " + miMateria + ".elegí con las flechas qué querés abrir y aceptá con enter. Estás en"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Decir "" 'para callar la voz al salir
    If swSalir = True Then
        If SalirDelPrograma = True Then
            chauPrograma
        Else
            Cancel = 1
            swSalir = False
        End If
        Exit Sub
    End If
    'Call contarFormularios(False)
End Sub
