VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmControl 
   Caption         =   "Configuración"
   ClientHeight    =   7545
   ClientLeft      =   285
   ClientTop       =   675
   ClientWidth     =   9975
   Icon            =   "frmControl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmControl.frx":08CA
   ScaleHeight     =   7545
   ScaleWidth      =   9975
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   8640
      TabIndex        =   31
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Guardar todos los cambios"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1080
      TabIndex        =   30
      Top             =   7080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog diálogo 
      Left            =   240
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Materias"
      TabPicture(0)   =   "frmControl.frx":326A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label6(3)"
      Tab(0).Control(1)=   "Image1(2)"
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(4)=   "Label17"
      Tab(0).Control(5)=   "Label18"
      Tab(0).Control(6)=   "Line2"
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(8)=   "List2"
      Tab(0).Control(9)=   "Command4(0)"
      Tab(0).Control(10)=   "Command4(1)"
      Tab(0).Control(11)=   "Text2"
      Tab(0).Control(12)=   "Command5"
      Tab(0).Control(13)=   "Command6"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Voz"
      TabPicture(1)   =   "frmControl.frx":3286
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameVoz1"
      Tab(1).Control(1)=   "chkVoz"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "Check8"
      Tab(1).Control(4)=   "Image1(0)"
      Tab(1).Control(5)=   "Label6(0)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Otros"
      TabPicture(2)   =   "frmControl.frx":32A2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label6(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Image1(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Check7"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Check6"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Check3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Check2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame21"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Check5"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame1"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Check1"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Check4"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Frame3"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Créditos"
      TabPicture(3)   =   "frmControl.frx":32BE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label24"
      Tab(3).Control(1)=   "Label23"
      Tab(3).Control(2)=   "Label2"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "Corrector ortográfico"
         Height          =   1815
         Left            =   600
         TabIndex        =   54
         Top             =   4560
         Width           =   3735
         Begin VB.ComboBox cmbIdiomasCorrector 
            Height          =   315
            Left            =   840
            TabIndex        =   56
            Text            =   "Combo2"
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Usar el corrector ortográfico"
            Height          =   375
            Left            =   240
            TabIndex        =   55
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblIdioma 
            Alignment       =   2  'Center
            Caption         =   "Para cambiar el idioma se  tiene que estar en el menú principal"
            Height          =   375
            Left            =   360
            TabIndex        =   58
            Top             =   1320
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label Label5 
            Caption         =   "Idioma:"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Mostrar evaluaciones sólo del año actual"
         Height          =   495
         Left            =   600
         TabIndex        =   16
         Top             =   600
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar las actividades al docente optimizadas para un lector de pantallas (por ej. Jaws)"
         Height          =   495
         Left            =   600
         TabIndex        =   22
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Letra y fondo de la carpeta"
         Height          =   2415
         Left            =   4560
         TabIndex        =   45
         Top             =   2640
         Width           =   4215
         Begin VB.ComboBox cmbFormColor 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1680
            Width           =   1335
         End
         Begin VB.ComboBox cmbFontColor 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox cmbFontName 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox cmbFontSize 
            Height          =   315
            Left            =   240
            TabIndex        =   28
            Text            =   "Combo4"
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "Tamaño de la letra:"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Color del fondo:"
            Height          =   255
            Left            =   2280
            TabIndex        =   48
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Color de la letra:"
            Height          =   255
            Left            =   2280
            TabIndex        =   47
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo de letra:"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Imprimir en tinta automáticamente sin mostrar las opciones de impresión"
         Height          =   495
         Left            =   600
         TabIndex        =   19
         Top             =   2220
         Width           =   3615
      End
      Begin VB.Frame Frame21 
         Caption         =   "Usuario"
         Height          =   1815
         Left            =   4560
         TabIndex        =   42
         Top             =   600
         Width           =   4215
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   120
            MaxLength       =   50
            TabIndex        =   23
            Top             =   720
            Width           =   3735
         End
         Begin VB.OptionButton Option10 
            Caption         =   "hombre"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   24
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton Option10 
            Caption         =   "mujer"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   25
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Nombre de la persona que va a usar la computadora:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label Label20 
            Caption         =   "Sexo:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1320
            Width           =   615
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Mostrar actividades sólo del año actual"
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Top             =   1140
         Width           =   3375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Mostrar hojas guardadas sólo del año actual"
         Height          =   495
         Left            =   600
         TabIndex        =   18
         Top             =   1680
         Width           =   3615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Usar música de fondo para identificar las distintas partes de la mochila"
         Height          =   375
         Left            =   600
         TabIndex        =   20
         Top             =   2760
         Width           =   3975
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Permitir abrir en las carpetas archivos externos a la mochila, archivos con extensiones .rtf o .txt"
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   3240
         Width           =   3975
      End
      Begin VB.Frame frameVoz1 
         Caption         =   "Elegir la voz:"
         Height          =   3015
         Left            =   -73680
         TabIndex        =   40
         Top             =   1800
         Width           =   2775
         Begin VB.OptionButton Option8 
            Caption         =   "Usar voces avanzadas"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Usar voces simples"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   1800
            Width           =   2055
         End
         Begin VB.ComboBox Combo9 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2160
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   960
            Width           =   2535
         End
      End
      Begin VB.CheckBox chkVoz 
         Caption         =   "Deshabilitar el uso de la voz"
         Height          =   495
         Left            =   -73680
         TabIndex        =   9
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         Caption         =   "Velocidad de la voz"
         Height          =   4455
         Left            =   -68760
         TabIndex        =   37
         Top             =   1080
         Width           =   1695
         Begin ComctlLib.Slider Slider1 
            Height          =   3615
            Left            =   720
            TabIndex        =   15
            Top             =   480
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   6376
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   1
            Min             =   -10
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "rápida"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   39
            Top             =   4080
            Width           =   435
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "lenta"
            Height          =   195
            Left            =   960
            TabIndex        =   38
            Top             =   240
            Width           =   345
         End
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Leer renglones en la carpeta, actividades y libros"
         Height          =   735
         Left            =   -73680
         TabIndex        =   14
         Top             =   5040
         Width           =   3975
      End
      Begin VB.CommandButton Command6 
         Height          =   735
         Left            =   -66840
         Picture         =   "frmControl.frx":32DA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Eliminar una materia"
         Top             =   2955
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Añadir"
         Height          =   495
         Left            =   -74040
         TabIndex        =   2
         ToolTipText     =   "Añadir una materia a la lista"
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   -74040
         MaxLength       =   30
         TabIndex        =   1
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CommandButton Command4 
         Height          =   735
         Index           =   1
         Left            =   -66840
         Picture         =   "frmControl.frx":3BA4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Bajar un lugar"
         Top             =   3885
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Height          =   735
         Index           =   0
         Left            =   -66840
         Picture         =   "frmControl.frx":446E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Subir un lugar"
         Top             =   2040
         Width           =   615
      End
      Begin VB.ListBox List2 
         Height          =   2595
         Left            =   -69480
         TabIndex        =   4
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Añadir una materia desde el historial"
         Height          =   495
         Left            =   -74010
         TabIndex        =   3
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contacto: guillermo_toscani@yahoo.com.ar - educacion@tiflonexos.com.ar"
         Height          =   195
         Left            =   -72240
         TabIndex        =   53
         Top             =   3720
         Width           =   5310
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Idea y Conceptualización: Mara Lis Villar - Guillermo Toscani"
         Height          =   195
         Left            =   -72240
         TabIndex        =   52
         Top             =   3060
         Width           =   4245
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Desarrollo y Programación: Guillermo Toscani"
         Height          =   195
         Left            =   -72240
         TabIndex        =   51
         Top             =   2400
         Width           =   3195
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   9000
         Picture         =   "frmControl.frx":4D38
         Top             =   6090
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Si necesita ayuda, haga click aquí:"
         Height          =   375
         Index           =   1
         Left            =   6720
         TabIndex        =   50
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   -66000
         Picture         =   "frmControl.frx":5602
         Top             =   6090
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Si necesita ayuda, haga click aquí:"
         Height          =   375
         Index           =   0
         Left            =   -68280
         TabIndex        =   41
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         X1              =   -70200
         X2              =   -70200
         Y1              =   1560
         Y2              =   5640
      End
      Begin VB.Label Label18 
         Caption         =   "Suba o baje las materias para que se muestren en el orden que usted desee"
         Height          =   615
         Left            =   -69480
         TabIndex        =   36
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Materias ya añadidas:"
         Height          =   375
         Left            =   -69480
         TabIndex        =   35
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label16 
         Caption         =   "Escriba aquí las materias que desea que muestre el cuaderno:"
         Height          =   615
         Left            =   -74040
         TabIndex        =   34
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar en el historial una materia añadida y borrada:"
         Height          =   495
         Left            =   -74040
         TabIndex        =   33
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   -66000
         Picture         =   "frmControl.frx":5ECC
         Top             =   6090
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Si necesita ayuda, haga click aquí:"
         Height          =   375
         Index           =   3
         Left            =   -68280
         TabIndex        =   32
         Top             =   6120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sw2 As Boolean 'para ver si habla en este form sapi4
Dim sw1 As Boolean 'para ver si habla en este form sapi5
Dim objPipe As clsPipe 'para comunicarse con aspell

Private Enum pestaña
    Créditos
    Voces
    Materias
    Otros
End Enum

Private Sub chkVoz_Click()
    If chkVoz.Value = 1 Then 'si está desactivada la voz
        frameVoz1.Enabled = False
        Option8(0).Enabled = False
        Option8(1).Enabled = False
        Combo1.Enabled = False
        Combo9.Enabled = False
        Frame2.Enabled = False
        Slider1.Enabled = False
        Check8.Enabled = False
        Label3.ForeColor = &HC0C0C0 'gris
        Label4.ForeColor = &HC0C0C0 'gris
    Else 'si está activada la voz
        frameVoz1.Enabled = True
        Option8(0).Enabled = True
        Option8(1).Enabled = True
        Combo1.Enabled = True
        Combo9.Enabled = True
        Frame2.Enabled = True
        Slider1.Enabled = True
        Check8.Enabled = True
        Label3.ForeColor = &H80000012 'negro
        Label4.ForeColor = &H80000012 'negro
    End If
End Sub


Private Sub Command1_Click()
    frmHistorial.Show 1
End Sub


Private Sub Command12_Click() 'botón aceptar
    On Error Resume Next
    
    Call guardarMateriasGeneral
    
    nombreUsuario = Trim(Text3.Text)
    
    If Check1.Value = 1 Then
        usuario.swVerActividadesConJaws = True
    Else
        usuario.swVerActividadesConJaws = False
    End If

    If Check2.Value = 1 Then
        swMostrarAñoEnActividades = False
    Else
        swMostrarAñoEnActividades = True
    End If
    
    If Check3.Value = 1 Then
        swMostrarAñoEnTareas = False
    Else
        swMostrarAñoEnTareas = True
    End If
    
    If Check4.Value = 1 Then 'evaluaciones sólo del año actual
        usuario.mostrarAñoEnEvaluaciones = False
    Else
        usuario.mostrarAñoEnEvaluaciones = True
    End If
    
    If Check5 = 1 Then
        swImprimirDirecto = True
    Else
        swImprimirDirecto = False
    End If
    
    If Check9 = 1 Then
        swUsarCorrectorOrtográfico = True
    Else
        swUsarCorrectorOrtográfico = False
    End If
    
    If Check8 = 1 Then
        swLeerRenglones = True
    Else
        swLeerRenglones = False
    End If
    
    If Check6 = 1 Then
        swMúsicaDeFondo = True
    Else
        swMúsicaDeFondo = False
    End If
    
    If Check7 = 1 Then
        swPermitirAbrirArchivos = True
    Else
        swPermitirAbrirArchivos = False
    End If
    
    If chkVoz = 1 Then
        swHablarVoz = False
    Else
        swHablarVoz = True
    End If

    If Option10(0).Value = True Then  'se elige si el usuario es hombre o mujer
        swUsuarioMujer = False
    Else
        swUsuarioMujer = True
    End If
    
    If Option8(0).Value = True Then
        swSapi5 = True
    Else
        swSapi5 = False
    End If
    
    NombreFuente = cmbFontName  'se ajusta la fuente del programa
    'se graba en una variable el color de la fuente del programa
    Select Case cmbFontColor.List(cmbFontColor.ListIndex)
        Case "Blanco"
           colorFuente = vbWhite
        Case "Azul"
           colorFuente = vbBlue
        Case "Negro"
           colorFuente = vbBlack
        Case "Verde"
           colorFuente = vbGreen
        Case "Rojo"
            colorFuente = vbRed
        Case "Amarillo"
            colorFuente = vbYellow
    End Select
    
    'se graba en una variable el color de la fuente del programa
    Select Case cmbFormColor.List(cmbFormColor.ListIndex)
        Case "Blanco"
           colorFondo = vbWhite
        Case "Azul"
           colorFondo = vbBlue
        Case "Negro"
           colorFondo = vbBlack
        Case "Verde"
           colorFondo = vbGreen
        Case "Rojo"
            colorFondo = vbRed
        Case "Amarillo"
            colorFondo = vbYellow
    End Select
    
    tamañoFuente = cmbFontSize.Text 'se ajusta el tamaño de la fuente
    
    If swCuadernoAbierto = False Or frmLectorEvaluaciones.swEstoyAbierto = False Then
        Select Case cmbIdiomasCorrector.List(cmbIdiomasCorrector.ListIndex)
            Case "Español"
                idiomaAspell = "es"
            Case "Inglés"
                idiomaAspell = "en"
            Case "Francés"
                idiomaAspell = "fr"
            Case "Portugués"
                idiomaAspell = "pt"
        End Select
        
        If swAspellInstalado = True Then
            If objPipe.Running = True Then Call objPipe.Terminate
            KillProcess ("cmd.exe")
            KillProcess ("aspell.exe")
            Call objPipe.Execute("CMD.EXE")
            Call Sleep(200)
            Call objPipe.Write_("c:" & vbCrLf)
            Call objPipe.Write_("cd " & rutaDeAspell & vbCrLf)
            Call Sleep(100)
            Call objPipe.Write_("aspell -a -d " & idiomaAspell & vbCrLf)
            Call Sleep(200)
            'Debug.Print objPipe.Read
        End If
    End If
    
    nombreSapi5 = Combo1.List(Combo1.ListIndex)
    nombreSapi4 = Combo9.List(Combo9.ListIndex)
    
    GuardarDatosUsuario 'se guardan las preferencias del usuario
    
    If swCuadernoAbierto = True Then frmCuaderno.Refresh 'se actualiza el cuaderno si está abierto
    If swLibroAbierto = True Then frmLectorLibro.Refresh 'si está abierto el lector de libros, se lo actualiza
    If swActividadAbierta = True Then frmLectorActividad.Refresh 'si está abierto el lector de actividad, se lo actualiza para ver si se pueden o no modificar las actividades
    Unload Me
End Sub


Private Sub Command4_Click(Index As Integer) 'subir una materia
    Dim auxÍndice
    Dim auxCadena
    
    If List2.List(List2.ListIndex) = "" Then
        frmMsgBox.cadenaAMostrar = "Primero elija una materia de la lista y luego pulse este botón"
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    If Index = 0 Then 'si se presiona el botón subir
        If List2.ListIndex <> 0 Then 'si la materia a subir no está ya arriba
            auxÍndice = List2.ListIndex - 1 'se cargan en los auxiliares los datos del elemento a bajar
            auxCadena = List2.List(List2.ListIndex - 1)
            List2.List(auxÍndice) = List2.List(List2.ListIndex) 'se hace el cambio para abajo
            List2.List(List2.ListIndex) = auxCadena
            List2.ListIndex = auxÍndice
            
            swHuboCambioEnMaterias = True 'para guardar los cambios
        End If
    Else 'si se presiona el botón bajar
        If List2.ListIndex <> List2.ListCount - 1 Then 'se chequea que no esté en el fin la materia a bajar
            auxÍndice = List2.ListIndex + 1 'se cargan en los auxiliares los datos del elemento a bajar
            auxCadena = List2.List(List2.ListIndex + 1)
            
            List2.List(auxÍndice) = List2.List(List2.ListIndex) 'se hace el cambio para abajo
            List2.List(List2.ListIndex) = auxCadena
            List2.ListIndex = auxÍndice
            
            swHuboCambioEnMaterias = True 'para guardar los cambios
        End If
    End If
    
End Sub



Private Sub Command5_Click() 'añadir una materia
    If Trim(Text2 = "") Then
        frmMsgBox.cadenaAMostrar = "Antes de apretar este botón escriba la materia a añadir."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    If List2.ListCount = 28 Then
        frmMsgBox.cadenaAMostrar = "Ya se han agregado las 28 materias que acepta este programa. Para añadir una materia distinta, por favor borre una ya guardada"
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 0 To List2.ListCount - 1 'se controla que no haya una materia ya añadida con el mismo nombre
        If Trim(Text2) = List2.List(i) Then
            frmMsgBox.cadenaAMostrar = "Ya hay una materia con el nombre " + Chr(34) + Trim(Text2) + Chr(34) + ". No se puede añadir."
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
            Text2.SelStart = 0
            Text2.SelLength = Len(Text2)
            Text2.SetFocus
            Exit Sub
        End If
    Next
    
    List2.AddItem Trim(Text2)
    Text2 = ""
    swHuboCambioEnMaterias = True 'para guardar los cambios
End Sub

Private Sub Command6_Click() 'eliminar una materia
    If List2.ListCount = 0 Then
        frmMsgBox.cadenaAMostrar = "No se puede eliminar ninguna materia pues la lista está vacía."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    If List2.List(List2.ListIndex) = "" Then
        frmMsgBox.cadenaAMostrar = "Antes de apretar este botón tiene que seleccionar alguna materia de la lista."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    Dim aux As Integer
    
    frmMsgBox.swMostrarCancelar = False
    frmMsgBox.cadenaAMostrar = "¿Realmente querés eliminar la materia?"
    frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
    frmMsgBox.Show 1
    If frmMsgBox.swResultadoMostrado = True Then
        If List2.ListIndex < List2.ListCount - 1 Then
            aux = List2.ListIndex
        Else
            aux = List2.ListIndex - 1
        End If
        List2.RemoveItem List2.ListIndex
        List2.SetFocus
        If List2.ListCount <> 0 Then List2.ListIndex = aux
        Call guardarHistorial(List2)  'se guardan los cambios en las materias y en el historial
        swHuboCambioEnMaterias = True 'para guardar los cambios
    End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    If KeyCode = vbKeyEscape Then Unload Me
    
    If KeyCode = vbKeyF1 Then ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/configuración.htm ", "", 1 'leer la ayuda
End Sub

Private Sub Form_Load()
    Dim i As Integer, archivolibre As Byte, cadena As String ', archivoLibre As Byte
    
    On Error GoTo manejoError
    Set objPipe = New clsPipe
    Call centrarFormulario(Me)
    Select Case SSTab1.Tab
        Case 0
            Call establecerOrdenTabulación(pestaña.Materias)
        Case 1
            Call establecerOrdenTabulación(pestaña.Voces)
        Case 2
            Call establecerOrdenTabulación(pestaña.Otros)
        Case 3
            Call establecerOrdenTabulación(pestaña.Créditos)
    End Select
    
    sw1 = False 'que sapi 5 no hable
    sw2 = False 'que no hable sapi4
        
    With usuario
        swSapi5 = .sapi5 'si se usa sapi 5 o sapi 4
        swHablarVoz = .usarVoz
        swMostrarAñoEnTareas = .mostrarTodasLasTareas
        swMostrarAñoEnActividades = .mostrarTodasLasActividades
        nombreUsuario = .nombre
        swUsuarioMujer = .usuarioMujer
        swImprimirDirecto = .imprimirDirecto
        colorFuente = .fuenteColor
        NombreFuente = Trim(.fuenteNombre)
        tamañoFuente = .fuenteTamaño
        colorFondo = .colorFondo
        velocidadVoz = .velocidadVoz
        swLeerRenglones = .swLeerRenglones
        swUsarCorrectorOrtográfico = .swUsarCorrectorOrtográfico
        nombreSapi4 = Trim(.nombreVozSapi4)
        nombreSapi5 = Trim(.nombreVozSapi5)
        swPermitirAbrirArchivos = .swPermitirAbrirArchivos
        If swEmpezóLaMochila = False Then swMúsicaDeFondo = .swMúsicaDeFondo
    End With
    
    
    Call llenarComboVoz(Combo1, Combo9)
    If Combo1.List(Combo1.ListIndex) = "No hay voces avanzadas (SAPI5) instaladas" And Combo9.List(Combo9.ListIndex) = "No hay voces simples (SAPI4) instaladas" Then
        chkVoz = 1
        chkVoz.Enabled = False
    End If
    Slider1.Value = velocidadVoz
    Call regularVelocidadVoz
    
    'se verifica sintetizadores de qué sapi están instalados en la máquina actual
    If Combo1.Enabled = False Then swSapi5 = False
    If Combo9.Enabled = False Then swSapi5 = True
    If Combo9.Enabled = False And Combo1.Enabled = False Then swHablarVoz = False
    
    Dim swFuenteExistente As Boolean
    swFuenteExistente = False
    For i = 0 To Screen.FontCount - 1 'se chequea que la fuente del registro esté en el sistema
        If Screen.Fonts(i) = NombreFuente Then
            swFuenteExistente = True
            Exit For
        End If
    Next
    If swFuenteExistente = False Then NombreFuente = Screen.Fonts(i) 'si no está se asigna la primer fuente encontrada en el sistema
    
    Text3 = Trim(nombreUsuario) 'se carga el nombre en el cuadro de texto respectivo
    
    If swUsuarioMujer = True Then
        Option10(1).Value = True 'se carga el sexo del usuario
    Else
        Option10(0).Value = True
    End If
       
    If swHablarVoz = True Then 'a ver si está activada la voz
        chkVoz.Value = 0
    Else
        chkVoz.Value = 1
    End If
    
    If usuario.swVerActividadesConJaws = True Then 'a ver si se ven las actividades para jaws o con el calendario
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    
    If swMúsicaDeFondo = True Then 'a ver si está activada la voz
        Check6.Value = 1
    Else
        Check6.Value = 0
    End If
    
    If swPermitirAbrirArchivos = True Then 'a ver si está activada la voz
        Check7.Value = 1
    Else
        Check7.Value = 0
    End If
    
    If usuario.mostrarAñoEnEvaluaciones = True Then
        Check4.Value = 0
    Else
        Check4.Value = 1
    End If
    
    If swMostrarAñoEnActividades = True Then
        Check2.Value = 0
    Else
        Check2.Value = 1
    End If

    If swMostrarAñoEnTareas = True Then
        Check3.Value = 0
    Else
        Check3.Value = 1
    End If

    
    If swImprimirDirecto = True Then
        Check5.Value = 1
    Else
        Check5.Value = 0
    End If
    
    If swLeerRenglones = True Then
        Check8.Value = 1
    Else
        Check8.Value = 0
    End If
    
    If swUsarCorrectorOrtográfico Then
        Check9.Value = 1
    Else
        Check9.Value = 0
    End If

    If swSapi5 = True Then
        Option8(0).Value = True
    Else
        Option8(1).Value = True
    End If
       
    archivolibre = FreeFile  'se abren las materias
    Open App.path + "\datos\materias.txt" For Input As archivolibre
    While Not EOF(archivolibre)
        Line Input #archivolibre, cadena
        List2.AddItem Trim(cadena) 'se añaden las materias al listado y al combo
    Wend
    Close #archivolibre
    sw1 = True 'que sapi 5 hable
    sw2 = True 'que hable sapi4
    
    '++++++++++++++++++++++++++++++++++
    Dim índice As Integer, CadenaAbuscar As String
    ' Agrega los colores a cmbFontColor.
    With cmbFontColor
        .AddItem "Negro"
        .AddItem "Azul"
        .AddItem "Rojo"
        .AddItem "Verde"
        .AddItem "Blanco"
        .AddItem "Amarillo"
        Select Case colorFuente
            Case vbBlack
               CadenaAbuscar = "Negro"
            Case vbBlue
               CadenaAbuscar = "Azul"
            Case vbRed
               CadenaAbuscar = "Rojo"
            Case vbGreen
               CadenaAbuscar = "Verde"
            Case vbWhite
                CadenaAbuscar = "Blanco"
            Case vbYellow
                CadenaAbuscar = "Amarillo"
        End Select
        índice = 0
        For i = 0 To .ListCount - 1
            If CadenaAbuscar = .List(i) Then
                índice = i
                Exit For
            End If
        Next
        .ListIndex = índice
    End With
    
    
    ' Agrega los colores a cmbFormColor.
    With cmbFormColor
        .AddItem "Blanco"
        .AddItem "Negro"
        .AddItem "Verde"
        .AddItem "Azul"
        .AddItem "Rojo"
        .AddItem "Amarillo"
        Select Case colorFondo
            Case vbBlack
               CadenaAbuscar = "Negro"
            Case vbBlue
               CadenaAbuscar = "Azul"
            Case vbRed
               CadenaAbuscar = "Rojo"
            Case vbGreen
               CadenaAbuscar = "Verde"
            Case vbWhite
                CadenaAbuscar = "Blanco"
            Case vbYellow
                CadenaAbuscar = "Amarillo"
        End Select
        índice = 0
        For i = 0 To .ListCount - 1
            If CadenaAbuscar = .List(i) Then
                índice = i
                Exit For
            End If
        Next
        .ListIndex = índice
    End With
    
    
    índice = 0
    With cmbFontName
       For i = 0 To Screen.FontCount - 1
            .AddItem Screen.Fonts(i)
            If NombreFuente = .List(i) Then índice = i
       Next i
       ' Establece ListIndex a la fuente que está guardada en datos usuario.
       .ListIndex = índice
    End With
    
    índice = 0
    With cmbFontSize
       ' Llena el control con tamaños en incrementos de 2.
       For i = 8 To 72 Step 2
          .AddItem i
       Next i
       
       For i = 0 To .ListCount - 1
            If tamañoFuente = CInt(cmbFontSize.List(i)) Then
                índice = i
                Exit For
            End If
        Next
       ' Establece ListIndex a 0
       .ListIndex = índice ' size 10.
    End With
        
    NombreFuente = cmbFontName  'se ajusta la fuente del programa
    'se graba en una variable el color de la fuente del programa
    Select Case cmbFontColor.List(cmbFontColor.ListIndex)
        Case "Blanco"
           colorFuente = vbWhite
        Case "Azul"
           colorFuente = vbBlue
        Case "Negro"
           colorFuente = vbBlack
        Case "Verde"
           colorFuente = vbGreen
        Case "Rojo"
            colorFuente = vbRed
        Case "Amarillo"
            colorFuente = vbYellow
    End Select

    tamañoFuente = cmbFontSize.Text 'se ajusta el tamaño de la fuente
    
    'se graba en una variable el color de la fuente del programa
    Select Case cmbFormColor.List(cmbFormColor.ListIndex)
        Case "Blanco"
           colorFondo = vbWhite
        Case "Azul"
           colorFondo = vbBlue
        Case "Negro"
           colorFondo = vbBlack
        Case "Verde"
           colorFondo = vbGreen
        Case "Rojo"
            colorFondo = vbRed
        Case "Amarillo"
            colorFondo = vbYellow
    End Select
    
'    Dim lectorRegistro As Object, swExisteDiccionario
    
    
    If swAspellInstalado = True Then 'si aspell está instalado, se lo prefiere
        '*************************
        'ver diccionarios de aspell
        Call objPipe.Execute("CMD.EXE")
        Call Sleep(200)
        Call objPipe.Write_("c:" & vbCrLf)
        Call objPipe.Write_("cd " & rutaDeAspell & vbCrLf)
        Call Sleep(100)
        Call objPipe.Write_("aspell dump dicts" & vbCrLf)
        Call Sleep(200)
        Call cargarIdiomasAspell(objPipe.Read)
        KillProcess ("cmd.exe")
        KillProcess ("aspell.exe")
    Else
        '*****************************
        'ver si están los diccionarios propios de idiomas
        
        
    End If
    
    'se selecciona en el combo el idioma que se ha elegido
    Dim idioma As String
    Select Case idiomaAspell
        Case "es"
            idioma = "Español"
        Case "en"
            idioma = "Inglés"
        Case "fr"
            idioma = "Francés"
        Case "pt"
            idioma = "Portugués"
    End Select
    
    For i = 0 To cmbIdiomasCorrector.ListCount - 1
        If cmbIdiomasCorrector.List(i) = idioma Then
            cmbIdiomasCorrector.ListIndex = i
            Exit For
        End If
    Next
    
    'si el cuaderno está abierto, no se habilita el combo
    If swCuadernoAbierto = True Or frmLectorEvaluaciones.swEstoyAbierto = True Then
        cmbIdiomasCorrector.Enabled = False
        lblIdioma.Visible = True
    Else
        cmbIdiomasCorrector.Enabled = True
        lblIdioma.Visible = False
    End If
    
    Exit Sub
manejoError:
    If Err.Number = 52 Then
        Open App.path + "\datos\materias.txt" For Output As #archivolibre 'se abre el trabajo ya guardado
        Close #archivolibre
        Resume
        Exit Sub
    End If
    frmMsgBox.cadenaAMostrar = "soy el controlador del form control. Error número: " + Str(Err.Number) + ", descripción: " + Err.Description
    frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
    frmMsgBox.Show 1
    Exit Sub
End Sub


Private Sub Image1_Click(Index As Integer)
    Select Case Index
        Case 0
            ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/config voz.htm", "", 1
        Case 1
            ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/config otros.htm", "", 1
        Case 2
            ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/configuración materias.htm", "", 1
    End Select
End Sub

Private Sub Option8_Click(Index As Integer) 'usar voz sapi 4 ó 5
    If Index = 0 Then
        Combo1.Enabled = True
        Combo9.Enabled = False
        swSapi5 = True
    Else
        Combo1.Enabled = False
        Combo9.Enabled = True
        swSapi5 = False
    End If
End Sub


Private Sub Option4_Click(Index As Integer) 'usar las banderas spVoice
    If Index = 0 Then
        banderasSPVoice = SVSFNLPSpeakPunc
    Else
        banderasSPVoice = SVSFDefault
    End If
End Sub

Private Sub Combo1_Click()
    On Error GoTo manejoError
    Set Voz.Voice = Voz.GetVoices().Item(Combo1.ListIndex)
    If sw1 = True Then Voz.Speak "Elegiste mi voz para hablarte"
    Call regularVelocidadVoz
    sw1 = True
    Exit Sub
manejoError:
    If Err.Number = 91 Then
        Option8(0).Enabled = False
        Combo1.Enabled = False
        Exit Sub
    Else
        frmMsgBox.cadenaAMostrar = "Soy el controlador de combo1_click. Error número: " + Err.Number + ". Descripción: " + Err.Description
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
    End If
End Sub

Private Sub Combo9_Click()
    On Error GoTo manejoError
    vozSapi4.CurrentMode = Combo9.ListIndex + 1
    If sw2 = True Then vozSapi4.Speak "Elegiste mi voz para hablarte"
    Call regularVelocidadVoz
    sw2 = True
    Exit Sub
manejoError:
    If Err.Number = 91 Then
        Option8(1).Enabled = False
        Combo9.Enabled = False
        Exit Sub
    Else
        frmMsgBox.cadenaAMostrar = "Soy el controlador de combo9_click. Error número: " + Err.Number + ". Descripción: " + Err.Description
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
    End If
End Sub


Private Sub Slider1_Scroll() 'regular la velocidad de la voz
    Dim aux As Integer
    
    aux = velocidadVoz
    velocidadVoz = Slider1.Value
    
    Call regularVelocidadVoz
    
    If aux > velocidadVoz Then Decir "más lento"
    If aux < velocidadVoz Then Decir "más rápido"
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.TabCaption(SSTab1.Tab)
        Case "Materias"
            establecerOrdenTabulación (pestaña.Materias)
        Case "Voz"
            establecerOrdenTabulación (pestaña.Voces)
        Case "Otros"
            establecerOrdenTabulación (pestaña.Otros)
        Case "Créditos"
            establecerOrdenTabulación (pestaña.Créditos)
    End Select
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call controlarCaracteresEspeciales(KeyCode, Text2)
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Command5_Click
    End If
End Sub


Private Function retString(ByVal seleccion As String) As String
    Dim position As Long
    Dim Pos As Long
    
    position = 1
    Do While position < Len(seleccion)
      Pos = InStr(position, seleccion, vbCr)
      If Pos = 0 Then
         Exit Do
      End If
      seleccion = Left$(seleccion, Pos - 1) & vbCrLf & Mid$(seleccion, Pos + 1)
      position = Pos + 2
    Loop
    
    retString = seleccion
End Function

Sub guardarMateriasGeneral()
    Dim i As Integer, j As Integer
    On Error Resume Next
    If swHuboCambioEnMaterias = True Then 'si hubo cambios en las materias se los guarda
        Call guardarMaterias(List2)
        Call guardarHistorial(List2)
        For i = 0 To List2.ListCount - 1 'se actualiza el listado de materias para las actividades
            MkDir (App.path + "\trabajos\" + List2.List(i))
            For j = 1 To 12
                MkDir (App.path + "\trabajos\" + List2.List(i) + "\" + Trim(Str(j)))
                MkDir (App.path + "\trabajos\" + List2.List(i) + "\" + Trim(Str(j)) + "\datosHojas")
            Next
            MkDir (App.path + "\trabajos\" + List2.List(i) + "\actividades")
            MkDir (App.path + "\trabajos\" + List2.List(i) + "\soporte")
            MkDir (App.path + "\trabajos\" + List2.List(i) + "\evaluaciones") 'carpeta para poner evaluaciones falsas por si los papás quieren modificar una evaluación ya hecha ];)
            For j = 1 To 12
                MkDir (App.path + "\trabajos\" + List2.List(i) + "\actividades\" + Trim(Str(j)))
                MkDir (App.path + "\trabajos\" + List2.List(i) + "\actividades\" + Trim(Str(j)) + "\datosActividades")
                MkDir (App.path + "\trabajos\" + List2.List(i) + "\soporte\" + Trim(Str(j)))
                MkDir (App.path + "\trabajos\" + List2.List(i) + "\soporte\" + Trim(Str(j)) + "\datosSoporte")
                MkDir (App.path + "\trabajos\" + List2.List(i) + "\evaluaciones\" + Trim(Str(j)))
            Next
            MkDir (App.path + "\trabajos\" + List2.List(i) + "\libros")
        Next
        If frmPrincipal.swEstoyAbierto = True Then  'si el form principal está abierto se lo descarga y vuelve a cargar para que actualice las materias
            Unload frmPrincipal
            frmPrincipal.Show
            frmControl.SetFocus
        End If
        swHuboCambioEnMaterias = False
    End If
End Sub

Sub establecerOrdenTabulación(quéPestaña As Byte)
    'se establece que no se tabule a ningún control
    'Materias:
    Text2.TabStop = False
    Command5.TabStop = False
    Command1.TabStop = False
    List2.TabStop = False
    Command4(0).TabStop = False
    Command6.TabStop = False
    Command4(1).TabStop = False
    
    'Voz:
    chkVoz.TabStop = False
    Option8(0).TabStop = False
    Combo1.TabStop = False
    Option8(1).TabStop = False
    Combo9.TabStop = False
    Check8.TabStop = False
    Slider1.TabStop = False
    
    'Otros:
    Check4.TabStop = False
    Check2.TabStop = False
    Check3.TabStop = False
    Check5.TabStop = False
    Check9.TabStop = False
    Check6.TabStop = False
    Check7.TabStop = False
    Check1.TabStop = False
    Text3.TabStop = False
    Option10(0).TabStop = False
    Option10(1).TabStop = False
    cmbFontName.TabStop = False
    cmbFontColor.TabStop = False
    cmbFontSize.TabStop = False
    cmbFormColor.TabStop = False
    
    'Créditos:
    Command12.TabStop = False
    
    'se activa la tabulación de los controles según la pestaña activa
    Select Case quéPestaña
        Case pestaña.Materias
            Text2.TabStop = True
            Command5.TabStop = True
            Command1.TabStop = True
            List2.TabStop = True
            Command4(0).TabStop = True
            Command6.TabStop = True
            Command4(1).TabStop = True
            Command12.TabStop = True
        Case pestaña.Voces
            chkVoz.TabStop = True
            Option8(0).TabStop = True
            Option8(1).TabStop = True
            Combo1.TabStop = True
            Combo9.TabStop = True
            Check8.TabStop = True
            Slider1.TabStop = True
            Command12.TabStop = True
        Case pestaña.Otros
            Check4.TabStop = True
            Check2.TabStop = True
            Check3.TabStop = True
            Check5.TabStop = True
            Check9.TabStop = True
            Check6.TabStop = True
            Check7.TabStop = True
            Check1.TabStop = True
            Text3.TabStop = True
            Option10(0).TabStop = True
            Option10(1).TabStop = True
            cmbFontName.TabStop = True
            cmbFontColor.TabStop = True
            cmbFontSize.TabStop = True
            cmbFormColor.TabStop = True
            Command12.TabStop = True
        Case pestaña.Créditos
            Command12.TabStop = True
    End Select
End Sub

Private Sub cargarIdiomasAspell(texto As String)
    Dim posición As Integer
    'se deja la devolución de aspell donde empiezan los diccionarios instalados
    texto = Right(texto, Len(texto) - InStr(1, texto, "C:\Archivos de programa\Aspell\bin>aspell dump dicts") - Len("C:\Archivos de programa\Aspell\bin>aspell dump dicts") - 1)
    
    'se buscan cuáles están instalados y se los incluye en el combo
    posición = InStr(1, texto, "es" & vbCrLf, vbTextCompare)
    If posición Then cmbIdiomasCorrector.AddItem "Español"
  
    posición = InStr(1, texto, "en" & vbCrLf, vbTextCompare)
    If posición Then cmbIdiomasCorrector.AddItem "Inglés"
    
    posición = InStr(1, texto, "fr" & vbCrLf, vbTextCompare)
    If posición Then cmbIdiomasCorrector.AddItem "Francés"
   
    posición = InStr(1, texto, "pt" & vbCrLf, vbTextCompare)
    If posición Then cmbIdiomasCorrector.AddItem "Portugués"
End Sub
