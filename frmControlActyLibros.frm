VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmControlActyLibros 
   Caption         =   "Actividades y Libros"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
   Icon            =   "frmControlActyLibros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmControlActyLibros.frx":08CA
   ScaleHeight     =   7770
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "Salir"
      Height          =   495
      Left            =   1650
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Actividades"
      TabPicture(0)   =   "frmControlActyLibros.frx":2922
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command7"
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(2)=   "Command13"
      Tab(0).Control(3)=   "Image1(2)"
      Tab(0).Control(4)=   "Label6(2)"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Libros"
      TabPicture(1)   =   "frmControlActyLibros.frx":293E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Image1(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command9"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton Command9 
         Caption         =   "Añadir un libro"
         Height          =   615
         Left            =   1170
         TabIndex        =   4
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Añadir un capítulo a un libro ya guardado"
         Height          =   615
         Left            =   1170
         TabIndex        =   5
         Top             =   2640
         Width           =   3495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Ver los libros y capítulos ya guardados"
         Height          =   615
         Left            =   1170
         TabIndex        =   6
         Top             =   3720
         Width           =   3495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Ver las actividades guardadas"
         Height          =   615
         Left            =   -73830
         TabIndex        =   2
         Top             =   3360
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ver las actividades guardadas usando un lector de pantallas"
         Height          =   615
         Left            =   -73830
         TabIndex        =   7
         Top             =   3360
         Width           =   3495
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Añadir una actividad"
         Height          =   615
         Left            =   -73830
         TabIndex        =   1
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   -70200
         Picture         =   "frmControlActyLibros.frx":295A
         Top             =   6000
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Si necesita ayuda, haga click aquí:"
         Height          =   375
         Index           =   2
         Left            =   -72480
         TabIndex        =   9
         Top             =   6015
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   4800
         Picture         =   "frmControlActyLibros.frx":3224
         Top             =   6000
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Si necesita ayuda, haga click aquí:"
         Height          =   375
         Index           =   3
         Left            =   2520
         TabIndex        =   8
         Top             =   6015
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmControlActyLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command12_Click()
    Unload Me
End Sub

Private Sub Command13_Click()
    frmAñadirActividad.swEditarActividades = False
    frmAñadirActividad.Show 1 'añadir una actividad
End Sub

Private Sub Command2_Click()
    frmActivDefVisual.Show 1 'ver actividades para def visual
End Sub

Private Sub Command7_Click() 'ver act guardadas
    frmCalendario.Show 1
End Sub


Private Sub Command9_Click() 'el botón de buscar libros
    frmAñadirLibro.Show 1
End Sub


Private Sub Command10_Click()
    frmAñadirCapítuloLibro.Show 1
End Sub

Private Sub Command11_Click() 'botón añadir libro
    frmVerLibros.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift And 7 = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    If usuario.swVerActividadesConJaws = True Then
        Command7.Visible = False
        Command2.Visible = True
    Else
        Command7.Visible = True
        Command2.Visible = False
    End If
End Sub

Private Sub Image1_Click(Index As Integer)
    If Index = 3 Then
        ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/añadir libros.htm", "", 1
    Else
        ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/Añadir actividad.htm", "", 1
    End If
End Sub
