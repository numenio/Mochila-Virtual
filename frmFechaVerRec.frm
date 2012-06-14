VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmFechaVerRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recordatorios"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmFechaVerRec.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFechaVerRec.frx":08CA
   ScaleHeight     =   3780
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Caption         =   "Ver recordatorios"
      EstiloDelBoton  =   1
      Picture         =   "frmFechaVerRec.frx":14F2
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes para ver los recordatorios:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A�o para ver los recordatorios:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   2160
   End
End
Attribute VB_Name = "frmFechaVerRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim swPuls�EnterParaAvanzar As Boolean

Private Sub Combo1_GotFocus()
    Decir "eleg� con las flechas el a�o del cual quer�s ver los recordatorios y acept� con enter. Est�s en " + Combo1.Text
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then Decir Combo1.Text
End Sub

Private Sub Combo2_GotFocus()
    Decir "ahora us� las flechas para elegir el mes del que quer�s ver los recordatorios guardados y acept� con enter. Est�s en " + Combo2.Text
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then Decir Combo2.Text
End Sub

Private Sub Command1_Click()
    frmCalendarioM�ltiple.tipoElemento = elemento.Recordatorio
    frmCalendarioM�ltiple.MesParaAbrir = Combo2.List(Combo2.ListIndex)
    frmCalendarioM�ltiple.numMesParaAbrir = Combo2.ListIndex + 1
    frmCalendarioM�ltiple.a�o = Int(Combo1.Text)
    frmCalendarioM�ltiple.Show
    swPuls�EnterParaAvanzar = True
    Unload Me
End Sub

Private Sub Command1_GotFocus()
    Decir "ver los recordatorios del mes " + Combo2.Text + " del a�o " + Combo1.Text + ". apret� enter para aceptar o tab para cambiar el a�o o el mes"
End Sub

Private Sub Command1_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift And 7 = vbAltMask And KeyCode = 18 Then 'se neutraliza el men� de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
'        If swCuadernoAbierto = True Then Decir "volviendo a tu carpeta"
'        If swCuadernoAbierto = False Then frmPrincipal.Show
        Unload Me
        Exit Sub
    End If
    If KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de m�sica, ten�s que estar en el men� principal o en una carpeta. ahora est�s en los recordatorios"
    If KeyCode = vbKeyReturn And TypeOf Me.ActiveControl Is ComboBox Then SendKeys "{tab}"
    If Shift And 7 = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.fechaVerRec
         frmAyuda.Show
         Exit Sub
    End If
    
    If Shift And 7 = vbCtrlMask Then Decir ""
End Sub

Private Sub Form_Load()
    Dim mes As Byte
    Dim i As Long, a�oActual As Long ', d�a As Byte
    
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swPuls�EnterParaAvanzar = False
    '++++++++++++++++++++++++++++++++++++++
    'Cargar la fecha
    Combo2.AddItem "Enero"
    Combo2.AddItem "Febrero"
    Combo2.AddItem "Marzo"
    Combo2.AddItem "Abril"
    Combo2.AddItem "Mayo"
    Combo2.AddItem "Junio"
    Combo2.AddItem "Julio"
    Combo2.AddItem "Agosto"
    Combo2.AddItem "Setiembre"
    Combo2.AddItem "Octubre"
    Combo2.AddItem "Noviembre"
    Combo2.AddItem "Diciembre"
    
    mes = Mid(Format(Date, "dd/mm/yyyy"), 4, 2) 'seleccionar el mes actual
    Combo2.ListIndex = mes - 1
    
    a�oActual = Year(Date) 'cargar 20 a�os desde el actual y seleccionar el a�o actual
    For i = a�oActual To a�oActual + 20
        Combo1.AddItem i
    Next
    
    Combo1.ListIndex = 0
    
    Decir "Eleg� con las flechas el a�o del que quer�s ver los recordatorios y despu�s apret� enter"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If swPuls�EnterParaAvanzar = False Then
        If swCuadernoAbierto = True Then Decir "volviendo a tu carpeta"
        If swCuadernoAbierto = False Then frmPrincipal.Show
    End If
    'Call contarFormularios(False)
End Sub
