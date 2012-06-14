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
      Caption         =   "Año para ver los recordatorios:"
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
Dim swPulsóEnterParaAvanzar As Boolean

Private Sub Combo1_GotFocus()
    Decir "elegí con las flechas el año del cual querés ver los recordatorios y aceptá con enter. Estás en " + Combo1.Text
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then Decir Combo1.Text
End Sub

Private Sub Combo2_GotFocus()
    Decir "ahora usá las flechas para elegir el mes del que querés ver los recordatorios guardados y aceptá con enter. Estás en " + Combo2.Text
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then Decir Combo2.Text
End Sub

Private Sub Command1_Click()
    frmCalendarioMúltiple.tipoElemento = elemento.Recordatorio
    frmCalendarioMúltiple.MesParaAbrir = Combo2.List(Combo2.ListIndex)
    frmCalendarioMúltiple.numMesParaAbrir = Combo2.ListIndex + 1
    frmCalendarioMúltiple.año = Int(Combo1.Text)
    frmCalendarioMúltiple.Show
    swPulsóEnterParaAvanzar = True
    Unload Me
End Sub

Private Sub Command1_GotFocus()
    Decir "ver los recordatorios del mes " + Combo2.Text + " del año " + Combo1.Text + ". apretá enter para aceptar o tab para cambiar el año o el mes"
End Sub

Private Sub Command1_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift And 7 = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
'        If swCuadernoAbierto = True Then Decir "volviendo a tu carpeta"
'        If swCuadernoAbierto = False Then frmPrincipal.Show
        Unload Me
        Exit Sub
    End If
    If KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en los recordatorios"
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
    Dim i As Long, añoActual As Long ', día As Byte
    
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swPulsóEnterParaAvanzar = False
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
    
    añoActual = Year(Date) 'cargar 20 años desde el actual y seleccionar el año actual
    For i = añoActual To añoActual + 20
        Combo1.AddItem i
    Next
    
    Combo1.ListIndex = 0
    
    Decir "Elegí con las flechas el año del que querés ver los recordatorios y después apretá enter"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If swPulsóEnterParaAvanzar = False Then
        If swCuadernoAbierto = True Then Decir "volviendo a tu carpeta"
        If swCuadernoAbierto = False Then frmPrincipal.Show
    End If
    'Call contarFormularios(False)
End Sub
