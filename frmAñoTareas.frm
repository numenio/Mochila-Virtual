VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmAñoTareas 
   Caption         =   "Hojas guardadas"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   Icon            =   "frmAñoTareas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAñoTareas.frx":08CA
   ScaleHeight     =   5520
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1410
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
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
      ItemData        =   "frmAñoTareas.frx":2922
      Left            =   150
      List            =   "frmAñoTareas.frx":2924
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   510
      TabIndex        =   2
      Top             =   4680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      Caption         =   "         Mostrar las hojas del año seleccionado"
      EstiloDelBoton  =   1
      Picture         =   "frmAñoTareas.frx":2926
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
Attribute VB_Name = "frmAñoTareas"
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
'        If swCuadernoAbierto = True Then
'            frmCuaderno.Show
'        Else
'            frmPrincipal.Show
'        End If
        Unload Me
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en las hojas ya escritas de tu carpeta"
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.añoTareas
         frmAyuda.Show
         Exit Sub
    End If
End Sub


Private Sub Form_Load()
    Dim año As Integer, cadena As String, mes As Byte, día As Byte, fechaTotal As String
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swPulsóEnterParaAvanzar = False
    Me.Caption = "Años que tienen hojas guardadas"
    
    Decir "Entrando en los años que tienen hojas escritas de " + miMateria + ". Elegí con las flechas cuál año querés abrir y aceptá con enter"

    Dim i As Integer, j As Integer, swAñoDuplicado As Boolean
    For i = 1 To 12
        File1.path = App.path + dirTrabajo + Trim(Str(i))
        If File1.ListCount <> 0 Then 'si el mes tiene tareas
            For j = 0 To File1.ListCount - 1
'                If Mid(File1.List(j), cantPrefijo + 4, 4) < year(date) Then 'si el año es menor al actual
                    swAñoDuplicado = controlarAñoDuplicado(Left(Right(File1.List(j), 8), 4))
                    If swAñoDuplicado = False Then List1.AddItem "Año " + Left(Right(File1.List(j), 8), 4)
'                End If
'
'                If Mid(File1.List(j), cantPrefijo + 4, 4) = year(date) Then 'si el año es igual al actual
'                    If i < month(date) Then 'si el mes es menor al actual
'                        swAñoDuplicado = controlarAñoDuplicado(Mid(File1.List(j), cantPrefijo + 4, 4))
'                        If swAñoDuplicado = False Then List1.AddItem "Año " + Mid(File1.List(j), cantPrefijo + 4, 4)
'                    ElseIf i = month(date) Then 'si el mes es igual al actual
'                        If Mid(File1.List(j), 4, 2) < day(date) Then 'si el día es menor al actual
'                            swAñoDuplicado = controlarAñoDuplicado(Mid(File1.List(j), cantPrefijo + 4, 4))
'                            If swAñoDuplicado = False Then List1.AddItem "Año " + Mid(File1.List(j), cantPrefijo + 4, 4)
'                        End If
'                    End If
'                End If
            Next
        End If
    Next
    If List1.ListCount = 0 Then List1.AddItem "No hay ninguna hoja guardada"
    List1.ListIndex = 0
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
    
    If swPulsóEnterParaAvanzar = False Then
        If swCuadernoAbierto = True Then
            frmCuaderno.Show
        Else
            frmPrincipal.Show
        End If
    End If
    'Call contarFormularios(False)
End Sub

Private Sub List1_DblClick()
    If List1.List(List1.ListIndex) <> ("No hay ninguna hoja guardada") And List1.ListIndex <> -1 Then
        frmTareasAnt.añoParaVerMeses = Trim(Right(List1.List(List1.ListIndex), 4))
        frmTareasAnt.Show
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

Private Function controlarAñoDuplicado(año As String) As Boolean
    Dim j As Integer, cadenaAux As String
    controlarAñoDuplicado = False
    For j = 0 To List1.ListCount - 1 'se controla que no esté ya el mes incluído
        cadenaAux = Right(List1.List(j), 4) 'se toma el año del listado
        If año = cadenaAux Then
            controlarAñoDuplicado = True 'si el año no está ya en la lista
            Exit For
        End If
    Next
End Function

