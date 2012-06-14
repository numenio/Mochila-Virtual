VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmLibroX 
   Caption         =   "Libro X"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   Icon            =   "frmLibroX.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmLibroX.frx":08CA
   ScaleHeight     =   6360
   ScaleWidth      =   7740
   Begin VB.ListBox List1 
      Height          =   4155
      ItemData        =   "frmLibroX.frx":2922
      Left            =   263
      List            =   "frmLibroX.frx":2924
      TabIndex        =   0
      Top             =   600
      Width           =   7215
   End
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   2303
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      Caption         =   "    Abrir el capítulo seleccionado"
      EstiloDelBoton  =   1
      Picture         =   "frmLibroX.frx":2926
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
Attribute VB_Name = "frmLibroX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public libroParaVer As String

Private Sub Command1_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    If swCuadernoAbierto = True Then Decir "" 'callar la voz si se vuelve al cuaderno
'End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
'        If swCuadernoAbierto = True Then
            Decir "volviendo a tu carpeta"
            Unload Me
'        Else
'            frmMateria.Show
'            Unload Me
'        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
       
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.libroX
         frmAyuda.Show
         Exit Sub
    End If
End Sub


Private Sub Form_Load()
    Dim cadena As String, archivolibre As Byte
    
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    frmLibroX.Caption = "Capítulos del libro " + libroParaVer
    
    Decir "Entrando en los capítulos del libro " + libroParaVer + ". Elegí con las flechas cuál capítulo buscás y abrilo con enter"

    Dim i As Integer, contador As Integer, cadenaAux As String  ', swMesDuplicado As Boolean
'    File1.Path = App.Path + dirTrabajo + "libros\" + libroParaVer + "\"
    
    contador = 0
'    For i = 0 To (File1.ListCount - 1) 'se añaden a la lista todas las carpetas (libros)
'        If Right(File1.List(i), 4) = ".rtf" Then 'si es un archivo rtf, o sea que es un capítulo
'            cadena = Right(File1.List(i), Len(File1.List(i)) - InStrRev(File1.List(i), "\")) 'se añaden los libros de la materia
'            cadena = Left(cadena, Len(cadena) - 4) 'se le saca el .rft
'            contador = contador + 1
'            cadenaAux = "Capítulo " + Trim(Str(contador)) + ": " + cadena
'            List1.AddItem cadenaAux
'        End If
'    Next i
                        
    'se carga en la lista los capítulos en orden
    cadenaAux = App.path + dirTrabajo + "libros\" + libroParaVer + "\ordenCapítulos"
    If existeCarpeta(cadenaAux) Then
        archivolibre = FreeFile
        Open cadenaAux For Input As #archivolibre 'se abre el trabajo ya guardado
        Do While Not EOF(archivolibre)
            Input #archivolibre, cadena
            If Trim(cadena) <> "" Then
                contador = contador + 1
                List1.AddItem "Capítulo " + Trim(Str(contador)) + ": " + cadena
            End If
        Loop
        Close #archivolibre
    End If
                
    If List1.ListCount = 0 Then List1.AddItem "No hay ningún capítulo guardado de " + libroParaVer
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
    'Call contarFormularios(False)
End Sub

Private Sub List1_DblClick()
    If List1.List(List1.ListIndex) <> ("No hay ningún libro guardado de " + miMateria) Then
        frmLectorLibro.archivoParaLeer = libroParaVer + "\" + Trim(Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - InStrRev(List1.List(List1.ListIndex), ":"))) + ".rtf"
        If swLibroAbierto = True Then frmLectorLibro.cargarLibro
        frmLectorLibro.Show
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

'Private Sub Command1_GotFocus()
'    Decir Command1.Caption
'    sonido = sndPlaySound(App.Path + "\sonidos\cb.wav", SND_ASYNC)
'End Sub

Private Sub Command1_Click()
    List1_DblClick
End Sub
