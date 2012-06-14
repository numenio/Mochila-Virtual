VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmMesTareasX 
   Caption         =   "Tareas: mes X"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   Icon            =   "frmMesActTareasX.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMesActTareasX.frx":08CA
   ScaleHeight     =   7020
   ScaleWidth      =   8685
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   6120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      Caption         =   "    Abrir la hoja seleccionada"
      EstiloDelBoton  =   0
      Picture         =   "frmMesActTareasX.frx":326A
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
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      ItemData        =   "frmMesActTareasX.frx":3B44
      Left            =   240
      List            =   "frmMesActTareasX.frx":3B46
      TabIndex        =   1
      Top             =   360
      Width           =   8175
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   6720
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmMesTareasX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MesParaAbrir As String
Public numMesParaAbrir As Byte
Dim cadena As String
Dim trabajos() As String


Private Sub Command1_Click()
    List1_DblClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    
    If KeyCode = vbKeyEscape Then
        frmTareasAnt.Show
        Unload Me
    End If
    
    If KeyCode = vbKeyF12 Then frmControl.Show
    
    shiftkey = Shift And 7
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = mesTareasX
         frmAyuda.Show
         Exit Sub
    End If
End Sub



Private Sub Form_Load()
    Call centrarFormulario(Me)
    Me.Caption = "Hojas anteriores del mes: " + MesParaAbrir
    Decir "abriendo las hojas de tu carpeta de " + miMateria + " del mes de " + MesParaAbrir
    Dim i As Integer, j As Integer, cadenaAux As String, contador As Integer
    
    File1.Path = App.Path + dirTrabajo + Trim(Str(numMesParaAbrir))
    
    For i = 0 To File1.ListCount - 1 'se ven todos los archivos del mes seleccionado
        If swMostrarAñoEnTareas = True Then 'si se muestran tareas de todos los años
            cadenaAux = decodificarArchivo(File1.List(i))
            List1.AddItem "Hoja del día " + Left(cadenaAux, 2) + " de " + MesParaAbrir + " de " + Right(cadenaAux, 4) 'se añaden todas las tareas del mes seleccionado
            ReDim Preserve trabajos(0 To contador)
            trabajos(contador) = File1.List(i)
            contador = contador + 1
        Else 'si se muestra sólo el año actual
            If Mid(File1.List(i), cantPrefijo + 4, 4) = Right(Date, 4) Then 'si el año es igual al actual
                cadenaAux = decodificarArchivo(File1.List(i))
                List1.AddItem "Hoja del día " + Left(cadenaAux, 2) + " de " + MesParaAbrir + " de " + Right(cadenaAux, 4) 'se añaden todas las tareas del mes seleccionado pero sólo del año actual
                ReDim Preserve trabajos(0 To contador)
                trabajos(contador) = File1.List(i)
                contador = contador + 1
            End If
        End If
    Next
    
    
    
    
    
    
    
    
    
'    For i = 0 To (File1.ListCount - 1) 'se examinan todos los archivos para clasificarlos en los anteriores o de hoy
'        cadena = Right(File1.List(i), Len(File1.List(i)) - cantPrefijo)
'        If swMostrarAñoEnTareas = True Then 'si se muestran todos los años
'            If Mid(cadena, 7, 4) <= Right(Date, 4) Then 'si el año es igual o menor al actual
'                If Mid(cadena, 7, 4) < Right(Date, 4) Then 'si el año es menor al actual
'                    If Mid(cadena, 4, 2) = NumMesParaAbrir Then 'si el mes es igual al seleccionado
'                        cadenaAux = Format(Left(cadena, 10))
'                        cadenaAux = Format(cadenaAux, "Long Date")
'                        cadenaAux = "Hoja del día " + cadenaAux ' + Str(NumMesParaAbrir) + " de " + MesParaAbrir 'transformarCadena(cadena)
'                        List1.AddItem cadenaAux
'                        ReDim Preserve trabajos(0 To contador)
'                        trabajos(contador) = File1.List(i)
'                        contador = contador + 1
'                    End If
'                Else 'si el año es igual al actual
'                    If Mid(cadena, 3, 2) <= Mid(Date, 3, 2) Then  'si el mes es menor o igual al actual
'                        If Mid(cadena, 4, 2) = NumMesParaAbrir Then 'si el mes es igual al seleccionado
'                            cadenaAux = Format(Left(cadena, 10))
'                            cadenaAux = Format(cadenaAux, "Long Date")
'                            cadenaAux = "Hoja del día " + cadenaAux ' + Str(NumMesParaAbrir) + " de " + MesParaAbrir 'transformarCadena(cadena)
'                            List1.AddItem cadenaAux
'                            ReDim Preserve trabajos(0 To contador)
'                            trabajos(contador) = File1.List(i)
'                            contador = contador + 1
'                        End If
'                    End If
'                End If
'            End If
'        Else 'si no se muestran todos los años
'            If Mid(cadena, 7, 4) = Right(Date, 4) Then 'si el año es igual al actual
'                If Mid(cadena, 4, 2) = NumMesParaAbrir Then 'si el mes es igual al seleccionado
'                    cadenaAux = Format(Left(cadena, 10))
'                    cadenaAux = Format(cadenaAux, "Long Date")
'                    cadenaAux = Left(cadenaAux, (Len(cadenaAux) - 8))
'                    cadenaAux = "Hoja del día " + cadenaAux ' + Left(cadena, 2) + " de " + MesParaAbrir 'transformarCadena(cadena)
'                    List1.AddItem cadenaAux
'                    ReDim Preserve trabajos(0 To contador)
'                    trabajos(contador) = File1.List(i)
'                    contador = contador + 1
'                End If
'            End If
'        End If
'    Next i
                        
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
    End If
End Sub

Private Sub List1_DblClick()
    'abrirMes = List1.List(List1.ListIndex)
'    frmLectorActividad.archivoParaLeer = trabajos(List1.ListIndex) 'cadena 'List1.List(List1.ListIndex)
'    frmLectorActividad.Show
    If List1.List(List1.ListIndex) <> "" Then
        frmCuaderno.nombreArchivo = trabajos(List1.ListIndex)
        frmCuaderno.nombreMesArchivo = numMesParaAbrir
        frmCuaderno.swContinuarArchivo = True
        frmCuaderno.swAbriendoHojaAnterior = True
        frmCuaderno.díaAbierto = Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 9)
        frmCuaderno.RichTextBox1.LoadFile App.Path + dirTrabajo + Trim(Str(numMesParaAbrir)) + "\" + frmCuaderno.nombreArchivo
        Decir "abriendo la hoja de " + miMateria + " del " + frmCuaderno.díaAbierto + ". podés seguir trabajando en ella"
        frmCuaderno.Show
        'frmCuaderno.l
        Unload Me
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then List1_DblClick
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    Decir List1.List(List1.ListIndex)
    sonido = sndPlaySound(App.Path + "\sonidos\td.wav", SND_ASYNC)
End Sub

Private Sub List1_GotFocus()
    Decir List1.List(List1.ListIndex)
End Sub

Private Sub Command1_GotFocus()
    Decir Command1.Caption
    sonido = sndPlaySound(App.Path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

