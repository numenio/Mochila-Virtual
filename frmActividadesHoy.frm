VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmActividadesHoy 
   Caption         =   "Actividades"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   Icon            =   "frmActividadesHoy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmActividadesHoy.frx":08CA
   ScaleHeight     =   5955
   ScaleWidth      =   7905
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      Caption         =   "    Abrir la actividad seleccionada"
      EstiloDelBoton  =   1
      Picture         =   "frmActividadesHoy.frx":2922
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
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   6480
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3960
      ItemData        =   "frmActividadesHoy.frx":31FC
      Left            =   345
      List            =   "frmActividadesHoy.frx":31FE
      TabIndex        =   0
      Top             =   600
      Width           =   7215
   End
End
Attribute VB_Name = "frmActividadesHoy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cadena As String
Dim trabajos() As String
Dim miRegistro As DatosActividad
Dim swPulsóEnterParaAvanzar As Boolean
'Private Sub btnActividad_Click(Index As Integer)
'    frmLectorActividad.Show
'    Unload Me
'End Sub

Private Sub Command1_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer, mes As String ', miActividad As DatosActividad
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
'        frmActividades.Show
        Unload Me
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en las actividades"
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    
    If shiftkey = 0 And KeyCode = vbKeyF1 Then 'f1 dice los comentarios
'        mes = Trim(Str(CInt(month(date))))
'        Open App.Path + "\datos\" + mes + "datosActividades.gui" For Random As #1 Len = Len(miRegistro)
'        Do While Not EOF(1)   ' Repite hasta el final del archivo.
'           Get #1, , miRegistro   ' Lee el registro siguiente.
'           If Trim(miRegistro.DirArchivo) = App.Path + dirTrabajo + "actividades\" + trabajos(List1.ListIndex) Then Exit Do
'        Loop
'        Close #1   ' Cierra el archivo.
        
        If List1.List(0) <> "No hay ninguna actividad guardada de " + miMateria + " para el día de hoy" Then
            Open App.path + dirTrabajo + "actividades\" + Trim(Str(Month(Date))) + "\datosActividades\" + Left(trabajos(List1.ListIndex), Len(trabajos(List1.ListIndex)) - 4) + ".gui" For Random As #2 Len = Len(miRegistro)
            Get #2, 1, miRegistro   ' Lee el regitro
            Close #2   ' Cierra el archivo.
            
            Decir "el tema de la actividad es " + Trim(miRegistro.tema) + ". Tiene como comentario lo siguiente: " + Trim(miRegistro.comentarios)
        Else
            Decir "No hay actividades guardadas para el día de hoy"
        End If
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.actividadesHoy
         frmAyuda.Show
         Exit Sub
    End If
    
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
End Sub


Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swActividadDeHoy = True
    swPulsóEnterParaAvanzar = False
    
    Me.Caption = "Actividades de hoy"
    
    Decir "Entrando en las actividades de " + miMateria + " de hoy. Elegí con las flechas cuál buscás y abrila con enter"
'    Label1.BackColor = Form4.BackColor
    Dim i As Integer, contador As Integer, cadenaAux As String, mes As String
    File1.path = App.path + dirTrabajo + "actividades\" + Trim(Str(Month(Date)))
    
    contador = 0
    For i = 0 To (File1.ListCount - 1) 'se examinan todos los archivos para clasificarlos en los anteriores o de hoy
        'cadena = File1.List(i)
        cadena = Left(Right(File1.List(i), Len(File1.List(i)) - cantPrefijo), Len(Right(File1.List(i), Len(File1.List(i)) - cantPrefijo)) - 4)
        If Left(cadena, 2) = Trim(Str(Day(Date))) And Right(cadena, 4) = Trim(Str(Year(Date))) Then 'si el archivo es de hoy
            contador = contador + 1
            cadenaAux = "Actividad " + Trim(Str(contador))
            
'            mes = Trim(Str(CInt(month(date))))
'            Open App.Path + "\datos\" + mes + "\datosActividades.gui" For Random As #1 Len = Len(miRegistro)
'            Do While Not EOF(1)   ' Repite hasta el final del archivo.
'               Get #1, , miRegistro   ' Lee el registro siguiente.
'               If Trim(miRegistro.DirArchivo) = App.Path + dirTrabajo + "actividades\" + Trim(Str(month(date))) + "\" + File1.List(i) Then Exit Do
'            Loop
'            Close #1   ' Cierra el archivo.

            If existeCarpeta(App.path + dirTrabajo + "actividades\" + Trim(Str(Month(Date))) + "\datosActividades\") Then
                Open App.path + dirTrabajo + "actividades\" + Trim(Str(Month(Date))) + "\datosActividades\" + Left(File1.List(i), Len(File1.List(i)) - 4) + ".gui" For Random As #2 Len = Len(miRegistro)
                Get #2, 1, miRegistro   ' Lee el regitro
                Close #2   ' Cierra el archivo.
            End If
                
            If Asc(Left(miRegistro.tema, 1)) Then
                cadenaAux = cadenaAux + ". El tema de la actividad es " + Trim(miRegistro.tema) + "."
            Else
                cadenaAux = cadenaAux + ". No se le ha escrito un tema a la actividad."
            End If
            
            List1.AddItem cadenaAux
            ReDim Preserve trabajos(0 To contador - 1)
            trabajos(contador - 1) = File1.List(i)
        End If
    Next i
                        
    If List1.ListCount = 0 Then List1.AddItem "No hay ninguna actividad guardada de " + miMateria + " para el día de hoy"
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
    If swPulsóEnterParaAvanzar = False Then frmActividades.Show
    'Call contarFormularios(False)
End Sub

Private Sub List1_DblClick()
    If List1.List(List1.ListIndex) <> ("No hay ninguna actividad guardada de " + miMateria + " para el día de hoy") And List1.List(List1.ListIndex) <> "" Then
'        frmLectorActividad.numMesParaAbrir = month(date)
        frmLectorActividad.archivoParaLeer = Trim(Str(Month(Date))) + "\" + trabajos(List1.ListIndex)
        frmLectorActividad.díaAbierto = List1.List(List1.ListIndex)
        frmLectorActividad.Show
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

'Private Sub Command1_GotFocus()
'    Decir Command1.Caption
'    sonido = sndPlaySound(App.Path + "\sonidos\cb.wav", SND_ASYNC)
'End Sub

Private Sub Command1_Click()
    List1_DblClick
End Sub

