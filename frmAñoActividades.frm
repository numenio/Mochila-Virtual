VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmA�oActividades 
   Caption         =   "Elegir el A�o de la Actividad"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   Icon            =   "frmA�oActividades.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmA�oActividades.frx":08CA
   ScaleHeight     =   5580
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
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
      ItemData        =   "frmA�oActividades.frx":2922
      Left            =   180
      List            =   "frmA�oActividades.frx":2924
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   540
      TabIndex        =   0
      Top             =   4680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      Caption         =   "    Mostrar actividades del a�o seleccionado"
      EstiloDelBoton  =   1
      Picture         =   "frmA�oActividades.frx":2926
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
Attribute VB_Name = "frmA�oActividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim swPuls�EnterParaAvanzar As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el men� de ventana si se aprieta alt
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
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de m�sica, ten�s que estar en el men� principal o en una carpeta. ahora est�s en las actividades"
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al men� de la aplicaci�n. Para leer los �tems de este men� necesit�s jaws u otro lector de pantallas. Para volver a la mochila, apret� escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda
         frmAyuda.formulario = formularios.a�oActividad
         frmAyuda.Show
         Exit Sub
    End If
End Sub


Private Sub Form_Load()
    Dim a�o As Integer, cadena As String, mes As Byte, d�a As Byte, fechaTotal As String
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swActividadDeHoy = True
    swPuls�EnterParaAvanzar = False
    Me.Caption = "A�os que tienen actividades"
    
    Decir "Entrando en los a�os que tienen actividades de " + miMateria + ". Eleg� con las flechas cu�l a�o quer�s abrir y acept� con enter"

    Dim i As Integer, j As Integer, swA�oDuplicado As Boolean
    For i = 1 To 12
        File1.path = App.path + dirTrabajo + "actividades\" + Trim(Str(i))
        If File1.ListCount <> 0 Then 'si el mes tiene actividades
            If swActividadAnterior = True Then 'si son actividades pasadas
                For j = 0 To File1.ListCount - 1
                    If Left(Right(File1.List(j), 8), 4) < Trim(Str(Year(Date))) Then 'si el a�o es menor al actual
                        swA�oDuplicado = controlarA�oDuplicado(Left(Right(File1.List(j), 8), 4))
                        If swA�oDuplicado = False Then List1.AddItem "A�o " + Left(Right(File1.List(j), 8), 4)
                    End If
                    
                    If Left(Right(File1.List(j), 8), 4) = Trim(Str(Year(Date))) Then 'si el a�o es igual al actual
                        If i < Month(Date) Then 'si el mes es menor al actual
                            swA�oDuplicado = controlarA�oDuplicado(Left(Right(File1.List(j), 8), 4))
                            If swA�oDuplicado = False Then List1.AddItem "A�o " + Left(Right(File1.List(j), 8), 4)
                        ElseIf i = Month(Date) Then 'si el mes es igual al actual
                            If Mid(File1.List(j), 4, 2) < Trim(Str(Day(Date))) Then 'si el d�a es menor al actual
                                swA�oDuplicado = controlarA�oDuplicado(Left(Right(File1.List(j), 8), 4))
                                If swA�oDuplicado = False Then List1.AddItem "A�o " + Left(Right(File1.List(j), 8), 4)
                            End If
                        End If
                    End If
                Next
            Else 'si son actividades futuras
                For j = 0 To File1.ListCount - 1
                    If Left(Right(File1.List(j), 8), 4) > Trim(Str(Year(Date))) Then 'si el a�o es menor al actual
                        swA�oDuplicado = controlarA�oDuplicado(Left(Right(File1.List(j), 8), 4))
                        If swA�oDuplicado = False Then List1.AddItem "A�o " + Left(Right(File1.List(j), 8), 4)
                    End If
                    
                    If Left(Right(File1.List(j), 8), 4) = Trim(Str(Year(Date))) Then 'si el a�o es igual al actual
                        If i > Month(Date) Then 'si el mes es menor al actual
                            swA�oDuplicado = controlarA�oDuplicado(Left(Right(File1.List(j), 8), 4))
                            If swA�oDuplicado = False Then List1.AddItem "A�o " + Left(Right(File1.List(j), 8), 4)
                        ElseIf i = Month(Date) Then 'si el mes es igual al actual
                            If Mid(File1.List(j), 4, 2) > Trim(Str(Day(Date))) Then 'si el d�a es menor al actual
                                swA�oDuplicado = controlarA�oDuplicado(Left(Right(File1.List(j), 8), 4))
                                If swA�oDuplicado = False Then List1.AddItem "A�o " + Left(Right(File1.List(j), 8), 4)
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next
    If List1.ListCount = 0 Then List1.AddItem "No hay ninguna actividad guardada"
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
    
    If swPuls�EnterParaAvanzar = False Then
        If swCuadernoAbierto = True Then
            frmCuaderno.Show
        Else
            frmPrincipal.Show
        End If
    End If
    
    'Call contarFormularios(False)
End Sub

Private Sub List1_DblClick()
    If List1.List(List1.ListIndex) <> ("No hay ninguna actividad guardada") And List1.ListIndex <> -1 Then
        frmActAntFut.a�oParaVerMeses = Trim(Right(List1.List(List1.ListIndex), 4))
        frmActAntFut.Show
        swPuls�EnterParaAvanzar = True
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

Private Function controlarA�oDuplicado(a�o As String) As Boolean
    Dim j As Integer, cadenaAux As String
    controlarA�oDuplicado = False
    For j = 0 To List1.ListCount - 1 'se controla que no est� ya el mes inclu�do
        cadenaAux = Right(List1.List(j), 4) 'se toma el a�o del listado
        If a�o = cadenaAux Then
            controlarA�oDuplicado = True 'si el a�o no est� ya en la lista
            Exit For
        End If
    Next
End Function
