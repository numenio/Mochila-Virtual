VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TRANSPARENTBUTTON.OCX"
Begin VB.Form frmMesActividades 
   Caption         =   "Actividades: mes x"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   Icon            =   "frmMesActividades.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMesActividades.frx":08CA
   ScaleHeight     =   5910
   ScaleWidth      =   7920
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   4920
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      Caption         =   "   Abrir la actividad seleccionada"
      EstiloDelBoton  =   1
      Picture         =   "frmMesActividades.frx":2922
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
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   4350
      ItemData        =   "frmMesActividades.frx":31FC
      Left            =   120
      List            =   "frmMesActividades.frx":31FE
      TabIndex        =   0
      Top             =   360
      Width           =   7695
   End
End
Attribute VB_Name = "frmMesActividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MesParaAbrir As String
Public numMesParaAbrir As Byte
Public año As Integer 'el año para ver el mes
Dim cadena As String
Dim trabajos() As String
Dim miRegistro As DatosActividad
Dim contador As Integer

Private Sub Command1_Click()
    List1_DblClick
End Sub

'Private Sub btnActMesX_Click(Index As Integer)
'    frmLectorActividad.Show
'    Unload Me
'End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer, cadena As String
    
    If KeyCode = vbKeyEscape Then
        frmActAntFut.Show
        Unload Me
    End If
    
    If KeyCode = vbKeyF12 Then frmControl.Show
    
    If KeyCode = vbKeyF1 Then 'f1 dice los comentarios
        Open App.Path + "\datos\" + Trim(Str(numMesParaAbrir)) + "\datosActividades.gui" For Random As #1 Len = Len(miRegistro)
        Do While Not EOF(1)   ' Repite hasta el final del archivo.
           Get #1, , miRegistro   ' Lee el registro siguiente.
           If Trim(miRegistro.DirArchivo) = App.Path + dirTrabajo + "actividades\" + trabajos(List1.ListIndex) Then Exit Do
        Loop
        Close #1   ' Cierra el archivo.
        
        If Asc(Left(miRegistro.tema, 1)) Then
            cadena = "el tema de la actividad es " + Trim(miRegistro.tema)
        Else
            cadena = "no le han guardado un tema a la actividad"
        End If
        
        If Trim(miRegistro.comentarios) <> "" Then
            cadena = cadena + ". Tiene como comentario lo siguiente: " + Trim(miRegistro.comentarios)
        Else
            If cadena = "no le han guardado un tema a la actividad" Then
                cadena = cadena + ". no le han guardado un comentario a la actividad"
            Else
                cadena = ", tampoco se le ha guardado un comentario"
            End If
        End If
        Decir cadena
    End If
    
    shiftkey = Shift And 7
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = mesActividades
         frmAyuda.Show
         Exit Sub
    End If
End Sub



Private Sub Form_Load()
    Call centrarFormulario(Me)
    
    contador = 0
    
    Me.Caption = "Actividades "
    If swActividadAnterior = True Then
        Me.Caption = Me.Caption + "pasadas "
    Else
        Me.Caption = Me.Caption + "futuras "
    End If
    Me.Caption = Me.Caption + "del mes: " + MesParaAbrir
    Decir "abriendo las " + Me.Caption + " de la materia " + miMateria + "buscá cuál querés abrir y aceptá con enter"
    
    Dim j As Integer
    File1.Path = App.Path + dirTrabajo + "actividades\" + Trim(Str(numMesParaAbrir))
    If swActividadAnterior = True Then 'si son actividades pasadas
        For j = 0 To File1.ListCount - 1
            If Mid(File1.List(j), 7, 4) = año Then
                If año < Right(Date, 4) Then 'si el año es menor al actual
                    Call añadirActividad(File1.List(j), Trim(Str(numMesParaAbrir)))
                ElseIf año = Right(Date, 4) Then 'si es el actual año
                    If numMesParaAbrir < Mid(Date, 4, 2) Then 'si el mes es menor al actual
                        Call añadirActividad(File1.List(j), Trim(Str(numMesParaAbrir)))
                    ElseIf numMesParaAbrir = Mid(Date, 4, 2) Then 'si el mes es el actual
                        If Mid(File1.List(j), 4, 2) < Left(Date, 2) Then 'si el día es menor al actual
                            Call añadirActividad(File1.List(j), Trim(Str(numMesParaAbrir)))
                        End If
                    End If
                End If
            End If
        Next j
    Else 'si son actividades futuras
        For j = 0 To File1.ListCount - 1
            If Mid(File1.List(j), 7, 4) = año Then
                If año > Right(Date, 4) Then 'si el año es mayor al actual
                    Call añadirActividad(File1.List(j), Trim(Str(numMesParaAbrir)))
                ElseIf año = Right(Date, 4) Then 'si es el actual año
                    If numMesParaAbrir > Mid(Date, 4, 2) Then 'si el mes es mayor al actual
                        Call añadirActividad(File1.List(j), Trim(Str(numMesParaAbrir)))
                    ElseIf numMesParaAbrir = Mid(Date, 4, 2) Then 'si el mes es el actual
                        If Mid(File1.List(j), 4, 2) > Left(Date, 2) Then 'si el día es mayor al actual
                            Call añadirActividad(File1.List(j), Trim(Str(numMesParaAbrir)))
                        End If
                    End If
                End If
            End If
        Next j
    End If
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
    If List1.List(List1.ListIndex) <> "" Then
        frmLectorActividad.archivoParaLeer = Trim(Str(numMesParaAbrir)) + "\" + trabajos(List1.ListIndex)
        frmLectorActividad.díaAbierto = List1.List(List1.ListIndex)
        frmLectorActividad.Show
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

Private Sub añadirActividad(archivo As String, mes As String)
    Dim cadena As String, cadenaAux As String
    
    cadena = Left(Right(archivo, Len(archivo) - cantPrefijo), Len(Right(archivo, Len(archivo) - cantPrefijo)) - 4)
    cadenaAux = Left(cadena, 3) + Str(numMesParaAbrir) + "-" + Right(cadena, 4)
    cadenaAux = Format(Left(cadenaAux, 10))
    cadenaAux = Format(cadenaAux, "Long Date")
    cadenaAux = "Actividad del día " + cadenaAux ' + Str(NumMesParaAbrir) + " de " + MesParaAbrir 'transformarCadena(cadena)
    
    Open App.Path + "\datos\" + mes + "\datosActividades.gui" For Random As #1 Len = Len(miRegistro)
    Do While Not EOF(1)   ' Repite hasta el final del archivo.
       Get #1, , miRegistro   ' Lee el registro siguiente.
       If Trim(Right(miRegistro.DirArchivo, Len(miRegistro.DirArchivo) - 3)) = Right(App.Path, Len(App.Path) - 3) + dirTrabajo + "actividades\" + mes + "\" + archivo Then Exit Do
    Loop
    Close #1   ' Cierra el archivo.
    
    If Asc(Left(miRegistro.tema, 1)) Then
        cadenaAux = cadenaAux + ". El tema de la actividad es " + Trim(miRegistro.tema) + "."
    Else
        cadenaAux = cadenaAux + ". No se le ha escrito un tema a la actividad."
    End If
    
    List1.AddItem cadenaAux
    ReDim Preserve trabajos(0 To contador)
    trabajos(contador) = archivo
    contador = contador + 1
End Sub

