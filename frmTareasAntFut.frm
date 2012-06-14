VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmTareasAnt 
   Caption         =   "Tareas Anteriores"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   Icon            =   "frmTareasAntFut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmTareasAntFut.frx":08CA
   ScaleHeight     =   6420
   ScaleWidth      =   8460
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5400
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1296
      Caption         =   "    Mostrar las hojas de carpeta del mes seleccionado"
      EstiloDelBoton  =   0
      Picture         =   "frmTareasAntFut.frx":326A
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      ItemData        =   "frmTareasAntFut.frx":3B44
      Left            =   240
      List            =   "frmTareasAntFut.frx":3B46
      TabIndex        =   0
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "frmTareasAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public añoParaVerMeses
Dim swEmpezando As Boolean

Private Sub Command1_Click()
    If List1.List(List1.ListIndex) = "" Then
'        MsgBox "Debe seleccionar un archivo de la lista antes de activar el botón", , "Cuidado!!"
        frmMsgBox.cadenaAMostrar = "Debe seleccionar un archivo de la lista antes de activar el botón"
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
    Else
        List1_DblClick
    End If
End Sub

Private Sub Command1_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
'        If swCuadernoAbierto = True Then 'si se vuelve al cuaderno
'            frmCuaderno.Show
''        Else
''            frmCarpeta.Show
'        End If
        Decir "volviendo a tu carpeta"
        Unload Me
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en las hojas ya escritas de tu carpeta"
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.tareasAnt
         frmAyuda.Show
         Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    Me.Caption = "Hojas anteriores de la carpeta de " + miMateria
    swEmpezando = True
    Dim i As Integer, j As Integer ', cadenaAux As String, swMesDuplicado As Boolean
    Dim swEscribirMes As Boolean
    For i = 1 To 12
        File1.path = App.path + dirTrabajo + Trim(Str(i))
        If File1.ListCount <> 0 Then 'si el mes tiene actividades
            For j = 0 To File1.ListCount - 1 'se chequean todos los archivos para ver si el día es anterior o igual al actual
                If Left(Right(File1.List(j), 8), 4) = Trim(Str(añoParaVerMeses)) Then 'si el año es igual al elegido
                    List1.AddItem NumMesACadena(i)   'si se escribe el mes, como ya se añadió se sigue con el sgte mes
                    Exit For
                End If
            Next
        End If
    Next
                        
    If List1.ListCount = 0 Then List1.AddItem "No se ha guardado ninguna hoja escrita de " + miMateria
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
    'If swCuadernoAbierto = True Then Decir "" 'callar la voz si se vuelve al cuaderno
End Sub

Private Sub List1_DblClick()
    If List1.List(List1.ListIndex) <> ("No se ha guardado ninguna hoja escrita de " + miMateria) And List1.List(List1.ListIndex) <> "" Then
'        frmMesTareasX.MesParaAbrir = List1.List(List1.ListIndex)
'        frmMesTareasX.numMesParaAbrir = pasarANúmero(List1.List(List1.ListIndex))
'        frmMesTareasX.Show
        
        frmCalendarioMúltiple.tipoElemento = elemento.tarea
        frmCalendarioMúltiple.MesParaAbrir = List1.List(List1.ListIndex)
        frmCalendarioMúltiple.numMesParaAbrir = pasarANúmero(List1.List(List1.ListIndex))
        frmCalendarioMúltiple.año = añoParaVerMeses
        frmCalendarioMúltiple.Show
        Unload Me
    End If
End Sub


Function NumMesACadena(numMes As Integer) As String
    
    Dim mes As String

    Select Case numMes
        Case "01"
            mes = "enero"
        Case "02"
            mes = "febrero"
        Case "03"
            mes = "marzo"
        Case "04"
            mes = "abril"
        Case "05"
            mes = "mayo"
        Case "06"
            mes = "junio"
        Case "07"
            mes = "julio"
        Case "08"
            mes = "agosto"
        Case "09"
            mes = "setiembre"
        Case "10"
            mes = "octubre"
        Case "11"
            mes = "noviembre"
        Case "12"
            mes = "diciembre"
    End Select
    
    NumMesACadena = mes
End Function

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then List1_DblClick
End Sub

Function pasarANúmero(cadena As String) As Byte
    Dim mes As Byte
    Select Case cadena
        Case "enero"
            mes = 1
        Case "febrero"
            mes = 2
        Case "marzo"
            mes = 3
        Case "abril"
            mes = 4
        Case "mayo"
            mes = 5
        Case "junio"
            mes = 6
        Case "julio"
            mes = 7
        Case "agosto"
            mes = 8
        Case "setiembre"
            mes = 9
        Case "octubre"
            mes = 10
        Case "noviembre"
            mes = 11
        Case "diciembre"
            mes = 12
    End Select
    pasarANúmero = mes
End Function

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn And KeyCode <> vbKeySpace Then Decir List1.List(List1.ListIndex)
    sonido = sndPlaySound(App.path + "\sonidos\td.wav", SND_ASYNC)
End Sub

Private Sub List1_GotFocus()
    If List1.List(List1.ListIndex) <> "No se ha guardado ninguna hoja escrita de " + miMateria Then
        If swEmpezando = True Then
            Decir "entrando a las hojas ya escritas de la materia " + miMateria + ". elegí con las flechas de qué mes es la hoja que querés abrir y seleccionalo con enter"
            swEmpezando = False
        Else
            Decir List1.List(List1.ListIndex)
        End If
    Else
        Decir "No se ha guardado ninguna hoja escrita de " + miMateria + ". Apretá escape para volver a la carpeta"
    End If
End Sub

'Private Sub Command1_GotFocus()
'    Decir Command1.Caption
'    sonido = sndPlaySound(App.Path + "\sonidos\cb.wav", SND_ASYNC)
'End Sub
