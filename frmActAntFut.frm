VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmActAntFut 
   Caption         =   "Actividades Ant/Fut"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   Icon            =   "frmActAntFut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmActAntFut.frx":08CA
   ScaleHeight     =   6330
   ScaleWidth      =   7890
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      Caption         =   "    Mostrar las actividades del mes seleccionado"
      EstiloDelBoton  =   1
      Picture         =   "frmActAntFut.frx":2922
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
      Height          =   4380
      ItemData        =   "frmActAntFut.frx":31FC
      Left            =   360
      List            =   "frmActAntFut.frx":31FE
      TabIndex        =   1
      Top             =   600
      Width           =   7215
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   6360
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmActAntFut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public añoParaVerMeses As Integer 'el año del que se mostrarán los meses
Dim swPulsóEnterParaAvanzar As Boolean

'Private Sub btnMesAct_Click(Index As Integer)
'    frmMesActividades.Show
'    Unload Me
'End Sub
'
Private Sub Command1_Click()
    If List1.List(List1.ListIndex) = "" Then
        If swHablarVoz = True Then '@ 1 debe selecc...
            Decir "Debe seleccionar un mes de la lista antes de activar el botón." ', True, True
        Else
            frmMsgBox.cadenaAMostrar = "Debe seleccionar un archivo de la lista antes de activar el botón"
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
        End If
        List1.SetFocus
    Else
        List1_DblClick
    End If
End Sub

'Private Sub Command1_GotFocus()
'    Decir Command1.Caption
'    sonido = sndPlaySound(App.Path + "\sonidos\cb.wav", SND_ASYNC)
'End Sub

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
    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    
    If KeyCode = vbKeyEscape Then Unload Me
    
    '@ 2 para abrir...
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en las actividades"
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.actAntFut
         frmAyuda.Show
         Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim cadena As String
    
    Call centrarFormulario(Me)
    swActividadDeHoy = False
    swPulsóEnterParaAvanzar = False
    
    If swMostrarAñoEnActividades = False Then añoParaVerMeses = Year(Date) 'si no se muestran los años, el año para ver en el form de actividades es el actual
    Me.Caption = "Meses con actividades del año " + Str(añoParaVerMeses) '@ 3 meses...
    
    Decir "Elegí con las flechas el mes de la actividad que busques, y abrilo con enter" '@ 4
    
    Dim i As Integer, j As Integer
    For i = 1 To 12
        File1.path = App.path + dirTrabajo + "actividades\" + Trim(Str(i))
        If File1.ListCount <> 0 Then 'si el mes tiene actividades
            For j = 0 To File1.ListCount - 1 'si alguna actividad es del mes seleccionado
                If Mid(File1.List(j), 7, 4) = añoParaVerMeses Then
                    List1.AddItem NumMesACadena(i) 'se añade el mes y se pasa al siguiente
                    Exit For 'se va al mes siguiente
                End If
            Next j
        End If
    Next
    
    '@ 5 no hay ninguna...
    If List1.ListCount = 0 Then List1.AddItem "No hay ninguna actividad guardada de " + miMateria
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
    If swPulsóEnterParaAvanzar = False Then frmActividades.Show
End Sub

Private Sub List1_DblClick()
    '@ 5 no hay..
    If List1.List(List1.ListIndex) <> ("No hay ninguna actividad guardada de " + miMateria) And List1.List(List1.ListIndex) <> "" Then
        frmCalendarioMúltiple.tipoElemento = elemento.actividad
        frmCalendarioMúltiple.MesParaAbrir = List1.List(List1.ListIndex)
        frmCalendarioMúltiple.numMesParaAbrir = pasarANúmero(List1.List(List1.ListIndex))
        frmCalendarioMúltiple.año = añoParaVerMeses
        frmCalendarioMúltiple.Show
        swPulsóEnterParaAvanzar = True
        Unload Me
    Else '@ 6 hno hay actividades, 7 a hoy guar..., 8 apretá escape para volver...
        Decir "no hay actividades guardadas de " + miMateria + ". apretá escape para volver al menú de actividades"
    End If
End Sub


Function transformarCadena(cadena As String)
    Dim mes As String
    Dim cadenaAux As String
    
    cadenaAux = Mid(cadena, 4, 2)
    Select Case cadenaAux '@ desde el 9 al 20 enero, febr...
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
    
    transformarCadena = mes
End Function

Private Sub List1_GotFocus()
    If List1.ListIndex <> -1 Then '@ 5
        If List1.List(List1.ListIndex) <> "No hay ninguna actividad guardada de " + miMateria Then
            Decir List1.List(List1.ListIndex)
        Else '@ 5, 21 apretá escape...
            Decir "No hay ninguna actividad guardada de " + miMateria + ". Apretá escape para volver a la carpeta"
        End If
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then List1_DblClick
End Sub

Function pasarANúmero(cadena As String) As Byte
    Dim mes As Byte
    Select Case cadena '@ de 9 al 21
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
    If List1.ListIndex <> -1 Then Decir List1.List(List1.ListIndex)
    sonido = sndPlaySound(App.path + "\sonidos\td.wav", SND_ASYNC)
End Sub


Function NumMesACadena(numMes As Integer) As String
    
    Dim mes As String

    Select Case numMes
        Case "01" '@ de 9 al 21
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

