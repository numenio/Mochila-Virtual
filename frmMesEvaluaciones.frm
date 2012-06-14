VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmMesEvaluaciones 
   Caption         =   "Evaluaciones realizadas"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   Icon            =   "frmMesEvaluaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMesEvaluaciones.frx":08CA
   ScaleHeight     =   6420
   ScaleWidth      =   8460
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   735
      Left            =   2483
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5400
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      Caption         =   "    Mostrar las evaluaciones del mes seleccionado"
      EstiloDelBoton  =   0
      Picture         =   "frmMesEvaluaciones.frx":326A
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
      ItemData        =   "frmMesEvaluaciones.frx":3B44
      Left            =   240
      List            =   "frmMesEvaluaciones.frx":3B46
      TabIndex        =   0
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "frmMesEvaluaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public añoParaVerMeses
Public swMateria As String
Dim swEmpezando As Boolean
Dim swPulsóEnterParaAvanzar As Boolean

Private Sub Command1_Click()
    If List1.List(List1.ListIndex) = "" Then
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
'        frmPrincipal.Show
        Unload Me
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en las evaluaciones"
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.mesEvaluaciones
         frmAyuda.Show
         Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    Me.Caption = "Evaluaciones realizadas de la materia " + Chr(34) + Me.swMateria + Chr(34)
    swEmpezando = True
    swPulsóEnterParaAvanzar = False
    Dim i As Integer, j As Integer
    Dim swEscribirMes As Boolean
    For i = 1 To 12
        File1.path = App.path + "\trabajos\" + swMateria + "\soporte\" + Trim(Str(i))
        If File1.ListCount <> 0 Then 'si el mes tiene evaluaciones
            For j = 0 To File1.ListCount - 1 'se chequean todos los archivos para ver si el año es el adecuado
                If Left(Right(File1.List(j), 8), 4) = Trim(Str(añoParaVerMeses)) Then 'si el año es el elegido
                'if mid(File1.List(j), cantPrefijo + 4, 4) = year(date) Then 'si el año es igual al actual
'                        If i < month(date) Then 'si el mes es menor al mes actual
                        List1.AddItem NumMesACadena(i)
                        Exit For
'                        ElseIf i = month(date) Then 'si el mes es el actual
'                            If Mid(File1.List(j), 4, 2) <= day(date) Then
'                                List1.AddItem NumMesACadena(i)   'si se escribe el mes, como ya se añadió se sigue con el sgte mes
'                                Exit For
'                            End If
'                        End If
                End If
            Next
        End If
    Next
                        
    If List1.ListCount = 0 Then List1.AddItem "No se ha guardado ninguna evaluación de " + swMateria
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
    
    If swPulsóEnterParaAvanzar = False Then frmPrincipal.Show
    
    'Call contarFormularios(False)
End Sub

Private Sub List1_DblClick()
    If List1.List(List1.ListIndex) <> ("No se ha guardado ninguna evaluación de " + swMateria) And List1.ListIndex <> -1 Then
        frmCalendarioMúltiple.tipoElemento = elemento.evaluación
        frmCalendarioMúltiple.MesParaAbrir = List1.List(List1.ListIndex)
        frmCalendarioMúltiple.numMesParaAbrir = pasarANúmero(List1.List(List1.ListIndex))
        frmCalendarioMúltiple.año = añoParaVerMeses
        frmCalendarioMúltiple.swMateriaEvaluaciones = swMateria
        frmCalendarioMúltiple.Show
        swPulsóEnterParaAvanzar = True
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
    If List1.List(List1.ListIndex) <> "No se ha guardado ninguna evaluación de " + swMateria Then
        If swEmpezando = True Then
            Decir "entrando a las evaluaciones de la materia " + swMateria + ". elegí con las flechas de qué mes es la hoja que querés abrir y seleccionalo con enter"
            swEmpezando = False
        Else
            Decir List1.List(List1.ListIndex)
        End If
    Else
        Decir "No se ha guardado ninguna evaluación de " + swMateria + ". Apretá escape para volver al menú principal"
    End If
End Sub
