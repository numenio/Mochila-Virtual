VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmControlAlumno 
   Caption         =   "Configuración de la Mochila"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   Icon            =   "frmControlAlumno.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmControlAlumno.frx":08CA
   ScaleHeight     =   8085
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo13 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4440
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"frmControlAlumno.frx":2922
   End
   Begin VB.ComboBox Combo12 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox Combo11 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.ComboBox Combo10 
      Height          =   315
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   6720
      Width           =   1935
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   6720
      Width           =   1935
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   6720
      Width           =   1935
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5880
      Width           =   1935
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   5880
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   5880
      TabIndex        =   28
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   5880
      Width           =   1935
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.ComboBox cmbFontSize 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3840
      Width           =   735
   End
   Begin VB.ComboBox cmbFontName 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3840
      Width           =   1815
   End
   Begin VB.ComboBox cmbFontColor 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ComboBox cmbFormColor 
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
   End
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   2760
      TabIndex        =   18
      Top             =   7320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   "     Aceptar"
      EstiloDelBoton  =   1
      Picture         =   "frmControlAlumno.frx":29A5
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
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido de los recordatorios:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   37
      Top             =   4440
      Width           =   1965
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Permitir ver otros archivos en las carpetas:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4440
      TabIndex        =   36
      Top             =   2400
      Width           =   2985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2040
      X2              =   6120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   1920
      X2              =   6000
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1920
      X2              =   6000
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usar música para identificar las partes de la mochila:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   35
      Top             =   5040
      Width           =   3690
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hojas anteriores:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5280
      TabIndex        =   34
      Top             =   6480
      Width           =   1185
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accesorios:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2640
      TabIndex        =   33
      Top             =   6480
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lector de libros:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   6480
      Width           =   1125
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lector de actividades:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5280
      TabIndex        =   31
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carpetas:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2640
      TabIndex        =   30
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menú principal:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   29
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Velocidad de la voz:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4440
      TabIndex        =   27
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de letra:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Color de la letra:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Color del fondo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño de la letra:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   23
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Elegí tu sexo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4560
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tu nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leer renglones en la carpeta, actividades y libros:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de voz:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   885
   End
End
Attribute VB_Name = "frmControlAlumno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim voces() As Byte
'Private Enum sapi
'    sapi4
'    sapi5
'End Enum
Dim índiceEnQueEmpiezaSapi4 As Integer
Dim swEmpezando As Boolean
Dim controlPresionado As Boolean

Private Sub Command1_Click() 'botón aceptar
    If Trim(RichTextBox1.Text = "") Then
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "Borraste tu nombre y no has escrito ninguno para reemplazar el que estaba. Ahora aceptá este mensaje y escribí tu nombre."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        RichTextBox1.SetFocus
        Exit Sub
    End If
    
    nombreUsuario = Trim(RichTextBox1.Text)

    If Combo2.Text = "Sí" Then
        swLeerRenglones = True
    Else
        swLeerRenglones = False
    End If
    
    If Combo11.Text = "Sí" Then
        swMúsicaDeFondo = True
    Else
        swMúsicaDeFondo = False
    End If

    If Combo12.Text = "Sí" Then
        swPermitirAbrirArchivos = True
    Else
        swPermitirAbrirArchivos = False
    End If

    If Combo3.Text = "Hombre" Then   'se elige si el usuario es hombre o mujer
        swUsuarioMujer = False
    Else
        swUsuarioMujer = True
    End If

    NombreFuente = cmbFontName  'se ajusta la fuente del programa
    'se graba en una variable el color de la fuente del programa
    Select Case cmbFontColor.List(cmbFontColor.ListIndex)
        Case "Blanco"
           colorFuente = vbWhite
        Case "Azul"
           colorFuente = vbBlue
        Case "Negro"
           colorFuente = vbBlack
        Case "Verde"
           colorFuente = vbGreen
        Case "Rojo"
            colorFuente = vbRed
        Case "Amarillo"
            colorFuente = vbYellow
    End Select
    
    'se graba en una variable el color de la fuente del programa
    Select Case cmbFormColor.List(cmbFormColor.ListIndex)
        Case "Blanco"
           colorFondo = vbWhite
        Case "Azul"
           colorFondo = vbBlue
        Case "Negro"
           colorFondo = vbBlack
        Case "Verde"
           colorFondo = vbGreen
        Case "Rojo"
            colorFondo = vbRed
        Case "Amarillo"
            colorFondo = vbYellow
    End Select

    tamañoFuente = cmbFontSize.Text 'se ajusta el tamaño de la fuente

    If Combo1.ListIndex < índiceEnQueEmpiezaSapi4 Then 'si la voz elegida es sapi5
        nombreSapi5 = Combo1.Text
    Else 'si es sapi4
        nombreSapi4 = Combo1.Text
    End If
    
    usuario.rutaMúsicaFormPrincipal = Combo5.Text
    usuario.rutaMúsicaFormCuaderno = Combo6.Text
    usuario.rutaMúsicaFormActividad = Combo7.Text
    usuario.rutaMúsicaFormLibros = Combo8.Text
    usuario.rutaMúsicaFormAccesorios = Combo9.Text
    usuario.rutaMúsicaFormTareas = Combo10.Text
    usuario.rutaSonidosRecordatorios = Combo13.Text

    GuardarDatosUsuario 'se guardan las preferencias del usuario

    If swCuadernoAbierto = True Then frmCuaderno.Refresh 'se actualiza el cuaderno si está abierto
    If swLibroAbierto = True Then frmLectorLibro.Refresh 'si está abierto el lector de libros, se lo actualiza
    If swActividadAbierta = True Then frmLectorActividad.Refresh 'si está abierto el lector de actividad, se lo actualiza para ver si se pueden o no modificar las actividades
    If frmLectorEvaluaciones.swEstoyAbierto = True Then frmLectorEvaluaciones.Refresh
    Decir "configuración guardada"
    Unload Me
End Sub

Private Sub Command1_GotFocus()
    Decir "botón aceptar, apretá la barra espaciadora o enter para guardar tu configuración"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    If KeyCode = vbKeyReturn Then
        If TypeOf Me.ActiveControl Is ComboBox Then
            Decir ""
            SendKeys "{tab}"
        End If
    End If
    
    If KeyCode = vbKeyEscape Then Unload Me
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.controlAlumno
         frmAyuda.Show
         Exit Sub
    End If
    
    If shiftkey = vbCtrlMask Then Decir ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    swEmpezando = True
    frmOculto.swContinuarReproducción = False
    frmOculto.media.Stop
    'Call contarFormularios(False)
End Sub

Private Sub richTextbox1_GotFocus()
    Dim cadena As String
    'SendKeys ("^{end}") 'se pasa al final del cuadro
    If swEmpezando = True Then
        cadena = "Abriendo la configuración de tu mochila. Aquí podés hacer que tu mochila esté a tu gusto. En primer lugar, podés ver o modificar tu nombre"
        swEmpezando = False
    Else
        cadena = "Aquí podés escribir tu nombre. Cuando termines apretá enter"
    End If
    If Trim(RichTextBox1.Text) <> "" Then cadena = cadena + ". Ya está escrito: " + RichTextBox1.Text
    Decir cadena
End Sub

Private Sub richTextbox1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    If shiftkey = vbCtrlMask Then controlPresionado = True
    If KeyCode = vbKeyReturn Then
        SendKeys "{BACKSPACE}"
        SendKeys "{tab}"
        Exit Sub
    End If
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then 'Or KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        Decir "está escrito " + RichTextBox1.Text
    End If
        
    Dim letra As String, auxString As String
    If KeyCode = vbKeyDelete Then
        If RichTextBox1.Text <> "" Then 'si no está vacío
            If RichTextBox1.SelStart <> Len(RichTextBox1.Text) Then 'y no está al final de la hoja
                If RichTextBox1.SelText <> "" Then 'si hay algo seleccionado
                    Decir "borrando el texto seleccionado"
                    Exit Sub
                End If
                letra = Mid(RichTextBox1.Text, RichTextBox1.SelStart + 1, 1)
                If letra = " " Then
                    Decir "borrando a la derecha el espacio", False
                ElseIf letra = Chr(9) Then
                    Decir "borrando a la derecha un salto"
                Else
                    auxString = traducirParaBorrar(letra)
                    Decir "borrando a la derecha " + auxString
                End If
            Else
                Decir "imposible borrar, estás al final del título"
            End If
        Else
            Decir "no se puede borrar a la derecha porque no hay nada escrito"
        End If
    End If
    
    If KeyCode = vbKeyBack Then
        If RichTextBox1.Text = "" Then
            Decir "Ya está todo borrado"
        Else
            If RichTextBox1.SelStart = 0 Then
                Decir "imposible borrar porque estás al principio de la hoja"
            Else
                If RichTextBox1.SelText = "" Then 'si no hay nada seleccionado
                    letra = Mid(RichTextBox1.Text, RichTextBox1.SelStart, 1)
                Else
                    Decir "borrando el texto seleccionado"
                    Exit Sub
                End If
    
                If letra = " " Then
                    Decir "borrando el espacio"
                ElseIf letra = Chr(9) Then
                    Decir "borrando un salto"
                Else
                    auxString = traducirParaBorrar(letra)
                    Decir "borrando " + auxString
                End If
            End If
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyEnd Then Decir "final del texto"
    If shiftkey = 0 And KeyCode = vbKeyHome Then Decir "principio del texto"
End Sub

Private Sub richTextbox1_KeyPress(KeyAscii As Integer)
    Dim cadena As String
    If Len(RichTextBox1.Text) < 64 Then
        If KeyAscii >= 32 And KeyAscii <= 255 And controlPresionado = False Then cadena = quéLetraSeApretó(KeyAscii)

        If KeyAscii = 9 Then cadena = "salto hacia adelante" 'tab
        If KeyAscii = 39 Then cadena = "apóstrofo"
        If KeyAscii = 123 Then cadena = "abre llave"
        If KeyAscii = 125 Then cadena = "cierra llave"
        If KeyAscii = 91 Then cadena = "abre corchete"
        If KeyAscii = 93 Then cadena = "cierra corchete"
        If KeyAscii = 64 Then cadena = "arroba"
        
        'leer la palabra al apretar espacio, punto, coma, etc.
        If KeyAscii = 32 Or KeyAscii = Asc(".") Or KeyAscii = Asc(",") Or KeyAscii = Asc(";") Or KeyAscii = Asc(":") _
        Or KeyAscii = Asc("-") Or KeyAscii = Asc("_") Then cadena = cadena + decirPalabraAnterior(RichTextBox1)

        If cadena <> "" Then
            If RichTextBox1.SelBold = True Then cadena = cadena + " en negrita"
            If RichTextBox1.SelUnderline = True Then cadena = cadena + " subrayada"
            Decir cadena
        End If
    Else
        Decir "ya se escribieron las 64 letras que le podés poner al título de tu hoja"
    End If
    controlPresionado = False 'se resetea la variable
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    
    shiftkey = Shift And 7
            
    If shiftkey = vbCtrlMask And KeyCode = vbKeyLeft Then 'leer por palabras retrocediendo
        Decir decirPalabraSiguiente(RichTextBox1) 'cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyRight Then 'leer por palabras avanzando
        Decir decirPalabraSiguiente(RichTextBox1) 'cadena
    End If
        
    If shiftkey = 0 And KeyCode = vbKeyRight Then 'avanzar de a caracter
        Decir decirLetraSiguiente(RichTextBox1)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyLeft Then 'retroceder de a caracter
        If RichTextBox1.SelStart = 0 And RichTextBox1.Text <> "" Then
            Decir "Estás en el principio del texto, delante de la letra " + decirLetraSiguiente(RichTextBox1)
        Else
            Decir decirLetraSiguiente(RichTextBox1)
        End If
    End If
    
    Dim TeclaShift, TeclaControl
    TeclaShift = (Shift And vbShiftMask) > 0
    TeclaControl = (Shift And vbCtrlMask) > 0

    Dim teclaApretada As Byte, control As Boolean, shift2 As Boolean
    Select Case KeyCode
        Case vbKeyA
            teclaApretada = tecla.a
        Case vbKeyUp
            teclaApretada = tecla.flechaArriba
        Case vbKeyDown
            teclaApretada = tecla.flechaAbajo
        Case vbKeyLeft
            teclaApretada = tecla.flechaIzquierda
        Case vbKeyRight
            teclaApretada = tecla.flechaDerecha
        Case vbKeyPageUp
            teclaApretada = tecla.avancePágina
        Case vbKeyPageDown
            teclaApretada = tecla.retrocesoPágina
        Case vbKeyHome
            teclaApretada = tecla.inicio
        Case vbKeyEnd
            teclaApretada = tecla.fin
        Case vbKeyBack
            teclaApretada = tecla.borrar
        Case vbKeyDelete
            teclaApretada = tecla.borrar
    End Select

    If TeclaControl Then
        control = True
    Else
        control = False
    End If

    If TeclaShift Then
        shift2 = True
    Else
        shift2 = False
    End If

    Call evaluarSelección(RichTextBox1, control, shift2, teclaApretada) 'se ve si hay selección
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copiar
        If RichTextBox1.SelText <> "" Then
            Decir "se copió el texto seleccionado"
        Else
            Decir "No se puede copiar porque no hay texto seleccionado. para seleccionar, usar shift más las flechas"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyX Then 'cortar
        If RichTextBox1.SelText <> "" Then
            Decir "se cortó el texto seleccionado"
        Else
            Decir "No se puede cortar porque no hay texto seleccionado. para seleccionar, usar shift más las flechas"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyV Then 'pegar
        If Clipboard.GetText <> "" Then
            Decir "texto pegado: " + Clipboard.GetText
        Else
            Decir "No se puede pegar porque no hay texto copiado o cortado. para copiar, usar control más ce. para cortar, usar control más équis"
        End If
    End If
    controlPresionado = False 'se resetea la variable
End Sub



Private Sub Combo1_Click()
    Decir ""
    If Combo1.ListIndex < índiceEnQueEmpiezaSapi4 Then 'si la voz elegida es sapi5
        Set Voz.Voice = Voz.GetVoices().Item(Combo1.ListIndex)
        velocidadVoz = 10
        Call regularVelocidadVoz
        If swEmpezando = False Then Voz.Speak "Elegiste mi voz para hablarte", SVSFPurgeBeforeSpeak Or SVSFlagsAsync
        swSapi5 = True
    Else 'si es la voz sapi4
        vozSapi4.CurrentMode = Combo1.ListIndex - índiceEnQueEmpiezaSapi4 + 1
        velocidadVoz = 1
        Call regularVelocidadVoz
        'vozSapi4.AudioReset
        If swEmpezando = False Then vozSapi4.Speak "Elegiste mi voz para hablarte"
        swSapi5 = False
    End If
    Call regularVelocidadVoz
End Sub


Private Sub Combo1_GotFocus()
    Decir "elegí con las flechas la voz que querés que te hable y aceptá con enter. Estás en " + Combo1.Text
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir Combo1.Text
End Sub

Private Sub Combo12_GotFocus()
    Decir "elegí con las flechas si querés permitir que puedas abrir cualquier archivo en las carpetas que sea r t f o t x t, y aceptá con enter. Estás en " + Combo12.Text
End Sub

Private Sub Combo12_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir Combo12.Text
End Sub

Private Sub Combo13_GotFocus()
    Decir "elegí con las flechas el sonido que querés escuchar cuando sea la hora de un recordatorio"
End Sub

Private Sub Combo13_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then Call reproducirSonidoRecordatorio(Combo13.Text)
End Sub

Private Sub Combo2_GotFocus()
    Decir "elegí con las flechas si querés que te lea los renglones en las carpetas, actividades y libros, y aceptá con enter. Estás en " + Combo2.Text
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir Combo2.Text
End Sub

Private Sub Combo3_GotFocus()
    Decir "elegí con las flechas si sos hombre o mujer y aceptá con enter. Estás en " + Combo3.Text
End Sub

Private Sub Combo3_Keyup(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir Combo3.Text
End Sub

Private Sub Combo4_GotFocus()
    Decir "elegí con las flechas la velocidad a la que querés que te hable y aceptá con enter. Estás en " + Combo4.Text
End Sub

Private Sub Combo4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        velocidadVoz = Int(Combo4.Text)
        velocidadVoz = velocidadVoz - 10
        Call regularVelocidadVoz
        Decir Combo4.Text
    End If
End Sub

Private Sub Combo5_GotFocus()
    Decir "elegí con las flechas la música que querés que suene cuando estés en el menú principal y aceptá con enter"
End Sub

Private Sub Combo5_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then reproducirMúsica (Combo5.Text)
End Sub


Private Sub Combo6_GotFocus()
    Decir "elegí con las flechas la música de tus carpetas y aceptá con enter"
End Sub

Private Sub Combo6_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then reproducirMúsica (Combo6.Text)
End Sub


Private Sub Combo7_GotFocus()
    Decir "elegí con las flechas la música del lector de actividades y aceptá con enter"
End Sub

Private Sub Combo7_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then reproducirMúsica (Combo7.Text)
End Sub


Private Sub Combo8_GotFocus()
    Decir "elegí con las flechas la música del lector de libros y aceptá con enter"
End Sub

Private Sub Combo8_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then reproducirMúsica (Combo8.Text)
End Sub


Private Sub Combo9_GotFocus()
    Decir "elegí con las flechas la música de los accesorios y aceptá con enter"
End Sub

Private Sub Combo9_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then reproducirMúsica (Combo9.Text)
End Sub


Private Sub Combo10_GotFocus()
    Decir "elegí con las flechas la música de las hojas que ya has escrito y aceptá con enter"
End Sub

Private Sub Combo10_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then reproducirMúsica (Combo10.Text)
End Sub


Private Sub Combo11_GotFocus()
    Decir "elegí con las flechas si queres que suene una música distinta en cada parte de la mochila, así podés reconocer esas partes más fácilmente, y aceptá con enter. Estás en " + Combo11.Text
End Sub

Private Sub Combo11_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir Combo11.Text
End Sub


Private Sub cmbFontName_GotFocus()
    Decir "elegí con las flechas el nombre de la letra con que querés escribir en tus carpetas y aceptá con enter. Estás en " + cmbFontName.Text
End Sub

Private Sub cmbFontName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir cmbFontName.Text
End Sub


Private Sub cmbFontSize_GotFocus()
    Decir "elegí con las flechas el tamaño de la letra que se va a ver en tus carpetas, libros y actividades, y aceptá con enter. Estás en " + cmbFontSize.Text
End Sub

Private Sub cmbFontSize_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir cmbFontSize.Text
End Sub


Private Sub cmbFontColor_GotFocus()
    Decir "elegí con las flechas el color de la letra que se va a ver en tus carpetas, libros y actividades, y aceptá con enter. Estás en " + cmbFontColor.Text
End Sub

Private Sub cmbFontColor_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir cmbFontColor.Text
End Sub


Private Sub cmbFormColor_GotFocus()
    Decir "elegí con las flechas el color de fondo que querés que se vea en tus carpetas, libros y actividades, y aceptá con enter. Estás en " + cmbFormColor.Text
End Sub

Private Sub cmbFormColor_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir cmbFormColor.Text
End Sub




Private Sub Form_Load()
    Dim i As Integer, modename As String, Token As ISpeechObjectToken, contador As Integer
    
    'Call contarFormularios(True)
    Call centrarFormulario(Me)

    swEmpezando = True
    Set Voz = Nothing
    Set vozSapi4 = Nothing
    Set Voz = New SpVoice
    Set vozSapi4 = New DirectSS
    
    If Trim(Left(nombreUsuario, 1)) <> Chr(0) Then
        RichTextBox1.Text = Trim(nombreUsuario)
    Else
        RichTextBox1.Text = "Usuario"
    End If
    
    For Each Token In Voz.GetVoices 'llenar voces SAPI 5
        Combo1.AddItem (Token.GetDescription())
'        ReDim Preserve voces(0 To contador)
'        voces(contador) = sapi.sapi5
        contador = contador + 1
    Next
    
    índiceEnQueEmpiezaSapi4 = contador

    For i = 1 To vozSapi4.CountEngines 'llenar voces SAPI 4
        modename = vozSapi4.modename(i)
        Combo1.AddItem modename
'        ReDim Preserve voces(0 To contador)
'        voces(contador) = sapi.sapi4
'        contador = contador + 1
    Next i
    
    If Combo1.ListCount = 0 Then 'se chequea si hay alguna voz instalada
        Set Voz = Nothing
        Set vozSapi4 = Nothing
    Else
        For i = 0 To Combo1.ListCount - 1 'se muestra la voz de sapi4 o sapi5
            If swSapi5 = True Then
                If nombreSapi5 = Combo1.List(i) Then
                    Combo1.ListIndex = i
                    Exit For
                End If
            Else
                If nombreSapi4 = Combo1.List(i) Then
                    vozSapi4.Speak ""
                    Combo1.ListIndex = i
                    Exit For
                End If
            End If
        Next
    End If
    
    Combo2.AddItem "Sí" 'leer renglones
    Combo2.AddItem "No"
    If usuario.swLeerRenglones = True Then
        Combo2.ListIndex = 0
    Else
        Combo2.ListIndex = 1
    End If
    
    Combo11.AddItem "Sí" 'usar música en forms
    Combo11.AddItem "No"
    If usuario.swMúsicaDeFondo = True Then
        Combo11.ListIndex = 0
    Else
        Combo11.ListIndex = 1
    End If
    
    Combo12.AddItem "Sí" 'permitir abrir archivos en carpetas
    Combo12.AddItem "No"
    If usuario.swPermitirAbrirArchivos = True Then
        Combo12.ListIndex = 0
    Else
        Combo12.ListIndex = 1
    End If
    
    Combo3.AddItem "Mujer" 'sexo
    Combo3.AddItem "Hombre"
    If usuario.usuarioMujer = True Then
        Combo3.ListIndex = 0
    Else
        Combo3.ListIndex = 1
    End If
    
    For i = 1 To 20 'velocidad de la voz
        Combo4.AddItem i
    Next
    If velocidadVoz + 10 <= 20 Then 'se muestra la velocidad que se está usando en el programa
        If velocidadVoz > 0 Then
            Combo4.ListIndex = velocidadVoz + 9 'se le suma 9 pues velvoz va de -10 a +10, y se le resta 1 pues la matriz de ítems del combo empieza en 0 (sería como hacer +10-1)
        Else
            Combo4.ListIndex = 0
        End If
    Else
        Combo4.ListIndex = Combo4.ListCount - 1
    End If
    
    Dim índice As Integer, CadenaAbuscar As String
    ' Agrega los colores a cmbFontColor.
    With cmbFontColor
        .AddItem "Negro"
        .AddItem "Azul"
        .AddItem "Rojo"
        .AddItem "Verde"
        .AddItem "Blanco"
        .AddItem "Amarillo"
        
        Select Case colorFuente
            Case vbBlack
               CadenaAbuscar = "Negro"
            Case vbBlue
               CadenaAbuscar = "Azul"
            Case vbRed
               CadenaAbuscar = "Rojo"
            Case vbGreen
               CadenaAbuscar = "Verde"
            Case vbWhite
                CadenaAbuscar = "Blanco"
            Case vbYellow
                CadenaAbuscar = "Amarillo"
        End Select
        índice = 0
        For i = 0 To .ListCount - 1
            If CadenaAbuscar = .List(i) Then
                índice = i
                Exit For
            End If
        Next
        .ListIndex = índice
    End With
    
    
    ' Agrega los colores a cmbFormColor.
    With cmbFormColor
        .AddItem "Blanco"
        .AddItem "Negro"
        .AddItem "Verde"
        .AddItem "Azul"
        .AddItem "Rojo"
        .AddItem "Amarillo"
        
        Select Case colorFondo
            Case vbBlack
               CadenaAbuscar = "Negro"
            Case vbBlue
               CadenaAbuscar = "Azul"
            Case vbRed
               CadenaAbuscar = "Rojo"
            Case vbGreen
               CadenaAbuscar = "Verde"
            Case vbWhite
                CadenaAbuscar = "Blanco"
            Case vbYellow
                CadenaAbuscar = "Amarillo"
        End Select
        índice = 0
        For i = 0 To .ListCount - 1
            If CadenaAbuscar = .List(i) Then índice = i
        Next
        .ListIndex = índice
    End With
    
    
    índice = 0
    With cmbFontName
       For i = 0 To Screen.FontCount - 1
            .AddItem Screen.Fonts(i)
            If NombreFuente = .List(i) Then índice = i
       Next i
       ' Establece ListIndex a la fuente que está guardada en datos usuario.
       .ListIndex = índice
    End With
    
    índice = 0
    With cmbFontSize
       ' Llena el control con tamaños en incrementos de 2.
       For i = 8 To 72 Step 2
          .AddItem i
       Next i
       
       For i = 0 To .ListCount - 1
            If tamañoFuente = CInt(cmbFontSize.List(i)) Then
                índice = i
                Exit For
            End If
        Next
       ' Establece ListIndex a 0
       .ListIndex = índice ' size 10.
    End With
    
    NombreFuente = cmbFontName  'se ajusta la fuente del programa
    'se graba en una variable el color de la fuente del programa
    Select Case cmbFontColor.List(cmbFontColor.ListIndex)
        Case "Blanco"
           colorFuente = vbWhite
        Case "Azul"
           colorFuente = vbBlue
        Case "Negro"
           colorFuente = vbBlack
        Case "Verde"
           colorFuente = vbGreen
        Case "Rojo"
            colorFuente = vbRed
        Case "Amarillo"
            colorFuente = vbYellow
    End Select

    tamañoFuente = cmbFontSize.Text 'se ajusta el tamaño de la fuente
    
    'se graba en una variable el color de la fuente del programa
    Select Case cmbFormColor.List(cmbFormColor.ListIndex)
        Case "Blanco"
           colorFondo = vbWhite
        Case "Azul"
           colorFondo = vbBlue
        Case "Negro"
           colorFondo = vbBlack
        Case "Verde"
           colorFondo = vbGreen
        Case "Rojo"
            colorFondo = vbRed
        Case "Amarillo"
            colorFondo = vbYellow
    End Select
    
    File1.Path = App.Path + "\sonidos\formularios\" 'se llenan los combo de música
    File1.Refresh
'    Combo5.AddItem "Buscar más música en la PC" 'para usar otra música diferentes de los que están en la mochila
'    Combo6.AddItem "Buscar más música en la PC"
'    Combo7.AddItem "Buscar más música en la PC"
'    Combo8.AddItem "Buscar más música en la PC"
'    Combo9.AddItem "Buscar más música en la PC"
'    Combo10.AddItem "Buscar más música en la PC"
    For i = 0 To File1.ListCount - 1
        Combo5.AddItem File1.List(i)
        Combo6.AddItem File1.List(i)
        Combo7.AddItem File1.List(i)
        Combo8.AddItem File1.List(i)
        Combo9.AddItem File1.List(i)
        Combo10.AddItem File1.List(i)
    Next
    
    For i = 0 To Combo5.ListCount - 1 'form principal
        If Combo5.List(i) = Trim(usuario.rutaMúsicaFormPrincipal) Then
            Combo5.ListIndex = i
            Exit For
        End If
    Next
    
    For i = 0 To Combo6.ListCount - 1 'form carpetas
        If Combo6.List(i) = Trim(usuario.rutaMúsicaFormCuaderno) Then
            Combo6.ListIndex = i
            Exit For
        End If
    Next
    
    For i = 0 To Combo7.ListCount - 1 'lector actividades
        If Combo7.List(i) = Trim(usuario.rutaMúsicaFormActividad) Then
            Combo7.ListIndex = i
            Exit For
        End If
    Next
    
    For i = 0 To Combo8.ListCount - 1 'libros
        If Combo8.List(i) = Trim(usuario.rutaMúsicaFormLibros) Then
            Combo8.ListIndex = i
            Exit For
        End If
    Next
    
    For i = 0 To Combo9.ListCount - 1 'accesorios
        If Combo9.List(i) = Trim(usuario.rutaMúsicaFormAccesorios) Then
            Combo9.ListIndex = i
            Exit For
        End If
    Next
    
    For i = 0 To Combo10.ListCount - 1 'hojas anteriores
        If Combo10.List(i) = Trim(usuario.rutaMúsicaFormTareas) Then
            Combo10.ListIndex = i
            Exit For
        End If
    Next
    
    File1.Path = App.Path + "\sonidos\recordatorios\"
    File1.Refresh
    For i = 0 To File1.ListCount - 1
        Combo13.AddItem File1.List(i)
    Next
    
    For i = 0 To Combo13.ListCount - 1 'sonidos recordatorios
        If Combo13.List(i) = Trim(usuario.rutaSonidosRecordatorios) Then
            Combo13.ListIndex = i
            Exit For
        End If
    Next
End Sub

Sub reproducirMúsica(quéMúsica As String)
    frmOculto.swContinuarReproducción = False
    frmOculto.media.Stop
    frmOculto.media.FileName = App.Path + "\sonidos\formularios\" + quéMúsica
    frmOculto.swContinuarReproducción = True
    frmOculto.media.Play
End Sub

Sub reproducirSonidoRecordatorio(quéSonido As String)
    sonido = sndPlaySound(App.Path + "\Sonidos\Recordatorios\" + quéSonido, SND_ASYNC)
End Sub
