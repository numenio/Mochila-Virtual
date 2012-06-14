VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmPrincipal 
   Caption         =   "Menú Principal"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5250
   Icon            =   "frmPrincipal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmPrincipal.frx":08CA
   ScaleHeight     =   8595
   ScaleWidth      =   5250
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox File2 
      Height          =   480
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TransparentButton.ButtonTransparent btnMateria 
      Height          =   615
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   5880
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   "Button dfdfTr"
      EstiloDelBoton  =   1
      Picture         =   "frmPrincipal.frx":29F9
      PictureSize     =   2
      PictureHover    =   "frmPrincipal.frx":32D3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   4920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   "Accesorios"
      EstiloDelBoton  =   1
      Picture         =   "frmPrincipal.frx":3BAD
      PictureSize     =   2
      PictureHover    =   "frmPrincipal.frx":4487
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin TransparentButton.ButtonTransparent command2 
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   3000
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   "Cuaderno de comunicaciones"
      EstiloDelBoton  =   1
      Picture         =   "frmPrincipal.frx":4D61
      PictureSize     =   2
      PictureHover    =   "frmPrincipal.frx":563B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin TransparentButton.ButtonTransparent btnConfiguración 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      Caption         =   ""
      EstiloDelBoton  =   1
      Picture         =   "frmPrincipal.frx":5F15
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
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   495
      Left            =   30
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Configurar las materias, voz y otros"
      Top             =   8040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      Caption         =   ""
      PicturePosition =   3
      EstiloDelBoton  =   1
      Picture         =   "frmPrincipal.frx":67EF
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
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent command3 
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   3840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Caption         =   " Evaluaciones"
      EstiloDelBoton  =   1
      Picture         =   "frmPrincipal.frx":70C9
      PictureSize     =   2
      PictureHover    =   "frmPrincipal.frx":79A3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Elegí de qué materia querés abrir la carpeta:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      X1              =   840
      X2              =   840
      Y1              =   1080
      Y2              =   8880
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   750
      Index           =   0
      Left            =   1440
      Top             =   7320
      Width           =   2670
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim columnas As Integer 'para saber cuántas columnas hay en el form con botones
'Dim swTerminar As Boolean 'para saber si al hacer escape se quiere terminar el programa
Dim swReproducir As Boolean 'para ver si se reproduce el sonido del botón
Dim cadenaVoz As String 'para lo que dice la voz al cargar el form
Dim hizoEnterparaAvanzar As Boolean
Public swEstoyAbierto As Boolean

Private Sub btnConfiguración_Click()
    frmControlActyLibros.Show
End Sub

Private Sub btnMateria_Click(Index As Integer)
    dirTrabajo = "\trabajos\" + Trim(btnMateria(Index).Caption) + "\"
    miMateria = Trim(btnMateria(Index).Caption)
'    If swEmpezarEnCuaderno = True Then
        frmCuaderno.swContinuarArchivo = False
        frmCuaderno.Show
'    Else
'        frmMateria.Show
'    End If
'    swTerminar = False
    hizoEnterparaAvanzar = True
    Unload Me
End Sub

Private Sub btnMateria_GotFocus(Index As Integer)
    If Index = 0 And swReproducir = False Then 'si recién se carga el form
        Decir cadenaVoz + btnMateria(Index).Caption, True, False
    Else
        Decir btnMateria(Index).Caption
    End If
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
    swReproducir = True 'que se reproduzca un sonido en todos los botones
End Sub

Private Sub btnMateria_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub btnMateria_MouseIn(Index As Integer, Shift As Integer)
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub ButtonTransparent1_Click()
    frmControl.Show
    'Call Form_KeyDown(vbKeyF12, 0)
End Sub

Private Sub Command1_Click() 'botón accesorios
'    swTerminar = False
    frmAccesorios.Show
    hizoEnterparaAvanzar = True
    Unload Me
End Sub


Private Sub Command1_GotFocus()
    Decir Command1.Caption, True, False
    If swReproducir = True Then
        sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
        swReproducir = True
    End If
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Command1_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Command2_Click() 'botón cuaderno de comunicaciones
'    swTerminar = False
    frmCuadernoComunicaciones.Show
    hizoEnterparaAvanzar = True
    Unload Me
End Sub

Private Sub Command2_GotFocus()
    Decir Command2.Caption, True, False
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
    swReproducir = True 'que se reproduzca un sonido en todos los botones
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub command2_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Command3_Click()
    frmMateriasEvaluaciones.Show
    hizoEnterparaAvanzar = True
    Unload Me
End Sub

Private Sub Command3_GotFocus()
    Decir Command3.Caption, True, False
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
    swReproducir = True 'que se reproduzca un sonido en todos los botones
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys ("{tab}")
    If KeyCode = vbKeyUp Then SendKeys ("+{tab}")
End Sub

Private Sub Command3_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Form_GotFocus()
    Call reproducirForm(formularios.principal)
    Command1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
'        swTerminar = False
'        swSalir = False
        If swHablarVoz = True Then
            Decir "Estás en el menú principal de tu mochila. para salir del programa apretá en cualquier momento alt+f4"
        Else
            'MsgBox "Para salir del programa apretá Alt + f4", , "Información"
            frmMsgBox.cadenaAMostrar = "Para salir del programa apretá Alt + f4"
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
        End If
    End If
    
    If KeyCode = vbKeyF7 And frmReproductorMúsica.swEstoyAbierto = True Then 'f7 vuelve a la mochila desde el reproductor
        Decir "Pasando al reproductor de música, para regresar a tu mochila otra vez, apretá f7"
        frmReproductorMúsica.Show
    End If
    
    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF12 Then frmControl.Show 'control f12 abre la configuración
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF11 Then frmControlActyLibros.Show 'control f11 abre las activ y libros
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    
    If shiftkey = vbCtrlMask Then Decir "" 'control calla la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.principal
         frmAyuda.Show
         Exit Sub
    End If
End Sub


Private Sub Form_Load()
    Dim archivolibre As Integer, auxUsuario As DatosUsuario
    
    On Error GoTo manejoErrorPpal
    'Call contarFormularios(True)
    Call centrarFormulario(Me)
    hizoEnterparaAvanzar = False
    
    miMateria = ""
    swEstoyAbierto = True
    swReproducir = False 'para que no reproduzca el sonido en el primer botón
    archivolibre = FreeFile
    Open App.path + "\datos\datos.gui" For Random As archivolibre Len = Len(usuario)
    Get archivolibre, 1, auxUsuario 'se toman los datos del usuario pasándolos a una variable con el mismo nombre
    Close #archivolibre
    
    If Trim(Left(auxUsuario.nombre, 1)) <> Chr(0) Then usuario = auxUsuario
    
    Load frmControl 'se carga el formulario de control
    
    Call reproducirForm(formularios.principal)

    Dim i As Integer
    Dim j As Integer, espacioEntreBotones As Integer, z As Integer, espacioHorizontalbotones As Byte
    Dim límite As Byte, límiteAnterior As Integer ', columnas As Integer
    
    cantPrefijo = 3
    espacioEntreBotones = 200
    espacioHorizontalbotones = 200
    límite = 10 'cantidad de materias por columna
    límiteAnterior = 1
    columnas = 1
    
    If frmControl.List2.List(0) <> "" Then 'si hay alguna materia guardada
        btnMateria(0).Caption = frmControl.List2.List(0) 'se carga en el botón 0 la materia 0
        While Len(btnMateria(0).Caption) < 20 'se les agregan espacios para que al tener todos los botones el mismo largo de cadena, los dibujos del los botones se vean alineados
            btnMateria(0).Caption = btnMateria(0).Caption + " "
        Wend
    Else
        btnMateria(0).Enabled = False 'si no hay materias guardadas
    End If
    
    Me.Width = 5200 'se establece el tamaño para una columna
    
    btnMateria(0).Top = 1300 '960
    btnMateria(0).Left = 1320
    
    Shape1(0).Top = btnMateria(0).Top - 60 'se igualan los rectángulos con los botones
    Shape1(0).Left = btnMateria(0).Left - 80
    
    i = 1 'btnMateria.Count
    'j empieza en 1 pues se saltea el primer btnMateria ya que el índice 0 se llenó arriba
    For j = 1 To frmControl.List2.ListCount - 1 'se crean tantos botones como materias hayan cargadas
        Load btnMateria(i)
        Load Shape1(i)
        If 0 = límiteAnterior Mod límite Then
            columnas = columnas + 1 'si se pasa el límite, se añade una columna
            Me.Width = Me.Width + 2600 'se estira el form si hay más de una col
        End If
        
        With btnMateria(i)
            If columnas > 1 Then 'si hay más de una columna
                z = i - (límite * (columnas - 1)) 'se empieza a contar el lugar de la columna de nuevo
            Else
                z = i 'si hay una sola columna, el lugar es igual al índice del vector a que pertenece el btnMateria
            End If
            
            If z <> 0 Then 'se ubican los botones en el form
                .Top = btnMateria(z - 1).Top + btnMateria(0).Height + espacioEntreBotones
                If columnas > 1 Then .Left = btnMateria(i - límite).Left + btnMateria(0).Width + espacioHorizontalbotones
            Else
                .Top = btnMateria(0).Top
                .Left = btnMateria(i - 1).Left + btnMateria(0).Width + espacioHorizontalbotones
            End If
            
            Shape1(i).Top = .Top - 60 'se igualan los rectángulos con los botones
            Shape1(i).Left = .Left - 80
            
            .Caption = frmControl.List2.List(j) 'se les da el título de las materias cargadas
            While Len(.Caption) < 20 'se les agregan espacios para que al tener todos los botones el mismo largo de cadena, los dibujos del los botones se vean alineados
                .Caption = .Caption + " "
            Wend
            Shape1(i).Visible = True
            .Visible = True
            .TabIndex = btnMateria(i - 1).TabIndex + 1
        End With
        
        i = i + 1 'se pasa al siguiente btnMateria
        límiteAnterior = i 'se sube el número de botones cargados
        
        If j = frmControl.List2.ListCount - 1 Then 'si se pusieron todas las materias, se agregan los botones de accesorios y cuaderno de comunic
            '++++++++++++++++++++
            'El botón evaluaciones
            If z <= 8 Then 'si en la fila hay 9 o menos botones, se pone en la misma col el botón accesorios
                Command3.Top = btnMateria(z).Top + btnMateria(0).Height + espacioEntreBotones 'se agrega accesorios
            Else
                Command3.Top = btnMateria(0).Top
                columnas = columnas + 1
                Me.Width = Me.Width + 2600 'se estira el form si hay más de una col
            End If
            
            If columnas > 1 Then
                Command3.Left = btnMateria(i - límite).Left + btnMateria(0).Width + espacioHorizontalbotones
            Else
                Command3.Left = btnMateria(0).Left
            End If
            Load Shape1(i + 1)
            Shape1(i + 1).Top = Command3.Top - 60 'se ubica el recuadro alrededor del botón
            Shape1(i + 1).Left = Command3.Left - 80
            Shape1(i + 1).Visible = True
            If z <= 8 Then
                z = z + 1
            Else
                z = 0
            End If
            i = i + 1
            Command3.TabIndex = btnMateria.UBound + 2
            
            '++++++++++++++++++++
            'El botón accesorios
            If z <= 8 Then 'si en la fila hay 9 o menos botones, se pone en la misma col el botón accesorios
                Command1.Top = Command3.Top + btnMateria(0).Height + espacioEntreBotones 'btnMateria(z).Top + btnMateria(0).Height + espacioEntreBotones 'se agrega accesorios
            Else
                Command1.Top = btnMateria(0).Top
                columnas = columnas + 1
                Me.Width = Me.Width + 2600 'se estira el form si hay más de una col
            End If
            
            If columnas > 1 Then
                Command1.Left = btnMateria(i - límite).Left + btnMateria(0).Width + espacioHorizontalbotones
            Else
                Command1.Left = btnMateria(0).Left
            End If
            Load Shape1(i + 1)
            Shape1(i + 1).Top = Command1.Top - 60 'se ubica el recuadro alrededor del botón
            Shape1(i + 1).Left = Command1.Left - 80
            Shape1(i + 1).Visible = True
            If z <= 8 Then
                z = z + 1
            Else
                z = 0
            End If
            i = i + 1
            Command1.TabIndex = Command3.TabIndex + 1
            
            '+++++++++++++++++++++++++++++++++++
            'El botón cuaderno de comunicaciones
            If z <= 8 Then 'si en la fila hay 9 o menos botones, se pone en la misma col el botón
                Command2.Top = Command1.Top + btnMateria(0).Height + espacioEntreBotones 'se agrega accesorios
            Else
                Command2.Top = btnMateria(0).Top
                columnas = columnas + 1
                Me.Width = Me.Width + 2600 'se estira el form si hay más de una col
            End If
            
            If columnas > 1 Then
                Command2.Left = btnMateria(i - límite).Left + btnMateria(0).Width + espacioHorizontalbotones
            Else
                Command2.Left = btnMateria(0).Left
            End If
            Load Shape1(i + 1)
            Shape1(i + 1).Top = Command2.Top - 60 'se ubica el recuadro alrededor del botón
            Shape1(i + 1).Left = Command2.Left - 80
            Shape1(i + 1).Visible = True
            z = z + 1
            i = i + 1
            Command2.TabIndex = Command1.TabIndex + 1
        End If
    Next
    
'    Command1.TabIndex = btnMateria.uBound + 2 'se les da el orden de tabulación al botón accesorios y cuaderno de com
'    command2.TabIndex = Command1.TabIndex + 1
    
'    If btnMateria.UBound > límite Then 'se ubica a la etiqueta inferior según los botones que haya
'        Label2.Top = btnMateria(límite - 1).Top + btnMateria(0).Height + espacioEntreBotones
'    Else
'        Label2.Top = btnMateria(btnMateria.UBound).Top + btnMateria(0).Height + espacioEntreBotones
'    End If
    
    Me.ScaleMode = 3
    Unload frmControl
    swEmpezóLaMochila = True
    If swYaEmpezóPrograma = False Then 'evalúa si recién empieza el programa o si se vuelve a este form desde otro
        If swUsuarioMujer = True Then
            cadenaVoz = "Bienvenida "
        Else
            cadenaVoz = "Bienvenido "
        End If
        cadenaVoz = cadenaVoz + Trim(nombreUsuario) + " a tu mochila. Elegí con las flechas qué carpeta querés abrir y hacelo con enter. Estás en "
        swYaEmpezóPrograma = True
    Else
        cadenaVoz = "volviendo al menú principal de tu mochila. Elegí qué carpeta querés abrir y hacelo con enter. Estás en "
    End If
    
    frmOculto.Timer1.Enabled = True
    Exit Sub
manejoErrorPpal:
'    MsgBox "soy el controlador de principal. Error número: " + Str(Err.Number) + ", descripción: " + Err.Description, , "Para mi creador"
    frmMsgBox.cadenaAMostrar = "soy el controlador de principal. Error número: " + Str(Err.Number) + ", descripción: " + Err.Description
    frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
    frmMsgBox.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hizoEnterparaAvanzar = False Then swSalir = True
    If swSalir = True Then
        If SalirDelPrograma = True Then
            chauPrograma
        Else
            swSalir = False
            Cancel = 1
        End If
        Exit Sub
    End If
    'Call contarFormularios(False)
    swEstoyAbierto = False
End Sub

Private Sub controlarEvaluaciones()
    Dim archivolibre As Byte, cadena As String, i As Integer, j As Integer
    Call eliminarEvaluacionesFalsas 'se borran todas las falsas para copiarlas de nuevo
    archivolibre = FreeFile 'se abren las materias
    Open App.path + "\datos\materias.txt" For Input As archivolibre
    While Not EOF(archivolibre)
        Line Input #archivolibre, cadena
        For i = 1 To 12
            File1.path = App.path + "\trabajos\" + Trim(cadena) + "\soporte\" + Trim(Str(i)) 'se añaden las materias al combo
'            If File1.ListCount > 0 Then
'                'arreglar que copie todos los archivos verdaderos cambiándole el dll por rtf
'            End If
        Next
    Wend
    Close #archivolibre
End Sub

Private Sub eliminarEvaluacionesFalsas()
    Dim i As Integer, j As Integer, archivolibre As Byte, cadena As String
    archivolibre = FreeFile 'se abren las materias
    Open App.path + "\datos\materias.txt" For Input As archivolibre
    While Not EOF(archivolibre)
        Line Input #archivolibre, cadena
        For i = 1 To 12 'se eliminan todas las evaluaciones
            File1.path = App.path + "\trabajos\" + Trim(cadena) + "\evaluaciones\" + Trim(Str(i)) 'se añaden las materias al combo
            If File1.ListCount > 0 Then
                For j = 0 To File1.ListCount - 1
                    Kill File1.List(j)
                Next
            End If
        Next
    Wend
    Close #archivolibre
End Sub
