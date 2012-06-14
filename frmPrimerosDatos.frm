VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmPrimerosDatos 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Primer uso de la mochila"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4140
   Icon            =   "frmPrimerosDatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPrimerosDatos.frx":08CA
   ScaleHeight     =   3555
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   615
      Left            =   870
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      Caption         =   "Ingresar los datos"
      EstiloDelBoton  =   1
      Picture         =   "frmPrimerosDatos.frx":14F2
      PictureHover    =   "frmPrimerosDatos.frx":1DCC
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Esta pantalla solamente va a aparecer en este primer uso de la mochila."
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "¿Sos hombre o mujer?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escribí tu nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frmPrimerosDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cadenaInicio As String
Dim swEmpezando As Boolean

Private Sub ButtonTransparent1_Click()
    Call Form_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub Combo1_GotFocus()
    Decir "ahora elegí con las flechas si sos hombre o mujer y apretá enter"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If TypeOf Me.ActiveControl Is ComboBox Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            Decir Combo1.List(Combo1.ListIndex)
        End If
    End If
    
'    If TypeOf Me.ActiveControl Is TextBox Then
'        If KeyCode <> vbKeyReturn Then Decir Text1
'    End If
    
    Dim archivolibre As Byte
    If KeyCode = vbKeyReturn Then
        If Trim(Text1) <> "" And Combo1.Text <> "" Then
            usuario.nombre = Trim(Text1)
            If Combo1.Text = "hombre" Then
                usuario.usuarioMujer = False
            Else
                usuario.usuarioMujer = True
            End If
            archivolibre = FreeFile 'se abre el archivo para guardar los datos de las partidas
            Open App.path + "\datos\datos.gui" For Random As archivolibre Len = Len(usuario)
            Put archivolibre, 1, usuario
            Close archivolibre
            frmPrincipal.Show
            Unload Me
        Else
            If Trim(Text1) = "" Then
                Text1.SetFocus
            Else
                Combo1.SetFocus
            End If
        End If
    End If
    
    If shiftkey = vbCtrlMask Then Decir ""
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    With usuario
'        .comenzarEnCarpeta = True
        .sapi5 = False 'si se usa sapi 5 o sapi 4
'        .permitirEditarActividades = False
        .usarVoz = True
        .mostrarTodasLasTareas = False
        .mostrarTodasLasActividades = False
        .mostrarAñoEnEvaluaciones = False
        .nombre = "Usuario"
        .usuarioMujer = True
        '.leerSignoPuntuación = False
        .imprimirDirecto = True
        .fuenteColor = vbBlack
'        Trim (.fuenteNombre)
        .fuenteTamaño = 12
        .colorFondo = vbWhite
        
'        .velocidadVoz
        .swLeerRenglones = True
        .swUsarCorrectorOrtográfico = False
'        Trim (.nombreVozSapi4)
'        Trim (.nombreVozSapi5)
        '.swInstalarVoz = True
        .swMúsicaDeFondo = False
        .swPermitirAbrirArchivos = False
        .rutaMúsicaFormPrincipal = "principal.mid"
        .rutaMúsicaFormCuaderno = "cuaderno.mid"
        .rutaMúsicaFormActividad = "actividades.mid"
        .rutaMúsicaFormLibros = "libros.mid"
        .rutaMúsicaFormAccesorios = "accesorios.mid"
        .rutaMúsicaFormTareas = "tareas.mid"
        .swVerActividadesConJaws = False
    End With
    Combo1.AddItem "hombre"
    Combo1.AddItem "mujer"
    Load frmControl
    Unload frmControl
    cadenaInicio = "Seas bienvenido o bienvenida a tu mochila personal, en la que vas a poder escribir todos tus trabajos y guardarlos en las carpetas de cada materia, hacer actividades, leer libros, escuchar música, y muchas cosas más. Esta bienvenida solamente va a aparecer esta única vez. Para empezar, necesito que me des algunos datos personales tuyos. En primer lugar escribí tu nombre, y luego presioná enter"
    swEmpezando = True
End Sub

Private Sub Text1_GotFocus()
    If swEmpezando = True Then
        Decir cadenaInicio
        swEmpezando = False
    Else
        Decir "escribí tu nombre y apretá la tecla enter"
    End If
End Sub


Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim archivolibre As Byte, shiftkey As Integer
    
    shiftkey = Shift And 7
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then Decir "está escrito " + Text1.Text
    
    Dim letra As String, auxString As String
    If KeyCode = vbKeyDelete Then
        If Text1.Text <> "" Then 'si no está vacío
            If Text1.SelStart <> Len(Text1.Text) Then 'y no está al final de la hoja
                If Text1.SelText <> "" Then 'si hay algo seleccionado
                    Decir "borrando el texto seleccionado"
                    Exit Sub
                End If
                letra = Mid(Text1.Text, Text1.SelStart + 1, 1)
                If letra = " " Then
                    Decir "borrando a la derecha el espacio", False
                ElseIf letra = Chr(9) Then
                    Decir "borrando a la derecha un salto"
                Else
                    auxString = traducirParaBorrar(letra)
                    Decir "borrando a la derecha " + auxString
                End If
            Else
                Decir "imposible borrar, estás al final del texto"
            End If
        Else
            Decir "no se puede borrar a la derecha porque no has escrito nada"
        End If
    End If
    
    If KeyCode = vbKeyBack Then
        If Text1.Text = "" Then
            Decir "Ya está todo borrado"
        Else
            If Text1.SelStart = 0 Then
                Decir "imposible borrar porque estás al principio del texto"
            Else
                If Text1.SelText = "" Then 'si no hay nada seleccionado
                    letra = Mid(Text1.Text, Text1.SelStart, 1)
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

Private Sub text1_KeyPress(KeyAscii As Integer)
    Dim cadena As String
    cadena = ""
    If KeyAscii >= 32 And KeyAscii <= 255 Then cadena = quéLetraSeApretó(KeyAscii)

    If KeyAscii = 9 Then cadena = "salto hacia adelante" 'tab
    If KeyAscii = 39 Then cadena = "apóstrofo"
    If KeyAscii = 123 Then cadena = "abre llave"
    If KeyAscii = 125 Then cadena = "cierra llave"
    If KeyAscii = 91 Then cadena = "abre corchete"
    If KeyAscii = 93 Then cadena = "cierra corchete"
    If KeyAscii = 64 Then cadena = "arroba"
    If KeyAscii = 32 Then cadena = "espacio"
    If KeyAscii = Asc(".") Then cadena = "punto"
    If KeyAscii = Asc(",") Then cadena = "coma"
    If KeyAscii = Asc(";") Then cadena = "punto y coma"
    If KeyAscii = Asc(":") Then cadena = "dos puntos"
    If KeyAscii = Asc("-") Then cadena = "guión"
    If KeyAscii = Asc("_") Then cadena = "guión bajo"

'    'leer la palabra al apretar espacio, punto, coma, etc.
    If KeyAscii = 32 Or KeyAscii = Asc(".") Or KeyAscii = Asc(",") Or KeyAscii = Asc(";") Or KeyAscii = Asc(":") _
    Or KeyAscii = Asc("-") Or KeyAscii = Asc("_") Then cadena = cadena + " " + Text1

    If cadena <> "" Then Decir cadena
End Sub


