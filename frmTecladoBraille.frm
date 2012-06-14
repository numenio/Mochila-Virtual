VERSION 5.00
Begin VB.Form frmTecladoBraille 
   Caption         =   "Teclado Perkins"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   5790
   ClientWidth     =   9495
   Icon            =   "frmTecladoBraille.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTecladoBraille.frx":08CA
   ScaleHeight     =   2565
   ScaleWidth      =   9495
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4920
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4320
      Top             =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTecladoBraille.frx":29F9
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Index           =   8
      Left            =   285
      TabIndex        =   8
      Top             =   1200
      Width           =   645
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   120
      Shape           =   3  'Circle
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Index           =   7
      Left            =   8520
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Espacio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Index           =   6
      Left            =   4320
      TabIndex        =   6
      Top             =   1680
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   5
      Left            =   7830
      TabIndex        =   5
      Top             =   1200
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   4
      Left            =   6855
      TabIndex        =   4
      Top             =   1200
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   3
      Left            =   5895
      TabIndex        =   3
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   1200
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   0
      Left            =   3435
      TabIndex        =   0
      Top             =   1200
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   3960
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Index           =   5
      Left            =   7560
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Index           =   4
      Left            =   6600
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Index           =   3
      Left            =   5640
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Index           =   2
      Left            =   1200
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Index           =   1
      Left            =   2160
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Index           =   0
      Left            =   3120
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmTecladoBraille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim puntos(6) As Boolean
'Public dóndeEscribir As RichTextBox
Dim letra As Byte
Dim swNúmero As Boolean
'Dim swParéntesisAbierto As Boolean
Dim swPreguntaAbierto As Boolean
Dim swAdmiraciónAbierto As Boolean
'Dim swCaracterEscritoAntesEsEspacio As Boolean 'para cerrar admirac y pregunta aunqeu no se hayan abierto si antes tienen una letra

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    Timer1.Enabled = True
    If KeyCode = vbKeyF Then puntos(0) = True
    If KeyCode = vbKeyD Then puntos(1) = True
    If KeyCode = vbKeyS Then puntos(2) = True
    If KeyCode = vbKeyJ Then puntos(3) = True
    If KeyCode = vbKeyK Then puntos(4) = True
    If KeyCode = vbKeyL Then puntos(5) = True
    
    If KeyCode = vbKeySpace Then
        If swNúmero = True Then swNúmero = False
        Call pegarTexto(" ")
'        frmCuaderno.RichTextBox1.Text = frmCuaderno.RichTextBox1.Text + " "
        Shape2.FillColor = vbRed
        Decir "espacio"
    End If
    
    If KeyCode = vbKeyReturn Then
        Call pegarTexto(vbNewLine)
'        frmCuaderno.RichTextBox1.Text = frmCuaderno.RichTextBox1.Text + vbNewLine
        Shape4.FillColor = vbRed
        Decir "bajando una línea"
        If swNúmero = True Then swNúmero = False
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en el teclado braille"
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    
    Dim auxString As String, swletra As Boolean, caracterBorrado As String
    If KeyCode = vbKeyBack Then
        If frmCuaderno.RichTextBox1.Text = "" Then
            Decir "Ya está todo borrado"
        Else
'            If frmCuaderno.RichTextBox1.SelStart = 0 Then
'                Decir "imposible borrar porque estás al principio de la hoja"
'            Else
                caracterBorrado = UCase(Right(frmCuaderno.RichTextBox1.Text, 1))
                If caracterBorrado = " " Then
                    Decir "borrando el espacio", False
                ElseIf caracterBorrado = Chr(10) Then
                    Decir "borrando la bajada de línea", False
                Else
                    swletra = True
                    Select Case caracterBorrado
                        Case "*"
                            auxString = " por "
                        Case "/"
                            auxString = " dividido "
                        Case "$"
                            auxString = " pesos "
                        Case "-"
                            auxString = " menos "
                        Case ","
                            auxString = " coma "
                        Case "+"
                            auxString = "más"
                        Case "-"
                            auxString = "menos"
                        Case "*"
                            auxString = "por"
                        Case "/"
                            auxString = "dividido"
                        Case "="
                            auxString = "igual"
                        Case ","
                            auxString = "coma"
                            swletra = False
                        Case "."
                            auxString = "punto"
                            swletra = False
                        Case ";"
                            auxString = "punto y coma"
                            swletra = False
                        Case ":"
                            auxString = "dos puntos"
                            swletra = False
                        Case Chr(34) '"'"
                            auxString = "comillas"
                            swletra = False
                        Case "¡"
                            auxString = "abre exclamación"
                            swletra = False
                        Case "!"
                            auxString = "cierra exclamación"
                            swletra = False
                        Case "¿"
                            auxString = "abre pregunta"
                            swletra = False
                        Case "?"
                            auxString = "cierra pregunta"
                            swletra = False
                        Case "$"
                            auxString = "pesos"
                            swletra = False
                        Case "%"
                            auxString = "porciento"
                            swletra = False
                        Case "("
                            auxString = "abre paréntesis"
                            swletra = False
                        Case ")"
                            auxString = "cierra paréntesis"
                            swletra = False
                        Case "á"
                            auxString = "a con acento"
                        Case "é"
                            auxString = "e con acento"
                        Case "í"
                            auxString = "i con acento"
                        Case "ó"
                            auxString = "o con acento"
                        Case "ú"
                            auxString = "u con acento"
                        Case "B"
                            auxString = "bé larga"
                        Case "C"
                            auxString = "cé"
                        Case "D"
                            auxString = "dé"
                        Case "F"
                            auxString = "éfe"
                        Case "G"
                            auxString = "gé"
                        Case "H"
                            auxString = "áche"
                        Case "J"
                            auxString = "jóta"
                        Case "K"
                            auxString = "ká"
                        Case "L"
                            auxString = "éle"
                        Case "M"
                            auxString = "éme"
                        Case "N"
                            auxString = "éne"
                        Case "Ñ"
                            auxString = "éñe"
                        Case "P"
                            auxString = "pé"
                        Case "Q"
                            auxString = "cú"
                        Case "R"
                            auxString = "érre"
                        Case "S"
                            auxString = "ése"
                        Case "T"
                            auxString = "té"
                        Case "V"
                            auxString = "vé corta"
                        Case "W"
                            auxString = "doble bé"
                        Case "X"
                            auxString = "équis"
                        Case "Y"
                            auxString = "i griega"
                        Case "Z"
                            auxString = "seta"
                        Case Else 'si es cualquier otro caracter
                            auxString = caracterBorrado
                    End Select
                    
                    If swletra = False Then
                        Decir "borrando el signo " + auxString, False, True
                    Else
                        Decir "borrando la letra " + auxString, False, True
                    End If
                End If
                'se borra la letra
                Call frmCuaderno.borrarUnCarácter
                Shape3.FillColor = vbRed
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        If mensajeSalir("¿Estás seguro que querés cerrar el teclado braille?") Then
            Unload Me
            Exit Sub
        End If
    End If
    
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.tecladoBraille
         frmAyuda.Show
         Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    Me.Move Me.Left, frmCuaderno.Top + frmCuaderno.ScaleHeight - Me.ScaleHeight
    swPreguntaAbierto = False
    swAdmiraciónAbierto = False
    swNúmero = False
    Decir "abriendo el teclado braille, ahora las únicas teclas que funcionan son ese, de, efe, jota, ka y ele. También enter, espacio, borrar, y escape para salir del teclado."
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Call contarFormularios(False)
    frmCuaderno.swVolviendodeBraille = True
    Decir "cerrando el teclado braille"
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Timer2.Enabled = True
    Dim i As Integer
    For i = 0 To 5
        If puntos(i) = True Then Shape1(i).FillColor = vbRed
    Next
    
    letra = encontrarLetra
    
    If frmCuaderno.RichTextBox1.Text <> "" Then
        If Right(frmCuaderno.RichTextBox1.Text, 1) <> Chr(32) _
        And Right(frmCuaderno.RichTextBox1.Text, 1) <> "¡" _
        And Right(frmCuaderno.RichTextBox1.Text, 1) <> "¿" Then 'si el caracter anterior es distinto de espacio, abre preg o abre admiración, se cierran los signos automáticamente
            swPreguntaAbierto = True
            swAdmiraciónAbierto = True
        Else
            swPreguntaAbierto = False
            swAdmiraciónAbierto = False
        End If
    End If
    
    If swPreguntaAbierto = True And letra = Asc("¿") Then 'se evalúa si hay que cerrar pregunta
        letra = Asc("?")
        swPreguntaAbierto = False
    End If

    If swAdmiraciónAbierto = True And letra = Asc("¡") Then 'se evalúa si hay que cerrar admiración
        letra = Asc("!")
        swAdmiraciónAbierto = False
    End If

    
    If letra <> 0 Then 'se escribe en el cuaderno si es un caracter válido
        Call pegarTexto(Chr(letra))
'        frmCuaderno.RichTextBox1.Text = frmCuaderno.RichTextBox1.Text + Chr(letra)
        Decir Chr(letra)
    End If
    
    If letra = Asc("¡") Then swAdmiraciónAbierto = True 'si se abren los signos, se avisa para que luego cierren
    If letra = Asc("¿") Then swPreguntaAbierto = True
    
    puntos(0) = False
    puntos(1) = False
    puntos(2) = False
    puntos(3) = False
    puntos(4) = False
    puntos(5) = False
End Sub

Function encontrarLetra() As Byte
    Static swMayúsculas As Boolean
    'se busca la tecla correspondiente a los puntos
    'Puntos presentes:
    '1
    If puntos(0) = True And puntos(1) = False And puntos(2) = False And puntos(3) = False And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("a")
    '16
    If puntos(0) = True And puntos(1) = False And puntos(2) = False And puntos(3) = False And puntos(4) = False And puntos(5) = True Then encontrarLetra = 0
    '15
    If puntos(0) = True And puntos(1) = False And puntos(2) = False And puntos(3) = False And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc("e")
    '156
    If puntos(0) = True And puntos(1) = False And puntos(2) = False And puntos(3) = False And puntos(4) = True And puntos(5) = True Then encontrarLetra = 0
    '14
    If puntos(0) = True And puntos(1) = False And puntos(2) = False And puntos(3) = True And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("c")
    '146
    If puntos(0) = True And puntos(1) = False And puntos(2) = False And puntos(3) = True And puntos(4) = False And puntos(5) = True Then encontrarLetra = 0
    '145
    If puntos(0) = True And puntos(1) = False And puntos(2) = False And puntos(3) = True And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc("d")
    '1456
    If puntos(0) = True And puntos(1) = False And puntos(2) = False And puntos(3) = True And puntos(4) = True And puntos(5) = True Then encontrarLetra = 0
    '13
    If puntos(0) = True And puntos(1) = False And puntos(2) = True And puntos(3) = False And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("k")
    '136
    If puntos(0) = True And puntos(1) = False And puntos(2) = True And puntos(3) = False And puntos(4) = False And puntos(5) = True Then encontrarLetra = Asc("u")
        '135
    If puntos(0) = True And puntos(1) = False And puntos(2) = True And puntos(3) = False And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc("o")
        '1356
    If puntos(0) = True And puntos(1) = False And puntos(2) = True And puntos(3) = False And puntos(4) = True And puntos(5) = True Then encontrarLetra = Asc("z")
        '134
    If puntos(0) = True And puntos(1) = False And puntos(2) = True And puntos(3) = True And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("m")
        '1346
    If puntos(0) = True And puntos(1) = False And puntos(2) = True And puntos(3) = True And puntos(4) = False And puntos(5) = True Then encontrarLetra = Asc("x")
        '1345
    If puntos(0) = True And puntos(1) = False And puntos(2) = True And puntos(3) = True And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc("n")
        '13456
    If puntos(0) = True And puntos(1) = False And puntos(2) = True And puntos(3) = True And puntos(4) = True And puntos(5) = True Then encontrarLetra = Asc("y")
        '12
    If puntos(0) = True And puntos(1) = True And puntos(2) = False And puntos(3) = False And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("b")
        '126
    If puntos(0) = True And puntos(1) = True And puntos(2) = False And puntos(3) = False And puntos(4) = False And puntos(5) = True Then
'        If swParéntesisAbierto = False Then
            encontrarLetra = Asc("(")
'            swParéntesisAbierto = True
'        Else
'            encontrarLetra = Asc(")")
'            swParéntesisAbierto = False
'        End If
        swNúmero = False
    End If
        '125
    If puntos(0) = True And puntos(1) = True And puntos(2) = False And puntos(3) = False And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc("h")
        '1256
    If puntos(0) = True And puntos(1) = True And puntos(2) = False And puntos(3) = False And puntos(4) = True And puntos(5) = True Then encontrarLetra = 0
        '124
    If puntos(0) = True And puntos(1) = True And puntos(2) = False And puntos(3) = True And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("f")
        '1246
    If puntos(0) = True And puntos(1) = True And puntos(2) = False And puntos(3) = True And puntos(4) = False And puntos(5) = True Then encontrarLetra = 0
        '1245
    If puntos(0) = True And puntos(1) = True And puntos(2) = False And puntos(3) = True And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc("g")
        '12456
    If puntos(0) = True And puntos(1) = True And puntos(2) = False And puntos(3) = True And puntos(4) = True And puntos(5) = True Then encontrarLetra = Asc("ñ")
        '123
    If puntos(0) = True And puntos(1) = True And puntos(2) = True And puntos(3) = False And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("l")
        '1236
    If puntos(0) = True And puntos(1) = True And puntos(2) = True And puntos(3) = False And puntos(4) = False And puntos(5) = True Then encontrarLetra = Asc("v")
        '1235
    If puntos(0) = True And puntos(1) = True And puntos(2) = True And puntos(3) = False And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc("r")
        '12356
    If puntos(0) = True And puntos(1) = True And puntos(2) = True And puntos(3) = False And puntos(4) = True And puntos(5) = True Then encontrarLetra = Asc("á")
        '1234
    If puntos(0) = True And puntos(1) = True And puntos(2) = True And puntos(3) = True And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("p")
        '12346
    If puntos(0) = True And puntos(1) = True And puntos(2) = True And puntos(3) = True And puntos(4) = False And puntos(5) = True Then encontrarLetra = 0
        '12345
    If puntos(0) = True And puntos(1) = True And puntos(2) = True And puntos(3) = True And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc("q")
        '123456 signo generador
'    If puntos(0) = True And puntos(1) = True And puntos(2) = True And puntos(3) = True And puntos(4) = True And puntos(5) = True Then encontrarLetra = Asc("signo generador")
    
        '2
    If puntos(0) = False And puntos(1) = True And puntos(2) = False And puntos(3) = False And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc(",")
        '26
    If puntos(0) = False And puntos(1) = True And puntos(2) = False And puntos(3) = False And puntos(4) = False And puntos(5) = True Then
'        If swPreguntaAbierto = False Then
            encontrarLetra = Asc("¿")
'            swPreguntaAbierto = True
'        Else
'            encontrarLetra = Asc("?")
'            swPreguntaAbierto = False
'        End If
    End If
        '25
    If puntos(0) = False And puntos(1) = True And puntos(2) = False And puntos(3) = False And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc(":")
        '256
    If puntos(0) = False And puntos(1) = True And puntos(2) = False And puntos(3) = False And puntos(4) = True And puntos(5) = True Then
        encontrarLetra = Asc("/")
        swNúmero = False
    End If
        '24
    If puntos(0) = False And puntos(1) = True And puntos(2) = False And puntos(3) = True And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("i")
        '246
    If puntos(0) = False And puntos(1) = True And puntos(2) = False And puntos(3) = True And puntos(4) = False And puntos(5) = True Then
        encontrarLetra = Asc("<")
        swNúmero = False
    End If
        '245
    If puntos(0) = False And puntos(1) = True And puntos(2) = False And puntos(3) = True And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc("j")
        '2456
    If puntos(0) = False And puntos(1) = True And puntos(2) = False And puntos(3) = True And puntos(4) = True And puntos(5) = True Then encontrarLetra = Asc("w")
        '23
    If puntos(0) = False And puntos(1) = True And puntos(2) = True And puntos(3) = False And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc(";")
        '236
    If puntos(0) = False And puntos(1) = True And puntos(2) = True And puntos(3) = False And puntos(4) = False And puntos(5) = True Then encontrarLetra = 34 'Asc("'")
        '235
    If puntos(0) = False And puntos(1) = True And puntos(2) = True And puntos(3) = False And puntos(4) = True And puntos(5) = False Then
'        If swAdmiraciónAbierto = False Then
            encontrarLetra = Asc("¡")
'            swAdmiraciónAbierto = True
'        Else
'            encontrarLetra = Asc("!")
'            swAdmiraciónAbierto = False
'        End If
    End If
        '2356
    If puntos(0) = False And puntos(1) = True And puntos(2) = True And puntos(3) = False And puntos(4) = True And puntos(5) = True Then
        encontrarLetra = Asc("=")
        swNúmero = False
    End If
        '234
    If puntos(0) = False And puntos(1) = True And puntos(2) = True And puntos(3) = True And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("s")
        '2346
    If puntos(0) = False And puntos(1) = True And puntos(2) = True And puntos(3) = True And puntos(4) = False And puntos(5) = True Then encontrarLetra = Asc("é")
        '2345
    If puntos(0) = False And puntos(1) = True And puntos(2) = True And puntos(3) = True And puntos(4) = True And puntos(5) = False Then encontrarLetra = Asc("t")
        '23456
    If puntos(0) = False And puntos(1) = True And puntos(2) = True And puntos(3) = True And puntos(4) = True And puntos(5) = True Then encontrarLetra = Asc("ú")
        'ningún punto
'    If puntos(0) = False And puntos(1) = False And puntos(2) = False And puntos(3) = False And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc(vbKeyA)
        '6
    If puntos(0) = False And puntos(1) = False And puntos(2) = False And puntos(3) = False And puntos(4) = False And puntos(5) = True Then encontrarLetra = 0
        '5
    If puntos(0) = False And puntos(1) = False And puntos(2) = False And puntos(3) = False And puntos(4) = True And puntos(5) = False Then encontrarLetra = 0
        '56
    If puntos(0) = False And puntos(1) = False And puntos(2) = False And puntos(3) = False And puntos(4) = True And puntos(5) = True Then encontrarLetra = 0
        '4
    If puntos(0) = False And puntos(1) = False And puntos(2) = False And puntos(3) = True And puntos(4) = False And puntos(5) = False Then encontrarLetra = 0
        '46
    If puntos(0) = False And puntos(1) = False And puntos(2) = False And puntos(3) = True And puntos(4) = False And puntos(5) = True Then
        encontrarLetra = 0
        swMayúsculas = True
        Decir "signo de mayúsculas"
        Exit Function
    End If
        '45
    If puntos(0) = False And puntos(1) = False And puntos(2) = False And puntos(3) = True And puntos(4) = True And puntos(5) = False Then encontrarLetra = 0
        '456
    If puntos(0) = False And puntos(1) = False And puntos(2) = False And puntos(3) = True And puntos(4) = True And puntos(5) = True Then encontrarLetra = 0
        '3
    If puntos(0) = False And puntos(1) = False And puntos(2) = True And puntos(3) = False And puntos(4) = False And puntos(5) = False Then
        encontrarLetra = Asc(".")
        swNúmero = False
    End If
        '36
    If puntos(0) = False And puntos(1) = False And puntos(2) = True And puntos(3) = False And puntos(4) = False And puntos(5) = True Then
        encontrarLetra = Asc("-")
        swNúmero = False
    End If
        '35
    If puntos(0) = False And puntos(1) = False And puntos(2) = True And puntos(3) = False And puntos(4) = True And puntos(5) = False Then encontrarLetra = 0
        '356
    If puntos(0) = False And puntos(1) = False And puntos(2) = True And puntos(3) = False And puntos(4) = True And puntos(5) = True Then encontrarLetra = 0
        '34
    If puntos(0) = False And puntos(1) = False And puntos(2) = True And puntos(3) = True And puntos(4) = False And puntos(5) = False Then encontrarLetra = Asc("í")
        '346
    If puntos(0) = False And puntos(1) = False And puntos(2) = True And puntos(3) = True And puntos(4) = False And puntos(5) = True Then encontrarLetra = Asc("ó")
        '345
    If puntos(0) = False And puntos(1) = False And puntos(2) = True And puntos(3) = True And puntos(4) = True And puntos(5) = False Then
        encontrarLetra = Asc(")")
        swNúmero = False
    End If
        '3456
    If puntos(0) = False And puntos(1) = False And puntos(2) = True And puntos(3) = True And puntos(4) = True And puntos(5) = True Then
        swNúmero = True
        Decir "signo de número"
    End If
    
    If encontrarLetra <> 0 Then
        If swMayúsculas = True Then
            encontrarLetra = encontrarLetra - 32
            swMayúsculas = False
        End If
    End If
    
    If swNúmero = True Then
        Select Case encontrarLetra
            Case Asc("a")
                encontrarLetra = Asc("1")
            Case Asc("b")
                encontrarLetra = Asc("2")
            Case Asc("c")
                encontrarLetra = Asc("3")
            Case Asc("d")
                encontrarLetra = Asc("4")
            Case Asc("e")
                encontrarLetra = Asc("5")
            Case Asc("f")
                encontrarLetra = Asc("6")
            Case Asc("g")
                encontrarLetra = Asc("7")
            Case Asc("h")
                encontrarLetra = Asc("8")
            Case Asc("i")
                encontrarLetra = Asc("9")
            Case Asc("j")
                encontrarLetra = Asc("0")
            Case Asc("o")
                encontrarLetra = Asc(">")
                swNúmero = False
            Case 34
                encontrarLetra = Asc("*")
                swNúmero = False
            Case Asc("¡")
                encontrarLetra = Asc("+")
                swNúmero = False
        End Select
    End If
End Function

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Dim i As Integer
    For i = 0 To 5
        Shape1(i).FillColor = vbBlack
    Next
    Shape2.FillColor = vbBlack
    Shape3.FillColor = vbBlack
    Shape4.FillColor = vbBlack
End Sub

Sub pegarTexto(texto As String)
    Clipboard.SetText texto
    frmCuaderno.RichTextBox1.SelText = Clipboard.GetText()
End Sub

