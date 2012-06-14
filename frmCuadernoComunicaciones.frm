VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmCuadernoComunicaciones 
   Caption         =   "Cuaderno de Comunicaciones"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   Icon            =   "frmCuadernoComunicaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmCuadernoComunicaciones.frx":08CA
   ScaleHeight     =   7335
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4920
      TabIndex        =   0
      Top             =   330
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2930
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2930
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3615
      Left            =   165
      TabIndex        =   5
      Top             =   3480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6376
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmCuadernoComunicaciones.frx":2922
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   16777215
      ScrollBars      =   2
      TextRTF         =   $"frmCuadernoComunicaciones.frx":29A5
   End
   Begin TransparentButton.ButtonTransparent btnGuardar 
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Caption         =   "Guardar comunicación"
      EstiloDelBoton  =   1
      Picture         =   "frmCuadernoComunicaciones.frx":2A28
      PictureHover    =   "frmCuadernoComunicaciones.frx":3302
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba su nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3480
      TabIndex        =   10
      Top             =   360
      Width           =   1350
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ver las comunicaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2040
      TabIndex        =   9
      Top             =   2280
      Width           =   3810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1178
      X2              =   6713
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Elegir el año:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5520
      TabIndex        =   8
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Añadir una comunicación:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   1845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Elegir el mes del que se quieren ver las comunicaciones:"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   2415
   End
End
Attribute VB_Name = "frmCuadernoComunicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim swEmpecé As Boolean
Dim controlPresionado As Boolean

Private Sub btnGuardar_Click()
    Dim mes As Byte, cadenaMes As String ', dóndeEmpezarSangría
    
    'si no puso su nombre
    If Trim(Text1) = "" Then
        frmMsgBox.cadenaAMostrar = "No ha escrito su nombre, por favor escríbalo para poder añadir la comunicación"
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Text1.SetFocus
        Exit Sub
    End If
    
    'si no escribió la comunicación
    If Trim(RichTextBox2.Text) = "" Then
        frmMsgBox.cadenaAMostrar = "No ha escrito la comunicación que quiere añadir, por favor escríbala y vuelva a apretar el botón " + Chr(34) + "Guardar comunicación" + Chr(34) + "."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        RichTextBox2.SetFocus
        Exit Sub
    End If
    
    RichTextBox1.Text = RichTextBox1.Text + vbNewLine + vbNewLine
    mes = Mid(Format(Date, "dd/mm/yyyy"), 4, 2)
    Select Case mes
        Case 1
            cadenaMes = "enero"
        Case 2
            cadenaMes = "febrero"
        Case 3
            cadenaMes = "marzo"
        Case 4
            cadenaMes = "abril"
        Case 5
            cadenaMes = "mayo"
        Case 6
            cadenaMes = "junio"
        Case 7
            cadenaMes = "julio"
        Case 8
            cadenaMes = "agosto"
        Case 9
            cadenaMes = "setiembre"
        Case 10
            cadenaMes = "octubre"
        Case 11
            cadenaMes = "noviembre"
        Case 12
            cadenaMes = "diciembre"
    End Select
    RichTextBox1.Text = RichTextBox1.Text + "El día " + Left(Format(Date, "dd/mm/yyyy"), 2) + " de " + cadenaMes + " " + Trim(Text1) + " escribió:" + vbNewLine
'    dóndeEmpezarSangría = Len(RichTextBox1.Text)
    RichTextBox1.Text = RichTextBox1.Text + Trim(RichTextBox2.Text)
'    RichTextBox1.SelStart = dóndeEmpezarSangría
'    RichTextBox1.SelLength = Len(RichTextBox2.Text)
'    RichTextBox1.SelIndent = 0.5
'    RichTextBox1.SelLength = 0
    Dim nombreArchivo As String
    nombreArchivo = Trim(Str(Mid(Format(Date, "dd/mm/yyyy"), 4, 2))) + "-" + Trim(Str(Right(Format(Date, "dd/mm/yyyy"), 4)))
    RichTextBox1.SaveFile App.path + "\comunicaciones\" + nombreArchivo + ".rtf"
    RichTextBox2.Text = ""
    Text1 = ""
'    Call chequearEspacioEnDisco(Left(App.Path, 2))
    frmMsgBox.cadenaAMostrar = "La comunicación se añadió satisfactoriamente"
    frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
    frmMsgBox.Show 1
    Text1.SetFocus
End Sub

Private Sub btnGuardar_GotFocus()
    Decir "botón " + btnGuardar.Caption + ". Apretá enter para añadir la comunicación"
End Sub

Private Sub Combo1_Click()
    Call cargarComunicaciones
End Sub

Private Sub Combo1_GotFocus()
    Decir "elegí con las flechas el mes del que querés ver las comunicaciones y aceptá con enter"
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then Decir Combo1.Text
End Sub

Private Sub Combo2_Click()
    Call cargarComunicaciones
End Sub

Private Sub Combo2_GotFocus()
    Decir "ahora elegí con las flechas el año para ver las comunicaciones y elegí con enter"
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then Decir Combo2.Text
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    Static swEstoyEnComunicaciones As Boolean
    
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.cuadernoComunicaciones
         frmAyuda.Show
         Exit Sub
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en el cuaderno de comunicaciones"
    
    If shiftkey = 0 And KeyCode = vbKeyF1 Then 'f1 intercambia las comunic añadidas y añadir una nueva
        If swEstoyEnComunicaciones = False Then
            swEstoyEnComunicaciones = True
            RichTextBox1.SetFocus
        Else
            swEstoyEnComunicaciones = False
            Text1.SetFocus
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        If TypeOf Me.ActiveControl Is ComboBox Or TypeOf Me.ActiveControl Is TextBox Or TypeOf Me.ActiveControl Is RichTextBox Then
            SendKeys ("{tab}")
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode <> vbKeyF1 Then Decir ""
End Sub

Private Sub Form_Load()
    Dim mes As Byte
    Combo1.AddItem "Enero"
    Combo1.AddItem "Febrero"
    Combo1.AddItem "Marzo"
    Combo1.AddItem "Abril"
    Combo1.AddItem "Mayo"
    Combo1.AddItem "Junio"
    Combo1.AddItem "Julio"
    Combo1.AddItem "Agosto"
    Combo1.AddItem "Setiembre"
    Combo1.AddItem "Octubre"
    Combo1.AddItem "Noviembre"
    Combo1.AddItem "Diciembre"
    
    mes = Mid(Format(Date, "dd/mm/yyyy"), 4, 2)
    
    Combo1.ListIndex = mes - 1
       
    Dim i As Byte, añoActual As Byte
    añoActual = Right(Format(Date, "dd/mm/yyyy"), 2)
    
    For i = 0 To añoActual
        If Len(añoActual) = 1 Then
            Combo2.AddItem "200" & i
        Else
            Combo2.AddItem "20" & i
        End If
    Next
    Combo2.ListIndex = Combo2.ListCount - 1
        
    swEmpecé = False
    Call cargarComunicaciones
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
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
    frmPrincipal.Show
    'Call contarFormularios(False)
End Sub


Private Sub richTextbox1_GotFocus()
    Decir "aquí están las comunicaciones que se hicieron del mes y año que seleccionaste. Para leerlas usá las flechas, o si querés volver a añadir una comunicación apretá la tecla tab"
End Sub

Private Sub RichTextBox2_GotFocus()
    Decir "escribí la comunicación que quieras agregar y después apretá la tecla tab"
End Sub

Sub cargarComunicaciones()
    On Error GoTo manejoError:
    If Combo2.ListIndex <> -1 And Combo1.ListIndex <> -1 Then
        RichTextBox1.LoadFile App.path + "\comunicaciones\" + Trim(Str(Combo1.ListIndex + 1) + "-" + Combo2.List(Combo2.ListIndex)) + ".rtf"
    End If
    Exit Sub
manejoError:
    RichTextBox1.Text = ""
    RichTextBox1.SaveFile App.path + "\comunicaciones\" + Trim(Str(Combo1.ListIndex + 1) + "-" + Combo2.List(Combo2.ListIndex)) + ".rtf"
    Resume
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    Dim renglón As Long
    
    shiftkey = Shift And 7
        
    If shiftkey = vbCtrlMask And KeyCode = vbKeyLeft Then 'leer por palabras retrocediendo
        Decir decirPalabraSiguiente(RichTextBox1) 'cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyRight Then 'leer por palabras avanzando
        Decir decirPalabraSiguiente(RichTextBox1) 'cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyHome Then 'control + inicio
        If RichTextBox1 <> "" Then
            Decir decirPalabraSiguiente(RichTextBox1)
        Else
            Decir "La hoja está en blanco, no hay nada escrito"
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyHome Then 'tecla inicio
        renglón = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
        Decir "principio del renglón " + Str(renglón)
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyEnd Then 'control + fin
        If RichTextBox1 <> "" Then
            Decir "final de la hoja. Estás detrás de la palabra " + decirPalabraAnterior(RichTextBox1)
        Else
            Decir "La hoja está en blanco, no hay nada escrito"
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyEnd Then 'tecla fin
        renglón = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
        Decir "final del renglón " + Str(renglón)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyPageDown Then 'tecla avance de página
        renglón = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
        Decir "saltando hacia adelante al renglón " + Str(renglón)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyPageUp Then 'tecla retroceso de página
        renglón = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
        Decir "saltando hacia atrás al renglón " + Str(renglón)
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
    
    If shiftkey = 0 And KeyCode = vbKeyDown Then 'leer por oración
        Decir decirOraciónSiguiente(RichTextBox1)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyUp Then 'leer por oración
        Decir decirOraciónSiguiente(RichTextBox1)
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyDown Then 'leer todo el texto
        If Trim(RichTextBox1.Text) <> "" Then
            Decir RichTextBox1.Text
        Else
            Decir "No se pueden leer todas las comunicaciones porque no se ha añadido ninguna en este mes"
        End If
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyUp Then 'leer desde el cursor hacia adelante
        If Trim(RichTextBox1.Text) <> "" Then
            If RichTextBox1.SelStart = 0 Then
                Decir Mid(RichTextBox1.Text, 1, Len(RichTextBox1.Text) - Len(Left(RichTextBox1.Text, RichTextBox1.SelStart))) 'leer desde el cursor hacia adelante
            Else
                Decir Mid(RichTextBox1.Text, RichTextBox1.SelStart, Len(RichTextBox1.Text) - Len(Left(RichTextBox1.Text, RichTextBox1.SelStart))) 'leer desde el cursor hacia adelante
            End If
        Else
            Decir "No se puede leer desde donde estás hacia adelante porque no hay ninguna comunicación guardada en este mes"
        End If
    End If
End Sub

Private Sub Text1_GotFocus()
    If swEmpecé = False Then
        Decir "Abriendo el cuaderno de comunicaciones. Podés escribir una comunicación o leer las que ya se han escrito. Para leer las comunicaciones de este mes, apretá F1; o para añadir una comunicación nueva escribí tu nombre y después apretá la tecla enter"
        swEmpecé = True
    Else
        Decir "escribí tu nombre y después apretá la tecla enter"
    End If
End Sub

Private Sub richtextbox2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim archivolibre As Byte, shiftkey As Integer
    
    shiftkey = Shift And 7
    If shiftkey = vbCtrlMask Then controlPresionado = True
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then Decir "está escrito " + RichTextBox2.Text
    
    Dim letra As String, auxString As String
    If KeyCode = vbKeyDelete Then
        If RichTextBox2.Text <> "" Then 'si no está vacío
            If RichTextBox2.SelStart <> Len(RichTextBox2.Text) Then 'y no está al final de la hoja
                If RichTextBox2.SelText <> "" Then 'si hay algo seleccionado
                    Decir "borrando el texto seleccionado"
                    Exit Sub
                End If
                letra = Mid(RichTextBox2.Text, RichTextBox2.SelStart + 1, 1)
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
        If RichTextBox2.Text = "" Then
            Decir "Ya está todo borrado"
        Else
            If RichTextBox2.SelStart = 0 Then
                Decir "imposible borrar porque estás al principio del texto"
            Else
                If RichTextBox2.SelText = "" Then 'si no hay nada seleccionado
                    letra = Mid(RichTextBox2.Text, RichTextBox2.SelStart, 1)
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

Private Sub richtextbox2_KeyPress(KeyAscii As Integer)
    Dim cadena As String
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
    Or KeyAscii = Asc("-") Or KeyAscii = Asc("_") Then cadena = cadena + decirPalabraAnterior(RichTextBox2)

    If cadena <> "" Then
        If RichTextBox2.SelBold = True Then cadena = cadena + " en negrita"
        If RichTextBox2.SelUnderline = True Then cadena = cadena + " subrayada"
        Decir cadena
    End If
    
    controlPresionado = False 'se resetea la variable
End Sub

Private Sub richtextbox2_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    
    shiftkey = Shift And 7
            
    If shiftkey = vbCtrlMask And KeyCode = vbKeyLeft Then 'leer por palabras retrocediendo
        Decir decirPalabraSiguiente(RichTextBox2) 'cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyRight Then 'leer por palabras avanzando
        Decir decirPalabraSiguiente(RichTextBox2) 'cadena
    End If
        
    If shiftkey = 0 And KeyCode = vbKeyRight Then 'avanzar de a caracter
        Decir decirLetraSiguiente(RichTextBox2)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyLeft Then 'retroceder de a caracter
        If RichTextBox2.SelStart = 0 And RichTextBox2.Text <> "" Then
            Decir "Estás en el principio del texto, delante de la letra " + decirLetraSiguiente(RichTextBox2)
        Else
            Decir decirLetraSiguiente(RichTextBox2)
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

    Call evaluarSelección(RichTextBox2, control, shift2, teclaApretada) 'se ve si hay selección
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copiar
        If RichTextBox2.SelText <> "" Then
            Decir "se copió el texto seleccionado"
        Else
            Decir "No se puede copiar porque no hay texto seleccionado. para seleccionar, usar shift más las flechas"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyX Then 'cortar
        If RichTextBox2.SelText <> "" Then
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

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim archivolibre As Byte, shiftkey As Integer
    
    shiftkey = Shift And 7
    If shiftkey = vbCtrlMask Then controlPresionado = True
    
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
    Or KeyAscii = Asc("-") Or KeyAscii = Asc("_") Then cadena = cadena + Text1

    If cadena <> "" Then Decir cadena
   
    controlPresionado = False 'se resetea la variable
End Sub

