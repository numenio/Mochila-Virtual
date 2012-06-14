VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmLectorLibro 
   Caption         =   "Libro X: unidad X"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   Icon            =   "frmLectorLibro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmLectorLibro.frx":08CA
   ScaleHeight     =   7245
   ScaleWidth      =   7815
   Begin MSComDlg.CommonDialog di�logo 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfLectorLibros 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11245
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmLectorLibro.frx":2922
   End
   Begin TransparentButton.ButtonTransparent btnImprimir 
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Imprimir"
      EstiloDelBoton  =   4
      Picture         =   "frmLectorLibro.frx":29A5
      PictureHover    =   "frmLectorLibro.frx":327F
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
End
Attribute VB_Name = "frmLectorLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public archivoParaLeer As String
Dim ImpresoraRich As ImpresoraRTF 'la impresora del rtf

Private Sub ImprimirConCuadroDi�logo()
   ' El control CommonDialog se llama "dlgPrint".
    di�logo.CancelError = True
    On Error GoTo manejoErrorImpresora
    di�logo.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If rtfLectorLibros.SelLength = 0 Then
       di�logo.Flags = di�logo.Flags + cdlPDAllPages
    Else
       di�logo.Flags = di�logo.Flags + cdlPDSelection
    End If
    di�logo.ShowPrinter
    rtfLectorLibros.SelPrint di�logo.hDC
manejoErrorImpresora: 'si el error es distinto a haber hecho click en cancelar, se muestra un msg
'    If Err.Number <> 32755 Then MsgBox "La impresora no est� lista para imprimir." + Chr(13) + "Por favor vuelva a intentar cuando est� lista.", , "Informaci�n"
    Exit Sub
End Sub


Private Sub btnImprimir_Click()
    Call Form_KeyDown(vbKeyF6, 0)
End Sub

Private Sub rtflectorlibros_GotFocus()
    Call reproducirForm(formularios.tareasAnt)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el men� de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
'    If swCuadernoAbierto = True Then
'        If KeyCode = vbKeyF2 Then
'            frmCuaderno.Show
'        End If
'    End If
    
    If KeyCode = vbKeyEscape Then
        If mensajeSalir("�Est�s seguro que quer�s cerrar el lector de libros?") Then
'            If swCuadernoAbierto = False Then frmLibroX.Show
            Unload Me
        End If
    End If
    
    If KeyCode = vbKeyF3 Then
        If swCuadernoAbierto = True Then
            Decir "volviendo a tu carpeta de " + miMateria + ", acordate que el libro sigue abierto para que puedas seguir leyendo" ', para volver al libro abierto apret� f3"
            frmCuaderno.Show
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de m�sica, ten�s que estar en el men� principal o en una carpeta. ahora est�s en tu lector de libros"
    
    If (shiftkey = vbCtrlMask And KeyCode = vbKeyP) Or KeyCode = vbKeyF6 Then 'imprimir con control + p
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "�Realmente quer�s imprimir este libro?"
        frmMsgBox.swS�No�Aceptar = True 'se elige que sea cuadro s�-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado = True Then
            If swImprimirDirecto = True Then
                With ImpresoraRich
                     
                     'Valores
                     'Encabezado y pie de p�gina
                     .Header = "Libro de la carpeta de " + miMateria + " impreso el d�a " + Format(Date, "dd/mm/yyyy") 'Text1
                     .Footer = "Libro impreso por " + Trim(nombreUsuario) 'Text2
                     
                     'Margenes
                     .MarginTop = 500 'Text3
                     .MarginLeft = 500 'Text4
                     .MarginRight = 500 'Text5
                     .MarginBottom = 500 'Text6
                     
                     'Imprimir el RichTextBox pasado como par�metro
                     .Imprimir rtfLectorLibros
                
                End With
            Else
                ImprimirConCuadroDi�logo 'se muestra el cuadro de di�logo de la impresora
            End If
        End If
        Exit Sub
   End If

    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show

    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.lectorLibros
         frmAyuda.Show
         Exit Sub
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al men� de la aplicaci�n. Para leer los �tems de este men� necesit�s jaws u otro lector de pantallas. Para volver a la mochila, apret� escape"
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    Call cargarLibro
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If swSalir = True Then
        swSalir = False
        If SalirDelPrograma = True Then
            chauPrograma
        Else
            Cancel = 1
            swSalir = False
        End If
        Exit Sub
    End If
    'Call contarFormularios(False)
    swLibroAbierto = False
    If swCuadernoAbierto = True Then Decir "Cerrando el lector de libros. Otra vez est�s en tu cuaderno" 'callar la voz si se vuelve al cuaderno
End Sub

Private Sub rtfLectorLibros_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 32 And KeyAscii <= 126 Then Decir "No se puede escribir en el libro, s�lo leer. Para escribir apret� f3 as� pas�s a tu carpeta"
End Sub

Private Sub rtfLectorLibros_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    Dim rengl�n As Long
    
    shiftkey = Shift And 7
       
    If shiftkey = vbCtrlMask And KeyCode = vbKeyLeft Then 'leer por palabras retrocediendo
        Decir decirPalabraSiguiente(rtfLectorLibros) 'cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyRight Then 'leer por palabras avanzando
        Decir decirPalabraSiguiente(rtfLectorLibros) 'cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyHome Then 'control + inicio
        If rtfLectorLibros <> "" Then
            Decir decirPalabraSiguiente(rtfLectorLibros)
        Else
            Decir "La hoja del libro est� en blanco, no hay nada escrito"
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyHome Then 'tecla inicio
        rengl�n = rtfLectorLibros.GetLineFromChar(rtfLectorLibros.SelStart) + 1
        Decir "principio del rengl�n " + Str(rengl�n)
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyEnd Then 'control + fin
        If rtfLectorLibros <> "" Then
            Decir "final de la hoja del libro. Est�s detr�s de la palabra " + decirPalabraAnterior(rtfLectorLibros)
        Else
            Decir "La hoja del libro est� en blanco, no hay nada escrito"
        End If
        Exit Sub
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyEnd Then 'tecla fin
        rengl�n = rtfLectorLibros.GetLineFromChar(rtfLectorLibros.SelStart) + 1
        Decir "final del rengl�n " + Str(rengl�n)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyPageDown Then 'tecla avance de p�gina
        rengl�n = rtfLectorLibros.GetLineFromChar(rtfLectorLibros.SelStart) + 1
        Decir "saltando hacia adelante al rengl�n " + Str(rengl�n)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyPageUp Then 'tecla retroceso de p�gina
        rengl�n = rtfLectorLibros.GetLineFromChar(rtfLectorLibros.SelStart) + 1
        Decir "saltando hacia atr�s al rengl�n " + Str(rengl�n)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyRight Then 'avanzar de a caracter
        Decir decirLetraSiguiente(rtfLectorLibros)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyLeft Then 'retroceder de a caracter
        If rtfLectorLibros.SelStart = 0 And rtfLectorLibros.Text <> "" Then
            Decir "Est�s en el principio del libro, delante de la letra " + decirLetraSiguiente(rtfLectorLibros)
        Else
            Decir decirLetraSiguiente(rtfLectorLibros)
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyDown Then 'leer por oraci�n
        Decir decirOraci�nSiguiente(rtfLectorLibros)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyUp Then 'leer por oraci�n
        Decir decirOraci�nSiguiente(rtfLectorLibros)
    End If
    
    Dim TeclaShift, TeclaControl
    TeclaShift = (Shift And vbShiftMask) > 0
    TeclaControl = (Shift And vbCtrlMask) > 0

    Dim teclaApretada As Byte, control As Boolean, shift2 As Boolean
    Select Case KeyCode 'para evaluar la selecci�n m�s abajo
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
            teclaApretada = tecla.avanceP�gina
        Case vbKeyPageDown
            teclaApretada = tecla.retrocesoP�gina
        Case vbKeyHome
            teclaApretada = tecla.inicio
        Case vbKeyEnd
            teclaApretada = tecla.fin
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

    Call evaluarSelecci�n(rtfLectorLibros, control, shift2, teclaApretada) 'se ve si hay selecci�n
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyUp Then 'retroceder de a p�rrafo
        If rtfLectorLibros.Text <> "" Then
            rengl�n = rtfLectorLibros.GetLineFromChar(rtfLectorLibros.SelStart) + 1
            If rengl�n = 1 Then
                Decir "principio de la hoja, rengl�n 1"
            Else
                Decir "retrocediendo un p�rrafo. rengl�n " + Str(rengl�n)
            End If
        Else
            Decir "No se puede retroceder de a p�rrafo porque la hoja del libro est� vac�a"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyDown Then 'avanzar de a p�rrafo
        If rtfLectorLibros.Text <> "" Then
            rengl�n = rtfLectorLibros.GetLineFromChar(rtfLectorLibros.SelStart) + 1
            If rtfLectorLibros.GetLineFromChar(Len(rtfLectorLibros.Text)) + 1 = rengl�n Then
                Decir "final del libro. rengl�n " + Str(rengl�n)
            Else
                Decir "avanzando un p�rrafo. rengl�n " + Str(rengl�n)
            End If
        Else
            Decir "No se puede avanzar de a p�rrafo porque el libro est� vac�o"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copiar
        If rtfLectorLibros.SelText <> "" Then
            Decir "se copi� el texto seleccionado. para pegarlo en otro lugar, usar control m�s ve corta"
        Else
            Decir "No se puede copiar porque no hay texto seleccionado. para seleccionar, usar shift m�s las flechas"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyX Then 'cortar
        If rtfLectorLibros.SelText <> "" Then
            Decir "se cort� el texto seleccionado. para pegarlo en otro lugar, usar control m�s ve corta"
        Else
            Decir "No se puede cortar porque no hay texto seleccionado. para seleccionar, usar shift m�s las flechas"
        End If
    End If
        
    If shiftkey = vbAltMask And KeyCode = vbKeyDown Then 'leer todo el texto
        If Trim(rtfLectorLibros.Text) <> "" Then
            Decir rtfLectorLibros.Text
        Else
            Decir "No se puede leer todo el texto porque la hoja est� vac�a"
        End If
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyUp Then 'leer desde el cursor hacia adelante
        If Trim(rtfLectorLibros.Text) <> "" Then
            If rtfLectorLibros.SelStart = 0 Then
                Decir Mid(rtfLectorLibros.Text, 1, Len(rtfLectorLibros.Text) - Len(Left(rtfLectorLibros.Text, rtfLectorLibros.SelStart))) 'leer desde el cursor hacia adelante
            Else
                Decir Mid(rtfLectorLibros.Text, rtfLectorLibros.SelStart, Len(rtfLectorLibros.Text) - Len(Left(rtfLectorLibros.Text, rtfLectorLibros.SelStart))) 'leer desde el cursor hacia adelante
            End If
        Else
            Decir "No se puede leer todo el texto porque el libro est� vac�o"
        End If
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyRight Then 'leer el texto seleccionado
        If Trim(rtfLectorLibros.SelText) <> "" Then
            If rtfLectorLibros.SelText = " " Then
                Decir "texto seleccionado: espacio"
            Else
                Decir "texto seleccionado: " + rtfLectorLibros.SelText
            End If
        Else
            Decir "No se puede leer la selecci�n porque no hay texto seleccionado"
        End If
    End If
End Sub

Public Sub cargarLibro()
    Dim libro As String, cap�tulo As String
    libro = Left(archivoParaLeer, InStr(archivoParaLeer, "\") - 1)
    cap�tulo = Right(archivoParaLeer, Len(archivoParaLeer) - InStrRev(archivoParaLeer, "\"))
    cap�tulo = Left(cap�tulo, Len(cap�tulo) - 4) 'se le saca el .rtf
    Me.Caption = "Libro " + libro + ", cap�tulo " + cap�tulo
    
    swLibroAbierto = True
    sonido = sndPlaySound(App.path + "\sonidos\abrir.wav", SND_ASYNC)
    
    Decir "abriste el cap�tulo de nombre " + cap�tulo + ", del libro " + libro + ". pod�s leerlo con las flechas. " 'para pasar a tu carpeta de " + miMateria + ", apret� f3"
    rtfLectorLibros.LoadFile App.path + dirTrabajo + "libros\" + archivoParaLeer
    
    'se actualizan las fuentes y los colores de los rtf del programa
    rtfLectorLibros.Font.Name = NombreFuente  'se ajusta la fuente del programa
    rtfLectorLibros.AutoVerbMenu = True
    'Selecciona todo
    rtfLectorLibros.SelStart = 0
    rtfLectorLibros.SelLength = Len(rtfLectorLibros)
    rtfLectorLibros.SelColor = colorFuente 'se ajusta el color de la fuente del programa
    rtfLectorLibros.SelLength = 0
    rtfLectorLibros.Font.Size = tama�oFuente 'se ajusta el tama�o de la fuente
    rtfLectorLibros.BackColor = colorFondo 'el color de fondo del rtf
End Sub
