VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmLectorEvaluaciones 
   Caption         =   "Evaluaci�n del X/X/X"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   Icon            =   "frmLectorEvaluaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmLectorEvaluaciones.frx":08CA
   ScaleHeight     =   7560
   ScaleWidth      =   7950
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   3248
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog di�logo 
      Left            =   4320
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfLectorEvaluaciones 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11456
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmLectorEvaluaciones.frx":40B9
   End
   Begin TransparentButton.ButtonTransparent btnImprimir 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Caption         =   "Imprimir"
      EstiloDelBoton  =   4
      Picture         =   "frmLectorEvaluaciones.frx":413B
      PictureHover    =   "frmLectorEvaluaciones.frx":4A15
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
   Begin TransparentButton.ButtonTransparent btnGuardar 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Guardar"
      EstiloDelBoton  =   4
      Picture         =   "frmLectorEvaluaciones.frx":52EF
      PictureHover    =   "frmLectorEvaluaciones.frx":5BC9
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
      Left            =   5520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "Llevarse la evaluaci�n"
      EstiloDelBoton  =   1
      Picture         =   "frmLectorEvaluaciones.frx":64A3
      PictureHover    =   "frmLectorEvaluaciones.frx":6D7D
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para el docente:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6090
      TabIndex        =   4
      Top             =   120
      Width           =   1170
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   5520
      X2              =   6000
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   7320
      X2              =   7800
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frmLectorEvaluaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public swMateriaParaAbrir As String
'Public swA�oParaAbrir As Integer
Public swNumMesParaAbrir As Byte
Public swArchivoParaLeer As String
Public swS�loLeer As Boolean 'para ver si es para hacer una evaluaci�n nueva, o leer una ya hecha
Public swD�aParaAbrir As Byte
Public swEstoyAbierto As Boolean
Dim ImpresoraRich As ImpresoraRTF 'la impresora del rtf
Dim swHuboCambio As Boolean
Dim controlPresionado As Boolean 'para mandar a Keascii si se est� apretando control
Dim swAbriendoEvaluaci�n As Boolean
Dim corrector As corrector_ortograf�a
Dim swListaCorrecci�nVisible As Boolean 'para ver si con escape se saca el cuaderno o s�lo la lista

Private Sub List1_DblClick()
    Call corregirConPalabraSeleccionada(rtfLectorEvaluaciones, List1.List(List1.ListIndex))
    Call List1_KeyDown(vbKeyEscape, 0)
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        List1.Visible = False
        swListaCorrecci�nVisible = False
        Decir "cerrando las sugerencias, est�s otra vez en tu evaluaci�n"
    End If
    
    If KeyCode = vbKeyReturn Then List1_DblClick
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        Decir List1.List(List1.ListIndex) + ". " + separarEnLetras(List1.List(List1.ListIndex))
    End If
End Sub

Private Sub ImprimirConCuadroDi�logo()
    di�logo.CancelError = True
    On Error GoTo manejoErrorImpresora
    di�logo.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If rtfLectorEvaluaciones.SelLength = 0 Then
       di�logo.Flags = di�logo.Flags + cdlPDAllPages
    Else
       di�logo.Flags = di�logo.Flags + cdlPDSelection
    End If
    di�logo.ShowPrinter
    rtfLectorEvaluaciones.SelPrint di�logo.hDC
manejoErrorImpresora:
    Exit Sub
End Sub

Private Sub btnImprimir_Click()
    Call Form_KeyDown(vbKeyF6, 0)
End Sub

Private Sub ButtonTransparent1_Click()
    di�logo.CancelError = True
    On Error GoTo manejoError
    di�logo.Filter = "Archivo de texto (*.rtf)|*.rtf"
    di�logo.ShowSave
    rtfLectorEvaluaciones.SaveFile di�logo.FileName
    Exit Sub
manejoError:
    Exit Sub ' El usuario ha hecho clic en el bot�n Cancelar
End Sub

Private Sub Form_Activate()
    'se actualizan las fuentes y los colores de los rtf del programa
    rtfLectorEvaluaciones.Font.Name = NombreFuente  'se ajusta la fuente del programa
    rtfLectorEvaluaciones.SelStart = 0 'Selecciona todo
    rtfLectorEvaluaciones.SelLength = Len(rtfLectorEvaluaciones)
    rtfLectorEvaluaciones.SelColor = colorFuente 'se ajusta el color de la fuente del programa
    rtfLectorEvaluaciones.SelLength = 0
    rtfLectorEvaluaciones.Font.Size = tama�oFuente 'se ajusta el tama�o de la fuente
    rtfLectorEvaluaciones.BackColor = colorFondo 'el color de fondo del rtf
End Sub

Private Sub Form_GotFocus()
    If List1.Visible = True Then
        List1.Visible = False
        swListaCorrecci�nVisible = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7

    Dim tecla(0 To 255) As Byte, EstadoNumLock As Boolean, EstadoMay�sculas As Boolean, EstadoScroll As Boolean
    GetKeyboardState tecla(0)
    EstadoNumLock = tecla(VK_NUMLOCK)
    EstadoMay�sculas = tecla(VK_CAPITAL)
    EstadoScroll = tecla(VK_SCROLL)

    If KeyCode = 20 Then 'si se presiona el boqueador de may�sculas
        If EstadoMay�sculas Then
            Decir "may�sculas activado"
        Else
            Decir "may�sculas desactivado"
        End If
    End If

    If KeyCode = 144 Then 'si se presiona el boqueador de n�meros
        If EstadoNumLock Then
            Decir "teclado num�rico activado"
        Else
            Decir "teclado num�rico desactivado"
        End If
    End If

    If KeyCode = 145 Then 'si se presiona el boqueador de desplazamiento
        If EstadoScroll Then
            Decir "bloqueador de desplazamiento activado"
        Else
            Decir "bloqueador de desplazamiento desactivado"
        End If
    End If

End Sub

'Private Sub Form_Paint()
'    'se actualizan las fuentes y los colores de los rtf del programa
'    rtfLectorEvaluaciones.Font.Name = NombreFuente  'se ajusta la fuente del programa
'    rtfLectorEvaluaciones.SelStart = 0 'Selecciona todo
'    rtfLectorEvaluaciones.SelLength = Len(rtfLectorEvaluaciones)
'    rtfLectorEvaluaciones.SelColor = colorFuente 'se ajusta el color de la fuente del programa
'    rtfLectorEvaluaciones.SelLength = 0
'    rtfLectorEvaluaciones.Font.Size = tama�oFuente 'se ajusta el tama�o de la fuente
'    rtfLectorEvaluaciones.BackColor = colorFondo 'el color de fondo del rtf
'End Sub
'
Private Sub rtfLectorEvaluaciones_Change()
    If swAbriendoEvaluaci�n = False Then
        swHuboCambio = True
    End If
    swAbriendoEvaluaci�n = False
End Sub

Private Sub rtfLectorEvaluaciones_GotFocus()
    Call reproducirForm(formularios.evaluaciones)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7

    If KeyCode = 222 Then Decir "acento agudo"

    If KeyCode = vbKeyEscape Then
        If swListaCorrecci�nVisible = False Then Unload Me
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir Trim(nombreUsuario) + ", para abrir o ir al reproductor de m�sica, ten�s que estar en el men� principal o en una carpeta. ahora est�s en una evaluaci�n"
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    If shiftkey = 0 And KeyCode = vbKeyF4 Then frmAccesorios.Show

    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa

    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.lectorEvaluaciones
         frmAyuda.Show
         Exit Sub
    End If

    Dim i As Integer, prefijo As String, contador As Integer ', extensi�n As String
    If Me.swS�loLeer = False Then
        If (shiftkey = vbCtrlMask And KeyCode = vbKeyG) Or KeyCode = vbKeyF5 Then 'se guarda el archivo con control + G � con f5
            If rtfLectorEvaluaciones.Text <> "" Then
                If swHuboCambio = True Then 'si hay cambios no guardados
                    If InStrRev(swArchivoParaLeer, ".", , vbTextCompare) = 11 Then 'se eval�a si tiene el archivo el .dll para sac�rselo en el nombre a guardar
                        swArchivoParaLeer = Left(swArchivoParaLeer, Len(swArchivoParaLeer) - 4)
                    End If
    
                    'se guarda la eval disfrazada como dll
                    rtfLectorEvaluaciones.SaveFile App.path + "\trabajos\" + swMateriaParaAbrir + "\soporte\" + Trim(Str(swNumMesParaAbrir)) + "\" + swArchivoParaLeer + ".dll"
                    'se guarda una copia falsa por si los pap�s quieren modificarla externamente
                    rtfLectorEvaluaciones.SaveFile App.path + "\trabajos\" + swMateriaParaAbrir + "\evaluaciones\" + Trim(Str(swNumMesParaAbrir)) + "\" + swArchivoParaLeer + ".rtf"
                    'se guarda el t�tulo de la evaluaci�n
                    frmT�tuloEvaluaci�n.nombreArchivo = App.path + "\trabajos\" + swMateriaParaAbrir + "\soporte\" + Trim(Str(swNumMesParaAbrir)) + "\datosSoporte\" + swArchivoParaLeer + ".gui"
    
                    If Not existeCarpeta(frmT�tuloEvaluaci�n.nombreArchivo) Then frmT�tuloEvaluaci�n.Show 1  'se ofrece guardar la hoja si a�n no se lo ha hecho
    '                Call chequearEspacioEnDisco(Left(App.Path, 2))
                    swHuboCambio = False 'se establece que no hay cambios sin guardar
                    If swHablarVoz = True Then
                        Decir "tu evaluaci�n est� guardada"
                    Else
                        frmMsgBox.cadenaAMostrar = "Tu evaluaci�n est� guardada"
                        frmMsgBox.swS�No�Aceptar = False 'se le dice que es un msg aceptar
                        frmMsgBox.Show 1
                    End If
                Else
                    If swHablarVoz = True Then
                        Decir "No has hecho cambios en tu evaluaci�n, no hay nada nuevo para guardar"
                    Else
                        frmMsgBox.cadenaAMostrar = "No has hecho cambios en tu evaluaci�n, no hay nada nuevo para guardar"
                        frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
                        frmMsgBox.Show 1
                    End If
                End If
            Else
                If swHablarVoz = True Then
                    Decir "La evaluaci�n est� vac�a, para guardar una evaluaci�n hay que escribir algo en ella"
                Else
                    frmMsgBox.cadenaAMostrar = "La hoja de tu evaluaci�n est� vac�a, para guardar una evaluaci�n hay que escribir algo en ella"
                    frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
                    frmMsgBox.Show 1
                End If
            End If
            Exit Sub
        End If
    
    
        If (shiftkey = vbCtrlMask And KeyCode = vbKeyN) Then 'control n para negrita
            If rtfLectorEvaluaciones.SelBold = False Then 'estoyEnNegrita = False Then
                rtfLectorEvaluaciones.SelBold = True
                Decir "negrita activado"
                'estoyEnNegrita = True
            Else
                rtfLectorEvaluaciones.SelBold = False
                Decir "negrita desactivado"
                'estoyennegrita=False
            End If
            Exit Sub
        End If
    
        If (shiftkey = vbCtrlMask And KeyCode = vbKeyS) Then 'control s para subrayar
            If rtfLectorEvaluaciones.SelUnderline = False Then   'estoyEnNegrita = False Then
                rtfLectorEvaluaciones.SelUnderline = True
                Decir "subrayado activado"
                'estoyEnNegrita = True
            Else
                rtfLectorEvaluaciones.SelUnderline = False
                Decir "subrayado desactivado"
                'estoyennegrita=False
            End If
            Exit Sub
        End If
    Else
        Decir "No se puede escribir en una evaluaci�n guardada, s�lo se puede leer"
    End If

    If shiftkey = 0 And KeyCode = vbKeyF8 Then 'f8, calculadora
        If frmCalculadora.swEstoyAbierto = True Then
            Decir "Pasando a la calculadora, para volver a tu evaluaci�n, apret� F8"
            frmCalculadora.Show
        Else
            frmCalculadora.Show
        End If
        Exit Sub
    End If

    If (shiftkey = vbCtrlMask And KeyCode = vbKeyP) Or KeyCode = vbKeyF6 Then 'imprimir con control + p
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "�Realmente quer�s imprimir esta evaluaci�n?"
        frmMsgBox.swS�No�Aceptar = True 'se elige que sea cuadro s�-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado = True Then
            If swImprimirDirecto = True Then
                With ImpresoraRich

                     'Valores
                     'Encabezado y pie de p�gina
                     .Header = "Evaluaci�n de la carpeta de " + swMateriaParaAbrir + " trabajada el d�a " + Format(Date, "dd/mm/yyyy")
                     .Footer = "Evaluaci�n realizada por " + Trim(nombreUsuario)

                     'Margenes
                     .MarginTop = 500 'Text3
                     .MarginLeft = 500 'Text4
                     .MarginRight = 500 'Text5
                     .MarginBottom = 500 'Text6

                     'Imprimir el RichTextBox pasado como par�metro
                     .Imprimir rtfLectorEvaluaciones
                End With
            Else
                ImprimirConCuadroDi�logo 'se muestra el cuadro de di�logo de la impresora
            End If
        End If
        Exit Sub
   End If

    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
End Sub

Private Sub Form_Load()
    Dim i As Integer, contador As Integer, prefijo As String, temp As String
    Call centrarFormulario(Me)
    Set corrector = New corrector_ortograf�a
    'si no est� aspell, se carga el diccionario propio
    If swAspellInstalado = False Then corrector.Cargar_Diccionario (App.path + "\datos\diccionario.txt")

    'Call contarFormularios(True)
    swEstoyAbierto = True
    swHuboCambio = False
    sonido = sndPlaySound(App.path + "\sonidos\abrir.wav", SND_ASYNC)
    If swS�loLeer = True Then
        rtfLectorEvaluaciones.Locked = True
        btnGuardar.Visible = False
        rtfLectorEvaluaciones.LoadFile App.path + "\trabajos\" + swMateriaParaAbrir + "\soporte\" + Trim(Str(swNumMesParaAbrir)) + "\" + swArchivoParaLeer '+ ".dll"
        swAbriendoEvaluaci�n = True
        Decir Trim(nombreUsuario) + ", abriste la evaluaci�n de la materia " + swMateriaParaAbrir + " del d�a " + Str(swD�aParaAbrir) + ". para leerla, us� las flechas"
        Me.Caption = "Evaluaci�n de la materia " + Chr(34) + swMateriaParaAbrir + Chr(34) + " del d�a " + Trim(Str(swD�aParaAbrir)) + " de " + decodificarMes(swNumMesParaAbrir)
    Else
        btnGuardar.Visible = True
        rtfLectorEvaluaciones.Locked = False
        Decir "empezando una evaluaci�n de la materia " + swMateriaParaAbrir + ". pod�s escribir en ella"
        Me.Caption = "Evaluaci�n de la materia " + Chr(34) + swMateriaParaAbrir + Chr(34) + " del d�a de hoy, " + Trim(Str(swD�aParaAbrir)) + " de " + decodificarMes(swNumMesParaAbrir)
        temp = Left(swArchivoParaLeer, InStr(swArchivoParaLeer, "-") - 1)
        If Len(Trim(Str(temp))) = 1 Then swArchivoParaLeer = "0" + swArchivoParaLeer
        For i = 1 To cantPrefijo 'se le agrega el prefijo al archivo
            swArchivoParaLeer = "0" + swArchivoParaLeer
        Next
        
        Do While existeCarpeta(App.path + "\trabajos\" + swMateriaParaAbrir + "\soporte\" + Trim(Str(swNumMesParaAbrir)) + "\" + swArchivoParaLeer + ".dll")
            swArchivoParaLeer = Right(swArchivoParaLeer, Len(swArchivoParaLeer) - cantPrefijo)
            contador = contador + 1
                    
            If contador < 10 Then prefijo = "10" + Trim(Str(contador))
            If contador >= 10 And contador < 100 Then prefijo = "1" + Str(contador)
            If contador >= 100 Then prefijo = Str(contador)
            
            swArchivoParaLeer = prefijo + swArchivoParaLeer
        Loop
    End If

    If swAspellInstalado = True Then 'si est� aspell, se deja listo el pipe para comunicarse con �l
        Call objPipe.Execute("CMD.EXE")
        Call Sleep(200)
        Call objPipe.Write_("c:" & vbCrLf)
        Call objPipe.Write_("cd " & rutaDeAspell & vbCrLf)
        Call Sleep(100)
        Call objPipe.Write_("aspell -a -d " & idiomaAspell & vbCrLf)
    End If
    'se actualizan las fuentes y los colores de los rtf del programa
    rtfLectorEvaluaciones.Font.Name = NombreFuente  'se ajusta la fuente del programa
    'rtfLectorEvaluaciones.AutoVerbMenu = True
    'Selecciona todo
    rtfLectorEvaluaciones.SelStart = 0
    rtfLectorEvaluaciones.SelLength = Len(rtfLectorEvaluaciones)
    rtfLectorEvaluaciones.SelColor = colorFuente 'se ajusta el color de la fuente del programa
    rtfLectorEvaluaciones.SelLength = 0
    rtfLectorEvaluaciones.Font.Size = tama�oFuente 'se ajusta el tama�o de la fuente
    rtfLectorEvaluaciones.BackColor = colorFondo 'el color de fondo del rtf
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
    
    If swS�loLeer = False Then 'si est� escribiendo una evaluaci�n
        If mensajeSalir("�Est�s seguro que quer�s cerrar la evaluaci�n, una vez cerrada ya no vas a poder modificar lo que has escrito en ella?") Then
            Call Form_KeyDown(vbKeyF5, 0)
            frmPrincipal.Show
            swEstoyAbierto = False
        Else
            Cancel = 1
        End If
    Else 'si est� revisando una evalaci�n a hecha
        frmPrincipal.Show
        swEstoyAbierto = False
    End If

    'swS�loLeer = False
End Sub

Private Sub rtfLectorEvaluaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim auxString As String, caracterAnteriorBorrado As String, letra As String
    Dim swEnterEnMedioDelRengl�n As Boolean, shiftkey As Integer, rengl�n As Long
    Dim palabra As String, temp() As String

    shiftkey = Shift And 7
    
    If KeyCode = vbKeyInsert Then KeyCode = 0 'si aprieta insert, se neutraliza as� no activa la sobreescritura
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el men� de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyReturn Or KeyCode = vbKeyBack Or KeyCode = 93 Then
        If swS�loLeer = False Then
            If KeyCode = 13 Then
                swEnterEnMedioDelRengl�n = medioDelRengl�n(rtfLectorEvaluaciones)
                If swEnterEnMedioDelRengl�n = False Then
                    Decir "bajada de l�nea. rengl�n " + Trim(Str(rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 2))
                Else
                    Decir "est�s haciendo una bajada de l�nea sin estar al final del rengl�n, si es un error pod�s corregirlo borrando la bajada de l�nea, yendo al final del rengl�n con la tecla fin, y ah� s� hacer la bajada de l�nea"
                End If
            End If
            
            If KeyCode = 93 Then 'si es el men� contextual
                palabra = buscarPalabraParaCorregir(rtfLectorEvaluaciones)
                If palabra = "" Then
                    Decir "no est�s en ninguna palabra, no puedo corregir"
                    KeyCode = 0
                Else
                    If Not corregirPalabra(palabra) Then 'si la palabra es incorrecta
                        If palabra <> "" Then
                            If swAspellInstalado = True Then 'si aspell est� instalado, se lo prefiere
                                '*************************
                                'corregir con aspell
                                Call objPipe.Write_(palabra & vbCrLf)
                                Call Sleep(200)
                                temp = arreglarCadena(objPipe.Read)
                            Else
                                '*****************************
                                'corregir con mi propio corrector (m�s lento)
                                temp = corrector.Controlar_Un_Error(palabra)
                            End If
                            Decir "us� flecha abajo para ver las palabras que te sugiero para corregir " + palabra
                            swListaCorrecci�nVisible = True
                            Call Cargar_Men�_En_Lista(List1, temp)
        '                    Call Cargar_Menu(mnuPalabras, temp)
        '                    Me.PopupMenu mnuContextual, 0, rtfLectorEvaluaciones.Width / 2, rtfLectorEvaluaciones.Top + 20
                        End If
                    Else
                        Decir "la palabra es correcta, no hay sugerencias"
                        KeyCode = 0
                    End If
                End If
            End If
        
            If KeyCode = vbKeyDelete Then
                If rtfLectorEvaluaciones.Text <> "" Then 'si no est� vac�o
                    If rtfLectorEvaluaciones.SelStart <> Len(rtfLectorEvaluaciones.Text) Then 'y no est� al final de la hoja
                        letra = Mid(rtfLectorEvaluaciones.Text, rtfLectorEvaluaciones.SelStart + 1, 1)
                        If letra = " " Then
                            Decir "borrando a la derecha el espacio", False
                        ElseIf letra = Chr(9) Then
                            Decir "borrando a la derecha un salto"
                        ElseIf letra = Chr(10) Or letra = Chr(13) Then
                            Decir "borrando a la derecha la bajada de l�nea. rengl�n " + Str(rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1), False
                        Else
                            auxString = traducirParaBorrar(letra)
                            Decir "borrando a la derecha " + auxString
                        End If
                    Else
                        Decir "imposible borrar, est�s al final de la hoja"
                    End If
                Else
                    Decir "no se puede borrar a la derecha porque la hoja est� vac�a"
                End If
            End If
        
            If KeyCode = vbKeyBack Then
                If rtfLectorEvaluaciones.Text = "" Then
                    Decir "Ya est� todo borrado"
                Else
                    If rtfLectorEvaluaciones.SelStart = 0 Then
                        Decir "imposible borrar porque est�s al principio de la hoja"
                    Else
                        If rtfLectorEvaluaciones.SelText = "" Then 'si no hay nada seleccionado
                            caracterAnteriorBorrado = Mid(rtfLectorEvaluaciones.Text, rtfLectorEvaluaciones.SelStart, 1)
                        Else
                            Decir "borrando el texto seleccionado"
                            Exit Sub
                        End If
            
                        If caracterAnteriorBorrado = " " Then
                            Decir "borrando el espacio", False
                        ElseIf caracterAnteriorBorrado = Chr(9) Then
                                Decir "borrando un salto"
                        ElseIf caracterAnteriorBorrado = Chr(10) Then
                            Decir "borrando la bajada de l�nea. rengl�n " + Str(rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart)), False
                        Else
                            If caracterAnteriorBorrado <> "" Then
                                auxString = traducirParaBorrar(caracterAnteriorBorrado)
                                Decir "borrando " + auxString, True, False
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Decir Trim(nombreUsuario) + ", no pod�s borrar en una evaluaci�n guardada, s�lo leer lo que est� escrito"
        End If
    End If

'    If shiftkey = vbCtrlMask And KeyCode = vbKeyHome Then 'control + inicio
'        If rtfLectorEvaluaciones <> "" Then
'            Decir decirPalabraSiguiente(rtfLectorEvaluaciones)
'        Else
'            Decir "La hoja est� en blanco, no hay nada escrito"
'        End If
'    End If

    If shiftkey = 0 And KeyCode = vbKeyHome Then 'tecla inicio
        rengl�n = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
        Decir "principio del rengl�n " + Str(rengl�n)
    End If

'    If shiftkey = vbCtrlMask And KeyCode = vbKeyEnd Then 'control + fin
''        If swVolviendodeBraille = False Then 'si no se dispara el evento al volver del teclado braille
'            If rtfLectorEvaluaciones <> "" Then
'                Decir "final de la hoja. Est�s detr�s de la palabra " + decirPalabraAnterior(rtfLectorEvaluaciones)
'            Else
'                Decir "La hoja est� en blanco, no hay nada escrito"
'            End If
''        End If
'    End If

    If shiftkey = 0 And KeyCode = vbKeyEnd Then 'tecla fin
        rengl�n = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
        Decir "final del rengl�n " + Str(rengl�n)
    End If

    If shiftkey = vbCtrlMask Then controlPresionado = True

End Sub

Private Sub rtfLectorEvaluaciones_KeyPress(KeyAscii As Integer)
    Dim cadena As String
    If swS�loLeer = False Then
        swHuboCambio = True
        If KeyAscii >= 32 And KeyAscii <= 255 And controlPresionado = False Then cadena = qu�LetraSeApret�(KeyAscii)

        If KeyAscii = 9 Then cadena = "salto hacia adelante" 'tab
        If KeyAscii = 39 Then cadena = "ap�strofo"
        If KeyAscii = 123 Then cadena = "abre llave"
        If KeyAscii = 125 Then cadena = "cierra llave"
        If KeyAscii = 91 Then cadena = "abre corchete"
        If KeyAscii = 93 Then cadena = "cierra corchete"
        If KeyAscii = 64 Then cadena = "arroba"

        'leer la palabra al apretar espacio, punto, coma, etc.
        If KeyAscii = 32 Or KeyAscii = Asc(".") Or KeyAscii = Asc(",") Or KeyAscii = Asc(";") Or KeyAscii = Asc(":") _
        Or KeyAscii = Asc("-") Then
            cadena = cadena + decirPalabraAnterior(rtfLectorEvaluaciones)
            If swUsarCorrectorOrtogr�fico = True Then 'si se usa el corrector, se dice si es incorrecta
                If Not corregirPalabra(buscarPalabraParaCorregir(rtfLectorEvaluaciones)) Then cadena = cadena + ", incorrecta"
            End If
        End If
        
        If cadena <> "" Then
            If rtfLectorEvaluaciones.SelBold = True Then cadena = cadena + " en negrita"
            If rtfLectorEvaluaciones.SelUnderline = True Then cadena = cadena + " subrayada"
            Decir cadena
        End If
    
        controlPresionado = False 'se resetea la variable
    Else
        If KeyAscii = vbKeyBack Then
            Decir Trim(nombreUsuario) + ", no pod�s borrar en una evaluaci�n guardada, s�lo leer lo que est� escrito"
        Else
            Decir "no pod�s escribir en una evaluaci�n guardada, s�lo leer lo que est� escrito"
        End If
    End If
End Sub

Private Sub rtfLectorEvaluaciones_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    Dim rengl�n As Long, cadena As String

    shiftkey = Shift And 7

    If shiftkey = vbCtrlMask And KeyCode = vbKeyLeft Then 'leer por palabras retrocediendo
        cadena = decirPalabraSiguiente(rtfLectorEvaluaciones)
        If Not esSigno(cadena) Then 'se ve si la cadena es solamente un signo ortogr�fico
            If Not corregirPalabra(cadena) Then cadena = cadena + ", incorrecta"
        End If
        Decir cadena
    End If

    If shiftkey = vbCtrlMask And KeyCode = vbKeyRight Then 'leer por palabras avanzando
        cadena = decirPalabraSiguiente(rtfLectorEvaluaciones)
        If Not esSigno(cadena) Then 'se ve si la cadena es solamente un signo ortogr�fico
            If Not corregirPalabra(cadena) Then cadena = cadena + ", incorrecta"
        End If
        Decir cadena
    End If

    If shiftkey = 0 And KeyCode = vbKeyRight Then 'avanzar de a caracter
        Decir decirLetraSiguiente(rtfLectorEvaluaciones)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyPageDown Then 'tecla avance de p�gina
        rengl�n = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
        Decir "saltando hacia adelante al rengl�n " + Str(rengl�n)
    End If

    If shiftkey = 0 And KeyCode = vbKeyPageUp Then 'tecla retroceso de p�gina
        rengl�n = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
        Decir "saltando hacia atr�s al rengl�n " + Str(rengl�n)
    End If

    Dim auxString As String
    If shiftkey = vbCtrlMask And KeyCode = vbKeyHome Then 'control + inicio
        If rtfLectorEvaluaciones.Text <> "" Then
            auxString = decirPalabraSiguiente(rtfLectorEvaluaciones)
            If Trim(auxString) <> Chr(10) And Trim(auxString) <> Chr(13) Then
                Decir "principio de la hoja." + auxString
            Else
                Decir "principio de la hoja. rengl�n en blanco"
            End If
        Else
            Decir "La hoja est� en blanco, no hay nada escrito"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyEnd Then 'control + fin
        If rtfLectorEvaluaciones.Text <> "" Then
            auxString = decirPalabraAnterior(rtfLectorEvaluaciones)
            If Trim(auxString) <> Chr(10) And Trim(auxString) <> Chr(13) Then
                If Len(Trim(auxString)) <> 0 Then
                    Decir "final de la hoja. Est�s detr�s de la palabra " + decirPalabraAnterior(rtfLectorEvaluaciones)
                Else
                    Decir "final de la hoja. s�lo hay escrito espacios en este rengl�n, ninguna letra"
                End If
            Else
                Decir "final de la hoja. rengl�n en blanco"
            End If
        Else
            Decir "La hoja est� en blanco, no hay nada escrito"
        End If
    End If

    If shiftkey = 0 And KeyCode = vbKeyLeft Then 'retroceder de a caracter
        If rtfLectorEvaluaciones.SelStart = 0 And rtfLectorEvaluaciones.Text <> "" Then
            Decir "Est�s en el principio de la hoja, delante de la letra " + decirLetraSiguiente(rtfLectorEvaluaciones)
        Else
            Decir decirLetraSiguiente(rtfLectorEvaluaciones)
        End If
    End If

    If shiftkey = 0 And KeyCode = vbKeyDown Then 'leer por oraci�n
        Decir decirOraci�nSiguiente(rtfLectorEvaluaciones)
    End If

    If shiftkey = 0 And KeyCode = vbKeyUp Then 'leer por oraci�n
        Decir decirOraci�nSiguiente(rtfLectorEvaluaciones)
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

    Call evaluarSelecci�n(rtfLectorEvaluaciones, control, shift2, teclaApretada) 'se ve si hay selecci�n
    

    If shiftkey = vbCtrlMask And KeyCode = vbKeyUp Then 'retroceder de a p�rrafo
        If rtfLectorEvaluaciones.Text <> "" Then
            rengl�n = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
            If rengl�n = 1 Then
                Decir "principio de la hoja, rengl�n 1"
            Else
                Decir "retrocediendo un p�rrafo. rengl�n " + Str(rengl�n)
            End If
        Else
            Decir "No se puede retroceder de a p�rrafo porque la hoja est� vac�a"
        End If
    End If

    If shiftkey = vbCtrlMask And KeyCode = vbKeyDown Then 'avanzar de a p�rrafo
        If rtfLectorEvaluaciones.Text <> "" Then
            rengl�n = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
            If rtfLectorEvaluaciones.GetLineFromChar(Len(rtfLectorEvaluaciones.Text)) + 1 = rengl�n Then
                Decir "final de la hoja. rengl�n " + Str(rengl�n)
            Else
                Decir "avanzando un p�rrafo. rengl�n " + Str(rengl�n)
            End If
        Else
            Decir "No se puede avanzar de a p�rrafo porque la hoja est� vac�a"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyZ Then 'deshacer
        Decir "deshaciendo la �ltima acci�n"
    End If

    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copiar
        If rtfLectorEvaluaciones.SelText <> "" Then
            Decir "se copi� el texto seleccionado. para pegarlo en otro lugar, usar control m�s ve corta"
        Else
            Decir "No se puede copiar porque no hay texto seleccionado. para seleccionar, usar shift m�s las flechas"
        End If
    End If

    If shiftkey = vbCtrlMask And KeyCode = vbKeyX Then 'cortar
        If swS�loLeer = False Then
            If rtfLectorEvaluaciones.SelText <> "" Then
                Decir "se cort� el texto seleccionado. para pegarlo en otro lugar, usar control m�s ve corta"
            Else
                Decir "No se puede cortar porque no hay texto seleccionado. para seleccionar, usar shift m�s las flechas"
            End If
        Else
            Decir "no pod�s cortar en una evaluaci�n guardada, s�lo leer lo que est� escrito"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyV Then 'pegar
        If swS�loLeer = False Then
            If Clipboard.GetText <> "" Then
                rengl�n = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
                Decir "texto pegado en el rengl�n " + Str(rengl�n)
            Else
                Decir "No se puede pegar porque no hay texto copiado o cortado. para copiar, usar control m�s c. para cortar, usar control m�s x"
            End If
        Else
            Decir "no pod�s pegar nada en una evaluaci�n guardada, s�lo leer lo que est� escrito"
        End If
    End If
    
'    Dim est�PalabraEnLista As Boolean 'para el corrector ortogr�fico
'    If swUsarCorrectorOrtogr�fico = True And swS�loLeer = False Then 'el corrector ortogr�fico
'        If shiftkey = 0 And (KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = Asc(",") Or KeyCode = Asc(".") Or KeyCode = Asc("!") Or KeyCode = Asc("?") Or KeyCode = Asc("-")) Then 'con espacio se corrige la palabra reci�n escrita
'            est�PalabraEnLista = corregirPalabra(rtfLectorEvaluaciones)
'            If est�PalabraEnLista = False Then Decir "palabra con posible error"
'        End If
'    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyLeft Then 'leer la oraci�n actual
        If Trim(rtfLectorEvaluaciones.Text) <> "" Then
            Decir "El rengl�n actual dice: " + decirOraci�nSiguiente(rtfLectorEvaluaciones)
        Else
            Decir "No se puede leer el rengl�n actual porque la evaluaci�n est� vac�a"
        End If
    End If

    If shiftkey = vbAltMask And KeyCode = vbKeyDown Then 'leer todo el texto
        If Trim(rtfLectorEvaluaciones.Text) <> "" Then
            Decir "toda la evaluaci�n dice: " + rtfLectorEvaluaciones.Text
        Else
            Decir Trim(nombreUsuario) + ", No se puede leer todo el texto porque la evaluaci�n est� vac�a"
        End If
    End If

    If shiftkey = vbAltMask And KeyCode = vbKeyUp Then 'leer desde el cursor hacia adelante
        If Trim(rtfLectorEvaluaciones.Text) <> "" Then
            If rtfLectorEvaluaciones.SelStart = 0 Then
                Decir "desde donde est�s hasta el final de la evaluaci�n dice: " + Mid(rtfLectorEvaluaciones.Text, 1, Len(rtfLectorEvaluaciones.Text) - Len(Left(rtfLectorEvaluaciones.Text, rtfLectorEvaluaciones.SelStart))) 'leer desde el cursor hacia adelante
            Else
                Decir "desde donde est�s hasta el final de la evaluaci�n dice: " + Mid(rtfLectorEvaluaciones.Text, rtfLectorEvaluaciones.SelStart, Len(rtfLectorEvaluaciones.Text) - Len(Left(rtfLectorEvaluaciones.Text, rtfLectorEvaluaciones.SelStart))) 'leer desde el cursor hacia adelante
            End If
        Else
            Decir "No se puede leer todo el texto porque la evaluaci�n est� vac�a"
        End If
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyRight Then 'leer el texto seleccionado
        If Trim(rtfLectorEvaluaciones.SelText) <> "" Then
            If rtfLectorEvaluaciones.SelText = " " Then
                Decir "texto seleccionado: espacio"
            Else
                Decir "texto seleccionado: " + rtfLectorEvaluaciones.SelText
            End If
        Else
            Decir "No se puede leer la selecci�n porque no hay texto seleccionado"
        End If
    End If

End Sub

Private Function decodificarMes(numMes As Byte) As String
    Dim cadenaMes As String
    Select Case numMes
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
    decodificarMes = cadenaMes
End Function

'++++++++++++++++++++++++++++++
'para pegar im�genes
'++++++++++++++++++++++++++++++
Public Sub pegarImagen()
    Clipboard.Clear
    Clipboard.SetData Me.Picture1
    rtfLectorEvaluaciones.SetFocus
    SendKeys "^(V)"
End Sub

Private Sub Form_Resize()
    Dim Alto_RTF As Single
    Dim posX As Single
    Dim posY As Single

  ' No es necesario ajustar cuando la ventana est� minimizada
    If WindowState = vbMinimized Then
            Exit Sub
    End If
       
    'el bot�n de arriba
    posX = Me.Width - ButtonTransparent1.Width - 400
    posY = 360

    ButtonTransparent1.Move posX, posY
    
    'las l�neas y etiqueta de arriba
    Line2.X2 = Me.Width - 400
    Line2.X1 = Line2.X2 - 480
    
    Label2.Left = Line2.X1 - 20 - Label2.Width
    
    Line1.X2 = Label2.Left - 20
    Line1.X1 = Line1.X2 - 480
    
    'el bot�n de abajo
'    posX = Me.Width - ButtonTransparent1.Width - 400
'    posY = Me.Height - ButtonTransparent1.Height - 700
'
'    ButtonTransparent1.Move posX, posY
'
'    Label1.Move Label1.Left, posY

     Alto_RTF = Me.ScaleHeight - 1500

     ' Esto chequea que el valor Height del text no sea negativo _
       ya que si no da error
     If Alto_RTF <= 0 Then Alto_RTF = 100 '- RichTextBox1.Top

     'Posiciona y redimensiona el rtf
     rtfLectorEvaluaciones.Move rtfLectorEvaluaciones.Left, rtfLectorEvaluaciones.Top, Me.ScaleWidth - 250 - rtfLectorEvaluaciones.Left, Alto_RTF
 End Sub



