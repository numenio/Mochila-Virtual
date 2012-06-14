VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmLectorEvaluaciones 
   Caption         =   "Evaluación del X/X/X"
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
   Begin MSComDlg.CommonDialog diálogo 
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
      Caption         =   "Llevarse la evaluación"
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
'Public swAñoParaAbrir As Integer
Public swNumMesParaAbrir As Byte
Public swArchivoParaLeer As String
Public swSóloLeer As Boolean 'para ver si es para hacer una evaluación nueva, o leer una ya hecha
Public swDíaParaAbrir As Byte
Public swEstoyAbierto As Boolean
Dim ImpresoraRich As ImpresoraRTF 'la impresora del rtf
Dim swHuboCambio As Boolean
Dim controlPresionado As Boolean 'para mandar a Keascii si se está apretando control
Dim swAbriendoEvaluación As Boolean
Dim corrector As corrector_ortografía
Dim swListaCorrecciónVisible As Boolean 'para ver si con escape se saca el cuaderno o sólo la lista

Private Sub List1_DblClick()
    Call corregirConPalabraSeleccionada(rtfLectorEvaluaciones, List1.List(List1.ListIndex))
    Call List1_KeyDown(vbKeyEscape, 0)
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        List1.Visible = False
        swListaCorrecciónVisible = False
        Decir "cerrando las sugerencias, estás otra vez en tu evaluación"
    End If
    
    If KeyCode = vbKeyReturn Then List1_DblClick
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        Decir List1.List(List1.ListIndex) + ". " + separarEnLetras(List1.List(List1.ListIndex))
    End If
End Sub

Private Sub ImprimirConCuadroDiálogo()
    diálogo.CancelError = True
    On Error GoTo manejoErrorImpresora
    diálogo.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If rtfLectorEvaluaciones.SelLength = 0 Then
       diálogo.Flags = diálogo.Flags + cdlPDAllPages
    Else
       diálogo.Flags = diálogo.Flags + cdlPDSelection
    End If
    diálogo.ShowPrinter
    rtfLectorEvaluaciones.SelPrint diálogo.hDC
manejoErrorImpresora:
    Exit Sub
End Sub

Private Sub btnImprimir_Click()
    Call Form_KeyDown(vbKeyF6, 0)
End Sub

Private Sub ButtonTransparent1_Click()
    diálogo.CancelError = True
    On Error GoTo manejoError
    diálogo.Filter = "Archivo de texto (*.rtf)|*.rtf"
    diálogo.ShowSave
    rtfLectorEvaluaciones.SaveFile diálogo.FileName
    Exit Sub
manejoError:
    Exit Sub ' El usuario ha hecho clic en el botón Cancelar
End Sub

Private Sub Form_Activate()
    'se actualizan las fuentes y los colores de los rtf del programa
    rtfLectorEvaluaciones.Font.Name = NombreFuente  'se ajusta la fuente del programa
    rtfLectorEvaluaciones.SelStart = 0 'Selecciona todo
    rtfLectorEvaluaciones.SelLength = Len(rtfLectorEvaluaciones)
    rtfLectorEvaluaciones.SelColor = colorFuente 'se ajusta el color de la fuente del programa
    rtfLectorEvaluaciones.SelLength = 0
    rtfLectorEvaluaciones.Font.Size = tamañoFuente 'se ajusta el tamaño de la fuente
    rtfLectorEvaluaciones.BackColor = colorFondo 'el color de fondo del rtf
End Sub

Private Sub Form_GotFocus()
    If List1.Visible = True Then
        List1.Visible = False
        swListaCorrecciónVisible = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7

    Dim tecla(0 To 255) As Byte, EstadoNumLock As Boolean, EstadoMayúsculas As Boolean, EstadoScroll As Boolean
    GetKeyboardState tecla(0)
    EstadoNumLock = tecla(VK_NUMLOCK)
    EstadoMayúsculas = tecla(VK_CAPITAL)
    EstadoScroll = tecla(VK_SCROLL)

    If KeyCode = 20 Then 'si se presiona el boqueador de mayúsculas
        If EstadoMayúsculas Then
            Decir "mayúsculas activado"
        Else
            Decir "mayúsculas desactivado"
        End If
    End If

    If KeyCode = 144 Then 'si se presiona el boqueador de números
        If EstadoNumLock Then
            Decir "teclado numérico activado"
        Else
            Decir "teclado numérico desactivado"
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
'    rtfLectorEvaluaciones.Font.Size = tamañoFuente 'se ajusta el tamaño de la fuente
'    rtfLectorEvaluaciones.BackColor = colorFondo 'el color de fondo del rtf
'End Sub
'
Private Sub rtfLectorEvaluaciones_Change()
    If swAbriendoEvaluación = False Then
        swHuboCambio = True
    End If
    swAbriendoEvaluación = False
End Sub

Private Sub rtfLectorEvaluaciones_GotFocus()
    Call reproducirForm(formularios.evaluaciones)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7

    If KeyCode = 222 Then Decir "acento agudo"

    If KeyCode = vbKeyEscape Then
        If swListaCorrecciónVisible = False Then Unload Me
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir Trim(nombreUsuario) + ", para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en una evaluación"
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    If shiftkey = 0 And KeyCode = vbKeyF4 Then frmAccesorios.Show

    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa

    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.lectorEvaluaciones
         frmAyuda.Show
         Exit Sub
    End If

    Dim i As Integer, prefijo As String, contador As Integer ', extensión As String
    If Me.swSóloLeer = False Then
        If (shiftkey = vbCtrlMask And KeyCode = vbKeyG) Or KeyCode = vbKeyF5 Then 'se guarda el archivo con control + G ó con f5
            If rtfLectorEvaluaciones.Text <> "" Then
                If swHuboCambio = True Then 'si hay cambios no guardados
                    If InStrRev(swArchivoParaLeer, ".", , vbTextCompare) = 11 Then 'se evalúa si tiene el archivo el .dll para sacárselo en el nombre a guardar
                        swArchivoParaLeer = Left(swArchivoParaLeer, Len(swArchivoParaLeer) - 4)
                    End If
    
                    'se guarda la eval disfrazada como dll
                    rtfLectorEvaluaciones.SaveFile App.path + "\trabajos\" + swMateriaParaAbrir + "\soporte\" + Trim(Str(swNumMesParaAbrir)) + "\" + swArchivoParaLeer + ".dll"
                    'se guarda una copia falsa por si los papás quieren modificarla externamente
                    rtfLectorEvaluaciones.SaveFile App.path + "\trabajos\" + swMateriaParaAbrir + "\evaluaciones\" + Trim(Str(swNumMesParaAbrir)) + "\" + swArchivoParaLeer + ".rtf"
                    'se guarda el título de la evaluación
                    frmTítuloEvaluación.nombreArchivo = App.path + "\trabajos\" + swMateriaParaAbrir + "\soporte\" + Trim(Str(swNumMesParaAbrir)) + "\datosSoporte\" + swArchivoParaLeer + ".gui"
    
                    If Not existeCarpeta(frmTítuloEvaluación.nombreArchivo) Then frmTítuloEvaluación.Show 1  'se ofrece guardar la hoja si aún no se lo ha hecho
    '                Call chequearEspacioEnDisco(Left(App.Path, 2))
                    swHuboCambio = False 'se establece que no hay cambios sin guardar
                    If swHablarVoz = True Then
                        Decir "tu evaluación está guardada"
                    Else
                        frmMsgBox.cadenaAMostrar = "Tu evaluación está guardada"
                        frmMsgBox.swSíNoóAceptar = False 'se le dice que es un msg aceptar
                        frmMsgBox.Show 1
                    End If
                Else
                    If swHablarVoz = True Then
                        Decir "No has hecho cambios en tu evaluación, no hay nada nuevo para guardar"
                    Else
                        frmMsgBox.cadenaAMostrar = "No has hecho cambios en tu evaluación, no hay nada nuevo para guardar"
                        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
                        frmMsgBox.Show 1
                    End If
                End If
            Else
                If swHablarVoz = True Then
                    Decir "La evaluación está vacía, para guardar una evaluación hay que escribir algo en ella"
                Else
                    frmMsgBox.cadenaAMostrar = "La hoja de tu evaluación está vacía, para guardar una evaluación hay que escribir algo en ella"
                    frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
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
        Decir "No se puede escribir en una evaluación guardada, sólo se puede leer"
    End If

    If shiftkey = 0 And KeyCode = vbKeyF8 Then 'f8, calculadora
        If frmCalculadora.swEstoyAbierto = True Then
            Decir "Pasando a la calculadora, para volver a tu evaluación, apretá F8"
            frmCalculadora.Show
        Else
            frmCalculadora.Show
        End If
        Exit Sub
    End If

    If (shiftkey = vbCtrlMask And KeyCode = vbKeyP) Or KeyCode = vbKeyF6 Then 'imprimir con control + p
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "¿Realmente querés imprimir esta evaluación?"
        frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado = True Then
            If swImprimirDirecto = True Then
                With ImpresoraRich

                     'Valores
                     'Encabezado y pie de página
                     .Header = "Evaluación de la carpeta de " + swMateriaParaAbrir + " trabajada el día " + Format(Date, "dd/mm/yyyy")
                     .Footer = "Evaluación realizada por " + Trim(nombreUsuario)

                     'Margenes
                     .MarginTop = 500 'Text3
                     .MarginLeft = 500 'Text4
                     .MarginRight = 500 'Text5
                     .MarginBottom = 500 'Text6

                     'Imprimir el RichTextBox pasado como parámetro
                     .Imprimir rtfLectorEvaluaciones
                End With
            Else
                ImprimirConCuadroDiálogo 'se muestra el cuadro de diálogo de la impresora
            End If
        End If
        Exit Sub
   End If

    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
End Sub

Private Sub Form_Load()
    Dim i As Integer, contador As Integer, prefijo As String, temp As String
    Call centrarFormulario(Me)
    Set corrector = New corrector_ortografía
    'si no está aspell, se carga el diccionario propio
    If swAspellInstalado = False Then corrector.Cargar_Diccionario (App.path + "\datos\diccionario.txt")

    'Call contarFormularios(True)
    swEstoyAbierto = True
    swHuboCambio = False
    sonido = sndPlaySound(App.path + "\sonidos\abrir.wav", SND_ASYNC)
    If swSóloLeer = True Then
        rtfLectorEvaluaciones.Locked = True
        btnGuardar.Visible = False
        rtfLectorEvaluaciones.LoadFile App.path + "\trabajos\" + swMateriaParaAbrir + "\soporte\" + Trim(Str(swNumMesParaAbrir)) + "\" + swArchivoParaLeer '+ ".dll"
        swAbriendoEvaluación = True
        Decir Trim(nombreUsuario) + ", abriste la evaluación de la materia " + swMateriaParaAbrir + " del día " + Str(swDíaParaAbrir) + ". para leerla, usá las flechas"
        Me.Caption = "Evaluación de la materia " + Chr(34) + swMateriaParaAbrir + Chr(34) + " del día " + Trim(Str(swDíaParaAbrir)) + " de " + decodificarMes(swNumMesParaAbrir)
    Else
        btnGuardar.Visible = True
        rtfLectorEvaluaciones.Locked = False
        Decir "empezando una evaluación de la materia " + swMateriaParaAbrir + ". podés escribir en ella"
        Me.Caption = "Evaluación de la materia " + Chr(34) + swMateriaParaAbrir + Chr(34) + " del día de hoy, " + Trim(Str(swDíaParaAbrir)) + " de " + decodificarMes(swNumMesParaAbrir)
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

    If swAspellInstalado = True Then 'si está aspell, se deja listo el pipe para comunicarse con él
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
    rtfLectorEvaluaciones.Font.Size = tamañoFuente 'se ajusta el tamaño de la fuente
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
    
    If swSóloLeer = False Then 'si está escribiendo una evaluación
        If mensajeSalir("¿Estás seguro que querés cerrar la evaluación, una vez cerrada ya no vas a poder modificar lo que has escrito en ella?") Then
            Call Form_KeyDown(vbKeyF5, 0)
            frmPrincipal.Show
            swEstoyAbierto = False
        Else
            Cancel = 1
        End If
    Else 'si está revisando una evalación a hecha
        frmPrincipal.Show
        swEstoyAbierto = False
    End If

    'swSóloLeer = False
End Sub

Private Sub rtfLectorEvaluaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim auxString As String, caracterAnteriorBorrado As String, letra As String
    Dim swEnterEnMedioDelRenglón As Boolean, shiftkey As Integer, renglón As Long
    Dim palabra As String, temp() As String

    shiftkey = Shift And 7
    
    If KeyCode = vbKeyInsert Then KeyCode = 0 'si aprieta insert, se neutraliza así no activa la sobreescritura
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyReturn Or KeyCode = vbKeyBack Or KeyCode = 93 Then
        If swSóloLeer = False Then
            If KeyCode = 13 Then
                swEnterEnMedioDelRenglón = medioDelRenglón(rtfLectorEvaluaciones)
                If swEnterEnMedioDelRenglón = False Then
                    Decir "bajada de línea. renglón " + Trim(Str(rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 2))
                Else
                    Decir "estás haciendo una bajada de línea sin estar al final del renglón, si es un error podés corregirlo borrando la bajada de línea, yendo al final del renglón con la tecla fin, y ahí sí hacer la bajada de línea"
                End If
            End If
            
            If KeyCode = 93 Then 'si es el menú contextual
                palabra = buscarPalabraParaCorregir(rtfLectorEvaluaciones)
                If palabra = "" Then
                    Decir "no estás en ninguna palabra, no puedo corregir"
                    KeyCode = 0
                Else
                    If Not corregirPalabra(palabra) Then 'si la palabra es incorrecta
                        If palabra <> "" Then
                            If swAspellInstalado = True Then 'si aspell está instalado, se lo prefiere
                                '*************************
                                'corregir con aspell
                                Call objPipe.Write_(palabra & vbCrLf)
                                Call Sleep(200)
                                temp = arreglarCadena(objPipe.Read)
                            Else
                                '*****************************
                                'corregir con mi propio corrector (más lento)
                                temp = corrector.Controlar_Un_Error(palabra)
                            End If
                            Decir "usá flecha abajo para ver las palabras que te sugiero para corregir " + palabra
                            swListaCorrecciónVisible = True
                            Call Cargar_Menú_En_Lista(List1, temp)
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
                If rtfLectorEvaluaciones.Text <> "" Then 'si no está vacío
                    If rtfLectorEvaluaciones.SelStart <> Len(rtfLectorEvaluaciones.Text) Then 'y no está al final de la hoja
                        letra = Mid(rtfLectorEvaluaciones.Text, rtfLectorEvaluaciones.SelStart + 1, 1)
                        If letra = " " Then
                            Decir "borrando a la derecha el espacio", False
                        ElseIf letra = Chr(9) Then
                            Decir "borrando a la derecha un salto"
                        ElseIf letra = Chr(10) Or letra = Chr(13) Then
                            Decir "borrando a la derecha la bajada de línea. renglón " + Str(rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1), False
                        Else
                            auxString = traducirParaBorrar(letra)
                            Decir "borrando a la derecha " + auxString
                        End If
                    Else
                        Decir "imposible borrar, estás al final de la hoja"
                    End If
                Else
                    Decir "no se puede borrar a la derecha porque la hoja está vacía"
                End If
            End If
        
            If KeyCode = vbKeyBack Then
                If rtfLectorEvaluaciones.Text = "" Then
                    Decir "Ya está todo borrado"
                Else
                    If rtfLectorEvaluaciones.SelStart = 0 Then
                        Decir "imposible borrar porque estás al principio de la hoja"
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
                            Decir "borrando la bajada de línea. renglón " + Str(rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart)), False
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
            Decir Trim(nombreUsuario) + ", no podés borrar en una evaluación guardada, sólo leer lo que está escrito"
        End If
    End If

'    If shiftkey = vbCtrlMask And KeyCode = vbKeyHome Then 'control + inicio
'        If rtfLectorEvaluaciones <> "" Then
'            Decir decirPalabraSiguiente(rtfLectorEvaluaciones)
'        Else
'            Decir "La hoja está en blanco, no hay nada escrito"
'        End If
'    End If

    If shiftkey = 0 And KeyCode = vbKeyHome Then 'tecla inicio
        renglón = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
        Decir "principio del renglón " + Str(renglón)
    End If

'    If shiftkey = vbCtrlMask And KeyCode = vbKeyEnd Then 'control + fin
''        If swVolviendodeBraille = False Then 'si no se dispara el evento al volver del teclado braille
'            If rtfLectorEvaluaciones <> "" Then
'                Decir "final de la hoja. Estás detrás de la palabra " + decirPalabraAnterior(rtfLectorEvaluaciones)
'            Else
'                Decir "La hoja está en blanco, no hay nada escrito"
'            End If
''        End If
'    End If

    If shiftkey = 0 And KeyCode = vbKeyEnd Then 'tecla fin
        renglón = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
        Decir "final del renglón " + Str(renglón)
    End If

    If shiftkey = vbCtrlMask Then controlPresionado = True

End Sub

Private Sub rtfLectorEvaluaciones_KeyPress(KeyAscii As Integer)
    Dim cadena As String
    If swSóloLeer = False Then
        swHuboCambio = True
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
        Or KeyAscii = Asc("-") Then
            cadena = cadena + decirPalabraAnterior(rtfLectorEvaluaciones)
            If swUsarCorrectorOrtográfico = True Then 'si se usa el corrector, se dice si es incorrecta
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
            Decir Trim(nombreUsuario) + ", no podés borrar en una evaluación guardada, sólo leer lo que está escrito"
        Else
            Decir "no podés escribir en una evaluación guardada, sólo leer lo que está escrito"
        End If
    End If
End Sub

Private Sub rtfLectorEvaluaciones_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    Dim renglón As Long, cadena As String

    shiftkey = Shift And 7

    If shiftkey = vbCtrlMask And KeyCode = vbKeyLeft Then 'leer por palabras retrocediendo
        cadena = decirPalabraSiguiente(rtfLectorEvaluaciones)
        If Not esSigno(cadena) Then 'se ve si la cadena es solamente un signo ortográfico
            If Not corregirPalabra(cadena) Then cadena = cadena + ", incorrecta"
        End If
        Decir cadena
    End If

    If shiftkey = vbCtrlMask And KeyCode = vbKeyRight Then 'leer por palabras avanzando
        cadena = decirPalabraSiguiente(rtfLectorEvaluaciones)
        If Not esSigno(cadena) Then 'se ve si la cadena es solamente un signo ortográfico
            If Not corregirPalabra(cadena) Then cadena = cadena + ", incorrecta"
        End If
        Decir cadena
    End If

    If shiftkey = 0 And KeyCode = vbKeyRight Then 'avanzar de a caracter
        Decir decirLetraSiguiente(rtfLectorEvaluaciones)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyPageDown Then 'tecla avance de página
        renglón = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
        Decir "saltando hacia adelante al renglón " + Str(renglón)
    End If

    If shiftkey = 0 And KeyCode = vbKeyPageUp Then 'tecla retroceso de página
        renglón = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
        Decir "saltando hacia atrás al renglón " + Str(renglón)
    End If

    Dim auxString As String
    If shiftkey = vbCtrlMask And KeyCode = vbKeyHome Then 'control + inicio
        If rtfLectorEvaluaciones.Text <> "" Then
            auxString = decirPalabraSiguiente(rtfLectorEvaluaciones)
            If Trim(auxString) <> Chr(10) And Trim(auxString) <> Chr(13) Then
                Decir "principio de la hoja." + auxString
            Else
                Decir "principio de la hoja. renglón en blanco"
            End If
        Else
            Decir "La hoja está en blanco, no hay nada escrito"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyEnd Then 'control + fin
        If rtfLectorEvaluaciones.Text <> "" Then
            auxString = decirPalabraAnterior(rtfLectorEvaluaciones)
            If Trim(auxString) <> Chr(10) And Trim(auxString) <> Chr(13) Then
                If Len(Trim(auxString)) <> 0 Then
                    Decir "final de la hoja. Estás detrás de la palabra " + decirPalabraAnterior(rtfLectorEvaluaciones)
                Else
                    Decir "final de la hoja. sólo hay escrito espacios en este renglón, ninguna letra"
                End If
            Else
                Decir "final de la hoja. renglón en blanco"
            End If
        Else
            Decir "La hoja está en blanco, no hay nada escrito"
        End If
    End If

    If shiftkey = 0 And KeyCode = vbKeyLeft Then 'retroceder de a caracter
        If rtfLectorEvaluaciones.SelStart = 0 And rtfLectorEvaluaciones.Text <> "" Then
            Decir "Estás en el principio de la hoja, delante de la letra " + decirLetraSiguiente(rtfLectorEvaluaciones)
        Else
            Decir decirLetraSiguiente(rtfLectorEvaluaciones)
        End If
    End If

    If shiftkey = 0 And KeyCode = vbKeyDown Then 'leer por oración
        Decir decirOraciónSiguiente(rtfLectorEvaluaciones)
    End If

    If shiftkey = 0 And KeyCode = vbKeyUp Then 'leer por oración
        Decir decirOraciónSiguiente(rtfLectorEvaluaciones)
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

    Call evaluarSelección(rtfLectorEvaluaciones, control, shift2, teclaApretada) 'se ve si hay selección
    

    If shiftkey = vbCtrlMask And KeyCode = vbKeyUp Then 'retroceder de a párrafo
        If rtfLectorEvaluaciones.Text <> "" Then
            renglón = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
            If renglón = 1 Then
                Decir "principio de la hoja, renglón 1"
            Else
                Decir "retrocediendo un párrafo. renglón " + Str(renglón)
            End If
        Else
            Decir "No se puede retroceder de a párrafo porque la hoja está vacía"
        End If
    End If

    If shiftkey = vbCtrlMask And KeyCode = vbKeyDown Then 'avanzar de a párrafo
        If rtfLectorEvaluaciones.Text <> "" Then
            renglón = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
            If rtfLectorEvaluaciones.GetLineFromChar(Len(rtfLectorEvaluaciones.Text)) + 1 = renglón Then
                Decir "final de la hoja. renglón " + Str(renglón)
            Else
                Decir "avanzando un párrafo. renglón " + Str(renglón)
            End If
        Else
            Decir "No se puede avanzar de a párrafo porque la hoja está vacía"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyZ Then 'deshacer
        Decir "deshaciendo la última acción"
    End If

    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copiar
        If rtfLectorEvaluaciones.SelText <> "" Then
            Decir "se copió el texto seleccionado. para pegarlo en otro lugar, usar control más ve corta"
        Else
            Decir "No se puede copiar porque no hay texto seleccionado. para seleccionar, usar shift más las flechas"
        End If
    End If

    If shiftkey = vbCtrlMask And KeyCode = vbKeyX Then 'cortar
        If swSóloLeer = False Then
            If rtfLectorEvaluaciones.SelText <> "" Then
                Decir "se cortó el texto seleccionado. para pegarlo en otro lugar, usar control más ve corta"
            Else
                Decir "No se puede cortar porque no hay texto seleccionado. para seleccionar, usar shift más las flechas"
            End If
        Else
            Decir "no podés cortar en una evaluación guardada, sólo leer lo que está escrito"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyV Then 'pegar
        If swSóloLeer = False Then
            If Clipboard.GetText <> "" Then
                renglón = rtfLectorEvaluaciones.GetLineFromChar(rtfLectorEvaluaciones.SelStart) + 1
                Decir "texto pegado en el renglón " + Str(renglón)
            Else
                Decir "No se puede pegar porque no hay texto copiado o cortado. para copiar, usar control más c. para cortar, usar control más x"
            End If
        Else
            Decir "no podés pegar nada en una evaluación guardada, sólo leer lo que está escrito"
        End If
    End If
    
'    Dim estáPalabraEnLista As Boolean 'para el corrector ortográfico
'    If swUsarCorrectorOrtográfico = True And swSóloLeer = False Then 'el corrector ortográfico
'        If shiftkey = 0 And (KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = Asc(",") Or KeyCode = Asc(".") Or KeyCode = Asc("!") Or KeyCode = Asc("?") Or KeyCode = Asc("-")) Then 'con espacio se corrige la palabra recién escrita
'            estáPalabraEnLista = corregirPalabra(rtfLectorEvaluaciones)
'            If estáPalabraEnLista = False Then Decir "palabra con posible error"
'        End If
'    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyLeft Then 'leer la oración actual
        If Trim(rtfLectorEvaluaciones.Text) <> "" Then
            Decir "El renglón actual dice: " + decirOraciónSiguiente(rtfLectorEvaluaciones)
        Else
            Decir "No se puede leer el renglón actual porque la evaluación está vacía"
        End If
    End If

    If shiftkey = vbAltMask And KeyCode = vbKeyDown Then 'leer todo el texto
        If Trim(rtfLectorEvaluaciones.Text) <> "" Then
            Decir "toda la evaluación dice: " + rtfLectorEvaluaciones.Text
        Else
            Decir Trim(nombreUsuario) + ", No se puede leer todo el texto porque la evaluación está vacía"
        End If
    End If

    If shiftkey = vbAltMask And KeyCode = vbKeyUp Then 'leer desde el cursor hacia adelante
        If Trim(rtfLectorEvaluaciones.Text) <> "" Then
            If rtfLectorEvaluaciones.SelStart = 0 Then
                Decir "desde donde estás hasta el final de la evaluación dice: " + Mid(rtfLectorEvaluaciones.Text, 1, Len(rtfLectorEvaluaciones.Text) - Len(Left(rtfLectorEvaluaciones.Text, rtfLectorEvaluaciones.SelStart))) 'leer desde el cursor hacia adelante
            Else
                Decir "desde donde estás hasta el final de la evaluación dice: " + Mid(rtfLectorEvaluaciones.Text, rtfLectorEvaluaciones.SelStart, Len(rtfLectorEvaluaciones.Text) - Len(Left(rtfLectorEvaluaciones.Text, rtfLectorEvaluaciones.SelStart))) 'leer desde el cursor hacia adelante
            End If
        Else
            Decir "No se puede leer todo el texto porque la evaluación está vacía"
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
            Decir "No se puede leer la selección porque no hay texto seleccionado"
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
'para pegar imágenes
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

  ' No es necesario ajustar cuando la ventana está minimizada
    If WindowState = vbMinimized Then
            Exit Sub
    End If
       
    'el botón de arriba
    posX = Me.Width - ButtonTransparent1.Width - 400
    posY = 360

    ButtonTransparent1.Move posX, posY
    
    'las líneas y etiqueta de arriba
    Line2.X2 = Me.Width - 400
    Line2.X1 = Line2.X2 - 480
    
    Label2.Left = Line2.X1 - 20 - Label2.Width
    
    Line1.X2 = Label2.Left - 20
    Line1.X1 = Line1.X2 - 480
    
    'el botón de abajo
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



