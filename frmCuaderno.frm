VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmCuaderno 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Cuaderno"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   Icon            =   "frmCuaderno.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmCuaderno.frx":08CA
   ScaleHeight     =   7725
   ScaleWidth      =   9270
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "frmCuaderno.frx":40B9
      Left            =   3923
      List            =   "frmCuaderno.frx":40BB
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      Picture         =   "frmCuaderno.frx":40BD
      ScaleHeight     =   435
      ScaleWidth      =   600
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   660
   End
   Begin TransparentButton.ButtonTransparent btnImprimir 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Imprimir"
      EstiloDelBoton  =   1
      Picture         =   "frmCuaderno.frx":E356F
      PictureHover    =   "frmCuaderno.frx":E3E49
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
   Begin MSComDlg.CommonDialog di�logo 
      Left            =   5040
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   5280
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10398
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmCuaderno.frx":E4723
   End
   Begin TransparentButton.ButtonTransparent btnGuardar 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Guardar"
      EstiloDelBoton  =   1
      Picture         =   "frmCuaderno.frx":E47A5
      PictureHover    =   "frmCuaderno.frx":E507F
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
   Begin TransparentButton.ButtonTransparent btnConfiguraci�n 
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "Actividades y Libros"
      EstiloDelBoton  =   1
      Picture         =   "frmCuaderno.frx":E5959
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
      Left            =   8280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      Caption         =   ""
      PicturePosition =   3
      EstiloDelBoton  =   1
      Picture         =   "frmCuaderno.frx":E6233
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
   Begin TransparentButton.ButtonTransparent ButtonTransparent2 
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "    Abrir"
      EstiloDelBoton  =   1
      Picture         =   "frmCuaderno.frx":E6B0D
      PictureHover    =   "frmCuaderno.frx":E73E7
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
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   8520
      X2              =   9000
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6720
      X2              =   7200
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para el docente:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7290
      TabIndex        =   6
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCuaderno.frx":E7CC1
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Configurar las materias, voz y otros"
      Top             =   6960
      Width           =   7935
   End
End
Attribute VB_Name = "frmCuaderno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nombreArchivo As String 'el archivo a abrir o guardar
Public nombreMesArchivo As String 'el mes del archivo a abrir
Public swContinuarArchivo As Boolean 'para saber si es archivo nuevo o si contin�a una hoja anterior
Public swAbriendoHojaAnterior As Boolean 'para controlar en el evento change del rtf que no se modifique swHuboCambio
Public d�aAbierto As String 'para que pase la cadena del d�a que se quiere abrir
Public swVolviendodeBraille As Boolean 'para saber si se vuelve del teclado braille
Public swArchivoExterno As Boolean 'para saber si se abri� un archivo externo
Dim swHuboCambio As Boolean 'para controlar si hubo cambio en un archivo para ofrecer guardar el documento al cerrar
Dim ImpresoraRich As ImpresoraRTF 'la impresora del rtf
Dim controlPresionado As Boolean 'para mandar a Keyascii si se est� apretando control
Dim corrector As corrector_ortograf�a
Dim swListaCorrecci�nVisible As Boolean 'para ver si con escape se saca el cuaderno o s�lo la lista
Dim swDeletrear As Boolean

Private Sub btnConfiguraci�n_Click()
    frmControlActyLibros.Show
End Sub

Private Sub btnGuardar_Click() 'bot�n guardar
    'Call Form_KeyDown(vbKeyF5, 0)
    di�logo.Filter = "Archivo RTF (*.rtf) |*.rtf"
    di�logo.ShowSave
    RichTextBox1.SaveFile di�logo.FileName
End Sub

Private Sub btnImprimir_Click() 'bot�n imprimir
    'Call Form_KeyDown(vbKeyF6, 0)
    Call ImprimirConCuadroDi�logo
End Sub

Private Sub ButtonTransparent2_Click() 'bot�n abrir
    'Call Form_KeyDown(vbKeyF1, 0)
    di�logo.Filter = "Archivos de Texto (*.rtf; *.txt)|*.rtf; *.txt"
    di�logo.ShowOpen
    RichTextBox1.LoadFile di�logo.FileName
End Sub


Private Sub Form_GotFocus()
    If List1.Visible = True Then
        List1.Visible = False
        swListaCorrecci�nVisible = False
    End If
End Sub

Private Sub List1_DblClick()
    Call corregirConPalabraSeleccionada(RichTextBox1, List1.List(List1.ListIndex))
    Call List1_KeyDown(vbKeyEscape, 0)
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        List1.Visible = False
        swListaCorrecci�nVisible = False
        Decir "cerrando las sugerencias, est�s otra vez en tu carpeta"
    End If
    
    If KeyCode = vbKeyReturn Then List1_DblClick
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        If swDeletrear = True Then
            Decir List1.List(List1.ListIndex) + ". " + separarEnLetras(List1.List(List1.ListIndex))
        Else
            Decir List1.List(List1.ListIndex)
        End If
    End If
End Sub

'Private Sub mnuPalabras_Click(Index As Integer)
'    Decir mnuPalabras(Index).Caption & ", " & separarEnLetras(mnuPalabras(Index).Caption)
'End Sub

Private Sub richTextbox1_GotFocus()
    Call reproducirForm(formularios.cuaderno)
    If swVolviendodeBraille = True Then
        SendKeys ("^{end}") 'cuando se vuelve del teclado braille se pasa al final de la hoja
    End If
    
    If List1.Visible = True Then
        List1.Visible = False
        swListaCorrecci�nVisible = False
    End If
End Sub

Private Sub ButtonTransparent1_Click() 'bot�n configurar
    Call Form_KeyDown(vbKeyF12, 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = 0 And KeyCode = 222 Then Decir "acento agudo"
'    If KeyCode = 96 Then Decir "acento grave"
'    If KeyCode = 94 Then Decir "acento circunflejo"
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    
    If shiftkey = 0 And KeyCode = vbKeyF1 Then
        frmDesdeCuaderno.Show '1
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF2 Then
        If swActividadAbierta = True Then
            Decir "pasando a la actividad abierta" ', para volver a tu carpeta apret� f2"
            frmLectorActividad.Show
        Else
            Decir Trim(nombreUsuario) + ", No hay ninguna actividad abierta. para abrir una, apret� f1"
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF3 Then
        If swLibroAbierto = True Then
            Decir "pasando al libro abierto" ', para volver a tu carpeta apret� f3"
            frmLectorLibro.Show
        Else
            Decir "No hay ning�n libro abierto " + Trim(nombreUsuario) + "+. para abrir uno, apret� f1"
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF4 Then frmAccesorios.Show '1
        
    If shiftkey = 0 And KeyCode = vbKeyEscape Then
        If swListaCorrecci�nVisible = False Then 'si la lista de sugerencias est� abierta, escape no cierra el cuaderno
            If swHuboCambio = True And RichTextBox1.Text <> "" Then 'se ofrece guardar si hay cambios
                frmMsgBox.swMostrarCancelar = True
                frmMsgBox.cadenaAMostrar = "Has hecho cambios en la hoja que no han sido guardados. �Quer�s que los guarde por vos?"
                frmMsgBox.swS�No�Aceptar = True 'se elige que sea cuadro aceptar
                frmMsgBox.Show 1
                If frmMsgBox.swResultadoMostrado = True Then
                    Call Form_KeyDown(vbKeyF5, 0)
                    swHuboCambio = False
                End If
            End If
            
            If frmMsgBox.swCancelar = False Then
                Unload Me
            End If
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF12 Then frmControl.Show
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF11 Then frmControlActyLibros.Show
    If shiftkey = 0 And KeyCode = vbKeyF10 Then
        If frmDiccionarioElegido.swEstoyAbierto Then
            Decir "pasando al diccionario"
            frmDiccionarioElegido.Show
            SendKeys ("%")
        Else
            Decir "no hay ning�n diccionario abierto. para abrir uno, apret� efe cuatro y buscalo en los accesorios"
            'RichTextBox1.SetFocus
            KeyCode = 0
        End If
        Exit Sub
    End If
        
    Dim i As Integer, prefijo As String, contador As Integer, extensi�n As String
    If (shiftkey = vbCtrlMask And KeyCode = vbKeyG) Or KeyCode = vbKeyF5 Then 'se guarda el archivo con control + G � con f5
        If RichTextBox1.Text <> "" Then
            If swHuboCambio = True Then 'si hay cambios no guardados
                
                If Right(nombreArchivo, 4) <> ".txt" Or Right(nombreArchivo, 4) <> ".rtf" Then
                    extensi�n = ".rtf"
                Else
                    extensi�n = Right(nombreArchivo, 4) 'se toma la extensi�n del archivo abierto antes de modificar la cadena
                End If
                If InStrRev(nombreArchivo, ".", , vbTextCompare) = 11 Then 'se eval�a si tiene el archivo el .rtf para sac�rselo en el nombre a guardar
                    nombreArchivo = Left(nombreArchivo, Len(nombreArchivo) - 4)
                End If
                    
                If swArchivoExterno = False Then
                    RichTextBox1.SaveFile App.path + dirTrabajo + Trim(Str(CInt(nombreMesArchivo))) + "\" + nombreArchivo + extensi�n
                    frmT�tuloHoja.nombreArchivo = App.path + dirTrabajo + Trim(Str(CInt(nombreMesArchivo))) + "\datosHojas\" + nombreArchivo + ".gui"
                    If Not existeCarpeta(frmT�tuloHoja.nombreArchivo) Then frmT�tuloHoja.Show 1  'se ofrece guardar el t�tulo de la hoja
'                    Call chequearEspacioEnDisco(Left(App.Path, 2))
                Else
                    RichTextBox1.SaveFile frmDi�logoAbrir.archivoDevuelto
                End If
                
                swHuboCambio = False 'se establece que no hay cambios sin guardar
                If swArchivoExterno = False Then
                    Decir "tu trabajo est� guardado"
                Else
                    Decir "se han guardado los cambios en el archivo abierto"
                End If
            Else
                Decir "No has hecho cambios en la hoja, no hay nada nuevo para guardar"
            End If
        Else
            Decir "La hoja de la carpeta est� vac�a, para guardar una hoja hay que escribir algo en ella"
        End If
        Exit Sub
    End If
       
    If (shiftkey = vbCtrlMask And KeyCode = vbKeyP) Or KeyCode = vbKeyF6 Then 'imprimir con control + p
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "�Realmente quer�s imprimir esta hoja de la carpeta?"
        frmMsgBox.swS�No�Aceptar = True 'se elige que sea cuadro s�-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado = True Then
            If swImprimirDirecto = True Then
                With ImpresoraRich
                     'Valores
                     'Encabezado y pie de p�gina
                     .Header = "Hoja de la carpeta de " + miMateria + " trabajada el d�a " + Format(Date, "dd/mm/yyyy") 'Text1
                     .Footer = "Trabajo realizado por " + Trim(nombreUsuario) 'Text2
                     
                     'Margenes
                     .MarginTop = 500 'Text3
                     .MarginLeft = 500 'Text4
                     .MarginRight = 500 'Text5
                     .MarginBottom = 500 'Text6
                     
                     'Imprimir el RichTextBox pasado como par�metro
                     .Imprimir RichTextBox1
                End With
            Else
                ImprimirConCuadroDi�logo 'se muestra el cuadro de di�logo de la impresora
            End If
        End If
        Exit Sub
    End If
    
    If (shiftkey = vbCtrlMask And KeyCode = vbKeyN) Then 'control n para negrita
        If RichTextBox1.SelBold = False Then 'estoyEnNegrita = False Then
            RichTextBox1.SelBold = True
            Decir "negrita activado"
        Else
            RichTextBox1.SelBold = False
            Decir "negrita desactivado"
        End If
        Exit Sub
    End If
    
    If (shiftkey = vbCtrlMask And KeyCode = vbKeyS) Then 'control s para subrayar
        If RichTextBox1.SelUnderline = False Then
            RichTextBox1.SelUnderline = True
            Decir "subrayado activado"
        Else
            RichTextBox1.SelUnderline = False
            Decir "subrayado desactivado"
        End If
        Exit Sub
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF9 Then 'f9 abre el teclado braille
        frmTecladoBraille.Show 1
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF8 Then 'f8, calculadora
        If frmCalculadora.swEstoyAbierto = True Then
            Decir "Pasando a la calculadora, para volver a tu carpeta, apret� F8"
            frmCalculadora.Show
        Else
            frmCalculadora.Show
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then 'f7 abre el reproductor de m�sica
        If frmReproductorM�sica.swEstoyAbierto = True Then
            Decir "Pasando al reproductor de m�sica, para volver a tu cuaderno, apret� F7"
            frmReproductorM�sica.SetFocus
        Else
            Decir "Abriendo el reproductor de m�sica, para volver a tu cuaderno, apret� F7"
            frmReproductorM�sica.Show
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.cuaderno
         frmAyuda.Show 1
         Exit Sub
    End If
    
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    
    If shiftkey = vbCtrlMask + vbShiftMask + vbAltMask And KeyCode = vbKeyH Then mostrarHuevo
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

Public Sub Form_Load()
    Dim posici�nDelReemplazo As Byte, contador As Integer, prefijo As String, temp As String
    Set ImpresoraRich = New ImpresoraRTF
    Set corrector = New corrector_ortograf�a
    'si no est� aspell, se carga el diccionario propio
    If swAspellInstalado = False Then corrector.Cargar_Diccionario (App.path + "\datos\diccionario.txt")
    
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    'Call reproducirForm(formularios.cuaderno)
    swVolviendodeBraille = False
    swCuadernoAbierto = True
    swArchivoExterno = False 'al abrir no se abre un archivo externo
    swHuboCambio = False 'se establece que no hay cambios sin guardar (pues reci�n se abre el cuaderno)
    Me.Caption = "Carpeta de " + miMateria
    Dim i As Integer, aux As String
    If swContinuarArchivo = False Then
        nombreArchivo = Format(Date, "dd/mm/yyyy")
        aux = Right(nombreArchivo, 5)
        aux = Left(nombreArchivo, 2) + aux
        nombreMesArchivo = Mid(nombreArchivo, 4, 2)
        nombreArchivo = aux
        Do  'se cambian los / por - para que se pueda guardar el archivo
            posici�nDelReemplazo = InStr(nombreArchivo, "/")
            Mid(nombreArchivo, posici�nDelReemplazo) = "-"
            posici�nDelReemplazo = InStr(nombreArchivo, "/")
        Loop Until posici�nDelReemplazo = 0
        
        For i = 1 To cantPrefijo 'se le agrega el prefijo al archivo
            nombreArchivo = "0" + nombreArchivo
        Next
        
        Do While existeCarpeta(App.path + dirTrabajo + Trim(Str(CInt(nombreMesArchivo))) + "\" + nombreArchivo + ".rtf")
            nombreArchivo = Right(nombreArchivo, Len(nombreArchivo) - cantPrefijo)
            contador = contador + 1
                    
            If contador < 10 Then prefijo = "10" + Trim(Str(contador))
            If contador >= 10 And contador < 100 Then prefijo = "1" + Str(contador)
            If contador >= 100 Then prefijo = Str(contador)
            
            nombreArchivo = prefijo + nombreArchivo
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
    'RichTextBox1.AutoVerbMenu = True
    Decir "abriendo la " + Me.Caption + ". pod�s escribir en esta hoja en blanco, o apretar f1 para abrir una hoja ya escrita, actividad, o libro. para abrir la ayuda apret� control m�s efe 1"
End Sub

Private Sub Form_Activate()
    'se actualizan las fuentes y los colores de los rtf del programa
    RichTextBox1.SelFontName = NombreFuente 'se ajusta la fuente del programa
    RichTextBox1.SelColor = colorFuente 'se ajusta el color de la fuente del programa
    RichTextBox1.SelFontSize = tama�oFuente 'se ajusta el tama�o de la fuente
    RichTextBox1.BackColor = colorFondo 'el color de fondo del rtf
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
    
    Me.swArchivoExterno = False
    If swActividadAbierta = True Then Unload frmLectorActividad
    If swLibroAbierto = True Then Unload frmLectorLibro
    If frmCalculadora.swEstoyAbierto = True Then Unload frmCalculadora
    If frmDiccionarioElegido.swEstoyAbierto = True Then Unload frmDiccionarioElegido
    Set corrector = Nothing
    frmPrincipal.Show
    swCuadernoAbierto = False
End Sub

Private Sub RichTextBox1_Change()
    If swAbriendoHojaAnterior = False Then
        swHuboCambio = True
    End If
    swAbriendoHojaAnterior = False
End Sub

Private Sub ImprimirConCuadroDi�logo()
   ' El control CommonDialog se llama "dlgPrint".
    di�logo.CancelError = True
    On Error GoTo manejoErrorImpresora
    di�logo.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If RichTextBox1.SelLength = 0 Then
       di�logo.Flags = di�logo.Flags + cdlPDAllPages
    Else
       di�logo.Flags = di�logo.Flags + cdlPDSelection
    End If
    di�logo.ShowPrinter
    RichTextBox1.SelPrint di�logo.hDC
manejoErrorImpresora: 'si el error es distinto a haber hecho click en cancelar, se muestra un msg
'    If Err.Number <> 32755 Then MsgBox "La impresora no est� lista para imprimir." + Chr(13) + "Por favor vuelva a intentar cuando est� lista.", , "Informaci�n"
    Exit Sub
End Sub

Private Sub richTextbox1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim auxString As String, caracterAnteriorBorrado As String, letra As String
    Dim swEnterEnMedioDelRengl�n As Boolean, shiftkey As Integer, rengl�n As Long
    'Dim palabrasCorregidas() As String
    Dim palabra As String
    
    shiftkey = Shift And 7
    
    If KeyCode = vbKeyInsert Then KeyCode = 0 'si aprieta insert, se neutraliza as� no activa la sobreescritura
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el men� de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    Dim temp() As String
    If KeyCode = 93 Then 'si es el men� contextual
        palabra = buscarPalabraParaCorregir(RichTextBox1)
        If palabra = "" Then
            Decir "no est�s en ninguna palabra, no puedo corregir"
            KeyCode = 0
        Else
            ReDim temp(0 To 0)
            If Not corregirPalabra(palabra) Then 'si la palabra es incorrecta
                'If palabra <> "" Then
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
                    
                    If temp(0) <> "" Then 'si hay alguna devoluci�n
                        Decir "us� flecha abajo para ver las palabras que te sugiero para corregir " + palabra
                        swListaCorrecci�nVisible = True
                        swDeletrear = True
                        Call Cargar_Men�_En_Lista(List1, temp)
                        Call Aplicar_ScrollBar(List1)
                    Else
                        Decir "no s� qu� palabra sugerirte para corregir"
                    End If
                'End If
            Else
                ReDim temp(0 To 2)
                Decir "como la palabra est� bien escrita, te doy su definici�n, sin�nimos y ant�nimos. Buscalos con las flechas o sal� con escape"
                temp(0) = buscarEntrada(palabra, "espa�ol.txt")
                If temp(0) = "" Then temp(0) = "No tengo definici�n en mi diccionario para " + palabra
                temp(1) = buscarEntrada(palabra, "sin�nimos.txt")
                If temp(1) = "" Then temp(1) = "No s� ning�n sin�nimo para " + palabra
                temp(2) = buscarEntrada(palabra, "ant�nimos.txt")
                If temp(2) = "" Then temp(2) = "No conozco ant�nimos para " + palabra
                
                Call Cargar_Men�_En_Lista(List1, temp)
                Call Aplicar_ScrollBar(List1)
                swListaCorrecci�nVisible = True
                swDeletrear = False
                KeyCode = 0
            End If
        End If
    End If

    
    If KeyCode = 13 Then
        swEnterEnMedioDelRengl�n = medioDelRengl�n(RichTextBox1)
        If swEnterEnMedioDelRengl�n = False Then
            Decir "bajada de l�nea. rengl�n " + Trim(Str(RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 2))
        Else
            Decir Trim(nombreUsuario) + ", est�s haciendo una bajada de l�nea sin estar al final del rengl�n, si es un error pod�s corregirlo borrando la bajada de l�nea, yendo al final del rengl�n con la tecla fin, y ah� s� hacer la bajada de l�nea"
        End If
    End If
    
    If KeyCode = vbKeyDelete Then
        If RichTextBox1.Text <> "" Then 'si no est� vac�o
            If RichTextBox1.SelStart <> Len(RichTextBox1.Text) Then 'y no est� al final de la hoja
                letra = Mid(RichTextBox1.Text, RichTextBox1.SelStart + 1, 1)
                If letra = " " Then
                    Decir "borrando a la derecha el espacio", False
                ElseIf letra = Chr(9) Then
                    Decir "borrando a la derecha un salto"
                ElseIf letra = Chr(10) Or letra = Chr(13) Then
                    Decir "borrando a la derecha la bajada de l�nea. rengl�n " + Str(RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1), False
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
        If RichTextBox1.Text = "" Then
            Decir "Ya est� todo borrado"
        Else
            If RichTextBox1.SelStart = 0 Then
                Decir "imposible borrar porque est�s al principio de la hoja"
            Else
                If RichTextBox1.SelText = "" Then 'si no hay nada seleccionado
                    caracterAnteriorBorrado = Mid(RichTextBox1.Text, RichTextBox1.SelStart, 1)
                Else
                    Decir "borrando el texto seleccionado"
                    Exit Sub
                End If
                
                If caracterAnteriorBorrado = " " Then
                    Decir "borrando el espacio", False
                ElseIf caracterAnteriorBorrado = Chr(9) Then
                        Decir "borrando un salto"
                ElseIf caracterAnteriorBorrado = Chr(10) Then
                    Decir "borrando la bajada de l�nea. rengl�n " + Str(RichTextBox1.GetLineFromChar(RichTextBox1.SelStart)), False
                Else
                    If caracterAnteriorBorrado <> "" Then
                        auxString = traducirParaBorrar(caracterAnteriorBorrado)
                        'arreglar, poner ac� que diga lo que queda de la palabra. Mirar si est� dentro de la palabra al borrar
                        Decir "borrando " + auxString, True, False
                    End If
                End If
            End If
        End If
    End If
    
    
    If shiftkey = 0 And KeyCode = vbKeyHome Then 'tecla inicio
        rengl�n = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
        Decir "principio del rengl�n " + Str(rengl�n)
    End If
       
    If shiftkey = 0 And KeyCode = vbKeyEnd Then 'tecla fin
        rengl�n = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
        Decir "final del rengl�n " + Str(rengl�n)
    End If
    
    If shiftkey = vbCtrlMask Then controlPresionado = True
End Sub

Private Sub richTextbox1_KeyPress(KeyAscii As Integer)
    swHuboCambio = True
    Dim cadena As String
    'Dim palabraEscrita As String 'para el corrector ortogr�fico
       
    'leer la tecla apretada
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
    Or KeyAscii = Asc("-") Then 'Or KeyAscii = Asc("_") Then
        cadena = cadena + decirPalabraAnterior(RichTextBox1)
        If swUsarCorrectorOrtogr�fico = True Then 'si se usa el corrector, se dice si es incorrecta
            'est�PalabraEnLista = corregirPalabra(obtenerPalabra(RichTextBox1))
            If Not corregirPalabra(buscarPalabraParaCorregir(RichTextBox1)) Then cadena = cadena + ", incorrecta"
        End If
    End If
    
    If cadena <> "" Then
        If RichTextBox1.SelBold = True Then cadena = cadena + " en negrita"
        If RichTextBox1.SelUnderline = True Then cadena = cadena + " subrayada"
        Decir cadena
    End If
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    Dim rengl�n As Long, cadena As String
    
    shiftkey = Shift And 7
            
    If shiftkey = vbCtrlMask And KeyCode = vbKeyLeft Then 'leer por palabras retrocediendo
        cadena = decirPalabraSiguiente(RichTextBox1)
        If Not esSigno(cadena) Then 'se ve si la cadena es solamente un signo ortogr�fico
            If Not corregirPalabra(cadena) Then cadena = cadena + ", incorrecta"
        End If
        Decir cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyRight Then 'leer por palabras avanzando
        cadena = decirPalabraSiguiente(RichTextBox1)
        If Not esSigno(cadena) Then 'se ve si la cadena es solamente un signo ortogr�fico
            If Not corregirPalabra(cadena) Then cadena = cadena + ", incorrecta"
        End If
        Decir cadena
    End If
        
    If shiftkey = 0 And KeyCode = vbKeyRight Then 'avanzar de a caracter
        Decir decirLetraSiguiente(RichTextBox1)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyLeft Then 'retroceder de a caracter
        If RichTextBox1.SelStart = 0 And RichTextBox1.Text <> "" Then
            Decir "Est�s en el principio de la hoja, delante de la letra " + decirLetraSiguiente(RichTextBox1)
        Else
            Decir decirLetraSiguiente(RichTextBox1)
        End If
    End If
    
    Dim auxString As String
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyHome Then 'control + inicio
        If RichTextBox1.Text <> "" Then
            auxString = decirPalabraSiguiente(RichTextBox1)
            If Trim(auxString) <> Chr(10) And Trim(auxString) <> Chr(13) Then
                Decir "principio de la hoja." + auxString 'decirPalabraAnterior(RichTextBox1)
            Else
                Decir "principio de la hoja. rengl�n en blanco"
            End If
        Else
            Decir "La hoja est� en blanco, no hay nada escrito"
        End If
    End If

    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyEnd Then 'control + fin
        If swVolviendodeBraille = False Then 'si no se dispara el evento al volver del teclado braille
            If RichTextBox1.Text <> "" Then
                auxString = decirPalabraAnterior(RichTextBox1)
                If Trim(auxString) <> Chr(10) And Trim(auxString) <> Chr(13) Then
                    If Len(Trim(auxString)) <> 0 Then
                        Decir "final de la hoja. Est�s detr�s de la palabra " + auxString 'decirPalabraAnterior(RichTextBox1)
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
        swVolviendodeBraille = False
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyDown Then 'leer por oraci�n
        Decir decirOraci�nSiguiente(RichTextBox1)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyUp Then 'leer por oraci�n
        Decir decirOraci�nSiguiente(RichTextBox1)
    End If
        
    If shiftkey = 0 And KeyCode = vbKeyPageDown Then 'tecla avance de p�gina
        rengl�n = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
        Decir "saltando hacia adelante al rengl�n " + Str(rengl�n)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyPageUp Then 'tecla retroceso de p�gina
        rengl�n = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
        Decir "saltando hacia atr�s al rengl�n " + Str(rengl�n)
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

    Call evaluarSelecci�n(RichTextBox1, control, shift2, teclaApretada) 'se ve si hay selecci�n
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyUp Then 'retroceder de a p�rrafo
        If RichTextBox1.Text <> "" Then
            rengl�n = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
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
        If RichTextBox1.Text <> "" Then
            rengl�n = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
            If RichTextBox1.GetLineFromChar(Len(RichTextBox1.Text)) + 1 = rengl�n Then
                Decir "final de la hoja. rengl�n " + Str(rengl�n)
            Else
                Decir "avanzando un p�rrafo. rengl�n " + Str(rengl�n)
            End If
        Else
            Decir "No se puede avanzar de a p�rrafo porque la hoja est� vac�a"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copiar
        If RichTextBox1.SelText <> "" Then
            Decir "se copi� el texto seleccionado. para pegarlo en otro lugar, usar control m�s ve corta"
        Else
            Decir "No se puede copiar porque no hay texto seleccionado. para seleccionar, usar shift m�s las flechas"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyZ Then 'deshacer
        Decir "deshaciendo la �ltima acci�n"
    End If

    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyX Then 'cortar
        If RichTextBox1.SelText <> "" Then
            Decir "se cort� el texto seleccionado. para pegarlo en otro lugar, usar control m�s ve corta"
        Else
            Decir "No se puede cortar porque no hay texto seleccionado. para seleccionar, usar shift m�s las flechas"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyV Then 'pegar
        If Clipboard.GetText <> "" Then
            rengl�n = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart) + 1
            Decir "texto pegado en el rengl�n " + Str(rengl�n)
        Else
            Decir "No se puede pegar porque no hay texto copiado o cortado. para copiar, usar control m�s ce. para cortar, usar control m�s equis"
        End If
    End If
      
    If shiftkey = vbAltMask And KeyCode = vbKeyDown Then 'leer todo el texto
        If Trim(RichTextBox1.Text) <> "" Then
            Decir "toda la hoja dice: " + RichTextBox1.Text
        Else
            Decir "No se puede leer todo el texto porque la hoja est� vac�a"
        End If
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyUp Then 'leer desde el cursor hacia adelante
        If Trim(RichTextBox1.Text) <> "" Then
            If RichTextBox1.SelStart = 0 Then
                Decir "desde donde est�s hasta el final de la hoja dice: " + Mid(RichTextBox1.Text, 1, Len(RichTextBox1.Text) - Len(Left(RichTextBox1.Text, RichTextBox1.SelStart))) 'leer desde el cursor hacia adelante
            Else
                Decir "desde donde est�s hasta el final de la hoja dice: " + Mid(RichTextBox1.Text, RichTextBox1.SelStart, Len(RichTextBox1.Text) - Len(Left(RichTextBox1.Text, RichTextBox1.SelStart))) 'leer desde el cursor hacia adelante
            End If
        Else
            Decir "No se puede leer desde donde est�s hasta el final porque la hoja est� vac�a"
        End If
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyLeft Then 'leer la oraci�n actual
        If Trim(RichTextBox1.Text) <> "" Then
            Decir "El rengl�n actual dice: " + decirOraci�nSiguiente(RichTextBox1)
        Else
            Decir "No se puede leer el rengl�n actual porque la hoja est� vac�a"
        End If
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyRight Then 'leer el texto seleccionado
        If Trim(RichTextBox1.SelText) <> "" Then
            If RichTextBox1.SelText = " " Then
                Decir "texto seleccionado: espacio"
            Else
                Decir "texto seleccionado: " + RichTextBox1.SelText
            End If
        Else
            Decir "No se puede leer la selecci�n porque no hay texto seleccionado"
        End If
    End If
    
    controlPresionado = False 'se resetea la variable
End Sub

'++++++++++++++++++++++++++++++
'para pegar im�genes
'++++++++++++++++++++++++++++++
Public Sub pegarImagen()
    Call pegarT�tuloImagen
    Clipboard.Clear
    Clipboard.SetData frmCuaderno.Picture1.Picture
    RichTextBox1.SetFocus
    SendKeys "^(V)"
    SendKeys "{enter}"
End Sub

Sub pegarT�tuloImagen()
    Dim ruta As String
    ruta = frmIm�genes.swImagenDevuelta
    ruta = Right(ruta, Len(ruta) - InStrRev(ruta, "\"))
    ruta = Left(ruta, InStr(ruta, ".") - 1) + Chr(13)
    Clipboard.SetText Chr(13) + "Imagen insertada: " + ruta
    RichTextBox1.SelText = Clipboard.GetText()
End Sub


Sub mostrarHuevo()
    Decir "Felicitaciones, encontraste este g�evo de pascua. Te regalo un chiste: se hab�an reunido una lechuguita, un tomatito y un g�evito. se acerca una manzana grande y le pregunta a todos: �qu� van a ser cuando sean grandes y la lechuguita responde, yo voy a ser un lechug�n. Y el tomatito responde, yo voy a ser un tomat�n. y el g�evito, triste, se puso a llorar. Fin de este g�evo de pascua."
End Sub

Public Sub borrarUnCar�cter()
    'Debug.Print Asc(Len(RichTextBox1.Text) - 1)
    RichTextBox1.SelStart = Len(RichTextBox1.Text) - 1 ' Comenzamos desde la cantidad de caracteres menos 1
    RichTextBox1.SelLength = 1 ' Con un maximo de un caracter.
    RichTextBox1.SelText = "" ' Borramos
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
    posX = Me.Width - btnConfiguraci�n.Width - 400
    posY = 360

    btnConfiguraci�n.Move posX, posY
    
    'las l�neas y etiqueta de arriba
    Line2.X2 = Me.Width - 400
    Line2.X1 = Line2.X2 - 480
    
    Label2.Left = Line2.X1 - 20 - Label2.Width
    
    Line1.X2 = Label2.Left - 20
    Line1.X1 = Line1.X2 - 480
    
    'el bot�n de abajo
    posX = Me.Width - ButtonTransparent1.Width - 400
    posY = Me.Height - ButtonTransparent1.Height - 700

    ButtonTransparent1.Move posX, posY
     
    Label1.Move Label1.Left, posY

     Alto_RTF = Me.ScaleHeight - 2000

     ' Esto chequea que el valor Height del text no sea negativo _
       ya que si no da error
     If Alto_RTF <= 0 Then Alto_RTF = 100 '- RichTextBox1.Top

     'Posiciona y redimensiona el rtf
     RichTextBox1.Move RichTextBox1.Left, RichTextBox1.Top, Me.ScaleWidth - 250 - RichTextBox1.Left, Alto_RTF

 End Sub

