VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmLectorActividad 
   Caption         =   "Actividad del X/X/X"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   Icon            =   "frmLectorActividad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmLectorActividad.frx":08CA
   ScaleHeight     =   7335
   ScaleWidth      =   7950
   Begin MSComDlg.CommonDialog diálogo 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfLectorActividad 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11456
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmLectorActividad.frx":2922
   End
   Begin TransparentButton.ButtonTransparent btnImprimir 
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Caption         =   "Imprimir"
      EstiloDelBoton  =   4
      Picture         =   "frmLectorActividad.frx":29A5
      PictureHover    =   "frmLectorActividad.frx":327F
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
Attribute VB_Name = "frmLectorActividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public numMesParaAbrir As Byte
Public archivoParaLeer As String
Public díaAbierto As String
Dim ImpresoraRich As ImpresoraRTF 'la impresora del rtf
Dim swHuboCambio As Boolean

Private Sub ImprimirConCuadroDiálogo()
   ' El control CommonDialog se llama "dlgPrint".
    diálogo.CancelError = True
    On Error GoTo manejoErrorImpresora
    diálogo.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If rtfLectorActividad.SelLength = 0 Then
       diálogo.Flags = diálogo.Flags + cdlPDAllPages
    Else
       diálogo.Flags = diálogo.Flags + cdlPDSelection
    End If
    diálogo.ShowPrinter
    rtfLectorActividad.SelPrint diálogo.hDC
manejoErrorImpresora: 'si el error es distinto a haber hecho click en cancelar, se muestra un msg
'    If Err.Number <> 32755 Then MsgBox "La impresora no está lista para imprimir." + Chr(13) + "Por favor vuelva a intentar cuando esté lista.", , "Información"
    Exit Sub
End Sub

Private Sub btnImprimir_Click()
    Call Form_KeyDown(vbKeyF6, 0)
End Sub


Private Sub rtflectoractividad_GotFocus()
    Call reproducirForm(formularios.actividades)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then
        If mensajeSalir("¿Estás seguro que querés cerrar el lector de actividades?") Then
            Unload Me
        End If
    End If
    
    If KeyCode = vbKeyF2 Then
        If swCuadernoAbierto = True Then
            Decir "volviendo a tu carpeta de " + miMateria + ", acordate que la actividad sigue abierta para que puedas seguir trabajando con ella" ', para volver a la actividad abierta apretá f2"
            frmCuaderno.Show
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en el lector de actividades"
    If shiftkey = 0 And KeyCode = vbKeyF12 Then frmControlAlumno.Show
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.lectorActividad
         frmAyuda.Show
         Exit Sub
    End If
    
    If (shiftkey = vbCtrlMask And KeyCode = vbKeyP) Or KeyCode = vbKeyF6 Then 'imprimir con control + p
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "¿Realmente querés imprimir esta actividad?"
        frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado = True Then
            If swImprimirDirecto = True Then
                With ImpresoraRich
                     
                     'Valores
                     'Encabezado y pie de página
                     .Header = "Actividad de la carpeta de " + miMateria + " trabajada el día " + Format(Date, "dd/mm/yyyy") 'Text1
                     .Footer = "Trabajo realizado por " + Trim(nombreUsuario) 'Text2
                     
                     'Margenes
                     .MarginTop = 500 'Text3
                     .MarginLeft = 500 'Text4
                     .MarginRight = 500 'Text5
                     .MarginBottom = 500 'Text6
                     
                     'Imprimir el RichTextBox pasado como parámetro
                     .Imprimir rtfLectorActividad
                
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
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swActividadAbierta = True
    rtfLectorActividad.Locked = True
    
    Decir "abriste la " + díaAbierto + ". para leerla, usá las flechas." '  para pasar a tu carpeta de " + miMateria + ", apretá f2"
    Call cargarActividad
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
    
    If swCuadernoAbierto = False Then
        If swActividadDeHoy = True Then
            frmActividadesHoy.Show
        Else
            frmActAntFut.Show
        End If
    End If
    
    'Call contarFormularios(False)
    swActividadAbierta = False
    If swCuadernoAbierto = True Then Decir "Cerrando el lector de actividades. Estás de nuevo en tu cuaderno" 'callar la voz si se vuelve al cuaderno
End Sub

Private Sub rtfLectorActividad_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 32 And KeyAscii <= 126 Then Decir "No se puede escribir en una actividad, sólo leer. Para escribir apretá f2 así pasás a tu carpeta"
End Sub


Private Sub rtflectoractividad_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    Dim renglón As Long
    
    shiftkey = Shift And 7
       
    If shiftkey = vbCtrlMask And KeyCode = vbKeyLeft Then 'leer por palabras retrocediendo
        Decir decirPalabraSiguiente(rtfLectorActividad) 'cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyRight Then 'leer por palabras avanzando
        Decir decirPalabraSiguiente(rtfLectorActividad) 'cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyHome Then 'control + inicio
        If rtfLectorActividad <> "" Then
            Decir decirPalabraSiguiente(rtfLectorActividad)
        Else
            Decir "La actividad está en blanco, no hay nada escrito"
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyHome Then 'tecla inicio
        renglón = rtfLectorActividad.GetLineFromChar(rtfLectorActividad.SelStart) + 1
        Decir "principio del renglón " + Str(renglón)
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyEnd Then 'control + fin
'        If swVolviendodeBraille = False Then 'si no se dispara el evento al volver del teclado braille
            If rtfLectorActividad <> "" Then
                Decir "final de la actividad. Estás detrás de la palabra " + decirPalabraAnterior(rtfLectorActividad)
            Else
                Decir "La actividad está en blanco, no hay nada escrito"
            End If
'        End If
        Exit Sub
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyEnd Then 'tecla fin
        renglón = rtfLectorActividad.GetLineFromChar(rtfLectorActividad.SelStart) + 1
        Decir "final del renglón " + Str(renglón)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyPageDown Then 'tecla avance de página
        renglón = rtfLectorActividad.GetLineFromChar(rtfLectorActividad.SelStart) + 1
        Decir "saltando hacia adelante al renglón " + Str(renglón)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyPageUp Then 'tecla retroceso de página
        renglón = rtfLectorActividad.GetLineFromChar(rtfLectorActividad.SelStart) + 1
        Decir "saltando hacia atrás al renglón " + Str(renglón)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyRight Then 'avanzar de a caracter
        Decir decirLetraSiguiente(rtfLectorActividad)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyLeft Then 'retroceder de a caracter
        If rtfLectorActividad.SelStart = 0 And rtfLectorActividad.Text <> "" Then
            Decir "Estás en el principio de la actividad, delante de la letra " + decirLetraSiguiente(rtfLectorActividad)
        Else
            Decir decirLetraSiguiente(rtfLectorActividad)
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyDown Then 'leer por oración
        Decir decirOraciónSiguiente(rtfLectorActividad)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyUp Then 'leer por oración
        Decir decirOraciónSiguiente(rtfLectorActividad)
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

    Call evaluarSelección(rtfLectorActividad, control, shift2, teclaApretada) 'se ve si hay selección
    
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyUp Then 'retroceder de a párrafo
        If rtfLectorActividad.Text <> "" Then
            renglón = rtfLectorActividad.GetLineFromChar(rtfLectorActividad.SelStart) + 1
            If renglón = 1 Then
                Decir "principio de la hoja, renglón 1"
            Else
                Decir "retrocediendo un párrafo. renglón " + Str(renglón)
            End If
        Else
            Decir "No se puede retroceder de a párrafo porque la actividad está vacía"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyDown Then 'avanzar de a párrafo
        If rtfLectorActividad.Text <> "" Then
            renglón = rtfLectorActividad.GetLineFromChar(rtfLectorActividad.SelStart) + 1
            If rtfLectorActividad.GetLineFromChar(Len(rtfLectorActividad.Text)) + 1 = renglón Then
                Decir "final de la actividad. renglón " + Str(renglón)
            Else
                Decir "avanzando un párrafo. renglón " + Str(renglón)
            End If
        Else
            Decir "No se puede avanzar de a párrafo porque la actividad está vacía"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copiar
        If rtfLectorActividad.SelText <> "" Then
            Decir "se copió el texto seleccionado. para pegarlo en otro lugar, usar control más ve corta"
        Else
            Decir "No se puede copiar porque no hay texto seleccionado. para seleccionar, usar shift más las flechas"
        End If
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyRight Then 'leer el texto seleccionado
        If Trim(rtfLectorActividad.SelText) <> "" Then
            If rtfLectorActividad.SelText = " " Then
                Decir "texto seleccionado: espacio"
            Else
                Decir "texto seleccionado: " + rtfLectorActividad.SelText
            End If
        Else
            Decir "No se puede leer la selección porque no hay texto seleccionado"
        End If
    End If

    If shiftkey = vbAltMask And KeyCode = vbKeyDown Then 'leer todo el texto
        If Trim(rtfLectorActividad.Text) <> "" Then
            Decir rtfLectorActividad.Text
        Else
            Decir "No se puede leer todo el texto porque la actividad está vacía"
        End If
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyUp Then 'leer desde el cursor hacia adelante
        If Trim(rtfLectorActividad.Text) <> "" Then
            If rtfLectorActividad.SelStart = 0 Then
                Decir Mid(rtfLectorActividad.Text, 1, Len(rtfLectorActividad.Text) - Len(Left(rtfLectorActividad.Text, rtfLectorActividad.SelStart))) 'leer desde el cursor hacia adelante
            Else
                Decir Mid(rtfLectorActividad.Text, rtfLectorActividad.SelStart, Len(rtfLectorActividad.Text) - Len(Left(rtfLectorActividad.Text, rtfLectorActividad.SelStart))) 'leer desde el cursor hacia adelante
            End If
        Else
            Decir "No se puede leer todo el texto porque la actividad está vacía"
        End If
    End If
End Sub

Public Sub cargarActividad()
    If swActividadDeHoy = True Then
        Me.Caption = "Actividad del día de hoy"
    Else
        Me.Caption = díaAbierto
    End If
    
    sonido = sndPlaySound(App.path + "\sonidos\abrir.wav", SND_ASYNC)
    rtfLectorActividad.LoadFile App.path + dirTrabajo + "actividades\" + archivoParaLeer
    
    'se actualizan las fuentes y los colores de los rtf del programa
    rtfLectorActividad.Font.Name = NombreFuente  'se ajusta la fuente del programa
    rtfLectorActividad.AutoVerbMenu = True
    'Selecciona todo
    rtfLectorActividad.SelStart = 0
    rtfLectorActividad.SelLength = Len(rtfLectorActividad)
    rtfLectorActividad.SelColor = colorFuente 'se ajusta el color de la fuente del programa
    rtfLectorActividad.SelLength = 0
    rtfLectorActividad.Font.Size = tamañoFuente 'se ajusta el tamaño de la fuente
    rtfLectorActividad.BackColor = colorFondo 'el color de fondo del rtf
End Sub

