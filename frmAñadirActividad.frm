VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAñadirActividad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Añadir actividad"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
   Icon            =   "frmAñadirActividad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAñadirActividad.frx":08CA
   ScaleHeight     =   7665
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnEliminar 
      Caption         =   "Eliminar la actividad"
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton btnModificar 
      Caption         =   "Abrir la actividad"
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   6600
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   7200
      TabIndex        =   13
      Top             =   960
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog diálogo 
      Left            =   3360
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo12 
      Height          =   315
      ItemData        =   "frmAñadirActividad.frx":29F9
      Left            =   4320
      List            =   "frmAñadirActividad.frx":29FB
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Añadir"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Buscar archivo"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   310
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1560
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   61341696
      CurrentDate     =   39582
   End
   Begin RichTextLib.RichTextBox rtfActividades 
      Height          =   4455
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAñadirActividad.frx":29FD
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      X1              =   7080
      X2              =   7080
      Y1              =   120
      Y2              =   7440
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ELIJA UNA ACTIVIDAD DE LA LISTA PARA VERLE SUS DETALLES:"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7200
      TabIndex        =   14
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de la actividad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   375
      Width           =   1695
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Materia a que pertenece:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4320
      TabIndex        =   10
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Tema de la actividad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba abajo la actividad a agregar, o haga clic aquí ---------------->"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Puede añadir un comentario aquí:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
End
Attribute VB_Name = "frmAñadirActividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public díaParaCargarActividades As Date
Public swEditarActividades As Boolean
Public materia As String
Public swCargarFecha As Boolean
Dim trabajos() As String
Dim miRegistro As DatosActividad, directorio As String
Dim swCambióActividad As Boolean 'para ver si se ha editado una actividad en la fecha o materia

Private Sub btnModificar_Click()
    Dim i As Integer
    If List1.ListIndex <> -1 Then
        DTPicker1.Value = díaParaCargarActividades
        rtfActividades.LoadFile App.path + directorio + "\actividades\" + Trim(Str(Mid(Format(díaParaCargarActividades), 4, 2))) + "\" + trabajos(List1.ListIndex)
        
        Open App.path + directorio + "\actividades\" + Trim(Str(Mid(Format(díaParaCargarActividades), 4, 2))) + "\datosActividades\" + Left(trabajos(List1.ListIndex), Len(trabajos(List1.ListIndex)) - 4) + ".gui" For Random As #2 Len = Len(miRegistro)
        Get #2, 1, miRegistro   ' Lee el regitro
        Close #2   ' Cierra el archivo.
                    
        If Asc(Left(miRegistro.tema, 1)) Then 'se carga el tema
            Text1 = Trim(miRegistro.tema)
        Else
            Text1 = "Sin tema"
        End If
        
        If Asc(Left(miRegistro.comentarios, 1)) Then  'se carga el comentario
            Text4 = Trim(miRegistro.comentarios)
        Else
            Text4 = "Sin comentarios"
        End If
        
        For i = 0 To Combo12.ListCount - 1 'se carga la materia
            If Combo12.List(i) = materia Then
                Combo12.ListIndex = i
                Exit For
            End If
        Next
    End If
End Sub

Private Sub Combo12_Click()
    swCambióActividad = True
End Sub

Private Sub btnEliminar_Click()
    If List1.ListIndex <> -1 Then
        frmMsgBox.cadenaAMostrar = "¿Realmente desea eliminar la actividad seleccionada?"
        frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado Then Call eliminarActividad(trabajos(List1.ListIndex))
        Text1 = ""
        rtfActividades.Text = ""
        Text4 = ""
        Call cargarActividades 'se actualizan las actividades
        btnModificar_Click 'se carga la primer actividad
        swCambióActividad = False
        If frmCalendario.swEstoyAbierto Then frmCalendario.actualizarCalendario
    End If
End Sub

Private Sub Command8_Click() 'abrir un archivo para añadir una actividad
    ' Establecer CancelError a True
    diálogo.CancelError = True
    On Error GoTo ErrHandler
    diálogo.Filter = "Archivos de Texto (*.txt; *.rtf)|*.txt;*.rtf|"
    diálogo.ShowOpen
    rtfActividades.LoadFile diálogo.FileName
    Exit Sub
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub

Private Sub Command3_Click() 'añadir una actividad
    Dim materia As String, contador As Integer, i As Integer, prefijo As String
    Dim fecha As String, mes As String, año As String
    Dim posiciónDelReemplazo As Long
    Dim actividad As DatosActividad

    
    If Trim(rtfActividades.Text = "") Then
'        MsgBox "La actividad que intenta añadir está vacía. No olvide escribirla en el cuadro de texto.", , "¡Atención!"
        frmMsgBox.cadenaAMostrar = "La actividad que intenta añadir está vacía. No olvide escribirla en el cuadro de texto."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    If Trim(Text1.Text = "") Then
'        MsgBox "Hay que escribirle un tema a la actividad.", , "¡Atención!"
        frmMsgBox.cadenaAMostrar = "Hay que escribirle un tema a la actividad."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    If Combo12.Text = "" Then
'        MsgBox "Hay que elegir una materia a la que pertenece la actividad.", , "¡Atención!"
        frmMsgBox.cadenaAMostrar = "Hay que elegir una materia a la que pertenece la actividad."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    materia = Combo12.Text 'se llena la materia
    
    File1.Refresh
    DTPicker1.Format = dtpShortDate
    fecha = Str(DTPicker1.Value) 'se llena fecha
    Do  'se cambian los / por - para que se pueda guardar el archivo
        posiciónDelReemplazo = InStr(fecha, "/")
        Mid(fecha, posiciónDelReemplazo) = "-"
        posiciónDelReemplazo = InStr(fecha, "/")
    Loop Until posiciónDelReemplazo = 0
    'DTPicker1.Format = dtpLongDate
    mes = Mid(fecha, 4, 2)
    año = Right(fecha, 4)
    
    
    If swEditarActividades = False Then 'si se está añadiendo una actividad
        File1.path = App.path + "\trabajos\" + materia + "\actividades\" + Trim(Str(CInt(mes))) 'se ve si no hay ya otro archivo con el mismo nombre
        contador = 0
        For i = 0 To File1.ListCount - 1 'se evalúan todos los archivos de la materia sin su prefijo
            If Left(Right(File1.List(i), Len(File1.List(i)) - cantPrefijo), Len(Right(File1.List(i), Len(File1.List(i)) - cantPrefijo)) - 4) = Left(fecha, 2) + "-" + año Then contador = contador + 1
        Next
        
        If contador < 10 Then prefijo = "10" + Trim(Str(contador))
        If contador >= 10 And contador < 100 Then prefijo = "1" + Trim(Str(contador))
        If contador >= 100 Then prefijo = Trim(Str(contador))
    Else
        If swCambióActividad = True Then
            File1.path = App.path + "\trabajos\" + materia + "\actividades\" + Trim(Str(CInt(mes))) 'se ve si no hay ya otro archivo con el mismo nombre
            contador = 0
            For i = 0 To File1.ListCount - 1 'se evalúan todos los archivos de la materia sin su prefijo
                If Left(Right(File1.List(i), Len(File1.List(i)) - cantPrefijo), Len(Right(File1.List(i), Len(File1.List(i)) - cantPrefijo)) - 4) = Left(fecha, 2) + "-" + año Then contador = contador + 1
            Next
            
            If contador < 10 Then prefijo = "10" + Trim(Str(contador))
            If contador >= 10 And contador < 100 Then prefijo = "1" + Trim(Str(contador))
            If contador >= 100 Then prefijo = Trim(Str(contador))
            'como se copió la actividad en otra materia o fecha, se borra la original
            eliminarActividad (trabajos(List1.ListIndex))
        Else
            prefijo = Left(trabajos(List1.ListIndex), cantPrefijo)
        End If
    End If
    
    rtfActividades.SaveFile App.path + "\trabajos\" + materia + "\actividades\" + Trim(Str(CInt(mes))) + "\" + prefijo + Left(fecha, 2) + "-" + año + ".rtf"
'    Call chequearEspacioEnDisco(Left(App.Path, 2))
    With actividad
        .tema = Trim(Text1)
        .comentarios = Trim(Text4)
    End With
    
    Call guardarDatosActividad(actividad, App.path + "\trabajos\" + materia + "\actividades\" + Trim(Str(CInt(mes))) + "\datosActividades\" + prefijo + Left(fecha, 2) + "-" + año + ".gui")
    
    'se resetean todos los campos de cargar actividades
    DTPicker1.Value = Date 'se le da al calendario la fecha de hoy
    Text1 = ""
    Combo12.Refresh
    
    Text4 = ""
    rtfActividades.Text = ""
    Text1.SetFocus
    'se muestra un cartel que avisa que todo anduvo bien
    If swEditarActividades = True Then
        frmMsgBox.cadenaAMostrar = "La actividad se modificó exitosamente"
    Else
        frmMsgBox.cadenaAMostrar = "La actividad se guardó exitosamente"
    End If
    frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
    frmMsgBox.Show 1
    swCambióActividad = False
    If frmCalendario.swEstoyAbierto = True Then frmCalendario.actualizarCalendario
    If swEditarActividades = False Then
        Unload Me 'se descarga el form
        Exit Sub
    End If
    Call cargarActividades 'se actualizan las actividades
    
    If List1.ListCount > 0 Then 'si hay alguna actividad en la lista
        btnModificar_Click 'se carga la primer actividad
    Else
        Unload Me 'se descarga el form
    End If
End Sub


Private Sub DTPicker1_Change()
    swCambióActividad = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    
    If KeyCode = vbKeyF1 Then ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/Añadir actividad.htm", "", 1
End Sub

Private Sub Form_Load()
    'Call contarFormularios(True)
    Call llenarComboMaterias(Combo12)
    directorio = "\trabajos\" + materia
    If swCargarFecha = False Then 'si no viene una fecha del calendario
        DTPicker1.Value = Date 'se le da al calendario la fecha de hoy
    Else
        DTPicker1.Value = díaParaCargarActividades
    End If
'    Label1 = "Actividades guardadas del día " + Format(díaParaCargarActividades) + ":"
    If swEditarActividades = True Then
        Me.Width = 10035 'se estira el form
        Command3.Caption = "Guardar los cambios"
        Call cargarActividades 'se cargan las actividades del día seleccionado
        btnModificar_Click 'se carga la actividad en el form
    Else
        Me.Width = 7140 'se encoje el form
    End If
    swCambióActividad = False
    Call centrarFormulario(Me)
End Sub

Sub cargarActividades()
    Dim i As Integer, contador As Integer, cadenaAux As String, mes As String, cadena As String
    List1.Clear
    File1.Refresh
    File1.path = App.path + directorio + "\actividades\" + Trim(Str(Mid(Format(díaParaCargarActividades), 4, 2)))
    contador = 0
    For i = 0 To (File1.ListCount - 1) 'se examinan todos los archivos
        cadena = Left(Right(File1.List(i), Len(File1.List(i)) - cantPrefijo), Len(Right(File1.List(i), Len(File1.List(i)) - cantPrefijo)) - 4)
        If Left(cadena, 2) = Left(Format(díaParaCargarActividades), 2) And Right(cadena, 4) = Right(Format(díaParaCargarActividades), 4) Then 'si el archivo es del día seleccionado
            contador = contador + 1
            cadenaAux = "Actividad " + Trim(Str(contador))
            
            Open App.path + directorio + "\actividades\" + Trim(Str(Mid(Format(díaParaCargarActividades), 4, 2))) + "\datosActividades\" + Left(File1.List(i), Len(File1.List(i)) - 4) + ".gui" For Random As #2 Len = Len(miRegistro)
            Get #2, 1, miRegistro   ' Lee el regitro
            Close #2   ' Cierra el archivo.
                
            If Asc(Left(miRegistro.tema, 1)) Then
                cadenaAux = cadenaAux + ". Tema: " + Trim(miRegistro.tema) + "."
            Else
                cadenaAux = cadenaAux + ". Sin tema."
            End If
            
            List1.AddItem cadenaAux
            ReDim Preserve trabajos(0 To contador - 1)
            trabajos(contador - 1) = File1.List(i)
        End If
    Next
    If List1.ListCount > 0 Then 'si hay alguna actividad en la lista
        List1.ListIndex = 0
    Else
        Me.Width = 7140 'se encoje el form
    End If
End Sub

Private Sub guardarDatosActividad(actividadActual As DatosActividad, nombreArchivo As String)
    Dim archivolibre As Byte, carpetaActividad As String
    
    carpetaActividad = Left(nombreArchivo, InStrRev(nombreArchivo, "\"))
    
    If existeCarpeta(carpetaActividad) Then
        archivolibre = FreeFile 'se abre el archivo para guardar los datos de las partidas
        Open nombreArchivo For Random As archivolibre Len = Len(actividadActual)
        Put archivolibre, 1, actividadActual
        Close archivolibre
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    swEditarActividades = False
    'Call contarFormularios(False)
End Sub


''Sub cargarDatosActividad(quéActividad As String)
'    Open App.Path + directorio + "\actividades\" + Trim(Str(Mid(Format(díaParaCargarActividades), 4, 2))) + "\datosActividades\" + Left(File1.List(i), Len(File1.List(i)) - 4) + ".gui" For Random As #2 Len = Len(miRegistro)
'    Get #2, 1, miRegistro   ' Lee el regitro
'    Close #2   ' Cierra el archivo.
'End Sub
'
Sub eliminarActividad(actividad As String)
    Kill App.path + directorio + "\actividades\" + Trim(Str(Mid(Format(díaParaCargarActividades), 4, 2))) + "\" + actividad
    Kill App.path + directorio + "\actividades\" + Trim(Str(Mid(Format(díaParaCargarActividades), 4, 2))) + "\datosActividades\" + Left(actividad, Len(actividad) - 4) + ".gui"
End Sub

Private Sub List1_DblClick()
    btnModificar_Click
End Sub
