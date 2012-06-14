VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmReproductorMúsica 
   Caption         =   "Reproductor MP3"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   Icon            =   "frmReproductorMúsica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmReproductorMúsica.frx":08CA
   ScaleHeight     =   8400
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   4680
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   6885
      Left            =   248
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.FileListBox File 
      Height          =   285
      Left            =   4440
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MediaPlayerCtl.MediaPlayer media 
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   7560
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscá con flecha arriba o abajo el archivo a abrir:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   285
      TabIndex        =   4
      Top             =   120
      Width           =   3510
   End
End
Attribute VB_Name = "frmReproductorMúsica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6
Dim Drivers() As String
Dim dóndeEstoyAhora As Byte
Dim sóloUnDriver As String
Dim rutaMisDoc As String
Dim rutaMiMúsica As String
Private Enum dóndeEstoy
    discos
    carpetas
'    miMúsica
End Enum

'Dim swEstoyEnMiMúsica As Boolean
Public swEstoyAbierto As Boolean
Dim swReproduciendo As Boolean
Dim swCambioTema As Boolean 'para dar play en un nuevo tema aunque esté sonando el anterior (sinó el botón toma que le tiene que dar stop al nuevo)
Dim swCambiarAutomático As Boolean
Public swPasarDelPrincipioAlFin As Boolean 'para ver si al reproducir el último tema se pasa al primero
Dim índicePrimerTema As Integer
Dim swCerréLaMúsicaDeFondo As Boolean 'para ver si el reproductor cerró la música de fondo, para que reinicie
Dim cantÍndicesConMisDocyMiMúsica As Byte
Dim swImposibleRetroceder As Boolean ' para ver si se puede retroceder con borrar


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyF7 Then
        If swCuadernoAbierto = True Then 'pasar al cuaderno
            Decir "Volviendo a tu carpeta, para regresar al reproductor, apretá F7"
            frmCuaderno.Show
        End If
        If frmPrincipal.swEstoyAbierto = True Then 'pasar al form principal
            Decir "Volviendo al menú principal, para regresar al reproductor, apretá F7"
            frmPrincipal.Show
        End If
    End If
    
    If KeyCode = vbKeyReturn Then List1_DblClick
    
    If shiftkey = vbShiftMask Then 'shift calla todo
        Decir ""
        media.Mute = Not media.Mute
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del reproductor
         frmAyuda.formulario = formularios.reproductorMúsica
         frmAyuda.Show
         Exit Sub
    End If
    
    If KeyCode = vbKey2 Then 'más volumen
        If media.Volume < 0 Then
            media.Volume = media.Volume + 50
            Decir "más volumen"
            'Debug.Print media.Volume
        Else
            Decir "El reproductor tiene el volumen al máximo"
        End If
    End If
    If KeyCode = vbKey1 Then 'menos volumen
        If media.Volume >= -2300 Then
            media.Volume = media.Volume - 50
            Decir "menos volumen"
            'Debug.Print media.Volume
        Else
            Decir "El reproductor tiene el volumen al mínimo"
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        If mensajeSalir("¿Estás seguro que querés cerrar el reproductor de música?") Then Unload Me
    End If
        
    If KeyCode = vbKeyBack Then 'ir a la carpeta anterior
        If swImposibleRetroceder = False Then
            List1.ListIndex = List1.ListCount - 2
            List1_DblClick
        Else
            Decir "imposible volver a la carpeta anterior porque llegaste a los discos de tu computadora, usá las flechas para ver cuál querés abrir, y aceptá con enter"
        End If
    End If
    
    If KeyCode = vbKey3 Then 'parar o reencender la música de fondo
        If swMúsicaDeFondo = True Then 'If frmOculto.media.PlayState = mpPlaying Then
            frmOculto.swContinuarReproducción = False
            frmOculto.media.Stop
            swMúsicaDeFondo = False
        Else
            frmOculto.swContinuarReproducción = True
            frmOculto.media.Play
            swMúsicaDeFondo = True
        End If
    End If
    
    If KeyCode = vbKeySpace Then 'pausar la reproducción
        If media.PlayState = mpPlaying Or media.PlayState = mpPaused Then
            If media.PlayState = mpPaused Then
                media.Play
                Decir "sacando la pausa"
            Else
                media.Pause
                Decir "pausa"
            End If
        Else
            Decir "no se puede pausar porque no se está reproduciendo música. Buscá el tema que quieras escuchar con las flechas y reproducilo con enter"
        End If
    End If
    
    If KeyCode = vbKeyDelete Then 'volver a los discos
        List1.ListIndex = List1.ListCount - 1
        List1_DblClick
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF1 Then 'leer la ayuda
         frmAyuda.formulario = formularios.reproductorMúsica
         frmAyuda.Show 1
         Exit Sub
    End If

    
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
End Sub

Private Sub Form_Load()
    'Dim lectorRegistro,
'    Dim versiónSistOp As String

    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    If frmOculto.media.PlayState = mpPlaying Then 'si está sonando la música de fondo
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "¿Querés que detenga la música de fondo de la mochila para que escuches bien la música del reproductor?"
        frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado Then
            swCerréLaMúsicaDeFondo = True
            Call Form_KeyDown(vbKey3, 0) 'si quiere para la música
        End If
    End If
    
    swEstoyAbierto = True
'    versiónSistOp = obtenerVersiónWindows
    File.Pattern = "*.mp3;*.wav;*.mid" ';*.wma"
'    Set lectorRegistro = CreateObject("WScript.Shell")
    'se ve cuál es la ruta de mis doc y mi música en el sistema para agregarla
'    rutaMisDoc = lectorRegistro.regRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal")
    rutaMisDoc = leerRegistro(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")
    rutaMiMúsica = leerRegistro(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "My Music")
'    If versiónSistOp = "windows xp" Then
'        rutaMiMúsica = lectorRegistro.regRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Music")
'    Else
'        rutaMiMúsica = rutaMisDoc + "\Mi música"
'    End If
'    Set lectorRegistro = Nothing
    Decir "Abriendo el reproductor de música, buscá con flecha arriba o abajo las carpetas o la música que quieras abrir"
    If cargarDrivers = False Then mensaje "Hubo un problema cargando las unidades del equipo"
    dóndeEstoyAhora = dóndeEstoy.discos 'está en los discos
    swCambioTema = False
    media.Volume = -1300
    swPasarDelPrincipioAlFin = True
    swImposibleRetroceder = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer, cadena As String
    shiftkey = Shift And 7
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        If List1.ListIndex = 0 Then 'si no se está arriba
            cadena = "principio de la lista, " + List1.List(List1.ListIndex)
        ElseIf List1.ListIndex = List1.ListCount - 1 Then 'si está abajo
            cadena = "final de la lista, " + List1.List(List1.ListIndex)
        Else 'cualquier otro caso
            cadena = List1.List(List1.ListIndex)
        End If
        If media.Mute = True Then
            If media.PlayState = mpPlaying Then
                cadena = cadena + ". El reproductor está corriendo un tema en silencio, para darle sonido otra vez, apretá shift"
            ElseIf media.PlayState = mpPaused Then
                cadena = cadena + ". El reproductor está silenciado y tiene un tema en pausa, para darle sonido otra vez apretá shift, y para sacar la pausa apretá espacio"
            End If
        End If
        Decir cadena
        swCambioTema = False
    End If
            
    If KeyCode = vbKeyM Or KeyCode = vbKeyC Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then Decir List1.List(List1.ListIndex) 'si se mueve por los archivos, carpetas, o con los controles de cursor
    If KeyCode = vbKeyEnd Then Decir "final de la lista. " + List1.List(List1.ListIndex)
    If KeyCode = vbKeyHome Then Decir "principio de la lista. " + List1.List(List1.ListIndex)
End Sub

'Private Sub Form_Paint()
'    List1.SetFocus
'End Sub

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
    
    If swCerréLaMúsicaDeFondo = True Then
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "¿Querés que vuelva a sonar la música de fondo que cerré cuando empezó el reproductor?"
        frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado Then Call Form_KeyDown(vbKey3, 0) 'si quiere reactivar la música
    End If
    
    If swCuadernoAbierto = True Then
        Decir "cerrando el reproductor. volviendo a tu carpeta"
    Else
        Decir "cerrando el reproductor. volviendo al menú principal"
        frmPrincipal.Show 'si no está abierto el cuaderno, se abre el principal
    End If
    
    'Call contarFormularios(False)
    swEstoyAbierto = False
End Sub

Private Sub List1_DblClick()
    Dim i As Integer, aux As String, carpetaAnterior As String, carpetaPrevia As String
    Static swIrADiscos As Boolean
    On Error GoTo manejoError:
    If List1.ListIndex <> -1 Then
        Select Case dóndeEstoyAhora
            Case dóndeEstoy.carpetas
                Select Case List1.List(List1.ListIndex)
                    Case "Cambiar de disco"
                        If cargarDrivers = False Then mensaje "Hubo un problema cargando las unidades del equipo"
                        dóndeEstoyAhora = dóndeEstoy.discos
                        Decir "Volviendo a los discos de tu computadora, buscalos con las flechas"
                        swImposibleRetroceder = True
                    Case "Volver a la carpeta anterior"
                        If swImposibleRetroceder = False Then 'si es posible retroceder
                            carpetaAnterior = Left(Dir1.path, InStrRev(Dir1.path, "\"))
                            If Len(carpetaAnterior) <= 3 Or Len(carpetaPrevia) <= 3 Then
                                If swIrADiscos = True Then
                                    If cargarDrivers = False Then mensaje "Hubo un problema cargando las unidades del equipo"
                                    dóndeEstoyAhora = dóndeEstoy.discos
                                    Decir "Volviendo a los discos de tu computadora"
                                    swImposibleRetroceder = True
                                Else
                                    Call cargarCarpetas(carpetaAnterior)
                                    carpetaPrevia = Left(Dir1.path, 1)
                                    Decir "Volviendo a la carpetas dentro del disco con letra: " + carpetaPrevia
                                    swImposibleRetroceder = False
                                End If
                                swIrADiscos = True
                            Else
                                Call cargarCarpetas(carpetaAnterior)
                                carpetaPrevia = Right(Dir1.path, Len(Dir1.path) - InStrRev(Dir1.path, "\"))
                                Decir "Volviendo a la carpeta: " + carpetaPrevia
                                swIrADiscos = False
                                swImposibleRetroceder = False
                            End If
                        Else
                            Decir "No se puede retroceder porque ya llegaste a los discos de tu computadora"
                        End If
                    Case Else
                        If List1.List(List1.ListIndex) <> "Aquí dentro no hay carpetas para mostrar" And _
                        List1.List(List1.ListIndex) <> "Tampoco hay música aquí dentro" And _
                        List1.List(List1.ListIndex) <> "Aquí dentro no hay música para mostrar" Then
                            swImposibleRetroceder = False
                            If chequearSiEsArchivo(List1.List(List1.ListIndex)) = True Then
    '                            archivoDevuelto = Dir1.Path
    '                            archivoDevuelto = Dir1.Path + "\" + Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 9)
                                carpetaPrevia = List1.List(List1.ListIndex)
                                If swCambioTema = False Then
                                    media.FileName = Dir1.path & "\" & Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 8)
                                    media.AutoStart = True
                                    swCambioTema = True
                                    swReproduciendo = True
                                    swCambiarAutomático = True
                                Else
                                    swCambiarAutomático = False
                                    media.Stop
                                    swReproduciendo = False
                                    swCambioTema = False
                                End If
                                If media.Mute = True Then
                                    If media.PlayState = mpPlaying Then
                                        Decir "se está reproduciendo un tema, pero el reproductor está en modo silencioso, para darle sonido otra vez, apretá shift"
                                    Else
                                        Decir "el reproductor está detenido, pero también está en modo silencioso, para darle sonido otra vez, apretá shift"
                                    End If
                                Else
                                    If media.PlayState = mpPlaying Then
                                        Decir "Reproduciendo el tema: " + Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 9) 'List1.List(List1.ListIndex)
                                    Else
                                        Decir "Tema detenido"
                                    End If
                                End If
                            Else
'                                If swEstoyEnMiMúsica = False Then
                                    carpetaPrevia = List1.List(List1.ListIndex)
'                                Else
                                    
                                Decir "Abriendo la " + List1.List(List1.ListIndex)
                                Call cargarCarpetas(Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 9))
                                swIrADiscos = False
                            End If
                        End If
                End Select
                
            Case dóndeEstoy.discos
                If List1.List(List1.ListIndex) = "Mis documentos" Then
                    Call cargarCarpetas(rutaMisDoc)
                    Decir "Abriendo mis documentos, usá las flechas para ver las carpetas y los archivos que contiene"
                    dóndeEstoyAhora = dóndeEstoy.carpetas
                    carpetaPrevia = "C"
                    swIrADiscos = True
                    swImposibleRetroceder = False
                ElseIf List1.List(List1.ListIndex) = "Mi música" Then
                    Call cargarCarpetas(rutaMiMúsica)
                    Decir "Abriendo mi música, usá las flechas para ver las carpetas y los archivos que contiene"
                    dóndeEstoyAhora = dóndeEstoy.carpetas
                    carpetaPrevia = "C"
                    swIrADiscos = True
                    swImposibleRetroceder = False
                Else
                    carpetaPrevia = List1.List(List1.ListIndex)
                    If Left(List1.List(List1.ListIndex), 1) = "D" Then 'si es un disco, que diga él, si no que diga la
                        Decir "Abriendo el " + List1.List(List1.ListIndex)
                    Else
                        Decir "Abriendo la " + List1.List(List1.ListIndex)
                    End If
                    Call cargarCarpetas(Drivers(List1.ListIndex - cantÍndicesConMisDocyMiMúsica))
                    swIrADiscos = True
                    dóndeEstoyAhora = dóndeEstoy.carpetas
                    swImposibleRetroceder = False
                End If
        End Select
    End If
    Exit Sub
manejoError:
    If Err.Number = 68 Then mensaje "El disco o disquete no está listo, puede que no haya un CD puesto o que aún no lo pueda leer la unidad"
    If Err.Number <> 68 Then mensaje Str(Err.Number) + " " + Err.Description
End Sub

Function chequearSiEsArchivo(quéCadena As String) As Boolean
    If Mid(quéCadena, Len(quéCadena) - 3, 1) = "." Then
        chequearSiEsArchivo = True
    Else
        chequearSiEsArchivo = False
    End If
End Function

'Function eliminarExtensiónArchivos(quéArchivo As String) As String
'    If swEliminarExtensión = True Then
'        If chequearSiEsArchivo(quéArchivo) = True Then
'            eliminarExtensiónArchivos = Left(quéArchivo, Len(quéArchivo) - 4)
'        End If
'    End If
'End Function

Function cargarDrivers() As Boolean
    Dim r As Integer, todosLosDrivers As String, posición As Double, tipoDeDrive As Integer, contador As Integer
    cantÍndicesConMisDocyMiMúsica = 0
    List1.Clear
    If existeCarpeta(rutaMisDoc) Then
        List1.AddItem "Mi música"
        cantÍndicesConMisDocyMiMúsica = cantÍndicesConMisDocyMiMúsica + 1
    End If
    If existeCarpeta(rutaMiMúsica) Then
        List1.AddItem "Mis documentos"
        cantÍndicesConMisDocyMiMúsica = cantÍndicesConMisDocyMiMúsica + 1
    End If
    
    todosLosDrivers = Space(64)
    r = GetLogicalDriveStrings(Len(todosLosDrivers), todosLosDrivers)
    todosLosDrivers = Left(todosLosDrivers, r)
    Do
        posición = InStr(todosLosDrivers, Chr$(0))
        If posición Then
            sóloUnDriver = Left(todosLosDrivers, posición)
            todosLosDrivers = Mid$(todosLosDrivers, posición + 1, Len(todosLosDrivers))
            tipoDeDrive = GetDriveType(sóloUnDriver)
            If tipoDeDrive = DRIVE_CDROM Then
                List1.AddItem "Unidad de CD, letra " + UCase(Left(sóloUnDriver, 1))
            ElseIf tipoDeDrive = DRIVE_FIXED Then
                List1.AddItem "Disco duro, letra " + UCase(Left(sóloUnDriver, 1))
            ElseIf tipoDeDrive = DRIVE_REMOVABLE Then
                If UCase(Left(sóloUnDriver, 1)) = "A" Or UCase(Left(sóloUnDriver, 1)) = "B" Then
                    List1.AddItem "Disco flexible, letra " + UCase(Left(sóloUnDriver, 1))
                Else
                    List1.AddItem "Disco extraíble, letra " + UCase(Left(sóloUnDriver, 1))
                End If
            ElseIf tipoDeDrive = DRIVE_REMOTE Then
                List1.AddItem "Disco remoto, unidad " + UCase(Left(sóloUnDriver, 1))
            ElseIf tipoDeDrive = DRIVE_RAMDISK Then
                List1.AddItem "Disco RAM, letra " + UCase(Left(sóloUnDriver, 1))
            End If
        End If
        ReDim Preserve Drivers(0 To contador)
        Drivers(contador) = sóloUnDriver
        contador = contador + 1
    Loop Until todosLosDrivers = ""
    cargarDrivers = True
    Exit Function
manejoError:
    cargarDrivers = False
End Function

Sub cargarCarpetas(quéDirectorio As String)
    Dim i As Integer
    
    Dir1.path = quéDirectorio
    File.path = Dir1.path
    Dir1.Refresh
    File.Refresh
    List1.Clear
'    List1.AddItem "Carpetas:"
    For i = 0 To Dir1.ListCount - 1
        List1.AddItem "Carpeta: " + Right(Dir1.List(i), Len(Dir1.List(i)) - InStrRev(Dir1.List(i), "\"))
    Next
    If Dir1.ListCount = 0 Then List1.AddItem "Aquí dentro no hay carpetas para mostrar"
'    List1.AddItem "Archivos:"
    For i = 0 To File.ListCount - 1
        List1.AddItem "Música: " + File.List(i)
        If i = 0 Then índicePrimerTema = List1.ListCount - 1
    Next
    If File.ListCount = 0 Then
        If Dir1.ListCount = 0 Then
            List1.AddItem "Tampoco hay música aquí dentro"
        Else
            List1.AddItem "Aquí dentro no hay música para mostrar"
        End If
    End If
    List1.AddItem "Volver a la carpeta anterior"
    List1.AddItem "Cambiar de disco"
    
End Sub

Sub mensaje(quéTexto)
    'se muestra un cartel que avisa que todo anduvo bien
    frmMsgBox.cadenaAMostrar = quéTexto
    frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
    frmMsgBox.Show 1
End Sub

Private Sub media_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    If OldState = mpPlaying And NewState = mpStopped And swCambiarAutomático = True Then
'        If media.PlayState = mpStopped And swCambiarAutomático = True Then
            If List1.ListIndex < List1.ListCount - 3 Then 'si no es el último tema
                swCambioTema = False
                List1.ListIndex = List1.ListIndex + 1
                List1_DblClick
            Else 'si es el último tema
                If swPasarDelPrincipioAlFin = True Then 'si se pasa del último tema al primero
                    swCambioTema = False
                    List1.ListIndex = índicePrimerTema
                    List1_DblClick
                End If
            End If
'        End If
    End If
End Sub
