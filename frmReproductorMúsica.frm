VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmReproductorM�sica 
   Caption         =   "Reproductor MP3"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   Icon            =   "frmReproductorM�sica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmReproductorM�sica.frx":08CA
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
      Caption         =   "Busc� con flecha arriba o abajo el archivo a abrir:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   285
      TabIndex        =   4
      Top             =   120
      Width           =   3510
   End
End
Attribute VB_Name = "frmReproductorM�sica"
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
Dim d�ndeEstoyAhora As Byte
Dim s�loUnDriver As String
Dim rutaMisDoc As String
Dim rutaMiM�sica As String
Private Enum d�ndeEstoy
    discos
    carpetas
'    miM�sica
End Enum

'Dim swEstoyEnMiM�sica As Boolean
Public swEstoyAbierto As Boolean
Dim swReproduciendo As Boolean
Dim swCambioTema As Boolean 'para dar play en un nuevo tema aunque est� sonando el anterior (sin� el bot�n toma que le tiene que dar stop al nuevo)
Dim swCambiarAutom�tico As Boolean
Public swPasarDelPrincipioAlFin As Boolean 'para ver si al reproducir el �ltimo tema se pasa al primero
Dim �ndicePrimerTema As Integer
Dim swCerr�LaM�sicaDeFondo As Boolean 'para ver si el reproductor cerr� la m�sica de fondo, para que reinicie
Dim cant�ndicesConMisDocyMiM�sica As Byte
Dim swImposibleRetroceder As Boolean ' para ver si se puede retroceder con borrar


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el men� de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyF7 Then
        If swCuadernoAbierto = True Then 'pasar al cuaderno
            Decir "Volviendo a tu carpeta, para regresar al reproductor, apret� F7"
            frmCuaderno.Show
        End If
        If frmPrincipal.swEstoyAbierto = True Then 'pasar al form principal
            Decir "Volviendo al men� principal, para regresar al reproductor, apret� F7"
            frmPrincipal.Show
        End If
    End If
    
    If KeyCode = vbKeyReturn Then List1_DblClick
    
    If shiftkey = vbShiftMask Then 'shift calla todo
        Decir ""
        media.Mute = Not media.Mute
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al men� de la aplicaci�n. Para leer los �tems de este men� necesit�s jaws u otro lector de pantallas. Para volver a la mochila, apret� escape"
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del reproductor
         frmAyuda.formulario = formularios.reproductorM�sica
         frmAyuda.Show
         Exit Sub
    End If
    
    If KeyCode = vbKey2 Then 'm�s volumen
        If media.Volume < 0 Then
            media.Volume = media.Volume + 50
            Decir "m�s volumen"
            'Debug.Print media.Volume
        Else
            Decir "El reproductor tiene el volumen al m�ximo"
        End If
    End If
    If KeyCode = vbKey1 Then 'menos volumen
        If media.Volume >= -2300 Then
            media.Volume = media.Volume - 50
            Decir "menos volumen"
            'Debug.Print media.Volume
        Else
            Decir "El reproductor tiene el volumen al m�nimo"
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        If mensajeSalir("�Est�s seguro que quer�s cerrar el reproductor de m�sica?") Then Unload Me
    End If
        
    If KeyCode = vbKeyBack Then 'ir a la carpeta anterior
        If swImposibleRetroceder = False Then
            List1.ListIndex = List1.ListCount - 2
            List1_DblClick
        Else
            Decir "imposible volver a la carpeta anterior porque llegaste a los discos de tu computadora, us� las flechas para ver cu�l quer�s abrir, y acept� con enter"
        End If
    End If
    
    If KeyCode = vbKey3 Then 'parar o reencender la m�sica de fondo
        If swM�sicaDeFondo = True Then 'If frmOculto.media.PlayState = mpPlaying Then
            frmOculto.swContinuarReproducci�n = False
            frmOculto.media.Stop
            swM�sicaDeFondo = False
        Else
            frmOculto.swContinuarReproducci�n = True
            frmOculto.media.Play
            swM�sicaDeFondo = True
        End If
    End If
    
    If KeyCode = vbKeySpace Then 'pausar la reproducci�n
        If media.PlayState = mpPlaying Or media.PlayState = mpPaused Then
            If media.PlayState = mpPaused Then
                media.Play
                Decir "sacando la pausa"
            Else
                media.Pause
                Decir "pausa"
            End If
        Else
            Decir "no se puede pausar porque no se est� reproduciendo m�sica. Busc� el tema que quieras escuchar con las flechas y reproducilo con enter"
        End If
    End If
    
    If KeyCode = vbKeyDelete Then 'volver a los discos
        List1.ListIndex = List1.ListCount - 1
        List1_DblClick
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF1 Then 'leer la ayuda
         frmAyuda.formulario = formularios.reproductorM�sica
         frmAyuda.Show 1
         Exit Sub
    End If

    
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
End Sub

Private Sub Form_Load()
    'Dim lectorRegistro,
'    Dim versi�nSistOp As String

    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    If frmOculto.media.PlayState = mpPlaying Then 'si est� sonando la m�sica de fondo
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "�Quer�s que detenga la m�sica de fondo de la mochila para que escuches bien la m�sica del reproductor?"
        frmMsgBox.swS�No�Aceptar = True 'se elige que sea cuadro s�-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado Then
            swCerr�LaM�sicaDeFondo = True
            Call Form_KeyDown(vbKey3, 0) 'si quiere para la m�sica
        End If
    End If
    
    swEstoyAbierto = True
'    versi�nSistOp = obtenerVersi�nWindows
    File.Pattern = "*.mp3;*.wav;*.mid" ';*.wma"
'    Set lectorRegistro = CreateObject("WScript.Shell")
    'se ve cu�l es la ruta de mis doc y mi m�sica en el sistema para agregarla
'    rutaMisDoc = lectorRegistro.regRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal")
    rutaMisDoc = leerRegistro(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")
    rutaMiM�sica = leerRegistro(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "My Music")
'    If versi�nSistOp = "windows xp" Then
'        rutaMiM�sica = lectorRegistro.regRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Music")
'    Else
'        rutaMiM�sica = rutaMisDoc + "\Mi m�sica"
'    End If
'    Set lectorRegistro = Nothing
    Decir "Abriendo el reproductor de m�sica, busc� con flecha arriba o abajo las carpetas o la m�sica que quieras abrir"
    If cargarDrivers = False Then mensaje "Hubo un problema cargando las unidades del equipo"
    d�ndeEstoyAhora = d�ndeEstoy.discos 'est� en los discos
    swCambioTema = False
    media.Volume = -1300
    swPasarDelPrincipioAlFin = True
    swImposibleRetroceder = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer, cadena As String
    shiftkey = Shift And 7
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        If List1.ListIndex = 0 Then 'si no se est� arriba
            cadena = "principio de la lista, " + List1.List(List1.ListIndex)
        ElseIf List1.ListIndex = List1.ListCount - 1 Then 'si est� abajo
            cadena = "final de la lista, " + List1.List(List1.ListIndex)
        Else 'cualquier otro caso
            cadena = List1.List(List1.ListIndex)
        End If
        If media.Mute = True Then
            If media.PlayState = mpPlaying Then
                cadena = cadena + ". El reproductor est� corriendo un tema en silencio, para darle sonido otra vez, apret� shift"
            ElseIf media.PlayState = mpPaused Then
                cadena = cadena + ". El reproductor est� silenciado y tiene un tema en pausa, para darle sonido otra vez apret� shift, y para sacar la pausa apret� espacio"
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
    
    If swCerr�LaM�sicaDeFondo = True Then
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "�Quer�s que vuelva a sonar la m�sica de fondo que cerr� cuando empez� el reproductor?"
        frmMsgBox.swS�No�Aceptar = True 'se elige que sea cuadro s�-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado Then Call Form_KeyDown(vbKey3, 0) 'si quiere reactivar la m�sica
    End If
    
    If swCuadernoAbierto = True Then
        Decir "cerrando el reproductor. volviendo a tu carpeta"
    Else
        Decir "cerrando el reproductor. volviendo al men� principal"
        frmPrincipal.Show 'si no est� abierto el cuaderno, se abre el principal
    End If
    
    'Call contarFormularios(False)
    swEstoyAbierto = False
End Sub

Private Sub List1_DblClick()
    Dim i As Integer, aux As String, carpetaAnterior As String, carpetaPrevia As String
    Static swIrADiscos As Boolean
    On Error GoTo manejoError:
    If List1.ListIndex <> -1 Then
        Select Case d�ndeEstoyAhora
            Case d�ndeEstoy.carpetas
                Select Case List1.List(List1.ListIndex)
                    Case "Cambiar de disco"
                        If cargarDrivers = False Then mensaje "Hubo un problema cargando las unidades del equipo"
                        d�ndeEstoyAhora = d�ndeEstoy.discos
                        Decir "Volviendo a los discos de tu computadora, buscalos con las flechas"
                        swImposibleRetroceder = True
                    Case "Volver a la carpeta anterior"
                        If swImposibleRetroceder = False Then 'si es posible retroceder
                            carpetaAnterior = Left(Dir1.path, InStrRev(Dir1.path, "\"))
                            If Len(carpetaAnterior) <= 3 Or Len(carpetaPrevia) <= 3 Then
                                If swIrADiscos = True Then
                                    If cargarDrivers = False Then mensaje "Hubo un problema cargando las unidades del equipo"
                                    d�ndeEstoyAhora = d�ndeEstoy.discos
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
                        If List1.List(List1.ListIndex) <> "Aqu� dentro no hay carpetas para mostrar" And _
                        List1.List(List1.ListIndex) <> "Tampoco hay m�sica aqu� dentro" And _
                        List1.List(List1.ListIndex) <> "Aqu� dentro no hay m�sica para mostrar" Then
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
                                    swCambiarAutom�tico = True
                                Else
                                    swCambiarAutom�tico = False
                                    media.Stop
                                    swReproduciendo = False
                                    swCambioTema = False
                                End If
                                If media.Mute = True Then
                                    If media.PlayState = mpPlaying Then
                                        Decir "se est� reproduciendo un tema, pero el reproductor est� en modo silencioso, para darle sonido otra vez, apret� shift"
                                    Else
                                        Decir "el reproductor est� detenido, pero tambi�n est� en modo silencioso, para darle sonido otra vez, apret� shift"
                                    End If
                                Else
                                    If media.PlayState = mpPlaying Then
                                        Decir "Reproduciendo el tema: " + Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 9) 'List1.List(List1.ListIndex)
                                    Else
                                        Decir "Tema detenido"
                                    End If
                                End If
                            Else
'                                If swEstoyEnMiM�sica = False Then
                                    carpetaPrevia = List1.List(List1.ListIndex)
'                                Else
                                    
                                Decir "Abriendo la " + List1.List(List1.ListIndex)
                                Call cargarCarpetas(Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 9))
                                swIrADiscos = False
                            End If
                        End If
                End Select
                
            Case d�ndeEstoy.discos
                If List1.List(List1.ListIndex) = "Mis documentos" Then
                    Call cargarCarpetas(rutaMisDoc)
                    Decir "Abriendo mis documentos, us� las flechas para ver las carpetas y los archivos que contiene"
                    d�ndeEstoyAhora = d�ndeEstoy.carpetas
                    carpetaPrevia = "C"
                    swIrADiscos = True
                    swImposibleRetroceder = False
                ElseIf List1.List(List1.ListIndex) = "Mi m�sica" Then
                    Call cargarCarpetas(rutaMiM�sica)
                    Decir "Abriendo mi m�sica, us� las flechas para ver las carpetas y los archivos que contiene"
                    d�ndeEstoyAhora = d�ndeEstoy.carpetas
                    carpetaPrevia = "C"
                    swIrADiscos = True
                    swImposibleRetroceder = False
                Else
                    carpetaPrevia = List1.List(List1.ListIndex)
                    If Left(List1.List(List1.ListIndex), 1) = "D" Then 'si es un disco, que diga �l, si no que diga la
                        Decir "Abriendo el " + List1.List(List1.ListIndex)
                    Else
                        Decir "Abriendo la " + List1.List(List1.ListIndex)
                    End If
                    Call cargarCarpetas(Drivers(List1.ListIndex - cant�ndicesConMisDocyMiM�sica))
                    swIrADiscos = True
                    d�ndeEstoyAhora = d�ndeEstoy.carpetas
                    swImposibleRetroceder = False
                End If
        End Select
    End If
    Exit Sub
manejoError:
    If Err.Number = 68 Then mensaje "El disco o disquete no est� listo, puede que no haya un CD puesto o que a�n no lo pueda leer la unidad"
    If Err.Number <> 68 Then mensaje Str(Err.Number) + " " + Err.Description
End Sub

Function chequearSiEsArchivo(qu�Cadena As String) As Boolean
    If Mid(qu�Cadena, Len(qu�Cadena) - 3, 1) = "." Then
        chequearSiEsArchivo = True
    Else
        chequearSiEsArchivo = False
    End If
End Function

'Function eliminarExtensi�nArchivos(qu�Archivo As String) As String
'    If swEliminarExtensi�n = True Then
'        If chequearSiEsArchivo(qu�Archivo) = True Then
'            eliminarExtensi�nArchivos = Left(qu�Archivo, Len(qu�Archivo) - 4)
'        End If
'    End If
'End Function

Function cargarDrivers() As Boolean
    Dim r As Integer, todosLosDrivers As String, posici�n As Double, tipoDeDrive As Integer, contador As Integer
    cant�ndicesConMisDocyMiM�sica = 0
    List1.Clear
    If existeCarpeta(rutaMisDoc) Then
        List1.AddItem "Mi m�sica"
        cant�ndicesConMisDocyMiM�sica = cant�ndicesConMisDocyMiM�sica + 1
    End If
    If existeCarpeta(rutaMiM�sica) Then
        List1.AddItem "Mis documentos"
        cant�ndicesConMisDocyMiM�sica = cant�ndicesConMisDocyMiM�sica + 1
    End If
    
    todosLosDrivers = Space(64)
    r = GetLogicalDriveStrings(Len(todosLosDrivers), todosLosDrivers)
    todosLosDrivers = Left(todosLosDrivers, r)
    Do
        posici�n = InStr(todosLosDrivers, Chr$(0))
        If posici�n Then
            s�loUnDriver = Left(todosLosDrivers, posici�n)
            todosLosDrivers = Mid$(todosLosDrivers, posici�n + 1, Len(todosLosDrivers))
            tipoDeDrive = GetDriveType(s�loUnDriver)
            If tipoDeDrive = DRIVE_CDROM Then
                List1.AddItem "Unidad de CD, letra " + UCase(Left(s�loUnDriver, 1))
            ElseIf tipoDeDrive = DRIVE_FIXED Then
                List1.AddItem "Disco duro, letra " + UCase(Left(s�loUnDriver, 1))
            ElseIf tipoDeDrive = DRIVE_REMOVABLE Then
                If UCase(Left(s�loUnDriver, 1)) = "A" Or UCase(Left(s�loUnDriver, 1)) = "B" Then
                    List1.AddItem "Disco flexible, letra " + UCase(Left(s�loUnDriver, 1))
                Else
                    List1.AddItem "Disco extra�ble, letra " + UCase(Left(s�loUnDriver, 1))
                End If
            ElseIf tipoDeDrive = DRIVE_REMOTE Then
                List1.AddItem "Disco remoto, unidad " + UCase(Left(s�loUnDriver, 1))
            ElseIf tipoDeDrive = DRIVE_RAMDISK Then
                List1.AddItem "Disco RAM, letra " + UCase(Left(s�loUnDriver, 1))
            End If
        End If
        ReDim Preserve Drivers(0 To contador)
        Drivers(contador) = s�loUnDriver
        contador = contador + 1
    Loop Until todosLosDrivers = ""
    cargarDrivers = True
    Exit Function
manejoError:
    cargarDrivers = False
End Function

Sub cargarCarpetas(qu�Directorio As String)
    Dim i As Integer
    
    Dir1.path = qu�Directorio
    File.path = Dir1.path
    Dir1.Refresh
    File.Refresh
    List1.Clear
'    List1.AddItem "Carpetas:"
    For i = 0 To Dir1.ListCount - 1
        List1.AddItem "Carpeta: " + Right(Dir1.List(i), Len(Dir1.List(i)) - InStrRev(Dir1.List(i), "\"))
    Next
    If Dir1.ListCount = 0 Then List1.AddItem "Aqu� dentro no hay carpetas para mostrar"
'    List1.AddItem "Archivos:"
    For i = 0 To File.ListCount - 1
        List1.AddItem "M�sica: " + File.List(i)
        If i = 0 Then �ndicePrimerTema = List1.ListCount - 1
    Next
    If File.ListCount = 0 Then
        If Dir1.ListCount = 0 Then
            List1.AddItem "Tampoco hay m�sica aqu� dentro"
        Else
            List1.AddItem "Aqu� dentro no hay m�sica para mostrar"
        End If
    End If
    List1.AddItem "Volver a la carpeta anterior"
    List1.AddItem "Cambiar de disco"
    
End Sub

Sub mensaje(qu�Texto)
    'se muestra un cartel que avisa que todo anduvo bien
    frmMsgBox.cadenaAMostrar = qu�Texto
    frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
    frmMsgBox.Show 1
End Sub

Private Sub media_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    If OldState = mpPlaying And NewState = mpStopped And swCambiarAutom�tico = True Then
'        If media.PlayState = mpStopped And swCambiarAutom�tico = True Then
            If List1.ListIndex < List1.ListCount - 3 Then 'si no es el �ltimo tema
                swCambioTema = False
                List1.ListIndex = List1.ListIndex + 1
                List1_DblClick
            Else 'si es el �ltimo tema
                If swPasarDelPrincipioAlFin = True Then 'si se pasa del �ltimo tema al primero
                    swCambioTema = False
                    List1.ListIndex = �ndicePrimerTema
                    List1_DblClick
                End If
            End If
'        End If
    End If
End Sub
