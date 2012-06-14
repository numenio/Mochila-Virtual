VERSION 5.00
Begin VB.Form frmDiálogoAbrir 
   Caption         =   "Abrir Archivo"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   Icon            =   "frmCuadroDiálogoAbrir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCuadroDiálogoAbrir.frx":08CA
   ScaleHeight     =   7770
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File 
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   6885
      Left            =   195
      TabIndex        =   2
      Top             =   600
      Width           =   5535
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscá con flecha arriba o abajo el archivo a abrir:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3510
   End
End
Attribute VB_Name = "frmDiálogoAbrir"
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
Dim cantÍndicesConMisDoc As Byte
Private Enum dóndeEstoy
    discos
    carpetas
'    misDocumentos
End Enum
Public quéArchivosFiltrar As String
Public archivoDevuelto As String
Dim swImposibleRetroceder As Boolean 'para ver si se puede retroceder con la tecla borrar

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift And 7 = vbCtrlMask Then Decir ""
    If Shift And 7 = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
'    Dim lectorRegistro

    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    File.Pattern = quéArchivosFiltrar '"*.mp3;*.wav;*.mid;*.wma"
    archivoDevuelto = ""
'    Set lectorRegistro = CreateObject("WScript.Shell")
'    'se ve cuál es la ruta de mis doc en el sistema para agregarla
'    rutaMisDoc = lectorRegistro.regRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal")
    rutaMisDoc = leerRegistro(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")
'    Set lectorRegistro = Nothing
    Decir "Entrando en el cuadro para abrir archivos, buscá con flecha arriba o abajo las carpetas o el archivo que quieras abrir" '. Estás en mis documentos"
    If cargarDrivers = False Then mensaje "Hubo un problema cargando las unidades del equipo"
'    If List1.ListCount <> 0 Then List1.ListIndex = 0
    dóndeEstoyAhora = dóndeEstoy.discos 'está en los discos
    swImposibleRetroceder = True
End Sub

Private Sub Form_Paint()
    List1.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If List1.ListIndex = 0 Then 'si no se está arriba
            Decir "principio de la lista, " + List1.List(List1.ListIndex)
        ElseIf List1.ListIndex = List1.ListCount - 1 Then 'si está abajo
            Decir "final de la lista, " + List1.List(List1.ListIndex)
        Else 'cualquier otro caso
            Decir List1.List(List1.ListIndex)
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta"
    
    If KeyCode = vbKeyEscape Then Unload Me
    
    If KeyCode = vbKeyReturn Then List1_DblClick
    
    If KeyCode = vbKeyBack Then 'ir a la carpeta anterior
        If swImposibleRetroceder = False Then
            List1.ListIndex = List1.ListCount - 2
            List1_DblClick
        Else
            Decir "imposible volver a la carpeta anterior porque llegaste a los discos de tu computadora, elegí con las flechas qué disco querés abrir y aceptá con enter"
        End If
    End If
    
    If KeyCode = vbKeyDelete Then 'volver a los discos
        List1.ListIndex = List1.ListCount - 1
        List1_DblClick
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda
         frmAyuda.formulario = formularios.diálogoAbrir
         frmAyuda.Show 1
         Exit Sub
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    
    If KeyCode = vbKeyA Or KeyCode = vbKeyC Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then Decir List1.List(List1.ListIndex) 'si se mueve por los archivos, carpetas, o con los controles de cursor
    If KeyCode = vbKeyEnd Then Decir "final de la lista. " + List1.List(List1.ListIndex)
    If KeyCode = vbKeyHome Then Decir "principio de la lista. " + List1.List(List1.ListIndex)
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
    'Call contarFormularios(False)
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
                        If swImposibleRetroceder = False Then 'si se puede retroceder
                            carpetaAnterior = Left(Dir1.path, InStrRev(Dir1.path, "\"))
                            If Len(carpetaAnterior) <= 3 Then
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
                            Decir "Imposible ir a una carpeta anterior porque llegaste a los discos de tu computadora"
                        End If
                    Case Else
                        If List1.List(List1.ListIndex) <> "Aquí dentro no hay carpetas para mostrar" And _
                        List1.List(List1.ListIndex) <> "Tampoco hay ningún archivo aquí dentro" And _
                        List1.List(List1.ListIndex) <> "Aquí dentro no hay archivos para mostrar" Then
                            If chequearSiEsArchivo(List1.List(List1.ListIndex)) = True Then
    '                            archivoDevuelto = Dir1.Path
                                archivoDevuelto = Dir1.path + "\" + Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 9)
                                Decir "Abriendo el " + List1.List(List1.ListIndex)
                                carpetaPrevia = List1.List(List1.ListIndex)
                                Unload Me
                                Exit Sub
                            Else
                                carpetaPrevia = List1.List(List1.ListIndex)
                                Decir "Abriendo la " + List1.List(List1.ListIndex)
                                Call cargarCarpetas(Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 9))
                                swIrADiscos = False
                            End If
                            swImposibleRetroceder = False
                        End If
                End Select
                
            Case dóndeEstoy.discos
                If List1.List(List1.ListIndex) = "Mis documentos" Then
                    Call cargarCarpetas(rutaMisDoc)
                    Decir "Abriendo mis documentos, usá las flechas para ver las carpetas y los archivos que contiene"
                    dóndeEstoyAhora = dóndeEstoy.carpetas
                    swImposibleRetroceder = False
                Else
                    carpetaPrevia = List1.List(List1.ListIndex)
                    If Left(List1.List(List1.ListIndex), 1) = "D" Then 'si es un disco, que diga él, si no que diga la
                        Decir "Abriendo el " + List1.List(List1.ListIndex)
                    Else
                        Decir "Abriendo la " + List1.List(List1.ListIndex)
                    End If
                    Call cargarCarpetas(Drivers(List1.ListIndex - cantÍndicesConMisDoc))
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
    
    cantÍndicesConMisDoc = 0
    List1.Clear
    If existeCarpeta(rutaMisDoc) Then
        List1.AddItem "Mis documentos"
        cantÍndicesConMisDoc = cantÍndicesConMisDoc + 1
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
        List1.AddItem "Archivo: " + File.List(i)
    Next
    If File.ListCount = 0 Then
        If Dir1.ListCount = 0 Then
            List1.AddItem "Tampoco hay ningún archivo aquí dentro"
        Else
            List1.AddItem "Aquí dentro no hay archivos para mostrar"
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

