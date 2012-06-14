Attribute VB_Name = "declaraciones"
Option Explicit

Public swEmpezóLaMochila As Boolean 'para ver si empezó el programa, por la música de fondo que se carga en el form control
'Public formulariosAbiertos As Byte 'para contar los forms, para ver si sólo queda el oculto y así terminar el programa
Public swMúsicaDeFondo As Boolean 'para ver si suena la música de fondo
'Public horaPrimerRecordatorio As Date
Public recordatoriosActivos() As Recordatorio
Public swActividadDeHoy As Boolean 'para saber si el lector de actividades abrió una de hoy, o ant ó fut
Public swActividadAnterior As Boolean 'para saber si es anterior o futura
Public swCuadernoAbierto As Boolean 'para que al estar el cuaderno abierto, las actividades o el libro no pasen al form principal
Public swActividadAbierta As Boolean
Public swLibroAbierto As Boolean
Public swHuboCambioEnMaterias As Boolean 'para guardar las materias si hubo cambios
Public miMateria As String 'para saber qué materia está en el cuaderno
Public dirTrabajo As String 'para saber el directorio donde guardar el trabajo
Public swMostrarAñoEnActividades As Boolean 'para ver si se muestran los años en las actividades
Public swMostrarAñoEnTareas As Boolean 'para ver si se muestran las tareas de años anteriores en las actividades
Public swImprimirDirecto As Boolean 'para imprimir sin mostrar el cuadro de diálogo de la impresora
Public colorFuente As ColorConstants 'el color de la fuente
Public NombreFuente As String  'la fuente
Public tamañoFuente As Byte 'el tamaño de la fuente
Public colorFondo As ColorConstants 'color de fondo de los rtf
Public velocidadVoz As Integer 'para regular la velocidad de la voz
'Public swInstalarVoz As Boolean 'para que el mensaje de instalar la voz se de una sola vez
Public renglónAnterior As Long 'para leer el renglón cuando avance o retroceda leyendo
Public swLeerRenglones As Boolean 'para ver si se leen los renglones
Public swUsarCorrectorOrtográfico As Boolean 'para ver si se usa el corrector ortográfico
Public swSalir As Boolean 'para ver si se quiere salir del programa
Public swPermitirAbrirArchivos As Boolean 'para ver si se pueden abrir archivos externos en las carpetas
Public idiomaAspell As String 'para saber en qué idioma corrige Aspell
Public swAspellInstalado As Boolean 'para saber si el corrector ortográfico aspell está instalado
Public rutaDeAspell As String
Public objPipe As clsPipe 'para comunicarse con aspell con un pipe

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LB_SETHORIZONTALEXTENT = &H194

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Long, ByVal uFlags As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Public Declare Function GetCursorPos Lib "user32.dll" ( _
'    ByRef lpPoint As POINT_API _
'    ) As Long
'
'Public Type POINT_API
'    X As Long
'    Y As Long
'End Type


'constantes para que la ventana sea always on top
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOACTIVATE = &H10
'Public Const SWP_SHOWWINDOW = &H40

'constantes para mantener la ventana always on top
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400

'constantes de tecla
Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91
Public Const VK_CAPITAL = &H14

'constantes para el sonido
Public Const SND_ASYNC = &H1
Public Const SND_SYNC = &H0

'constantes para leer el registro
Public Const REG_SZ = 1 ' Cadena Unicode terminada en valor nulo
Public Const KEY_ALL_ACCESS = &H3F '((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

Public swSapi5 As Boolean 'true es que está elegido sapi5, false es sapi4 en form2
Public sonido As Long
Public banderasSPVoice As SpeechVoiceSpeakFlags
Public swHablarVoz As Boolean 'para ver si se usa la voz del programa o no

'Public swEmpezarEnCuaderno As Boolean 'para ver si se empieza directo en el cuaderno
Public swYaEmpezóPrograma As Boolean 'para ver qué dice el form principal cdo se vuelve a él retrocediendo de otros form
'Public swPermitirCambioEnActividades As Boolean 'para controlar si se permite que se edite en el lector de actividades
Public cantPrefijo As Byte 'cuántos archivos con el mismo nombre pueden haber en la misma carpeta

Public nombreUsuario As String 'el nombre de quien usa la mochila
Public swUsuarioMujer As Boolean 'para saber si es hombre o mujer
'Public swLeerSignosPuntuación As Boolean 'para ver si se leen los signos ortográficos
Public nombreSapi4 As String
Public nombreSapi5 As String

Public Type DatosUsuario
    swVerActividadesConJaws As Boolean 'para ver si se muestran las actividades en el calendario o en el árbol
    usarVoz As Boolean 'si está o no habilitada la voz
    'modoLectura As Byte 'letras, palabra, oración, prárrafo o todo
    sapi5 As Boolean 'si se elije hablar con sapi 5
'    leerSignoPuntuación As Boolean
    nombre As String * 50 'se les da 50 caracteres para que se escriba el nombre
    usuarioMujer As Boolean
    'comenzarEnCarpeta As Boolean 'para ver si empieza directo en el cuaderno
    mostrarTodasLasActividades As Boolean 'para ver si muestra act de años anteriores o futuros
    mostrarTodasLasTareas As Boolean 'idem pero con tareas
    mostrarAñoEnEvaluaciones As Boolean 'idem para evaluaciones
'    permitirEditarActividades As Boolean 'para ver si se puede escribir en el lector de actividades
    imprimirDirecto As Boolean 'para ver si se imprime sin mostrar el cuadro de diálogo de la impresora
    fuenteNombre As String * 8 'el nombre de la fuente
    fuenteColor As Long 'el color de la fuente
    colorFondo As Long 'el color de fondo de los rtf
    fuenteTamaño As Byte 'el tamaño de la fuente
    velocidadVoz As Integer 'para regular la velocidad de la voz
    swLeerRenglones As Boolean 'para ver si leen los renglones
    swUsarCorrectorOrtográfico As Boolean
    nombreVozSapi4 As String * 50
    nombreVozSapi5 As String * 50
'    swInstalarVoz As Boolean
    swMúsicaDeFondo As Boolean 'para ver si suena la música de los form
    swPermitirAbrirArchivos As Boolean
    rutaMúsicaFormPrincipal As String * 64
    rutaMúsicaFormCuaderno As String * 64
    rutaMúsicaFormActividad As String * 64
    rutaMúsicaFormTareas As String * 64
    rutaMúsicaFormLibros As String * 64
    rutaMúsicaFormAccesorios As String * 64
    rutaSonidosRecordatorios As String * 64
End Type

Public usuario As DatosUsuario

Public Type DatosActividad
    'fecha As Date
    tema As String * 50
    'materia As String * 50
    'DirArchivo As String * 300
    comentarios As String * 50 '50 caracteres para escribir la descripción
End Type

'******************************
'para conocer la versión de windows
'Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As osVersionInfo) As Long

'Public Type osVersionInfo
'    dwosversioninfosize As Long
'    dwmajorversion As Long
'    dwminorversion As Long
'    dwbuildnumber As Long
'    dwplatformid As Long
'    szcsdversion As String * 128
'End Type

'Public osInfo As osVersionInfo

Public Enum formularios
    accesorios
    actAntFut
'    activDefVisual
    actividades
    actividadesHoy
'    añadirCapítuloLibro
    añadirRecordatorio
'    añadirActividad
'    añadirLibro
    añadirRecordatorios
    añoActividad
    añoTareas
    añoEvaluaciones
    calculadora
    CalendarioMúltiple
    controlAlumno
    cuaderno
    cuadernoComunicaciones
    desdeCuaderno
    diálogoAbrir
    evaluaciones
    fechaVerRec
    historial
    imágenes
    lectorActividad
    lectorLibros
    lectorEvaluaciones
    libros
    libroX
    materiasEvaluaciones
    mesEvaluaciones
'    ordenarCapítulos
    principal
    recordatorios
    reproductorMúsica
    tareasAnt
    tecladoBraille
    títuloEvaluación
    títuloHoja
End Enum

Public Type Recordatorio
    texto As String * 64
    fecha As Date
    hora As Date
    sonido As String * 64
    yaAnunciado As Boolean
    índiceEnArchivo As Long
End Type

'Public Enum repetir 'para ver cuándo se repite un recordatorio
'    diario
'    semanal
'    mensual
'    anual
'    nunca
'End Enum

Type contadorRecordatorios
    día As Byte
    cantidad As Integer
End Type

Public Enum elemento
    tarea
    actividad
    Recordatorio
    evaluación
End Enum

Public Enum tecla
    a = 1
    flechaArriba
    flechaAbajo
    inicio
    fin
    avancePágina
    retrocesoPágina
    flechaDerecha
    flechaIzquierda
    borrar
End Enum

Public Enum comparación
    primeroMayor
    primeroMenor
    iguales
End Enum
