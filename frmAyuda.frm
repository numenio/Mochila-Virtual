VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmAyuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
   Icon            =   "frmAyuda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAyuda.frx":08CA
   ScaleHeight     =   5430
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "Ver toda la ayuda"
      EstiloDelBoton  =   1
      Picture         =   "frmAyuda.frx":2922
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
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
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   225
      TabIndex        =   0
      Top             =   600
      Width           =   6360
   End
End
Attribute VB_Name = "frmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public formulario As formularios
Dim swReci�nAbierto As Boolean
Dim cadenaVoz As String
Dim d�ndeEstoy As String
Private Enum categor�aCuaderno
'    carpeta
    actividades
    libros
    hojas
    comandos
    lectura
    escritura
    selecci�n
'    ningunaCategor�a
End Enum

Private Sub ButtonTransparent1_Click()
    ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/introducci�n.htm", "", 1 'leer la ayuda
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swReci�nAbierto = True
    cadenaVoz = ""
    d�ndeEstoy = ""
    Select Case formulario
        Case formularios.cuaderno
            d�ndeEstoy = " la carpeta "
            cadenaVoz = "entrando en la ayuda de tu carpeta. Para leer lo que escribiste en tu carpeta, lo m�s f�cil es subir o bajar con las flechas. Para leer todo el texto, us� alt m�s flecha abajo, y para leer desde donde est�s hasta el final de la hoja, us� alt m�s flecha arriba. Para guardar tu hoja, apret� efe cinco. Para abrir una actividad, libro u hoja anterior, us� efe uno. si la quer�s imprimir en tinta, apret� efe seis."
            Call llenarCategor�asCuaderno
        Case formularios.accesorios
            d�ndeEstoy = " los accesorios "
            cadenaVoz = "entrando en la ayuda de los accesorios. us� flechas arriba o abajo para elegir el accesorio que quieras abrir y acept� con enter"
            List1.AddItem "Moverse por las opciones: Flechas arriba o abajo"
            List1.AddItem "Aceptar una opci�n: Enter"
        Case formularios.actAntFut
            d�ndeEstoy = " los meses que tienen actividades guardadas "
            cadenaVoz = "entrando en la ayuda de las actividades. movete con flecha arriba o abajo por los meses que tienen actividades y acept� con enter. si un mes no est� en la lista quiere decir que no tiene actividades guardadas."
            List1.AddItem "Moverse por los meses: Flechas arriba o abajo"
            List1.AddItem "Ver las actividades del mes seleccionado: Enter"
        Case formularios.actividades
            d�ndeEstoy = " las actividades "
            cadenaVoz = "entrando en la ayuda de las actividades. movete con flechas arriba o abajo para elegir una opci�n y aceptala con enter. "
            List1.AddItem "Moverse por las opciones: flechas arriba o abajo"
            List1.AddItem "Aceptar una opci�n: Enter"
        Case formularios.actividadesHoy
            d�ndeEstoy = " la lista de actividades para hoy "
            cadenaVoz = "entrando en la ayuda de las actividades de hoy. movete con flechas arriba o abajo para elegir una actividad y abrila con enter. "
            List1.AddItem "Elegir una actividad: flechas arriba o abajo"
            List1.AddItem "Abrir la actividad seleccionada: Enter"
        Case formularios.a�adirRecordatorio
            d�ndeEstoy = " a�adir recordatorios "
            cadenaVoz = "entrando en la ayuda de a�adir un recordatorio. movete con tab o enter por las opciones y modific� cada una de ellas con flechas arriba o abajo. "
            List1.AddItem "Moverse por las opciones: Tab o enter"
            List1.AddItem "A�adir recordatorio: Enter en el bot�n A�adir Recordatorio"
        Case formularios.a�oActividad
            d�ndeEstoy = " los a�os con actividades "
            cadenaVoz = "entrando en la ayuda de los a�os que tienen actividades. Us� flechas arriba o abajo para moverte por los a�os y enter para seleccionar el que desees. "
            List1.AddItem "Moverse por los a�os con actividades: flechas arriba o abajo"
            List1.AddItem "Mostrar las actividades del a�o seleccionado: Enter"
        Case formularios.a�oEvaluaciones
            d�ndeEstoy = " los a�os con evaluaciones "
            cadenaVoz = "entrando en la ayuda de los a�os que tienen evaluaciones. Us� flechas arriba o abajo para moverte por los a�os y enter para seleccionar el que desees. "
            List1.AddItem "Moverse por los a�os con evaluaciones: flechas arriba o abajo"
            List1.AddItem "Mostrar las evaluaciones del a�o seleccionado: Enter"
        Case formularios.a�oTareas
            d�ndeEstoy = " los a�os con hojas guardadas "
            cadenaVoz = "entrando en la ayuda de los a�os que tienen hojas guardadas. Us� flechas arriba o abajo para moverte por los a�os y enter para seleccionar el que desees. "
            List1.AddItem "Moverse por los a�os con hojas guardadas: flechas arriba o abajo"
            List1.AddItem "Mostrar las hojas guardadas del a�o seleccionado: Enter"
        Case formularios.calculadora
            d�ndeEstoy = " la calculadora "
            cadenaVoz = "entrando en la ayuda de la calculadora. escrib� tu c�lculo con los n�meros y apret� enter para tener el resultado. para que la mochila te repita el n�mero que est� en pantalla, apret� efe uno o efe dos. para borrar un n�mero apret� la tecla de borrar, o para borrar todo el c�lculo, apret� suprimir. "
            List1.AddItem "Borrar todo el c�lculo: suprimir"
            List1.AddItem "Borrar el �ltimo n�mero escrito: borrar"
            List1.AddItem "Leer el n�mero en pantalla: F1"
            List1.AddItem "Leer el n�mero en pantalla por cifras: F2"
            List1.AddItem "Escribir n�meros: los n�meros del teclado"
            List1.AddItem "Signo suma: m�s"
            List1.AddItem "Signo resta: gui�n"
            List1.AddItem "Signo multiplicaci�n: asterisco"
            List1.AddItem "Signo divisi�n: barra diagonal"
            List1.AddItem "Signo decimal: punto"
            List1.AddItem "Signo negativo: gui�n"
            List1.AddItem "Copiar lo escrito para pegarlo en otra hoja: Control + c"
            List1.AddItem "Pasar a la carpeta o evaluaci�n sin cerrar la calculadora: F8"
        Case formularios.CalendarioM�ltiple
            d�ndeEstoy = " el calendario "
            cadenaVoz = "Entrando en la ayuda del calendario. para moverte por los d�as del calendario, us� flecha arriba o abajo, y acept� con enter. Si quer�s ver todos los d�as, aunque est�n vac�os, us� espacio. Para volver atr�s, us� borrar. Para cerrar el calendario, us� escape"
            List1.AddItem "Moverse por el calendario: Flechas arriba o abajo"
            List1.AddItem "Abrir un elemento: Enter"
            List1.AddItem "Volver al calendario: Borrar"
            List1.AddItem "Mostrar todos los d�as del mes: Espacio"
            List1.AddItem "Ocultar los d�as del mes vac�os: Espacio"
        Case formularios.cuadernoComunicaciones
            d�ndeEstoy = " el cuaderno de comunicaciones "
            cadenaVoz = "entrando en la ayuda del cuaderno de comunicaciones. para a�adir una comunicaci�n, seguir las instrucciones del propio cuaderno. Para pasar r�pido a las comunicaciones ya escritas, usar efe uno. "
            List1.AddItem "Moverse por las opciones: Tab"
            List1.AddItem "Pasar r�pido a las comunicaciones guardadas: F1"
        Case formularios.desdeCuaderno
            d�ndeEstoy = " el cuadro para abrir actividades, libros y hojas ya escritas "
            cadenaVoz = "entrando en la ayuda del cuaderno de comunicaciones. para a�adir una comunicaci�n, seguir las instrucciones del propio cuaderno. Para pasar r�pido a las comunicaciones ya escritas, usar efe uno. "
            List1.AddItem "Moverse por las opciones: Flechas arriba o abajo"
            List1.AddItem "Abrir la opci�n seleccionada: Enter"
        Case formularios.di�logoAbrir
            d�ndeEstoy = " el cuadro para abrir un archivo "
            cadenaVoz = "Entrando en la ayuda del cuadro para abrir un archivo. para buscar el archivo que quieras abrir, buscalo con flechas arriba o abajo y acept� con enter. Si escuch�s que el cuadro te dice carpeta, eso indica que si das enter vas a abrir esa carpeta. Si en cambio dice archivo, eso indica que es uno de los archivos de texto que pod�s abrir. si apret�s la tecla suprimir, vas directamente a los discos de tu computadora, mientras que si apret�s borrar, volv�s a la carpeta anterior a la que est�s ahora."
            List1.AddItem "Volver directamente a los discos: Suprimir, o la opci�n Cambiar de disco"
            List1.AddItem "Volver a la carpeta que contiene a la actual: Tecla Borrar, o la opci�n Volver a la carpeta anterior"
            List1.AddItem "Moverse de a un elemento: flecha arriba o abajo"
            List1.AddItem "Abrir un disco, carpeta o archivo: Enter"
            List1.AddItem "Pasar r�pido a las carpetas: letra C"
            List1.AddItem "Pasar r�pido a los archivos: letra A"
            List1.AddItem "Ir al principio de la lista: Inicio"
            List1.AddItem "Ir al final de la lista: Fin"
            List1.AddItem "Saltar hacia adelante en la lista: Avance de P�gina"
            List1.AddItem "Saltar hacia atr�s en la p�gina: Retroceso de P�gina"
        Case formularios.controlAlumno
            d�ndeEstoy = " la configuraci�n de tu mochila "
            cadenaVoz = "Entrando en la ayuda del la configuraci�n de tu mochila. movete con tab o enter por las opciones y modificalas con flechas arriba o abajo. "
            List1.AddItem "Moverse por las opciones: Tab o enter"
            List1.AddItem "Modificar una opci�n: Flechas arriba o abajo"
        Case formularios.evaluaciones
            d�ndeEstoy = " las evaluaciones "
            cadenaVoz = "Entrando en la ayuda de las evaluaciones. para moverte por las opciones, us� flechas arriba o abajo. para aceptar una opci�n, us� enter. "
            List1.AddItem "Moverse por las opciones: Flechas arriba o abajo"
            List1.AddItem "Abrir la opci�n seleccionada: Enter"
        Case formularios.fechaVerRec
            d�ndeEstoy = " los recordatorios ya guardados "
            cadenaVoz = "Entrando en la ayuda de los recordatorios ya guardados. para pasar de una opci�n a otra us� tab o enter, y para modificar estas opciones, us� flechas arriba o abajo. "
            List1.AddItem "Pasar de una opci�n a otra: Enter o tab"
            List1.AddItem "Cambiar los meses o los a�os: Flechas arriba o abajo"
        Case formularios.im�genes
            d�ndeEstoy = " insertar una imagen "
            cadenaVoz = "Entrando en la ayuda de las im�genes. movete por las im�genes con flechas arriba o abajo. para insertar la imagen seleccionada, apret� enter. "
            List1.AddItem "Moverse por las im�genes: Flechas arriba o abajo"
            List1.AddItem "Insertar una imagen en la carpeta: Enter"
        Case formularios.lectorActividad
            d�ndeEstoy = " el lector de actividades "
            cadenaVoz = "Entrando en la ayuda del lector de actividades. Para leer la actividad, lo m�s f�cil es subir o bajar con las flechas. Para leer todo el texto, us� alt m�s flecha abajo, y para leer desde donde est�s hasta el final de la hoja, us� alt m�s flecha arriba. si quer�s leer de a palabras us� control m�s las flechas derecha o izquierda. Para imprimir, apret� efe cinco. "
            List1.AddItem "Volver a las carpetas: F2"
            List1.AddItem "Imprimir: F5"
            List1.AddItem "Leer todo el texto: Alt + flecha abajo"
            List1.AddItem "Leer desde el cursor hacia adelanta: Alt + flecha arriba"
            List1.AddItem "Leer por renglones: Flechas arriba o abajo"
            List1.AddItem "Leer por palabras: Control + flechas izquierda o derecha"
            List1.AddItem "Ir al principio del rengl�n: Inicio"
            List1.AddItem "Ir al fin del rengl�n: Fin"
            List1.AddItem "Ir r�pido al comienzo de la actividad: Control + inicio"
            List1.AddItem "Ir r�pido al fin de la actividad: Control + fin"
            List1.AddItem "Leer de a p�rrafos: Control + flechas arriba o abajo"
            List1.AddItem "Leer el texto seleccionado: Alt + flecha derecha"
        Case formularios.lectorEvaluaciones
            d�ndeEstoy = " las evaluaciones "
            cadenaVoz = "Entrando en la ayuda del lector de evaluaciones. Para leer la evaluaci�n, lo m�s f�cil es subir o bajar con las flechas. Para leer todo el texto, us� alt m�s flecha abajo, y para leer desde donde est�s hasta el final de la hoja, us� alt m�s flecha arriba. si quer�s leer de a palabras us� control m�s las flechas derecha o izquierda. Para imprimir, apret� efe cinco. "
            List1.AddItem "Imprimir: F5"
            List1.AddItem "Leer todo el texto: Alt + flecha abajo"
            List1.AddItem "Leer desde el cursor hacia adelanta: Alt + flecha arriba"
            List1.AddItem "Leer por renglones: Flechas arriba o abajo"
            List1.AddItem "Leer por palabras: Control + flechas izquierda o derecha"
            List1.AddItem "Ir al principio del rengl�n: Inicio"
            List1.AddItem "Ir al fin del rengl�n: Fin"
            List1.AddItem "Ir r�pido al comienzo de la evaluaci�n: Control + inicio"
            List1.AddItem "Ir r�pido al fin de la evaluaci�n: Control + fin"
            List1.AddItem "Leer de a p�rrafos: Control + flechas arriba o abajo"
            List1.AddItem "Leer el texto seleccionado: Alt + flecha derecha"
        Case formularios.lectorLibros
            d�ndeEstoy = " el lector de libros "
            cadenaVoz = "Entrando en la ayuda del "
            List1.AddItem "Volver a las carpetas: F3"
            List1.AddItem "Imprimir: F5"
            List1.AddItem "Leer todo el texto: Alt + flecha abajo"
            List1.AddItem "Leer desde el cursor hacia adelanta: Alt + flecha arriba"
            List1.AddItem "Leer por renglones: Flechas arriba o abajo"
            List1.AddItem "Leer por palabras: Control + flechas izquierda o derecha"
            List1.AddItem "Ir al principio del rengl�n: Inicio"
            List1.AddItem "Ir al fin del rengl�n: Fin"
            List1.AddItem "Ir r�pido al comienzo del libro: Control + inicio"
            List1.AddItem "Ir r�pido al fin del libro: Control + fin"
            List1.AddItem "Leer de a p�rrafos: Control + flechas arriba o abajo"
            List1.AddItem "Leer el texto seleccionado: Alt + flecha derecha"
        Case formularios.libros
            d�ndeEstoy = " los libros "
            cadenaVoz = "Entrando en la ayuda de los libros. seleccion� con flechas arriba o abajo el libro que desees y abr� sus cap�tulos con enter. "
            List1.AddItem "Moverse por los libros: Flechas arriba o abajo"
            List1.AddItem "Abrir los cap�tulos del libro seleccionado: Enter"
        Case formularios.libroX
            d�ndeEstoy = " los cap�tulos de libros "
            cadenaVoz = "Entrando en la ayuda de los cap�tulos del libro que elegiste. para moverte por los cap�tulos us� flechas arriba o abajo. para abrir el cap�tulo seleccionado apret� enter. "
            List1.AddItem "Moverse por los cap�tulos: Flechas arriba o abajo"
            List1.AddItem "Abrir un cap�tulo: Enter"
        Case formularios.materiasEvaluaciones
            d�ndeEstoy = " las materias para evaluar "
            cadenaVoz = "Entrando en la ayuda de las evaluaciones. eleg� con flechas arriba o abajo la materia de la que quieras hacer una evaluaci�n y acept� con enter. "
            List1.AddItem "Moverse por las materias: Flechas arriba o abajo"
            List1.AddItem "Seleccionar una materia: Enter"
        Case formularios.mesEvaluaciones
            d�ndeEstoy = " los meses con evaluaciones "
            cadenaVoz = "Entrando en la ayuda de los meses con evaluaciones guardadas. para moverte por los meses, us� las flechas arriba o abajo. para ver las hojas de el mes seleccionado, apret� enter. Si un mes no aparece en la lista, es que ese mes no tiene hojas guardadas. "
            List1.AddItem "Moverse por los meses: Flechas arriba o abajo"
            List1.AddItem "Mostrar las evaluaciones del mes seleccionado: Enter"
        Case formularios.principal
            d�ndeEstoy = " el men� principal "
            cadenaVoz = "Entrando en la ayuda del men� principal. Us� flechas arriba o abajo para buscar la materia que quieras abrir y hacelo con enter. "
            List1.AddItem "Moverse por las materias: Flechas arriba o abajo"
            List1.AddItem "Abrir una materia: Enter"
        Case formularios.recordatorios
            d�ndeEstoy = " los recordatorios "
            cadenaVoz = "Entrando en la ayuda de los recordatorios. para moverte por las opciones us� flechas arriba o abajo y para aceptar una de ellas, us� enter. "
            List1.AddItem "Moverse por las opciones: Flechas arriba o abajo"
            List1.AddItem "Aceptar una opci�n: Enter"
        Case formularios.reproductorM�sica
            d�ndeEstoy = " el reproductor de m�sica "
            cadenaVoz = "Entrando en la ayuda del reproductor de m�sica. para buscar el tema que quieras escuchar, buscalo con flechas arriba o abajo y acept� con enter. Si escuch�s que el reproductor te dice carpeta, eso indica que si das enter vas a abrir esa carpeta. Si en cambio dice m�sica, eso indica que es un tema, y que si das enter va a empezar a sonar. Para pausar una reproducci�n us� espacio, y para silenciar totalmente al reproductor, apret� shift. Para bajar el volumen apret� el n�mero uno, y para subirlo us� el dos. Cuando est�s buscando el tema para escuchar, si apret�s la tecla suprimir, vas directamente a los discos de tu computadora, mientras que si apret�s borrar, volv�s a la carpeta anterior a la que est�s ahora. Para cerrar el reproductor apret� escape y para dejar el reproductor abierto, pero seguir trabajando en tu mochila, apret� f 7."
            List1.AddItem "Buscar temas o carpetas: flechas arriba o abajo y enter"
            List1.AddItem "Reproducir: enter cuando est�s en un tema"
            List1.AddItem "Pausa: espacio"
            List1.AddItem "Parar la reproducci�n: enter"
            List1.AddItem "Silenciar el reproductor: shift"
            List1.AddItem "Subir el volumen: n�mero uno"
            List1.AddItem "Bajar el volumen: n�mero dos"
            List1.AddItem "Ir a los discos de tu computadora: suprimir"
            List1.AddItem "Volver a la carpeta anterior: borrar"
            List1.AddItem "Cerrar el reproductor: escape"
            List1.AddItem "Pasar a la carpeta: F7"
        Case formularios.tareasAnt
            d�ndeEstoy = " los meses que tienen hojas ya guardadas "
            cadenaVoz = "Entrando en la ayuda de los meses con hojas ya escritas. para moverte por los meses, us� las flechas arriba o abajo. para ver las hojas de el mes seleccionado, apret� enter. Si un mes no aparece en la lista, es que ese mes no tiene hojas guardadas. "
            List1.AddItem "Moverse por los meses: Flechas arriba o abajo"
            List1.AddItem "Mostrar las hojas del mes seleccionado: Enter"
        Case formularios.tecladoBraille
            d�ndeEstoy = " el teclado braille "
            cadenaVoz = "Entrando en la ayuda del teclado braille. para los puntos del braille, us� solamente las teclas efe, de, ese, jota, ka y ele. para borrar us� la tecla de borrar, y para bajar una l�nea us� enter. "
            List1.AddItem "Punto 1: F"
            List1.AddItem "Punto 2: D"
            List1.AddItem "Punto 3: S"
            List1.AddItem "Punto 4: J"
            List1.AddItem "Punto 5: K"
            List1.AddItem "Punto 6: L"
            List1.AddItem "Bajar una l�nea: Enter"
            List1.AddItem "Borrar una letra: Borrar"
            List1.AddItem "Espacio: Espacio"
        Case formularios.t�tuloEvaluaci�n
            d�ndeEstoy = " el cuadro para ponerle un t�tulo a tu evaluaci�n "
            cadenaVoz = "Entrando en la ayuda del t�tulo de tu evaluaci�n. para leer lo que has escrito, us� flechas arriba o abajo y para guardar el t�tulo us� enter. "
            List1.AddItem "Leer lo escrito: Flechas arriba o abajo"
            List1.AddItem "Aceptar el t�tulo: Enter"
        Case formularios.t�tuloHoja
            d�ndeEstoy = " el cuadro para ponerle un t�tulo a tu hoja "
            cadenaVoz = "Entrando en la ayuda del t�tulo de tu hoja. para leer lo que has escrito, us� flechas arriba o abajo y para guardar el t�tulo us� enter. "
            List1.AddItem "Leer lo escrito: Flechas arriba o abajo"
            List1.AddItem "Aceptar el t�tulo: Enter"
    End Select
    List1.AddItem "Abrir la configuraci�n de la Mochila: F12"
    
    cadenaVoz = cadenaVoz + ". Si quer�s repasar lentamente la ayuda que te acabo de dar, us� flecha arriba o abajo. Para cerrar esta ayuda apreta escape."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el men� de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then Unload Me
    
    If KeyCode = vbKeyF12 Then Call ButtonTransparent1_Click
'    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de m�sica, ten�s que estar en el men� principal o en una carpeta. ahora est�s en la ayuda"
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al men� de la aplicaci�n. Para leer los �tems de este men� necesit�s jaws u otro lector de pantallas. Para volver a la mochila, apret� escape"
    
End Sub

Private Sub Form_Activate()
    List1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Decir "saliendo de la ayuda de " + d�ndeEstoy
    'Call contarFormularios(False)
End Sub

Private Sub List1_GotFocus()
    If swReci�nAbierto = True Then
        Decir cadenaVoz + List1.List(List1.ListIndex)
        swReci�nAbierto = False
    Else
        Decir List1.List(List1.ListIndex)
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer, cadena As String
    shiftkey = Shift And 7
    If shiftkey <> vbCtrlMask And KeyCode <> vbKeyReturn Then
        Decir "est�s en la ayuda de " + d�ndeEstoy + " . Movete con flecha arriba o abajo. Para salir, apret� Escape."
    End If
    
    If formulario = formularios.cuaderno Then
        If shiftkey = 0 And KeyCode = vbKeyReturn Then
            If List1.ListIndex <> -1 Then
                Select Case List1.List(List1.ListIndex)
                    Case "Ayuda para trabajar con actividades"
                        cadena = "la ayuda para abrir, leer y realizar actividades"
                        Call abrirCategor�a(categor�aCuaderno.actividades)
                    Case "Ayuda para leer libros"
                        cadena = "la ayuda para leer libros"
                        Call abrirCategor�a(categor�aCuaderno.libros)
                    Case "Ayuda con las hojas ya escritas de tu carpeta"
                        cadena = "la ayuda para abrir, releer y escribir en hojas ya escritas de tu carpeta"
                        Call abrirCategor�a(categor�aCuaderno.hojas)
                    Case "Ayuda para leer un texto"
                        cadena = "la ayuda para leer el texto que est� escrito en la carpeta"
                        Call abrirCategor�a(categor�aCuaderno.lectura)
                    Case "Ayuda para escribir un texto"
                        cadena = "la ayuda para escribir texto en tu carpeta"
                        Call abrirCategor�a(categor�aCuaderno.escritura)
                    Case "Ayuda para seleccionar parte de un texto"
                        cadena = "la ayuda para seleccionar texto que has escrito en la hoja"
                        Call abrirCategor�a(categor�aCuaderno.selecci�n)
                    Case "Ayuda con otros comandos"
                        cadena = " una lista de varios comandos �tiles"
                        Call abrirCategor�a(categor�aCuaderno.comandos)
                    Case "Volver a la lista de categor�as"
                        Call llenarCategor�asCuaderno
                End Select
                
                If List1.List(List1.ListIndex) = "Volver a la lista de categor�as" Then
                    Decir "Volviendo a todas las categor�as de ayuda de tu carpeta, us� flecha abajo para elegir una y acept� con enter"
                Else
                    Decir "abriendo " + cadena + ". us� las flechas para leerla o eleg� volver a la lista de categor�as para cambiar de ayuda"
                End If
            End If
        End If
        
        If shiftkey = 0 And KeyCode = vbKeyBack Then Call llenarCategor�asCuaderno
    End If
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
'        If swReci�nAbierto = True Then
'            Decir cadenaVoz + List1.List(List1.ListIndex)
'            swReci�nAbierto = False
'        Else
            Decir List1.List(List1.ListIndex)
            Exit Sub
'        End If
    End If
    
    
'    sonido = sndPlaySound(App.Path + "\sonidos\td.wav", SND_ASYNC)
End Sub

Sub llenarCategor�asCuaderno()
    List1.Clear
    List1.AddItem "Ayuda para trabajar con actividades"
    List1.AddItem "Ayuda para leer libros"
    List1.AddItem "Ayuda con las hojas ya escritas de tu carpeta"
    List1.AddItem "Ayuda para leer un texto"
    List1.AddItem "Ayuda para escribir un texto"
    List1.AddItem "Ayuda para seleccionar parte de un texto"
    List1.AddItem "Ayuda con otros comandos"
End Sub

Sub abrirCategor�a(cu�lCategor�a As Byte)
    List1.Clear
    Select Case cu�lCategor�a
        Case categor�aCuaderno.escritura
            List1.AddItem "Guardar la hoja: Control + G � F5"
            List1.AddItem "Imprimir: Control + P � F6"
            List1.AddItem "Poner la letra en negrita: Control + N"
            List1.AddItem "Poner la letra en subrayado: Control + S"
            List1.AddItem "Copiar texto seleccionado: Control + C"
            List1.AddItem "Cortar texto seleccionado: Control + X"
            List1.AddItem "Pegar el texto copiado o cortado: Control + V"
            List1.AddItem "Abrir el teclado Braille: F9"
            List1.AddItem "Ir al principio de la hoja: Control + Inicio"
            List1.AddItem "Ir al final de la hoja: Control + Fin"
            List1.AddItem "Ir al principio del rengl�n: Inicio"
            List1.AddItem "Ir al final del rengl�n: Fin"
            List1.AddItem "Avanzar varios renglones en la hoja de un salto: Avance de P�gina"
            List1.AddItem "Retroceder varios renglones en la hoja de un salto: Retroceso de P�gina"
            List1.AddItem "Retroceder un p�rrafo: Control + Flecha Arriba"
            List1.AddItem "Avazar un p�rrafo: Control + Flecha Abajo"
        Case categor�aCuaderno.lectura
            List1.AddItem "Leer de a letras avanzando: Flecha derecha"
            List1.AddItem "Leer de a letras retrocediendo: Flecha izquierda"
            List1.AddItem "Leer de a palabras avanzando: Control + Flecha derecha"
            List1.AddItem "Leer de a palabras retrocediendo: Control + Flecha izquierda"
            List1.AddItem "Leer de a renglones bajando en la hoja: Flecha abajo"
            List1.AddItem "Leer de a renglones subiendo en la hoja: Flecha arriba"
            List1.AddItem "Leer todo el texto: Alt + Flecha Abajo"
            List1.AddItem "Leer desde el cursor hacia adelante: Alt + Flecha Arriba"
        Case categor�aCuaderno.selecci�n
            List1.AddItem "Leer el texto seleccionado: Alt + Flecha derecha"
            List1.AddItem "Seleccionar de a palabras avanzando: Shift + Control + Flecha derecha"
            List1.AddItem "Seleccionar de a palabras retrocediendo: Shift + Control + Flecha izquierda"
            List1.AddItem "Seleccionar todo el texto desde donde uno est� hasta el principio de la hoja: Shift + Control + Inicio"
            List1.AddItem "Seleccionar todo el texto desde donde uno est� hasta el final de la hoja: Shift + Control + Fin"
            List1.AddItem "Seleccionar desde donde uno est� hasta el principio del p�rrafo: Shift + Control + Flecha Arriba"
            List1.AddItem "Seleccionar desde donde uno est� hasta el final del p�rrafo: Shift + Control + Flecha Abajo"
            List1.AddItem "Seleccionar varios renglones hacia abajo: Shift + Control + Retroceso de P�gina"
            List1.AddItem "Seleccionar varios renglones hacia arriba: Shift + Control + Avance de P�gina"
            List1.AddItem "Seleccionar hasta el principio del rengl�n: Shift + Inicio"
            List1.AddItem "Seleccionar hasta el final del rengl�n: Shift + Fin"
            List1.AddItem "Seleccionar desde donde uno est� hasta el rengl�n inferior: Shift + Flecha Abajo"
            List1.AddItem "Seleccionar desde donde uno est� hasta el rengl�n superior: Shift + Flecha Arriba"
            List1.AddItem "Seleccionar varios renglones hacia arriba: Shift + Avance de P�gina"
            List1.AddItem "Seleccionar varios renglones hacia abajo: Shift + Retroceso de P�gina"
            List1.AddItem "Seleccionar toda la hoja: Control + A"
        Case categor�aCuaderno.comandos
            List1.AddItem "Callar la voz: Control"
            List1.AddItem "Abrir los accesorios: F4"
            List1.AddItem "Abrir el reproductor de m�sica: F7"
            List1.AddItem "Salir de la Carpeta: Escape"
            List1.AddItem "Salir del programa: Alt + F4"
            List1.AddItem "Abrir la configuraci�n: F12"
        Case categor�aCuaderno.actividades
            List1.AddItem "Abrir una actividad: F1"
            List1.AddItem "Cambiar a una actividad abierta: F2"
            List1.AddItem "Leer una actividad: igual que como se lee en la carpeta"
        Case categor�aCuaderno.hojas
            List1.AddItem "Abrir una hoja ya escrita: F1"
        Case categor�aCuaderno.libros
            List1.AddItem "Abrir un libro: F1"
            List1.AddItem "Cambiar a un libro abierto: F3"
            List1.AddItem "Leer un libro: igual que como se lee en la carpeta"
    End Select
    List1.AddItem "Volver a la lista de categor�as"
End Sub
