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
Dim swReciénAbierto As Boolean
Dim cadenaVoz As String
Dim dóndeEstoy As String
Private Enum categoríaCuaderno
'    carpeta
    actividades
    libros
    hojas
    comandos
    lectura
    escritura
    selección
'    ningunaCategoría
End Enum

Private Sub ButtonTransparent1_Click()
    ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/introducción.htm", "", 1 'leer la ayuda
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swReciénAbierto = True
    cadenaVoz = ""
    dóndeEstoy = ""
    Select Case formulario
        Case formularios.cuaderno
            dóndeEstoy = " la carpeta "
            cadenaVoz = "entrando en la ayuda de tu carpeta. Para leer lo que escribiste en tu carpeta, lo más fácil es subir o bajar con las flechas. Para leer todo el texto, usá alt más flecha abajo, y para leer desde donde estás hasta el final de la hoja, usá alt más flecha arriba. Para guardar tu hoja, apretá efe cinco. Para abrir una actividad, libro u hoja anterior, usá efe uno. si la querés imprimir en tinta, apretá efe seis."
            Call llenarCategoríasCuaderno
        Case formularios.accesorios
            dóndeEstoy = " los accesorios "
            cadenaVoz = "entrando en la ayuda de los accesorios. usá flechas arriba o abajo para elegir el accesorio que quieras abrir y aceptá con enter"
            List1.AddItem "Moverse por las opciones: Flechas arriba o abajo"
            List1.AddItem "Aceptar una opción: Enter"
        Case formularios.actAntFut
            dóndeEstoy = " los meses que tienen actividades guardadas "
            cadenaVoz = "entrando en la ayuda de las actividades. movete con flecha arriba o abajo por los meses que tienen actividades y aceptá con enter. si un mes no está en la lista quiere decir que no tiene actividades guardadas."
            List1.AddItem "Moverse por los meses: Flechas arriba o abajo"
            List1.AddItem "Ver las actividades del mes seleccionado: Enter"
        Case formularios.actividades
            dóndeEstoy = " las actividades "
            cadenaVoz = "entrando en la ayuda de las actividades. movete con flechas arriba o abajo para elegir una opción y aceptala con enter. "
            List1.AddItem "Moverse por las opciones: flechas arriba o abajo"
            List1.AddItem "Aceptar una opción: Enter"
        Case formularios.actividadesHoy
            dóndeEstoy = " la lista de actividades para hoy "
            cadenaVoz = "entrando en la ayuda de las actividades de hoy. movete con flechas arriba o abajo para elegir una actividad y abrila con enter. "
            List1.AddItem "Elegir una actividad: flechas arriba o abajo"
            List1.AddItem "Abrir la actividad seleccionada: Enter"
        Case formularios.añadirRecordatorio
            dóndeEstoy = " añadir recordatorios "
            cadenaVoz = "entrando en la ayuda de añadir un recordatorio. movete con tab o enter por las opciones y modificá cada una de ellas con flechas arriba o abajo. "
            List1.AddItem "Moverse por las opciones: Tab o enter"
            List1.AddItem "Añadir recordatorio: Enter en el botón Añadir Recordatorio"
        Case formularios.añoActividad
            dóndeEstoy = " los años con actividades "
            cadenaVoz = "entrando en la ayuda de los años que tienen actividades. Usá flechas arriba o abajo para moverte por los años y enter para seleccionar el que desees. "
            List1.AddItem "Moverse por los años con actividades: flechas arriba o abajo"
            List1.AddItem "Mostrar las actividades del año seleccionado: Enter"
        Case formularios.añoEvaluaciones
            dóndeEstoy = " los años con evaluaciones "
            cadenaVoz = "entrando en la ayuda de los años que tienen evaluaciones. Usá flechas arriba o abajo para moverte por los años y enter para seleccionar el que desees. "
            List1.AddItem "Moverse por los años con evaluaciones: flechas arriba o abajo"
            List1.AddItem "Mostrar las evaluaciones del año seleccionado: Enter"
        Case formularios.añoTareas
            dóndeEstoy = " los años con hojas guardadas "
            cadenaVoz = "entrando en la ayuda de los años que tienen hojas guardadas. Usá flechas arriba o abajo para moverte por los años y enter para seleccionar el que desees. "
            List1.AddItem "Moverse por los años con hojas guardadas: flechas arriba o abajo"
            List1.AddItem "Mostrar las hojas guardadas del año seleccionado: Enter"
        Case formularios.calculadora
            dóndeEstoy = " la calculadora "
            cadenaVoz = "entrando en la ayuda de la calculadora. escribí tu cálculo con los números y apretá enter para tener el resultado. para que la mochila te repita el número que está en pantalla, apretá efe uno o efe dos. para borrar un número apretá la tecla de borrar, o para borrar todo el cálculo, apretá suprimir. "
            List1.AddItem "Borrar todo el cálculo: suprimir"
            List1.AddItem "Borrar el último número escrito: borrar"
            List1.AddItem "Leer el número en pantalla: F1"
            List1.AddItem "Leer el número en pantalla por cifras: F2"
            List1.AddItem "Escribir números: los números del teclado"
            List1.AddItem "Signo suma: más"
            List1.AddItem "Signo resta: guión"
            List1.AddItem "Signo multiplicación: asterisco"
            List1.AddItem "Signo división: barra diagonal"
            List1.AddItem "Signo decimal: punto"
            List1.AddItem "Signo negativo: guión"
            List1.AddItem "Copiar lo escrito para pegarlo en otra hoja: Control + c"
            List1.AddItem "Pasar a la carpeta o evaluación sin cerrar la calculadora: F8"
        Case formularios.CalendarioMúltiple
            dóndeEstoy = " el calendario "
            cadenaVoz = "Entrando en la ayuda del calendario. para moverte por los días del calendario, usá flecha arriba o abajo, y aceptá con enter. Si querés ver todos los días, aunque estén vacíos, usá espacio. Para volver atrás, usá borrar. Para cerrar el calendario, usá escape"
            List1.AddItem "Moverse por el calendario: Flechas arriba o abajo"
            List1.AddItem "Abrir un elemento: Enter"
            List1.AddItem "Volver al calendario: Borrar"
            List1.AddItem "Mostrar todos los días del mes: Espacio"
            List1.AddItem "Ocultar los días del mes vacíos: Espacio"
        Case formularios.cuadernoComunicaciones
            dóndeEstoy = " el cuaderno de comunicaciones "
            cadenaVoz = "entrando en la ayuda del cuaderno de comunicaciones. para añadir una comunicación, seguir las instrucciones del propio cuaderno. Para pasar rápido a las comunicaciones ya escritas, usar efe uno. "
            List1.AddItem "Moverse por las opciones: Tab"
            List1.AddItem "Pasar rápido a las comunicaciones guardadas: F1"
        Case formularios.desdeCuaderno
            dóndeEstoy = " el cuadro para abrir actividades, libros y hojas ya escritas "
            cadenaVoz = "entrando en la ayuda del cuaderno de comunicaciones. para añadir una comunicación, seguir las instrucciones del propio cuaderno. Para pasar rápido a las comunicaciones ya escritas, usar efe uno. "
            List1.AddItem "Moverse por las opciones: Flechas arriba o abajo"
            List1.AddItem "Abrir la opción seleccionada: Enter"
        Case formularios.diálogoAbrir
            dóndeEstoy = " el cuadro para abrir un archivo "
            cadenaVoz = "Entrando en la ayuda del cuadro para abrir un archivo. para buscar el archivo que quieras abrir, buscalo con flechas arriba o abajo y aceptá con enter. Si escuchás que el cuadro te dice carpeta, eso indica que si das enter vas a abrir esa carpeta. Si en cambio dice archivo, eso indica que es uno de los archivos de texto que podés abrir. si apretás la tecla suprimir, vas directamente a los discos de tu computadora, mientras que si apretás borrar, volvés a la carpeta anterior a la que estás ahora."
            List1.AddItem "Volver directamente a los discos: Suprimir, o la opción Cambiar de disco"
            List1.AddItem "Volver a la carpeta que contiene a la actual: Tecla Borrar, o la opción Volver a la carpeta anterior"
            List1.AddItem "Moverse de a un elemento: flecha arriba o abajo"
            List1.AddItem "Abrir un disco, carpeta o archivo: Enter"
            List1.AddItem "Pasar rápido a las carpetas: letra C"
            List1.AddItem "Pasar rápido a los archivos: letra A"
            List1.AddItem "Ir al principio de la lista: Inicio"
            List1.AddItem "Ir al final de la lista: Fin"
            List1.AddItem "Saltar hacia adelante en la lista: Avance de Página"
            List1.AddItem "Saltar hacia atrás en la página: Retroceso de Página"
        Case formularios.controlAlumno
            dóndeEstoy = " la configuración de tu mochila "
            cadenaVoz = "Entrando en la ayuda del la configuración de tu mochila. movete con tab o enter por las opciones y modificalas con flechas arriba o abajo. "
            List1.AddItem "Moverse por las opciones: Tab o enter"
            List1.AddItem "Modificar una opción: Flechas arriba o abajo"
        Case formularios.evaluaciones
            dóndeEstoy = " las evaluaciones "
            cadenaVoz = "Entrando en la ayuda de las evaluaciones. para moverte por las opciones, usá flechas arriba o abajo. para aceptar una opción, usá enter. "
            List1.AddItem "Moverse por las opciones: Flechas arriba o abajo"
            List1.AddItem "Abrir la opción seleccionada: Enter"
        Case formularios.fechaVerRec
            dóndeEstoy = " los recordatorios ya guardados "
            cadenaVoz = "Entrando en la ayuda de los recordatorios ya guardados. para pasar de una opción a otra usá tab o enter, y para modificar estas opciones, usá flechas arriba o abajo. "
            List1.AddItem "Pasar de una opción a otra: Enter o tab"
            List1.AddItem "Cambiar los meses o los años: Flechas arriba o abajo"
        Case formularios.imágenes
            dóndeEstoy = " insertar una imagen "
            cadenaVoz = "Entrando en la ayuda de las imágenes. movete por las imágenes con flechas arriba o abajo. para insertar la imagen seleccionada, apretá enter. "
            List1.AddItem "Moverse por las imágenes: Flechas arriba o abajo"
            List1.AddItem "Insertar una imagen en la carpeta: Enter"
        Case formularios.lectorActividad
            dóndeEstoy = " el lector de actividades "
            cadenaVoz = "Entrando en la ayuda del lector de actividades. Para leer la actividad, lo más fácil es subir o bajar con las flechas. Para leer todo el texto, usá alt más flecha abajo, y para leer desde donde estás hasta el final de la hoja, usá alt más flecha arriba. si querés leer de a palabras usá control más las flechas derecha o izquierda. Para imprimir, apretá efe cinco. "
            List1.AddItem "Volver a las carpetas: F2"
            List1.AddItem "Imprimir: F5"
            List1.AddItem "Leer todo el texto: Alt + flecha abajo"
            List1.AddItem "Leer desde el cursor hacia adelanta: Alt + flecha arriba"
            List1.AddItem "Leer por renglones: Flechas arriba o abajo"
            List1.AddItem "Leer por palabras: Control + flechas izquierda o derecha"
            List1.AddItem "Ir al principio del renglón: Inicio"
            List1.AddItem "Ir al fin del renglón: Fin"
            List1.AddItem "Ir rápido al comienzo de la actividad: Control + inicio"
            List1.AddItem "Ir rápido al fin de la actividad: Control + fin"
            List1.AddItem "Leer de a párrafos: Control + flechas arriba o abajo"
            List1.AddItem "Leer el texto seleccionado: Alt + flecha derecha"
        Case formularios.lectorEvaluaciones
            dóndeEstoy = " las evaluaciones "
            cadenaVoz = "Entrando en la ayuda del lector de evaluaciones. Para leer la evaluación, lo más fácil es subir o bajar con las flechas. Para leer todo el texto, usá alt más flecha abajo, y para leer desde donde estás hasta el final de la hoja, usá alt más flecha arriba. si querés leer de a palabras usá control más las flechas derecha o izquierda. Para imprimir, apretá efe cinco. "
            List1.AddItem "Imprimir: F5"
            List1.AddItem "Leer todo el texto: Alt + flecha abajo"
            List1.AddItem "Leer desde el cursor hacia adelanta: Alt + flecha arriba"
            List1.AddItem "Leer por renglones: Flechas arriba o abajo"
            List1.AddItem "Leer por palabras: Control + flechas izquierda o derecha"
            List1.AddItem "Ir al principio del renglón: Inicio"
            List1.AddItem "Ir al fin del renglón: Fin"
            List1.AddItem "Ir rápido al comienzo de la evaluación: Control + inicio"
            List1.AddItem "Ir rápido al fin de la evaluación: Control + fin"
            List1.AddItem "Leer de a párrafos: Control + flechas arriba o abajo"
            List1.AddItem "Leer el texto seleccionado: Alt + flecha derecha"
        Case formularios.lectorLibros
            dóndeEstoy = " el lector de libros "
            cadenaVoz = "Entrando en la ayuda del "
            List1.AddItem "Volver a las carpetas: F3"
            List1.AddItem "Imprimir: F5"
            List1.AddItem "Leer todo el texto: Alt + flecha abajo"
            List1.AddItem "Leer desde el cursor hacia adelanta: Alt + flecha arriba"
            List1.AddItem "Leer por renglones: Flechas arriba o abajo"
            List1.AddItem "Leer por palabras: Control + flechas izquierda o derecha"
            List1.AddItem "Ir al principio del renglón: Inicio"
            List1.AddItem "Ir al fin del renglón: Fin"
            List1.AddItem "Ir rápido al comienzo del libro: Control + inicio"
            List1.AddItem "Ir rápido al fin del libro: Control + fin"
            List1.AddItem "Leer de a párrafos: Control + flechas arriba o abajo"
            List1.AddItem "Leer el texto seleccionado: Alt + flecha derecha"
        Case formularios.libros
            dóndeEstoy = " los libros "
            cadenaVoz = "Entrando en la ayuda de los libros. seleccioná con flechas arriba o abajo el libro que desees y abrí sus capítulos con enter. "
            List1.AddItem "Moverse por los libros: Flechas arriba o abajo"
            List1.AddItem "Abrir los capítulos del libro seleccionado: Enter"
        Case formularios.libroX
            dóndeEstoy = " los capítulos de libros "
            cadenaVoz = "Entrando en la ayuda de los capítulos del libro que elegiste. para moverte por los capítulos usá flechas arriba o abajo. para abrir el capítulo seleccionado apretá enter. "
            List1.AddItem "Moverse por los capítulos: Flechas arriba o abajo"
            List1.AddItem "Abrir un capítulo: Enter"
        Case formularios.materiasEvaluaciones
            dóndeEstoy = " las materias para evaluar "
            cadenaVoz = "Entrando en la ayuda de las evaluaciones. elegí con flechas arriba o abajo la materia de la que quieras hacer una evaluación y aceptá con enter. "
            List1.AddItem "Moverse por las materias: Flechas arriba o abajo"
            List1.AddItem "Seleccionar una materia: Enter"
        Case formularios.mesEvaluaciones
            dóndeEstoy = " los meses con evaluaciones "
            cadenaVoz = "Entrando en la ayuda de los meses con evaluaciones guardadas. para moverte por los meses, usá las flechas arriba o abajo. para ver las hojas de el mes seleccionado, apretá enter. Si un mes no aparece en la lista, es que ese mes no tiene hojas guardadas. "
            List1.AddItem "Moverse por los meses: Flechas arriba o abajo"
            List1.AddItem "Mostrar las evaluaciones del mes seleccionado: Enter"
        Case formularios.principal
            dóndeEstoy = " el menú principal "
            cadenaVoz = "Entrando en la ayuda del menú principal. Usá flechas arriba o abajo para buscar la materia que quieras abrir y hacelo con enter. "
            List1.AddItem "Moverse por las materias: Flechas arriba o abajo"
            List1.AddItem "Abrir una materia: Enter"
        Case formularios.recordatorios
            dóndeEstoy = " los recordatorios "
            cadenaVoz = "Entrando en la ayuda de los recordatorios. para moverte por las opciones usá flechas arriba o abajo y para aceptar una de ellas, usá enter. "
            List1.AddItem "Moverse por las opciones: Flechas arriba o abajo"
            List1.AddItem "Aceptar una opción: Enter"
        Case formularios.reproductorMúsica
            dóndeEstoy = " el reproductor de música "
            cadenaVoz = "Entrando en la ayuda del reproductor de música. para buscar el tema que quieras escuchar, buscalo con flechas arriba o abajo y aceptá con enter. Si escuchás que el reproductor te dice carpeta, eso indica que si das enter vas a abrir esa carpeta. Si en cambio dice música, eso indica que es un tema, y que si das enter va a empezar a sonar. Para pausar una reproducción usá espacio, y para silenciar totalmente al reproductor, apretá shift. Para bajar el volumen apretá el número uno, y para subirlo usá el dos. Cuando estás buscando el tema para escuchar, si apretás la tecla suprimir, vas directamente a los discos de tu computadora, mientras que si apretás borrar, volvés a la carpeta anterior a la que estás ahora. Para cerrar el reproductor apretá escape y para dejar el reproductor abierto, pero seguir trabajando en tu mochila, apretá f 7."
            List1.AddItem "Buscar temas o carpetas: flechas arriba o abajo y enter"
            List1.AddItem "Reproducir: enter cuando estás en un tema"
            List1.AddItem "Pausa: espacio"
            List1.AddItem "Parar la reproducción: enter"
            List1.AddItem "Silenciar el reproductor: shift"
            List1.AddItem "Subir el volumen: número uno"
            List1.AddItem "Bajar el volumen: número dos"
            List1.AddItem "Ir a los discos de tu computadora: suprimir"
            List1.AddItem "Volver a la carpeta anterior: borrar"
            List1.AddItem "Cerrar el reproductor: escape"
            List1.AddItem "Pasar a la carpeta: F7"
        Case formularios.tareasAnt
            dóndeEstoy = " los meses que tienen hojas ya guardadas "
            cadenaVoz = "Entrando en la ayuda de los meses con hojas ya escritas. para moverte por los meses, usá las flechas arriba o abajo. para ver las hojas de el mes seleccionado, apretá enter. Si un mes no aparece en la lista, es que ese mes no tiene hojas guardadas. "
            List1.AddItem "Moverse por los meses: Flechas arriba o abajo"
            List1.AddItem "Mostrar las hojas del mes seleccionado: Enter"
        Case formularios.tecladoBraille
            dóndeEstoy = " el teclado braille "
            cadenaVoz = "Entrando en la ayuda del teclado braille. para los puntos del braille, usá solamente las teclas efe, de, ese, jota, ka y ele. para borrar usá la tecla de borrar, y para bajar una línea usá enter. "
            List1.AddItem "Punto 1: F"
            List1.AddItem "Punto 2: D"
            List1.AddItem "Punto 3: S"
            List1.AddItem "Punto 4: J"
            List1.AddItem "Punto 5: K"
            List1.AddItem "Punto 6: L"
            List1.AddItem "Bajar una línea: Enter"
            List1.AddItem "Borrar una letra: Borrar"
            List1.AddItem "Espacio: Espacio"
        Case formularios.títuloEvaluación
            dóndeEstoy = " el cuadro para ponerle un título a tu evaluación "
            cadenaVoz = "Entrando en la ayuda del título de tu evaluación. para leer lo que has escrito, usá flechas arriba o abajo y para guardar el título usá enter. "
            List1.AddItem "Leer lo escrito: Flechas arriba o abajo"
            List1.AddItem "Aceptar el título: Enter"
        Case formularios.títuloHoja
            dóndeEstoy = " el cuadro para ponerle un título a tu hoja "
            cadenaVoz = "Entrando en la ayuda del título de tu hoja. para leer lo que has escrito, usá flechas arriba o abajo y para guardar el título usá enter. "
            List1.AddItem "Leer lo escrito: Flechas arriba o abajo"
            List1.AddItem "Aceptar el título: Enter"
    End Select
    List1.AddItem "Abrir la configuración de la Mochila: F12"
    
    cadenaVoz = cadenaVoz + ". Si querés repasar lentamente la ayuda que te acabo de dar, usá flecha arriba o abajo. Para cerrar esta ayuda apreta escape."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then Unload Me
    
    If KeyCode = vbKeyF12 Then Call ButtonTransparent1_Click
'    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en la ayuda"
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    
End Sub

Private Sub Form_Activate()
    List1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Decir "saliendo de la ayuda de " + dóndeEstoy
    'Call contarFormularios(False)
End Sub

Private Sub List1_GotFocus()
    If swReciénAbierto = True Then
        Decir cadenaVoz + List1.List(List1.ListIndex)
        swReciénAbierto = False
    Else
        Decir List1.List(List1.ListIndex)
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer, cadena As String
    shiftkey = Shift And 7
    If shiftkey <> vbCtrlMask And KeyCode <> vbKeyReturn Then
        Decir "estás en la ayuda de " + dóndeEstoy + " . Movete con flecha arriba o abajo. Para salir, apretá Escape."
    End If
    
    If formulario = formularios.cuaderno Then
        If shiftkey = 0 And KeyCode = vbKeyReturn Then
            If List1.ListIndex <> -1 Then
                Select Case List1.List(List1.ListIndex)
                    Case "Ayuda para trabajar con actividades"
                        cadena = "la ayuda para abrir, leer y realizar actividades"
                        Call abrirCategoría(categoríaCuaderno.actividades)
                    Case "Ayuda para leer libros"
                        cadena = "la ayuda para leer libros"
                        Call abrirCategoría(categoríaCuaderno.libros)
                    Case "Ayuda con las hojas ya escritas de tu carpeta"
                        cadena = "la ayuda para abrir, releer y escribir en hojas ya escritas de tu carpeta"
                        Call abrirCategoría(categoríaCuaderno.hojas)
                    Case "Ayuda para leer un texto"
                        cadena = "la ayuda para leer el texto que está escrito en la carpeta"
                        Call abrirCategoría(categoríaCuaderno.lectura)
                    Case "Ayuda para escribir un texto"
                        cadena = "la ayuda para escribir texto en tu carpeta"
                        Call abrirCategoría(categoríaCuaderno.escritura)
                    Case "Ayuda para seleccionar parte de un texto"
                        cadena = "la ayuda para seleccionar texto que has escrito en la hoja"
                        Call abrirCategoría(categoríaCuaderno.selección)
                    Case "Ayuda con otros comandos"
                        cadena = " una lista de varios comandos útiles"
                        Call abrirCategoría(categoríaCuaderno.comandos)
                    Case "Volver a la lista de categorías"
                        Call llenarCategoríasCuaderno
                End Select
                
                If List1.List(List1.ListIndex) = "Volver a la lista de categorías" Then
                    Decir "Volviendo a todas las categorías de ayuda de tu carpeta, usá flecha abajo para elegir una y aceptá con enter"
                Else
                    Decir "abriendo " + cadena + ". usá las flechas para leerla o elegí volver a la lista de categorías para cambiar de ayuda"
                End If
            End If
        End If
        
        If shiftkey = 0 And KeyCode = vbKeyBack Then Call llenarCategoríasCuaderno
    End If
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
'        If swReciénAbierto = True Then
'            Decir cadenaVoz + List1.List(List1.ListIndex)
'            swReciénAbierto = False
'        Else
            Decir List1.List(List1.ListIndex)
            Exit Sub
'        End If
    End If
    
    
'    sonido = sndPlaySound(App.Path + "\sonidos\td.wav", SND_ASYNC)
End Sub

Sub llenarCategoríasCuaderno()
    List1.Clear
    List1.AddItem "Ayuda para trabajar con actividades"
    List1.AddItem "Ayuda para leer libros"
    List1.AddItem "Ayuda con las hojas ya escritas de tu carpeta"
    List1.AddItem "Ayuda para leer un texto"
    List1.AddItem "Ayuda para escribir un texto"
    List1.AddItem "Ayuda para seleccionar parte de un texto"
    List1.AddItem "Ayuda con otros comandos"
End Sub

Sub abrirCategoría(cuálCategoría As Byte)
    List1.Clear
    Select Case cuálCategoría
        Case categoríaCuaderno.escritura
            List1.AddItem "Guardar la hoja: Control + G ó F5"
            List1.AddItem "Imprimir: Control + P ó F6"
            List1.AddItem "Poner la letra en negrita: Control + N"
            List1.AddItem "Poner la letra en subrayado: Control + S"
            List1.AddItem "Copiar texto seleccionado: Control + C"
            List1.AddItem "Cortar texto seleccionado: Control + X"
            List1.AddItem "Pegar el texto copiado o cortado: Control + V"
            List1.AddItem "Abrir el teclado Braille: F9"
            List1.AddItem "Ir al principio de la hoja: Control + Inicio"
            List1.AddItem "Ir al final de la hoja: Control + Fin"
            List1.AddItem "Ir al principio del renglón: Inicio"
            List1.AddItem "Ir al final del renglón: Fin"
            List1.AddItem "Avanzar varios renglones en la hoja de un salto: Avance de Página"
            List1.AddItem "Retroceder varios renglones en la hoja de un salto: Retroceso de Página"
            List1.AddItem "Retroceder un párrafo: Control + Flecha Arriba"
            List1.AddItem "Avazar un párrafo: Control + Flecha Abajo"
        Case categoríaCuaderno.lectura
            List1.AddItem "Leer de a letras avanzando: Flecha derecha"
            List1.AddItem "Leer de a letras retrocediendo: Flecha izquierda"
            List1.AddItem "Leer de a palabras avanzando: Control + Flecha derecha"
            List1.AddItem "Leer de a palabras retrocediendo: Control + Flecha izquierda"
            List1.AddItem "Leer de a renglones bajando en la hoja: Flecha abajo"
            List1.AddItem "Leer de a renglones subiendo en la hoja: Flecha arriba"
            List1.AddItem "Leer todo el texto: Alt + Flecha Abajo"
            List1.AddItem "Leer desde el cursor hacia adelante: Alt + Flecha Arriba"
        Case categoríaCuaderno.selección
            List1.AddItem "Leer el texto seleccionado: Alt + Flecha derecha"
            List1.AddItem "Seleccionar de a palabras avanzando: Shift + Control + Flecha derecha"
            List1.AddItem "Seleccionar de a palabras retrocediendo: Shift + Control + Flecha izquierda"
            List1.AddItem "Seleccionar todo el texto desde donde uno está hasta el principio de la hoja: Shift + Control + Inicio"
            List1.AddItem "Seleccionar todo el texto desde donde uno está hasta el final de la hoja: Shift + Control + Fin"
            List1.AddItem "Seleccionar desde donde uno está hasta el principio del párrafo: Shift + Control + Flecha Arriba"
            List1.AddItem "Seleccionar desde donde uno está hasta el final del párrafo: Shift + Control + Flecha Abajo"
            List1.AddItem "Seleccionar varios renglones hacia abajo: Shift + Control + Retroceso de Página"
            List1.AddItem "Seleccionar varios renglones hacia arriba: Shift + Control + Avance de Página"
            List1.AddItem "Seleccionar hasta el principio del renglón: Shift + Inicio"
            List1.AddItem "Seleccionar hasta el final del renglón: Shift + Fin"
            List1.AddItem "Seleccionar desde donde uno está hasta el renglón inferior: Shift + Flecha Abajo"
            List1.AddItem "Seleccionar desde donde uno está hasta el renglón superior: Shift + Flecha Arriba"
            List1.AddItem "Seleccionar varios renglones hacia arriba: Shift + Avance de Página"
            List1.AddItem "Seleccionar varios renglones hacia abajo: Shift + Retroceso de Página"
            List1.AddItem "Seleccionar toda la hoja: Control + A"
        Case categoríaCuaderno.comandos
            List1.AddItem "Callar la voz: Control"
            List1.AddItem "Abrir los accesorios: F4"
            List1.AddItem "Abrir el reproductor de música: F7"
            List1.AddItem "Salir de la Carpeta: Escape"
            List1.AddItem "Salir del programa: Alt + F4"
            List1.AddItem "Abrir la configuración: F12"
        Case categoríaCuaderno.actividades
            List1.AddItem "Abrir una actividad: F1"
            List1.AddItem "Cambiar a una actividad abierta: F2"
            List1.AddItem "Leer una actividad: igual que como se lee en la carpeta"
        Case categoríaCuaderno.hojas
            List1.AddItem "Abrir una hoja ya escrita: F1"
        Case categoríaCuaderno.libros
            List1.AddItem "Abrir un libro: F1"
            List1.AddItem "Cambiar a un libro abierto: F3"
            List1.AddItem "Leer un libro: igual que como se lee en la carpeta"
    End Select
    List1.AddItem "Volver a la lista de categorías"
End Sub
