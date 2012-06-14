VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmCalendarioM�ltiple 
   Caption         =   "Actividades guardadas"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   Icon            =   "frmCalendarioUsosM�ltiples.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmCalendarioUsosM�ltiples.frx":08CA
   ScaleHeight     =   6735
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   6270
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TransparentButton.ButtonTransparent Command1 
      Height          =   615
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6000
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      Caption         =   "   Abrir"
      EstiloDelBoton  =   1
      Picture         =   "frmCalendarioUsosM�ltiples.frx":2922
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
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
      Caption         =   "Mostrar todos los d�as del mes"
      EstiloDelBoton  =   1
      Picture         =   "frmCalendarioUsosM�ltiples.frx":31FC
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
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
End
Attribute VB_Name = "frmCalendarioM�ltiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MesParaAbrir As String
Public numMesParaAbrir As Byte
Private Type trabajo
    directorio As String
    d�a As Integer
    �ndiceEnListBox As Integer
End Type
Public a�o As Integer
Dim trabajos() As trabajo
Public tipoElemento As Byte
Public swMateriaEvaluaciones As String
Dim elementoPlural As String
Dim recordat() As Recordatorio
Dim qu�Elemento As String
'Dim swTodosLosA�os As Boolean
Dim swTodosLosD�as As Boolean
Dim swPuls�EnterParaAvanzar As Boolean

Private Sub ButtonTransparent1_Click()
    If swTodosLosD�as = False Then
        swTodosLosD�as = True
        ButtonTransparent1.Caption = "Mostrar s�lo d�as con " + qu�Elemento
        Decir "mostrando todos los d�as del mes, tanto los que tienen " + qu�Elemento + " y los que no tienen"
    Else
        swTodosLosD�as = False
        ButtonTransparent1.Caption = "Mostrar todos los d�as"
        Decir "mostrando solamente los d�as con " + qu�Elemento
    End If
    
    Select Case tipoElemento
        Case elemento.Recordatorio
            Call llenarCalendarioRecordatorios(a�o, numMesParaAbrir)
        Case elemento.actividad
            Call llenarCalendario(numMesParaAbrir, a�o, elemento.actividad, "actividad", "actividades", swTodosLosD�as)
        Case elemento.tarea
            Call llenarCalendario(numMesParaAbrir, a�o, elemento.tarea, "tarea", "tareas", swTodosLosD�as)
        Case elemento.evaluaci�n
            Call llenarCalendario(numMesParaAbrir, a�o, elemento.evaluaci�n, "evaluaci�n", "evaluaciones", swTodosLosD�as)
    End Select
End Sub

Private Sub Command1_Click()
    Call Form_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub Command1_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.Path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cadena As String, cadenaAux As String, cadenaAux2 As String, d�aACargar As Byte, shiftkey As Byte
    shiftkey = Shift And 7
    
    If KeyCode = vbKeyEscape Then
'        If swCuadernoAbierto = True Then
'            Decir "volviendo a tu carpeta"
'            frmCuaderno.Show
'        End If
'
'        If tipoElemento = elemento.evaluaci�n Or tipoElemento = elemento.Recordatorio Then
'            Decir "volviendo al men� principal, eleg� con las flechas qu� materia quer�s abrir"
'            frmPrincipal.Show
'        End If
        Unload Me
        Exit Sub
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.CalendarioM�ltiple
         frmAyuda.Show
         Exit Sub
    End If
    
    If KeyCode = vbKeyBack Then 'borrar regresa al calendario
        If Left(List1.List(List1.ListCount - 1), 1) <> "D" Then
            List1.ListIndex = List1.ListCount - 1
            Call Form_KeyDown(vbKeyReturn, 0)
        End If
    End If
    
    If KeyCode = vbKeySpace Then ButtonTransparent1_Click
    
    
    If KeyCode = vbKeyReturn Then 'si se da enter
        If List1.ListIndex <> -1 Then 'si hay algo seleccionado
            Select Case List1.List(List1.ListIndex)
                Case "Volver al calendario"
                    Select Case tipoElemento
                        Case elemento.Recordatorio
                            Call llenarCalendarioRecordatorios(a�o, numMesParaAbrir)
                        Case elemento.actividad
                            Call llenarCalendario(numMesParaAbrir, a�o, elemento.actividad, "actividad", "actividades", swTodosLosD�as)
                        Case elemento.actividad
                            Call llenarCalendario(numMesParaAbrir, a�o, elemento.tarea, "tarea", "tareas", swTodosLosD�as)
                        Case elemento.evaluaci�n
                            Call llenarCalendario(numMesParaAbrir, a�o, elemento.evaluaci�n, "evaluaci�n", "evaluaciones", swTodosLosD�as)
                    End Select
                    Decir "volviendo al calendario. eleg� con las flechas un d�a y acept� con enter"
                Case Else
                    If Left(List1.List(List1.ListIndex), 1) = "D" Then 'si est� en en el calendario
                        If Right(List1.List(List1.ListIndex), 5) <> "vac�o" Then
                            cadenaAux = Mid(List1.List(List1.ListIndex), 5, 2)
                            If InStrRev(cadenaAux, ":") Then cadenaAux = Mid(List1.List(List1.ListIndex), 4, 2)
                            d�aACargar = CByte(cadenaAux)
                            
                            If CInt(cadenaAux) = 1 Then
                                cadenaAux2 = " uno "
                            Else
                                cadenaAux2 = cadenaAux
                            End If
                            
                            If tipoElemento = elemento.Recordatorio Then
                                Decir "abriendo los recordatorios del d�a " + cadenaAux2 + ". us� las flechas para elegir y acept� con enter"
                                Call cargarRecordatorios(d�aACargar, numMesParaAbrir, a�o)
                            Else
                                Decir "abriendo las " + qu�Elemento + " del d�a " + cadenaAux2 + ". us� las flechas para elegir y acept� con enter"
                                If tipoElemento = elemento.actividad Then
                                    cadena = "Actividad"
                                ElseIf tipoElemento = elemento.tarea Then
                                    cadena = "Hoja"
                                ElseIf tipoElemento = elemento.evaluaci�n Then
                                    cadena = "Evaluaci�n"
                                End If
                                
                                Call cargarElementos(tipoElemento, CByte(cadenaAux), cadena)
                            End If
                        Else
                            If tipoElemento = elemento.actividad Then
                                cadenaAux = "actividad"
                            Else
                                cadenaAux = "hoja de " + miMateria
                            End If
                            Decir "el d�a est� vac�o. no hay ninguna " + qu�Elemento + " guardada"
                        End If
                    Else 'si est� en los elementos
                        If tipoElemento = elemento.Recordatorio Then
                            Call abrirRecordatorio
                        Else
                            Call abrirElemento(tipoElemento)
                        End If
                    End If
            End Select
        End If
    End If
    
    If shiftkey = vbCtrlMask Then Decir ""
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    
    swPuls�EnterParaAvanzar = False
    
    'Dim rec As Recordatorio
    File1.Refresh
    Select Case tipoElemento
        Case elemento.actividad
            qu�Elemento = "actividades"
            Me.Caption = "Actividades del mes: " + MesParaAbrir
            Decir "abriendo las actividades de " + miMateria + " del mes de " + MesParaAbrir + ". us� las flechas para elegir el d�a en que est� la actividad que quieras abrir y acept� con enter"
            File1.Path = App.Path + dirTrabajo + "actividades\" + Trim(Str(numMesParaAbrir))
            Call llenarCalendario(numMesParaAbrir, a�o, elemento.actividad, "actividad", "actividades", swTodosLosD�as)
        Case elemento.tarea
            qu�Elemento = "hojas guardadas"
            Me.Caption = "Hojas guardadas del mes: " + MesParaAbrir
            Decir "abriendo las hojas de tu carpeta de " + miMateria + " del mes de " + MesParaAbrir + ". us� las flechas para elegir el d�a en que est� la hoja que quieras abrir y acept� con enter"
            File1.Path = App.Path + dirTrabajo + Trim(Str(numMesParaAbrir))
            Call llenarCalendario(numMesParaAbrir, a�o, elemento.tarea, "tarea", "tareas", swTodosLosD�as)
        Case elemento.Recordatorio
            qu�Elemento = "recordatorios"
            Me.Caption = "Recordatorios del mes: " + MesParaAbrir
            Decir "abriendo los recordatorios del mes de " + MesParaAbrir + ". Eleg� con las flechas el d�a en que quer�s ver los recordatorios que has a�adido"
            Call llenarCalendarioRecordatorios(a�o, numMesParaAbrir)
        Case elemento.evaluaci�n
            qu�Elemento = "evaluaciones"
            Me.Caption = swMateriaEvaluaciones + ": Evaluaciones del mes " + MesParaAbrir
            Decir "abriendo las hojas de tu carpeta de " + miMateria + " del mes de " + MesParaAbrir + ". us� las flechas para elegir el d�a en que est� la hoja que quieras abrir y acept� con enter"
            File1.Path = App.Path + "\trabajos\" + swMateriaEvaluaciones + "\soporte\" + Trim(Str(numMesParaAbrir)) + "\"
            Call llenarCalendario(numMesParaAbrir, a�o, elemento.evaluaci�n, "evaluaci�n", "evaluaciones", swTodosLosD�as)
     End Select
End Sub

Sub abrirRecordatorio()
'    Dim hora As Date
'    hora = recordat(List1.ListIndex).hora
'    hora = Format(hora, "HH:mm")
    If List1.List(List1.ListIndex) <> "No hay ning�n recordatorio en " + MesParaAbrir Then
        frmA�adirRecordatorio.swEditar = True
        frmA�adirRecordatorio.swA�o = a�o
        frmA�adirRecordatorio.swD�a = Left(recordat(List1.ListIndex).fecha, 2)
        frmA�adirRecordatorio.swMes = numMesParaAbrir
        frmA�adirRecordatorio.swHora = Left(Format(recordat(List1.ListIndex).hora, "HH:mm"), 2)
        frmA�adirRecordatorio.swMinutos = Right(Format(recordat(List1.ListIndex).hora, "HH:mm"), 2)
        frmA�adirRecordatorio.swTexto = recordat(List1.ListIndex).texto
        frmA�adirRecordatorio.Show
        swPuls�EnterParaAvanzar = True
        Unload Me
    End If
End Sub

Sub cargarRecordatorios(d�a As Byte, mes As Byte, a�o As Integer)
    Dim archivo As Byte, auxRecordatorio As Recordatorio, contador As Integer
    
    contador = 0
    List1.Clear
    archivo = FreeFile 'se abre el archivo para guardar los datos de las partidas
    Open App.Path + "\recordatorios\" + Trim(a�o) + "\" + Trim(Str(mes)) + "\recordatorios.gui" For Random As #archivo Len = Len(auxRecordatorio)
    Do While Not EOF(archivo)   ' Repite hasta el final del archivo.
       Get #archivo, , auxRecordatorio   ' Lee el registro siguiente.
       If Right(Format(auxRecordatorio.fecha, "dd/mm/yyyy"), 4) <> "1899" Then '#12:00:00 AM# Then
            If d�a = Left(Format(auxRecordatorio.fecha, "dd/mm/yyyy"), 2) Then 'si es del d�a seleccionado
                List1.AddItem "Recordatorio " + Str(contador) + ". Texto: " + auxRecordatorio.texto
                ReDim Preserve recordat(0 To contador)
                recordat(contador) = auxRecordatorio
                contador = contador + 1
            End If
        End If
    Loop
    List1.AddItem "Volver al calendario"
    Close #archivo
End Sub

Sub llenarCalendario(mes As Byte, a�o As Integer, tipoElemento As Byte, elementoSingular As String, elementoPlural As String, mostrarTodoslosD�as As Boolean)
    Dim i As Integer, j As Integer, cadenaAux As String, contador As Integer, cadena As String
    Dim contadorElementos As Integer
    List1.Clear
    List1.Refresh
    File1.Refresh
    For j = 1 To cantD�asMes(mes, a�o)
        cadena = "D�a " + Trim(Str(j)) + ": "
        For i = 0 To File1.ListCount - 1 'se ven todos los archivos del mes seleccionado
            If Left(Right(File1.List(i), 8), 4) = a�o Then 'si el a�o es igual al par�metro
                If Mid(File1.List(i), cantPrefijo + 1, 2) = j Then 'si el d�a es el mismo que el que se est� mirando
                    contadorElementos = contadorElementos + 1
                    ReDim Preserve trabajos(0 To contador)
                    trabajos(contador).directorio = File1.List(i)
                    trabajos(contador).d�a = j
                    trabajos(contador).�ndiceEnListBox = 25000
                    contador = contador + 1
                End If
            End If
        Next
            
        If contadorElementos <> 0 Then
            If contadorElementos = 1 Then
                cadena = cadena + " una " + elementoSingular 'Str(contadorElementos) + " " + elementoSingular
            Else
                cadena = cadena + Str(contadorElementos) + " " + elementoPlural
            End If
        Else
            cadena = cadena + "vac�o" '"sin " + elemento
        End If
        
        If mostrarTodoslosD�as = False Then
            If contadorElementos Then List1.AddItem cadena
        Else
            List1.AddItem cadena
        End If
        contadorElementos = 0
    Next
End Sub

Sub cargarElementos(tipoElemento As Byte, qu�D�a As Byte, qu�EscribirElemento)
    Dim j As Integer, contador As Integer, cadena As String, cadenaAux As String * 64
    Dim miRegistro As DatosActividad, archivolibre As Byte
    
    archivolibre = FreeFile
    List1.Clear
    List1.Refresh
    contador = 1
    For j = 0 To UBound(trabajos)
        If qu�D�a = trabajos(j).d�a Then
            'se carga la actividad
            cadena = qu�EscribirElemento + " " + Trim(Str(contador)) + " del d�a " + Trim(Str(trabajos(j).d�a))
            
            'si es una actividad, se le carga el tema
            If tipoElemento = elemento.actividad Then
'                Open App.Path + "\datos\" + Trim(Str(numMesParaAbrir)) + "\datosActividades.gui" For Random As #1 Len = Len(miRegistro)
'                Do While Not EOF(1)   ' Repite hasta el final del archivo.
'                   Get #1, , miRegistro   ' Lee el registro siguiente.
'                   If Right(Trim(miRegistro.DirArchivo), Len(Trim(miRegistro.DirArchivo)) - 3) = Right(App.Path + dirTrabajo + "actividades\" + Trim(Str(numMesParaAbrir)) + "\" + trabajos(j).directorio, Len(App.Path + dirTrabajo + "actividades\" + Trim(Str(numMesParaAbrir)) + "\" + trabajos(j).directorio) - 3) Then Exit Do
'                Loop
'                Close #1   ' Cierra el archivo.
                
                Open App.Path + dirTrabajo + "actividades\" + Trim(Str(numMesParaAbrir)) + "\datosActividades\" + Left(trabajos(j).directorio, Len(trabajos(j).directorio) - 4) + ".gui" For Random As #archivolibre Len = Len(miRegistro)
                Get #archivolibre, 1, miRegistro   ' Lee el regitro
                Close #archivolibre   ' Cierra el archivo.
'                cadenaAux = Right(cadenaAux, Len(cadenaAux) - InStr(1, cadenaAux, Chr(0)))
'                cadenaAux = Left(cadenaAux, InStr(1, cadenaAux, Chr(0)) - 1)
'                cadena = cadena + ". T�tulo: " + Trim(cadenaAux)
                
                If Asc(Left(miRegistro.tema, 1)) Then
                    cadena = cadena + ". Tema: " + Trim(miRegistro.tema)
                Else
                    cadena = cadena + ". Sin tema"
                End If
            End If
            
            'si es una tarea, se le carga el t�tulo
            If tipoElemento = elemento.tarea Then
                Open App.Path + dirTrabajo + Trim(Str(numMesParaAbrir)) + "\datosHojas\" + Left(trabajos(j).directorio, Len(trabajos(j).directorio) - 4) + ".gui" For Random As #archivolibre Len = 64
                Get #archivolibre, 1, cadenaAux   ' Lee el registro
                Close #archivolibre   ' Cierra el archivo.
                cadenaAux = Right(cadenaAux, Len(cadenaAux) - InStr(1, cadenaAux, Chr(0)))
                cadenaAux = Left(cadenaAux, InStr(1, cadenaAux, Chr(0)) - 1)
                cadena = cadena + ". T�tulo: " + Trim(cadenaAux)
            End If
            
            'si es una evaluaci�n, tambi�n se le carga el t�tulo
            If tipoElemento = elemento.evaluaci�n Then
                Open App.Path + "\trabajos\" + swMateriaEvaluaciones + "\soporte\" + Trim(Str(numMesParaAbrir)) + "\datosSoporte\" + Left(trabajos(j).directorio, Len(trabajos(j).directorio) - 4) + ".gui" For Random As #archivolibre Len = 64
                Get #archivolibre, 1, cadenaAux   ' Lee el registro
                Close #archivolibre   ' Cierra el archivo.
                cadenaAux = Right(cadenaAux, Len(cadenaAux) - InStr(1, cadenaAux, Chr(0)))
                cadenaAux = Left(cadenaAux, InStr(1, cadenaAux, Chr(0)) - 1)
                cadena = cadena + ". T�tulo: " + Trim(cadenaAux)
            End If
            
            List1.AddItem cadena
            trabajos(j).�ndiceEnListBox = contador - 1
            contador = contador + 1
        End If
    Next
    List1.AddItem "Volver al calendario"
End Sub

Sub abrirElemento(tipoElemento As Byte)
    Dim i As Integer, d�ndeEmpezar As Byte
    
    For i = 0 To UBound(trabajos)
        If List1.ListIndex = trabajos(i).�ndiceEnListBox Then Exit For
    Next
    Select Case tipoElemento
        Case elemento.actividad
            frmLectorActividad.archivoParaLeer = Trim(Str(numMesParaAbrir)) + "\" + trabajos(i).directorio
            frmLectorActividad.d�aAbierto = List1.List(List1.ListIndex)
            If swActividadAbierta = True Then frmLectorActividad.cargarActividad 'si ya estaba abierto el lector de actividades, que cargue la actividad que se le indica aunque no entre al load
            frmLectorActividad.Show
        Case elemento.tarea
            frmCuaderno.nombreArchivo = trabajos(i).directorio
            frmCuaderno.nombreMesArchivo = numMesParaAbrir
            frmCuaderno.swContinuarArchivo = True
            frmCuaderno.swAbriendoHojaAnterior = True
            d�ndeEmpezar = InStr(1, List1.Text, "d�a") - 1
            frmCuaderno.d�aAbierto = Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - d�ndeEmpezar)
            frmCuaderno.RichTextBox1.LoadFile App.Path + dirTrabajo + Trim(Str(numMesParaAbrir)) + "\" + frmCuaderno.nombreArchivo
            Decir "abriendo la hoja de " + miMateria + " del " + frmCuaderno.d�aAbierto + ". pod�s seguir trabajando en ella"
            frmCuaderno.Show
        Case elemento.evaluaci�n
            frmLectorEvaluaciones.swArchivoParaLeer = trabajos(i).directorio
            frmLectorEvaluaciones.swD�aParaAbrir = trabajos(i).d�a
            frmLectorEvaluaciones.swMateriaParaAbrir = swMateriaEvaluaciones
            frmLectorEvaluaciones.swNumMesParaAbrir = numMesParaAbrir
            frmLectorEvaluaciones.swS�loLeer = True
            frmLectorEvaluaciones.Show
    End Select
    swPuls�EnterParaAvanzar = True
    Unload Me
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
    
    If swPuls�EnterParaAvanzar = False Then
        If swCuadernoAbierto = True Then
            Decir "volviendo a tu carpeta"
            frmCuaderno.Show
        End If
        
        If tipoElemento = elemento.evaluaci�n Or tipoElemento = elemento.Recordatorio Then
            Decir "volviendo al men� principal, eleg� con las flechas qu� materia quer�s abrir"
            frmPrincipal.Show
        End If
    End If

    'Call contarFormularios(False)
End Sub

Sub llenarCalendarioRecordatorios(a�o As Integer, mes As Byte)
    Dim archivo As Byte, contador As Integer, auxRecordatorio As Recordatorio
    Dim d�a As Byte, j As Integer, contadorRepetidos As Integer, contadorRegistros As Integer
    Dim contadorRec() As Byte 'contadorRecordatorios
    
    List1.Clear
    
    For j = 1 To cantD�asMes(mes, a�o)
        'cadena = "D�a " + Trim(Str(j)) + ": vac�o"
        List1.AddItem "D�a " + Trim(Str(j)) + ": vac�o" 'cadena
    Next
    
    On Error GoTo manejoError
    archivo = FreeFile 'se abre el archivo para guardar los datos de las partidas
    Open App.Path + "\recordatorios\" + Trim(a�o) + "\" + Trim(Str(mes)) + "\recordatorios.gui" For Random As #archivo Len = Len(auxRecordatorio)
    Do While Not EOF(archivo)   ' Repite hasta el final del archivo.
       Get #archivo, , auxRecordatorio   ' Lee el registro siguiente.
       If Right(Format(auxRecordatorio.fecha, "dd/mm/yyyy"), 4) <> "1899" Then
            d�a = Left(Format(auxRecordatorio.fecha, "dd/mm/yyyy"), 2)
            
            ReDim Preserve contadorRec(0 To contadorRegistros)
            For j = 0 To UBound(contadorRec)
                If contadorRec(j) = d�a Then contadorRepetidos = contadorRepetidos + 1
            Next
            
            contadorRec(contadorRegistros) = d�a
'            ReDim Preserve recordat(0 To contadorRegistros)
'            recordat(contadorRegistros) = auxRecordatorio
            contadorRegistros = contadorRegistros + 1
            contadorRepetidos = contadorRepetidos + 1 'se le suma uno pues si est� repetido hay uno previo
            
            If contadorRepetidos = 1 Then
                List1.List(d�a - 1) = "D�a " + Trim(Str(d�a)) + ": tiene 1 recordatorio"
            Else
                List1.List(d�a - 1) = "D�a " + Trim(Str(d�a)) + ": tiene " + Trim(Str(contadorRepetidos)) + " recordatorios"
            End If
            
        End If
        contadorRepetidos = 0
    Loop
    Close #archivo
    
    If swTodosLosD�as = False Then
        For j = cantD�asMes(mes, a�o) To 1 Step -1  'se sacan los d�as que no tienen recordatorios, desde el �ltimo hasta el primero
            If Right(List1.List(j - 1), 5) = "vac�o" Then List1.RemoveItem j - 1
        Next
        Select Case tipoElemento
            Case elemento.Recordatorio
                If List1.ListCount = 0 Then List1.AddItem "No hay ning�n recordatorio en " + MesParaAbrir
        End Select
    End If
    Exit Sub
manejoError:
    If Err.Number = 76 Then
        List1.Clear
        List1.AddItem "No hay ning�n recordatorio en " + MesParaAbrir
    End If
End Sub

Private Sub List1_DblClick()
    Call Form_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim cadenaAux As String
    If List1.ListIndex <> -1 Then
'        cadenaAux = Mid(List1.List(List1.ListIndex), 5, 2)
'        If InStrRev(cadenaAux, ":") Then cadenaAux = Mid(List1.List(List1.ListIndex), 4, 2)
'
'        If CInt(cadenaAux) = 1 Then cadenaAux = " una "
'
        
        Decir List1.Text
    End If
End Sub
