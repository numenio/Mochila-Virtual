VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmCalendarioMúltiple 
   Caption         =   "Actividades guardadas"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   Icon            =   "frmCalendarioUsosMúltiples.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmCalendarioUsosMúltiples.frx":08CA
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
      Picture         =   "frmCalendarioUsosMúltiples.frx":2922
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
      Caption         =   "Mostrar todos los días del mes"
      EstiloDelBoton  =   1
      Picture         =   "frmCalendarioUsosMúltiples.frx":31FC
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
Attribute VB_Name = "frmCalendarioMúltiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MesParaAbrir As String
Public numMesParaAbrir As Byte
Private Type trabajo
    directorio As String
    día As Integer
    índiceEnListBox As Integer
End Type
Public año As Integer
Dim trabajos() As trabajo
Public tipoElemento As Byte
Public swMateriaEvaluaciones As String
Dim elementoPlural As String
Dim recordat() As Recordatorio
Dim quéElemento As String
'Dim swTodosLosAños As Boolean
Dim swTodosLosDías As Boolean
Dim swPulsóEnterParaAvanzar As Boolean

Private Sub ButtonTransparent1_Click()
    If swTodosLosDías = False Then
        swTodosLosDías = True
        ButtonTransparent1.Caption = "Mostrar sólo días con " + quéElemento
        Decir "mostrando todos los días del mes, tanto los que tienen " + quéElemento + " y los que no tienen"
    Else
        swTodosLosDías = False
        ButtonTransparent1.Caption = "Mostrar todos los días"
        Decir "mostrando solamente los días con " + quéElemento
    End If
    
    Select Case tipoElemento
        Case elemento.Recordatorio
            Call llenarCalendarioRecordatorios(año, numMesParaAbrir)
        Case elemento.actividad
            Call llenarCalendario(numMesParaAbrir, año, elemento.actividad, "actividad", "actividades", swTodosLosDías)
        Case elemento.tarea
            Call llenarCalendario(numMesParaAbrir, año, elemento.tarea, "tarea", "tareas", swTodosLosDías)
        Case elemento.evaluación
            Call llenarCalendario(numMesParaAbrir, año, elemento.evaluación, "evaluación", "evaluaciones", swTodosLosDías)
    End Select
End Sub

Private Sub Command1_Click()
    Call Form_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub Command1_MouseIn(Shift As Integer)
    sonido = sndPlaySound(App.Path + "\sonidos\cb.wav", SND_ASYNC)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cadena As String, cadenaAux As String, cadenaAux2 As String, díaACargar As Byte, shiftkey As Byte
    shiftkey = Shift And 7
    
    If KeyCode = vbKeyEscape Then
'        If swCuadernoAbierto = True Then
'            Decir "volviendo a tu carpeta"
'            frmCuaderno.Show
'        End If
'
'        If tipoElemento = elemento.evaluación Or tipoElemento = elemento.Recordatorio Then
'            Decir "volviendo al menú principal, elegí con las flechas qué materia querés abrir"
'            frmPrincipal.Show
'        End If
        Unload Me
        Exit Sub
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.CalendarioMúltiple
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
                            Call llenarCalendarioRecordatorios(año, numMesParaAbrir)
                        Case elemento.actividad
                            Call llenarCalendario(numMesParaAbrir, año, elemento.actividad, "actividad", "actividades", swTodosLosDías)
                        Case elemento.actividad
                            Call llenarCalendario(numMesParaAbrir, año, elemento.tarea, "tarea", "tareas", swTodosLosDías)
                        Case elemento.evaluación
                            Call llenarCalendario(numMesParaAbrir, año, elemento.evaluación, "evaluación", "evaluaciones", swTodosLosDías)
                    End Select
                    Decir "volviendo al calendario. elegí con las flechas un día y aceptá con enter"
                Case Else
                    If Left(List1.List(List1.ListIndex), 1) = "D" Then 'si está en en el calendario
                        If Right(List1.List(List1.ListIndex), 5) <> "vacío" Then
                            cadenaAux = Mid(List1.List(List1.ListIndex), 5, 2)
                            If InStrRev(cadenaAux, ":") Then cadenaAux = Mid(List1.List(List1.ListIndex), 4, 2)
                            díaACargar = CByte(cadenaAux)
                            
                            If CInt(cadenaAux) = 1 Then
                                cadenaAux2 = " uno "
                            Else
                                cadenaAux2 = cadenaAux
                            End If
                            
                            If tipoElemento = elemento.Recordatorio Then
                                Decir "abriendo los recordatorios del día " + cadenaAux2 + ". usá las flechas para elegir y aceptá con enter"
                                Call cargarRecordatorios(díaACargar, numMesParaAbrir, año)
                            Else
                                Decir "abriendo las " + quéElemento + " del día " + cadenaAux2 + ". usá las flechas para elegir y aceptá con enter"
                                If tipoElemento = elemento.actividad Then
                                    cadena = "Actividad"
                                ElseIf tipoElemento = elemento.tarea Then
                                    cadena = "Hoja"
                                ElseIf tipoElemento = elemento.evaluación Then
                                    cadena = "Evaluación"
                                End If
                                
                                Call cargarElementos(tipoElemento, CByte(cadenaAux), cadena)
                            End If
                        Else
                            If tipoElemento = elemento.actividad Then
                                cadenaAux = "actividad"
                            Else
                                cadenaAux = "hoja de " + miMateria
                            End If
                            Decir "el día está vacío. no hay ninguna " + quéElemento + " guardada"
                        End If
                    Else 'si está en los elementos
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
    
    swPulsóEnterParaAvanzar = False
    
    'Dim rec As Recordatorio
    File1.Refresh
    Select Case tipoElemento
        Case elemento.actividad
            quéElemento = "actividades"
            Me.Caption = "Actividades del mes: " + MesParaAbrir
            Decir "abriendo las actividades de " + miMateria + " del mes de " + MesParaAbrir + ". usá las flechas para elegir el día en que está la actividad que quieras abrir y aceptá con enter"
            File1.Path = App.Path + dirTrabajo + "actividades\" + Trim(Str(numMesParaAbrir))
            Call llenarCalendario(numMesParaAbrir, año, elemento.actividad, "actividad", "actividades", swTodosLosDías)
        Case elemento.tarea
            quéElemento = "hojas guardadas"
            Me.Caption = "Hojas guardadas del mes: " + MesParaAbrir
            Decir "abriendo las hojas de tu carpeta de " + miMateria + " del mes de " + MesParaAbrir + ". usá las flechas para elegir el día en que está la hoja que quieras abrir y aceptá con enter"
            File1.Path = App.Path + dirTrabajo + Trim(Str(numMesParaAbrir))
            Call llenarCalendario(numMesParaAbrir, año, elemento.tarea, "tarea", "tareas", swTodosLosDías)
        Case elemento.Recordatorio
            quéElemento = "recordatorios"
            Me.Caption = "Recordatorios del mes: " + MesParaAbrir
            Decir "abriendo los recordatorios del mes de " + MesParaAbrir + ". Elegí con las flechas el día en que querés ver los recordatorios que has añadido"
            Call llenarCalendarioRecordatorios(año, numMesParaAbrir)
        Case elemento.evaluación
            quéElemento = "evaluaciones"
            Me.Caption = swMateriaEvaluaciones + ": Evaluaciones del mes " + MesParaAbrir
            Decir "abriendo las hojas de tu carpeta de " + miMateria + " del mes de " + MesParaAbrir + ". usá las flechas para elegir el día en que está la hoja que quieras abrir y aceptá con enter"
            File1.Path = App.Path + "\trabajos\" + swMateriaEvaluaciones + "\soporte\" + Trim(Str(numMesParaAbrir)) + "\"
            Call llenarCalendario(numMesParaAbrir, año, elemento.evaluación, "evaluación", "evaluaciones", swTodosLosDías)
     End Select
End Sub

Sub abrirRecordatorio()
'    Dim hora As Date
'    hora = recordat(List1.ListIndex).hora
'    hora = Format(hora, "HH:mm")
    If List1.List(List1.ListIndex) <> "No hay ningún recordatorio en " + MesParaAbrir Then
        frmAñadirRecordatorio.swEditar = True
        frmAñadirRecordatorio.swAño = año
        frmAñadirRecordatorio.swDía = Left(recordat(List1.ListIndex).fecha, 2)
        frmAñadirRecordatorio.swMes = numMesParaAbrir
        frmAñadirRecordatorio.swHora = Left(Format(recordat(List1.ListIndex).hora, "HH:mm"), 2)
        frmAñadirRecordatorio.swMinutos = Right(Format(recordat(List1.ListIndex).hora, "HH:mm"), 2)
        frmAñadirRecordatorio.swTexto = recordat(List1.ListIndex).texto
        frmAñadirRecordatorio.Show
        swPulsóEnterParaAvanzar = True
        Unload Me
    End If
End Sub

Sub cargarRecordatorios(día As Byte, mes As Byte, año As Integer)
    Dim archivo As Byte, auxRecordatorio As Recordatorio, contador As Integer
    
    contador = 0
    List1.Clear
    archivo = FreeFile 'se abre el archivo para guardar los datos de las partidas
    Open App.Path + "\recordatorios\" + Trim(año) + "\" + Trim(Str(mes)) + "\recordatorios.gui" For Random As #archivo Len = Len(auxRecordatorio)
    Do While Not EOF(archivo)   ' Repite hasta el final del archivo.
       Get #archivo, , auxRecordatorio   ' Lee el registro siguiente.
       If Right(Format(auxRecordatorio.fecha, "dd/mm/yyyy"), 4) <> "1899" Then '#12:00:00 AM# Then
            If día = Left(Format(auxRecordatorio.fecha, "dd/mm/yyyy"), 2) Then 'si es del día seleccionado
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

Sub llenarCalendario(mes As Byte, año As Integer, tipoElemento As Byte, elementoSingular As String, elementoPlural As String, mostrarTodoslosDías As Boolean)
    Dim i As Integer, j As Integer, cadenaAux As String, contador As Integer, cadena As String
    Dim contadorElementos As Integer
    List1.Clear
    List1.Refresh
    File1.Refresh
    For j = 1 To cantDíasMes(mes, año)
        cadena = "Día " + Trim(Str(j)) + ": "
        For i = 0 To File1.ListCount - 1 'se ven todos los archivos del mes seleccionado
            If Left(Right(File1.List(i), 8), 4) = año Then 'si el año es igual al parámetro
                If Mid(File1.List(i), cantPrefijo + 1, 2) = j Then 'si el día es el mismo que el que se está mirando
                    contadorElementos = contadorElementos + 1
                    ReDim Preserve trabajos(0 To contador)
                    trabajos(contador).directorio = File1.List(i)
                    trabajos(contador).día = j
                    trabajos(contador).índiceEnListBox = 25000
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
            cadena = cadena + "vacío" '"sin " + elemento
        End If
        
        If mostrarTodoslosDías = False Then
            If contadorElementos Then List1.AddItem cadena
        Else
            List1.AddItem cadena
        End If
        contadorElementos = 0
    Next
End Sub

Sub cargarElementos(tipoElemento As Byte, quéDía As Byte, quéEscribirElemento)
    Dim j As Integer, contador As Integer, cadena As String, cadenaAux As String * 64
    Dim miRegistro As DatosActividad, archivolibre As Byte
    
    archivolibre = FreeFile
    List1.Clear
    List1.Refresh
    contador = 1
    For j = 0 To UBound(trabajos)
        If quéDía = trabajos(j).día Then
            'se carga la actividad
            cadena = quéEscribirElemento + " " + Trim(Str(contador)) + " del día " + Trim(Str(trabajos(j).día))
            
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
'                cadena = cadena + ". Título: " + Trim(cadenaAux)
                
                If Asc(Left(miRegistro.tema, 1)) Then
                    cadena = cadena + ". Tema: " + Trim(miRegistro.tema)
                Else
                    cadena = cadena + ". Sin tema"
                End If
            End If
            
            'si es una tarea, se le carga el título
            If tipoElemento = elemento.tarea Then
                Open App.Path + dirTrabajo + Trim(Str(numMesParaAbrir)) + "\datosHojas\" + Left(trabajos(j).directorio, Len(trabajos(j).directorio) - 4) + ".gui" For Random As #archivolibre Len = 64
                Get #archivolibre, 1, cadenaAux   ' Lee el registro
                Close #archivolibre   ' Cierra el archivo.
                cadenaAux = Right(cadenaAux, Len(cadenaAux) - InStr(1, cadenaAux, Chr(0)))
                cadenaAux = Left(cadenaAux, InStr(1, cadenaAux, Chr(0)) - 1)
                cadena = cadena + ". Título: " + Trim(cadenaAux)
            End If
            
            'si es una evaluación, también se le carga el título
            If tipoElemento = elemento.evaluación Then
                Open App.Path + "\trabajos\" + swMateriaEvaluaciones + "\soporte\" + Trim(Str(numMesParaAbrir)) + "\datosSoporte\" + Left(trabajos(j).directorio, Len(trabajos(j).directorio) - 4) + ".gui" For Random As #archivolibre Len = 64
                Get #archivolibre, 1, cadenaAux   ' Lee el registro
                Close #archivolibre   ' Cierra el archivo.
                cadenaAux = Right(cadenaAux, Len(cadenaAux) - InStr(1, cadenaAux, Chr(0)))
                cadenaAux = Left(cadenaAux, InStr(1, cadenaAux, Chr(0)) - 1)
                cadena = cadena + ". Título: " + Trim(cadenaAux)
            End If
            
            List1.AddItem cadena
            trabajos(j).índiceEnListBox = contador - 1
            contador = contador + 1
        End If
    Next
    List1.AddItem "Volver al calendario"
End Sub

Sub abrirElemento(tipoElemento As Byte)
    Dim i As Integer, dóndeEmpezar As Byte
    
    For i = 0 To UBound(trabajos)
        If List1.ListIndex = trabajos(i).índiceEnListBox Then Exit For
    Next
    Select Case tipoElemento
        Case elemento.actividad
            frmLectorActividad.archivoParaLeer = Trim(Str(numMesParaAbrir)) + "\" + trabajos(i).directorio
            frmLectorActividad.díaAbierto = List1.List(List1.ListIndex)
            If swActividadAbierta = True Then frmLectorActividad.cargarActividad 'si ya estaba abierto el lector de actividades, que cargue la actividad que se le indica aunque no entre al load
            frmLectorActividad.Show
        Case elemento.tarea
            frmCuaderno.nombreArchivo = trabajos(i).directorio
            frmCuaderno.nombreMesArchivo = numMesParaAbrir
            frmCuaderno.swContinuarArchivo = True
            frmCuaderno.swAbriendoHojaAnterior = True
            dóndeEmpezar = InStr(1, List1.Text, "día") - 1
            frmCuaderno.díaAbierto = Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - dóndeEmpezar)
            frmCuaderno.RichTextBox1.LoadFile App.Path + dirTrabajo + Trim(Str(numMesParaAbrir)) + "\" + frmCuaderno.nombreArchivo
            Decir "abriendo la hoja de " + miMateria + " del " + frmCuaderno.díaAbierto + ". podés seguir trabajando en ella"
            frmCuaderno.Show
        Case elemento.evaluación
            frmLectorEvaluaciones.swArchivoParaLeer = trabajos(i).directorio
            frmLectorEvaluaciones.swDíaParaAbrir = trabajos(i).día
            frmLectorEvaluaciones.swMateriaParaAbrir = swMateriaEvaluaciones
            frmLectorEvaluaciones.swNumMesParaAbrir = numMesParaAbrir
            frmLectorEvaluaciones.swSóloLeer = True
            frmLectorEvaluaciones.Show
    End Select
    swPulsóEnterParaAvanzar = True
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
    
    If swPulsóEnterParaAvanzar = False Then
        If swCuadernoAbierto = True Then
            Decir "volviendo a tu carpeta"
            frmCuaderno.Show
        End If
        
        If tipoElemento = elemento.evaluación Or tipoElemento = elemento.Recordatorio Then
            Decir "volviendo al menú principal, elegí con las flechas qué materia querés abrir"
            frmPrincipal.Show
        End If
    End If

    'Call contarFormularios(False)
End Sub

Sub llenarCalendarioRecordatorios(año As Integer, mes As Byte)
    Dim archivo As Byte, contador As Integer, auxRecordatorio As Recordatorio
    Dim día As Byte, j As Integer, contadorRepetidos As Integer, contadorRegistros As Integer
    Dim contadorRec() As Byte 'contadorRecordatorios
    
    List1.Clear
    
    For j = 1 To cantDíasMes(mes, año)
        'cadena = "Día " + Trim(Str(j)) + ": vacío"
        List1.AddItem "Día " + Trim(Str(j)) + ": vacío" 'cadena
    Next
    
    On Error GoTo manejoError
    archivo = FreeFile 'se abre el archivo para guardar los datos de las partidas
    Open App.Path + "\recordatorios\" + Trim(año) + "\" + Trim(Str(mes)) + "\recordatorios.gui" For Random As #archivo Len = Len(auxRecordatorio)
    Do While Not EOF(archivo)   ' Repite hasta el final del archivo.
       Get #archivo, , auxRecordatorio   ' Lee el registro siguiente.
       If Right(Format(auxRecordatorio.fecha, "dd/mm/yyyy"), 4) <> "1899" Then
            día = Left(Format(auxRecordatorio.fecha, "dd/mm/yyyy"), 2)
            
            ReDim Preserve contadorRec(0 To contadorRegistros)
            For j = 0 To UBound(contadorRec)
                If contadorRec(j) = día Then contadorRepetidos = contadorRepetidos + 1
            Next
            
            contadorRec(contadorRegistros) = día
'            ReDim Preserve recordat(0 To contadorRegistros)
'            recordat(contadorRegistros) = auxRecordatorio
            contadorRegistros = contadorRegistros + 1
            contadorRepetidos = contadorRepetidos + 1 'se le suma uno pues si está repetido hay uno previo
            
            If contadorRepetidos = 1 Then
                List1.List(día - 1) = "Día " + Trim(Str(día)) + ": tiene 1 recordatorio"
            Else
                List1.List(día - 1) = "Día " + Trim(Str(día)) + ": tiene " + Trim(Str(contadorRepetidos)) + " recordatorios"
            End If
            
        End If
        contadorRepetidos = 0
    Loop
    Close #archivo
    
    If swTodosLosDías = False Then
        For j = cantDíasMes(mes, año) To 1 Step -1  'se sacan los días que no tienen recordatorios, desde el último hasta el primero
            If Right(List1.List(j - 1), 5) = "vacío" Then List1.RemoveItem j - 1
        Next
        Select Case tipoElemento
            Case elemento.Recordatorio
                If List1.ListCount = 0 Then List1.AddItem "No hay ningún recordatorio en " + MesParaAbrir
        End Select
    End If
    Exit Sub
manejoError:
    If Err.Number = 76 Then
        List1.Clear
        List1.AddItem "No hay ningún recordatorio en " + MesParaAbrir
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
