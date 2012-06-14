Attribute VB_Name = "funciones"
Option Explicit

Public Voz As SpVoice
Public vozSapi4 As DirectSS 'As TextToSpeech
Private Enum selección
    creció
    disminuyó
    igual
End Enum

'funciones para registrar los ocx
'Public Declare Function RegFlash Lib "Flash9f.ocx" Alias "DllRegisterServer" () As Long
'Public Declare Function UnRegFlash Lib "Flash9f.ocx" Alias "DllUnregisterServer" () As Long
'Public Declare Function RegBotónTransp Lib "TransparentButton.ocx" Alias "DllRegisterServer" () As Long
'Public Declare Function UnRegBotónTransp Lib "TransparentButton.ocx" Alias "DllUnregisterServer" () As Long

Public Function guardarOrdenCapítulosDesdeMatriz(materia As String, libro As String, ParamArray listaCapítulos()) As Boolean
    Dim capítulo As Variant, auxMatriz() As String, i As Integer, contador As Integer, archivolibre As Byte
    On Error GoTo error
    archivolibre = FreeFile
    Open App.path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos" For Output As #archivolibre
    contador = 0
    For Each capítulo In listaCapítulos
        For i = 0 To UBound(capítulo)
            ReDim Preserve auxMatriz(0 To contador)
            auxMatriz(contador) = capítulo(contador)
            contador = contador + 1
        Next
    Next capítulo
    
    For i = 0 To UBound(auxMatriz)
        Print #archivolibre, auxMatriz(i)
    Next i
    
    Close #archivolibre
    
    guardarOrdenCapítulosDesdeMatriz = True
    Exit Function
error:
    guardarOrdenCapítulosDesdeMatriz = False
End Function


Public Sub guardarMaterias(listaMaterias As ListBox)
    Dim i As Integer
    Open App.path + "\datos\materias.txt" For Output As #1 'se abre el trabajo ya guardado
    listaMaterias.Refresh
    For i = 0 To listaMaterias.ListCount - 1
        Print #1, listaMaterias.List(i)
    Next
    Close #1
End Sub

Public Sub guardarOrdenCapítulosLibro(listaCapítulos As ListBox, materia As String, libro As String)
    Dim i As Integer
    Open App.path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos" For Output As #1
    listaCapítulos.Refresh
    For i = 0 To listaCapítulos.ListCount - 1
        Print #1, listaCapítulos.List(i)
    Next
    Close #1
End Sub


Public Sub guardarHistorial(listaMaterias As ListBox)
    Dim swArchivoRepetido As Boolean, nextline As String, i As Integer
    
    On Error GoTo manejoError
    'se guardan las materias en el historial
    For i = 0 To listaMaterias.ListCount - 1 'se chequea que cada materia no esté ya guardada
        Open App.path + "\datos\historialMaterias.txt" For Input As #1 'se abre el trabajo ya guardado para leerlo
        Do While Not EOF(1) 'chequeamos que no esté en la lista ya el registro del archivo a guardar
            Line Input #1, nextline
            If nextline = listaMaterias.List(i) Then
                swArchivoRepetido = True
                Exit Do
            End If
        Loop
        Close #1
        
        If swArchivoRepetido = False Then
            Open App.path + "\datos\historialMaterias.txt" For Append As #1 'se abre el historial para añadir las materias
            Print #1, listaMaterias.List(i)
            swArchivoRepetido = False
            Close #1
        End If
        
        swArchivoRepetido = False
    Next
    Exit Sub
    
manejoError:
    Open App.path + "\datos\historialMaterias.txt" For Output As #1
    Close #1
    Resume
End Sub


Sub llenarComboVoz(Combo1 As ComboBox, Combo2 As ComboBox)
    On Error GoTo manejoErrorSapi
    
    Set Voz = Nothing
    Set vozSapi4 = Nothing
    Set Voz = New SpVoice
    Set vozSapi4 = New DirectSS 'TextToSpeech
    
    Dim Token As ISpeechObjectToken

    For Each Token In Voz.GetVoices
        Combo1.AddItem (Token.GetDescription())
    Next
    
    Dim i As Integer, modename As String
    
    For i = 1 To vozSapi4.CountEngines
'        swYaesSapi5 = False
        modename = vozSapi4.modename(i)
'        For j = 0 To combo1.ListCount
'            If combo1.List(j) = modename Then swYaesSapi5 = True
'        Next j
'        If swYaesSapi5 = False Then
        Combo2.AddItem modename
    Next i
    
    If Combo1.ListCount = 0 Then
        Set Voz = Nothing
        Combo1.AddItem "No hay voces avanzadas (SAPI5) instaladas"
        Combo1.ListIndex = 0
        Combo1.Enabled = False
        frmControl.Option8(0).Enabled = False
    Else
        For i = 0 To Combo1.ListCount - 1 'se activa la voz sapi5 que el usuario usó en la sesión anterior
            If Trim(nombreSapi5) = Combo1.List(i) Then
                Combo1.ListIndex = i
                Exit For
            End If
        Next
        If Combo1.ListIndex = -1 Then Combo1.ListIndex = 0 'si la voz no estaba, la voz es la primera del combo
    End If
    
    If Combo2.ListCount = 0 Then
        Set vozSapi4 = Nothing
        Combo2.AddItem "No hay voces simples (SAPI4) instaladas"
        Combo2.ListIndex = 0
        Combo2.Enabled = False
        frmControl.Option8(1).Enabled = False
'        If swInstalarVoz = False Then 'se intenta instalar la voz del programa una sola vez
'            MsgBox "Se va a instalar la voz del programa. Por favor, presione el botón " + Chr(34) + "Sí" + Chr(34) + " en el cuadro que va a aparecer.", , "Información"
'            Call ejecutar(App.Path + "\ejecutables\TTS3000.exe")
'            swInstalarVoz = True
'        End If
    Else
        For i = 0 To Combo2.ListCount - 1 'se activa la voz sapi4 que el usuario usó en la sesión anterior
            If Trim(nombreSapi4) = Combo2.List(i) Then
                vozSapi4.Speak ""
                Combo2.ListIndex = i
                Exit For
            End If
        Next
        If Combo1.ListIndex = -1 Then Combo1.ListIndex = 0 'si la voz no estaba, la voz es la primera del combo
    End If
    Exit Sub
manejoErrorSapi:
'    Dim aceptar As Byte
    If Err.Number = 424 Then
        MsgBox "El programa necesita tener instaladas Sapi4 y Sapi5. Una de ellas o ambas faltan. Por favor instálelas y reinicie el programa. Ambas sapi están en la carpeta " + Chr(34) + "Ejecutables" + Chr(34) + "  que viene con el programa. Si por cualquier eventualidad no estuviesen allí, se pueden descargar gratuitamente desde la página de Microsoft. Tenga presente instalar la sapi5 que es propia de su Windows, ya sea sapi5 para Windows XP, ó sapi5 para versiones anteriores. Sapi4 es idéntica para cualquier versión de Windows.", , "No se encuentra una SAPI"
        End
    End If
    
    If Err.Number = 53 Then
        'On Error GoTo manejoerrorCancelar
'        aceptar =
        MsgBox "No se encuentra el instalador de la voz que viene con el programa. Instálelo usted manualmente desde la carpeta " + Chr(34) + "Ejecutables" + Chr(34) + "  que viene con el programa. Si por cualquier eventualidad no estuviese allí, se puede descargar gratuitamente desde la página de Microsoft con el nombre de TTS3000.", , "Imposible instalar la voz del programa"
'        If aceptar = vbYes Then
'            frmControl.diálogo.CancelError = True
'            frmControl.diálogo.Filter = "Archivos Ejecutables (*.exe);*.rtf"
'            frmControl.diálogo.ShowOpen
'            If Err.Number = cdlCancel Then Exit Sub
'            Shell frmControl.diálogo.FileName
'        End If
        Exit Sub
    End If
    
'    If Err.Number = 429 Then MsgBox "soy el controlador del módulo funciones, error 429", , "Para mi creador"
    
'    MsgBox "soy el controlador del la función llenarComboVoz. Error número: " + Str(Err.Number) + ", descripción: " + Err.Description, , "Para mi creador"
    frmMsgBox.cadenaAMostrar = "Soy el controlador del la función llenarComboVoz. Error número: " + Str(Err.Number) + ", descripción: " + Err.Description
    frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
    frmMsgBox.Show 1
    Exit Sub
'    Exit Sub
'manejoerrorCancelar:
'    Exit Sub
End Sub

'Sub guardarTrabajo(NombreArchivo, textoAGuardar)
'    Dim nextLine As String 'aquí se almacena el contenido del registro de los guardados para chequear que no se repita un nombre
'    Dim cadena As String 'para almacenar la lista de guardados cuando se trabaja con tareas viejas
'
'    On Error GoTo manejoError
''    If swNuevoContinuar = False Then 'si se trabaja en un archivo nuevo
'        Open App.Path + NombreArchivo For Output As #1   ' Abre el archivo para operaciones de salida.
''    Else 'si se está trabajando con una tarea vieja
''        Open MiTrabajo For Output As #1 'se abre el trabajo ya guardado
''    End If
'    Print #1, textoAGuardar
'    Close #1
'
'    Open App.Path + "\trabajos\listadeGuardados.txt" For Input As #2
'    Do Until EOF(2) 'chequeamos que no esté en la lista ya el registro del archivo a guardar
'        Line Input #2, nextLine
'        If swNuevoContinuar = False Then
'            If nextLine = NombreArchivo Then
'                Close #2
'                Exit Sub
'            End If
'        Else
'            If nextLine <> MiTrabajo Then cadena = cadena + nextLine
'        End If
'        'contenidoGuardados = LinesFromFile + nextline + Chr(13) + Chr(10)
'    Loop
'    Close #2
'    Open App.Path + "\trabajos\listadeGuardados.txt" For Append As #2 'abrimos el registro de los archivos guardados para añadir este que hemos guardado ahora
'    If swNuevoContinuar = False Then
'        Print #2, NombreArchivo
'    Else
'        Print #2, cadena
'    End If
'    Close #2
'    Exit Sub
'manejoError:
'    Open App.Path + NombreArchivo For Append As #1
'    Close #1
'    Open App.Path + "\trabajos\listadeGuardados.txt" For Append As #2
'    Close #2
'    Resume
'End Sub

'Public Sub GuardarRTF(nombreRuta As String, cuálRTF As RichTextBox)
'        cuálRTF.SaveFile nombreRuta, rtfRTF
'End Sub
'
'
Public Sub GuardarDatosUsuario()
    Dim archivolibre As Byte
    With usuario
        '.comenzarEnCarpeta = swEmpezarEnCuaderno
        .sapi5 = swSapi5 'si se usa sapi 5 o sapi 4
'        .permitirEditarActividades = swPermitirCambioEnActividades
        .usarVoz = swHablarVoz
        .mostrarTodasLasTareas = swMostrarAñoEnTareas
        .mostrarTodasLasActividades = swMostrarAñoEnActividades
        .nombre = nombreUsuario
        .usuarioMujer = swUsuarioMujer
        '.leerSignoPuntuación = swLeerSignosPuntuación
        .imprimirDirecto = swImprimirDirecto
        .colorFondo = colorFondo
        .fuenteColor = colorFuente
        .fuenteNombre = NombreFuente
        .fuenteTamaño = tamañoFuente
        .velocidadVoz = velocidadVoz
        .swLeerRenglones = swLeerRenglones
        .swUsarCorrectorOrtográfico = swUsarCorrectorOrtográfico
        .nombreVozSapi4 = nombreSapi4
        .nombreVozSapi5 = nombreSapi5
        '.swInstalarVoz = swInstalarVoz
        .swMúsicaDeFondo = swMúsicaDeFondo
        .swPermitirAbrirArchivos = swPermitirAbrirArchivos
    End With

    archivolibre = FreeFile 'se abre el archivo para guardar los datos de las partidas
    Open App.path + "\datos\datos.gui" For Random As archivolibre Len = Len(usuario)
'    contador = 0
'    While Not EOF(archivolibre)
'        contador = contador + 1
'        Get archivolibre, contador, auxErr
'    Wend
    Put archivolibre, 1, usuario
    Close archivolibre
End Sub

Public Sub Decir(ByVal qué As String, Optional usarbanderasSpVoice As Boolean = True, Optional esperarSapi4 As Boolean)
    On Error Resume Next 'esto se lo agregué por la compu de franco, que tira un error siempre
    If swHablarVoz = True Then 'variable general del programa
        If swSapi5 = True Then
            If usarbanderasSpVoice = True Then
                Voz.Speak qué, SVSFPurgeBeforeSpeak Or SVSFlagsAsync ' Or SVSFNLPSpeakPunc
            Else
                Voz.Speak qué, SVSFPurgeBeforeSpeak Or SVSFlagsAsync
            End If
        Else
            If esperarSapi4 = False Then
                vozSapi4.AudioReset
                vozSapi4.Speak qué
            Else
                vozSapi4.Speak qué
            End If
        End If
    End If
End Sub

'Public Sub guardarMaterias(listaMaterias As ListBox)
'    Dim i As Integer
'    Open App.Path + "\datos\materias.txt" For Output As #1 'se abre el trabajo ya guardado
'    For i = 0 To listaMaterias.ListCount - 1
'        Print #1, listaMaterias.List(i)
'    Next
'    Close #1
'End Sub

'Public Function decirPalabra(cuadroRTF As RichTextBox) As String
'    Dim cont As Long, cadena As String
'    cont = cuadroRTF.SelStart
'    cont = cont + 1
'    cadena = Mid(cuadroRTF.Text, cont, 1) 'si se está en el comienzo del cuadro de texto
'    While Right(cadena, 1) <> " " And Right(cadena, 1) <> "," And _
'    Right(cadena, 1) <> "." And Right(cadena, 1) <> ";" And _
'    Right(cadena, 1) <> ":" And Right(cadena, 1) <> "?" And _
'    Right(cadena, 1) <> "!" And Len(cuadroRTF.Text) <> cont
'        cont = cont + 1
'        cadena = cadena + Mid(cuadroRTF.Text, cont, 1)
'    Wend
'
'    If Len(cuadroRTF.Text) <> cont Then 'si no se está al final de la hoja
'        decirPalabra = cadena
'    Else
'        decirPalabra = cadena '"estás al final de tu hoja"
'    End If
'    ponerPuntoInserciónEn = cont
'End Function

Public Function decirPalabraSiguiente(cuadroRTF As RichTextBox) As String
    Dim cont As Long, cadena As String, renglónActual As Long
    cont = cuadroRTF.SelStart
    If swLeerRenglones = True Then renglónActual = cuadroRTF.GetLineFromChar(cuadroRTF.SelStart)
    If Len(cuadroRTF.Text) <> cont Then 'si no se está al final de la hoja
        cont = cont + 1
        cadena = Mid(cuadroRTF.Text, cont, 1) 'si se está en el comienzo del cuadro de texto
        While Right(cadena, 1) <> " " And Right(cadena, 1) <> "," And _
        Right(cadena, 1) <> "." And Right(cadena, 1) <> ";" And _
        Right(cadena, 1) <> ":" And Right(cadena, 1) <> "?" And _
        Right(cadena, 1) <> "!" And Right(cadena, 1) <> Chr(13) And _
        Right(cadena, 1) <> "-" And Right(cadena, 1) <> "+" And _
        Len(cuadroRTF.Text) <> cont
            cont = cont + 1
            cadena = cadena + Mid(cuadroRTF.Text, cont, 1)
        Wend
        
        If Len(cadena) = 1 And cadena = Chr(13) Then cadena = "aparte"
        If cadena = "." Then cadena = "punto"
        If cadena = "," Then cadena = "coma"
        If cadena = ";" Then cadena = "punto y coma"
        If cadena = ":" Then cadena = "dos puntos"
        If cadena = Chr(34) Then cadena = "comillas"
'        if cadena="(" then cadena="abre paréntesis"
        
        If cuadroRTF.SelStart <> 0 Then 'si no se está al principio de la hoja
            decirPalabraSiguiente = cadena
            If cuadroRTF.SelBold = True Then decirPalabraSiguiente = decirPalabraSiguiente + " en negrita"
            If IsNull(cuadroRTF.SelBold) Then decirPalabraSiguiente = decirPalabraSiguiente + " parte en negrita"
            If cuadroRTF.SelUnderline = True Then decirPalabraSiguiente = decirPalabraSiguiente + " subrayada"
            If IsNull(cuadroRTF.SelUnderline) Then decirPalabraSiguiente = decirPalabraSiguiente + " parte subrayada"
        Else
            If cadena = "." Or cadena = "," Or cadena = ";" Or cadena = "aparte" Then
                decirPalabraSiguiente = "estás al principio de la hoja, delante del signo. " + cadena
            Else
                decirPalabraSiguiente = "estás al principio de la hoja, delante de la palabra. " + cadena
                If cuadroRTF.SelBold = True Then decirPalabraSiguiente = decirPalabraSiguiente + " en negrita"
                If IsNull(cuadroRTF.SelBold) Then decirPalabraSiguiente = decirPalabraSiguiente + " parte en negrita"
                If cuadroRTF.SelUnderline = True Then decirPalabraSiguiente = decirPalabraSiguiente + " subrayada"
                If IsNull(cuadroRTF.SelUnderline) Then decirPalabraSiguiente = decirPalabraSiguiente + " parte subrayada"
            End If
        End If
        
        If swLeerRenglones = True Then
            If renglónActual <> renglónAnterior Then decirPalabraSiguiente = decirPalabraSiguiente & ". renglón " & Str(renglónActual + 1)
        End If
    Else
        If cuadroRTF.Text <> "" Then
            decirPalabraSiguiente = "llegaste al final de la hoja"
        Else
            decirPalabraSiguiente = "no hay nada escrito, la hoja está vacía"
        End If
    End If
    decirPalabraSiguiente = controlarCadena(decirPalabraSiguiente)
    renglónAnterior = renglónActual
End Function

Public Function decirPalabraAnterior(cuadroRTF As RichTextBox) As String
    Dim cont As Long, cadena As String
    If cuadroRTF.SelStart <> 0 Then 'si no se está al principio de la carpeta
        If cuadroRTF.Text <> "" Then
            cont = cuadroRTF.SelStart
            'If cont <> Len(cuadroRtf.Text) Then cont = cont - 1
            cadena = Mid(cuadroRTF.Text, cont, 1) 'si se está en el comienzo del cuadro de texto
            While Left(cadena, 1) <> " " And Left(cadena, 1) <> "," And _
            Left(cadena, 1) <> "." And Left(cadena, 1) <> ";" And _
            Left(cadena, 1) <> ":" And Left(cadena, 1) <> "?" And Left(cadena, 1) <> Chr(10) And _
            Left(cadena, 1) <> Chr(13) And Left(cadena, 1) <> "!" And cont <> 1
                cont = cont - 1
                cadena = Mid(cuadroRTF.Text, cont, 1) + cadena
            Wend
            
            'si es un número con . que no lea una parte del número sino todo él
            If IsNumeric(Right(cadena, Len(cadena) - 1)) And Left(cadena, 1) = "." And cont <> 1 Then
                'se toma el siguiente número
                cont = cont - 1
                cadena = Mid(cuadroRTF.Text, cont, 1) + cadena
                While Left(cadena, 1) <> " " And Left(cadena, 1) <> "," And _
                Left(cadena, 1) <> "." And Left(cadena, 1) <> ";" And _
                Left(cadena, 1) <> ":" And Left(cadena, 1) <> "?" And Left(cadena, 1) <> Chr(10) And _
                Left(cadena, 1) <> Chr(13) And Left(cadena, 1) <> "!" And cont <> 1
                    cont = cont - 1
                    cadena = Mid(cuadroRTF.Text, cont, 1) + cadena
                Wend
            End If
            
            decirPalabraAnterior = cadena
            
            If cuadroRTF.SelBold = True Then decirPalabraAnterior = decirPalabraAnterior + " en negrita"
            If IsNull(cuadroRTF.SelBold) Then decirPalabraAnterior = decirPalabraAnterior + " parte en negrita"
            If cuadroRTF.SelUnderline = True Then decirPalabraAnterior = decirPalabraAnterior + " subrayada"
            If IsNull(cuadroRTF.SelUnderline) Then decirPalabraAnterior = decirPalabraAnterior + " parte subrayada"
        Else
            decirPalabraAnterior = "no hay nada escrito en la hoja"
        End If
        
        decirPalabraAnterior = controlarCadena(decirPalabraAnterior)
    End If
End Function

Public Function decirLetraSiguiente(cuadroRTF As RichTextBox) As String
    Dim cont As Long, cadena As String, renglónActual As Long
    If swLeerRenglones = True Then renglónActual = cuadroRTF.GetLineFromChar(cuadroRTF.SelStart)
    cont = cuadroRTF.SelStart
    If Len(cuadroRTF.Text) <> cont Then
        cont = cont + 1
        cadena = Mid(cuadroRTF.Text, cont, 1) 'si se está en el comienzo del cuadro de texto
        If cadena = " " Then
            decirLetraSiguiente = "espacio"
        ElseIf cadena = Chr(13) Then
            decirLetraSiguiente = "bajada de línea del renglón " + Str(cuadroRTF.GetLineFromChar(cuadroRTF.SelStart) + 1)
        ElseIf cadena = ":" Then
            decirLetraSiguiente = "dos puntos"
        Else
            decirLetraSiguiente = cadena '"estás al final de tu hoja"
        End If
        
        If cuadroRTF.SelBold = True Then decirLetraSiguiente = decirLetraSiguiente + " en negrita"
        If cuadroRTF.SelUnderline = True Then decirLetraSiguiente = decirLetraSiguiente + " subrayada"

        If swLeerRenglones = True Then
            If renglónActual <> renglónAnterior Then decirLetraSiguiente = decirLetraSiguiente & ". renglón " & Str(renglónActual + 1)
        End If
    Else
        If cuadroRTF.Text <> "" Then
            decirLetraSiguiente = "llegaste al final de la hoja"
        Else
            decirLetraSiguiente = "no hay nada escrito, la hoja está vacía"
        End If
    End If
    decirLetraSiguiente = controlarCadena(decirLetraSiguiente)
    renglónAnterior = renglónActual
End Function

'Public Function decirLetraAnterior(cuadroRTF As RichTextBox) As String
'    Dim cont As Long, cadena As String
'    cont = cuadroRTF.SelStart
'    If cont <> Len(cuadroRTF.Text) Then cont = cont - 1
'    cadena = Mid(cuadroRTF.Text, cont, 1) 'si se está en el comienzo del cuadro de texto
'
''    If Len(cuadroRTF.Text) <> cont Then 'si no se está al final de la hoja
''        decirPalabraAnterior = cadena
''    Else
'        decirLetraAnterior = cadena '"estás al final de tu hoja"
''    End If
'End Function
'

Public Function decirOraciónSiguiente(cuadroRTF As RichTextBox) As String
    Dim cont As Long, cadena As String, swEmpezando As Boolean
    Dim líneaInicio As Long, líneaActual As Long, TotalDeLíneasEnRTF As Long
    
    líneaInicio = cuadroRTF.GetLineFromChar(cuadroRTF.SelStart)
    TotalDeLíneasEnRTF = cuadroRTF.GetLineFromChar(Len(cuadroRTF.Text))
    
    cont = cuadroRTF.SelStart
    
    If cont <> 0 Then 'si no se está ya al comienzo de la línea se busca el punto de inicio
        líneaActual = líneaInicio
        Do While líneaActual = líneaInicio
            cont = cont - 1
            líneaActual = cuadroRTF.GetLineFromChar(cont)
            If cont = 0 Then Exit Do
        Loop
    End If
    
    cont = cont + 1
    If líneaInicio = 0 Then swEmpezando = True
    líneaActual = líneaInicio
    While líneaActual = líneaInicio And cont <= Len(cuadroRTF.Text)
        If Mid(cuadroRTF.Text, cont, 1) <> Chr(10) And Mid(cuadroRTF.Text, cont, 1) <> Chr(13) Then
            cadena = cadena + Mid(cuadroRTF.Text, cont, 1)
        End If
        cont = cont + 1
        líneaActual = cuadroRTF.GetLineFromChar(cont)
    Wend
            
    If TotalDeLíneasEnRTF <> líneaInicio Then 'se evalúa si se está en el último renglón
        If swEmpezando = False Then 'si no se está al principio de la hoja
            If swLeerRenglones = True Then decirOraciónSiguiente = "renglón " + Str(líneaInicio + 1)
            If cadena = "" Then
                decirOraciónSiguiente = decirOraciónSiguiente + ". renglón en blanco"
            Else
                If swLeerRenglones = True Then
                    decirOraciónSiguiente = decirOraciónSiguiente + " dice." + cadena
                Else
                    decirOraciónSiguiente = decirOraciónSiguiente + cadena
                End If
            End If
        Else
            If swLeerRenglones = True Then decirOraciónSiguiente = "estás en el primer renglón de la hoja "
            If cadena = "" Then
                decirOraciónSiguiente = decirOraciónSiguiente + ". el renglón está en blanco"
            Else
                If swLeerRenglones = True Then
                    decirOraciónSiguiente = decirOraciónSiguiente + ". el renglón dice. " + cadena
                Else
                    decirOraciónSiguiente = decirOraciónSiguiente + cadena
                End If
            End If
        End If
    Else
        If cuadroRTF.Text <> "" Then
            If swLeerRenglones = True Then decirOraciónSiguiente = "llegaste a el último renglón de la hoja"
            If cadena = "" Then
                decirOraciónSiguiente = decirOraciónSiguiente + ". el renglón está en blanco"
            Else
                If swLeerRenglones = True Then
                    decirOraciónSiguiente = decirOraciónSiguiente + ". el renglón dice. " + cadena
                Else
                    decirOraciónSiguiente = decirOraciónSiguiente + cadena
                End If
            End If
        Else
            decirOraciónSiguiente = "no hay nada escrito, la hoja está vacía"
        End If
    End If
    
    decirOraciónSiguiente = controlarCadena(decirOraciónSiguiente)
End Function

Public Function controlarCadena(cadena As String) As String
    Dim carácter(7) As String, posiciónCaracter As Long, cadenaFinal As String
    Dim i As Byte, swEntróAlFor As Boolean, swYaEmpezó As Boolean
    
    If cadena <> "" Then
        carácter(0) = "("
        carácter(1) = ")"
        carácter(2) = "-"
        carácter(3) = "*"
        carácter(4) = "/"
        carácter(5) = "{"
        carácter(6) = "}"
        carácter(7) = " 1 "
        
        swEntróAlFor = False
        
        For i = 0 To UBound(carácter)
            If swYaEmpezó = False Then
                posiciónCaracter = InStr(1, cadena, carácter(i))
            Else
                posiciónCaracter = InStr(1, cadenaFinal, carácter(i))
            End If
            
            Do While posiciónCaracter <> 0
                If swEntróAlFor = False Then
                    cadenaFinal = corregirCadena(cadena, posiciónCaracter, carácter(i))
                Else
                    cadenaFinal = corregirCadena(cadenaFinal, posiciónCaracter, carácter(i))
                End If
                posiciónCaracter = InStr(posiciónCaracter + 1, cadenaFinal, carácter(i))
                swEntróAlFor = True
                swYaEmpezó = True
            Loop
        Next
        
        If cadenaFinal = "" Then cadenaFinal = cadena
        
        carácter(0) = " m "
        carácter(1) = " s "
        carácter(2) = " l "
        carácter(3) = " h "
        carácter(4) = " p "
        carácter(5) = "$"
        carácter(6) = "_"

        swEntróAlFor = False
        swYaEmpezó = False
        
        For i = 0 To 6
'            If swYaEmpezó = False Then
'                posiciónCaracter = InStr(1, cadena, carácter(i))
'            Else
                posiciónCaracter = InStr(1, cadenaFinal, carácter(i))
'            End If
            
            Do While posiciónCaracter <> 0
'                If swEntróAlFor = False Then
'                    cadenaFinal = corregirCadena(cadena, posiciónCaracter, carácter(i))
'                Else
                    cadenaFinal = corregirCadena(cadenaFinal, posiciónCaracter, carácter(i))
'                End If
                posiciónCaracter = InStr(posiciónCaracter + 1, cadenaFinal, carácter(i))
'                swEntróAlFor = True
'                swYaEmpezó = True
            Loop
        Next
        
'        If cadenaFinal = "" Then cadenaFinal = cadena
        controlarCadena = cadenaFinal
    End If
End Function

Public Function corregirCadena(cadena As String, posiciónCaracter As Long, carácter As String) As String
    Dim cadenaIzq As String, cadenaDer As String, cadenaTotal As String
    Dim carácterCorregido As String
    
    cadenaIzq = Left(cadena, posiciónCaracter - 1)
    cadenaDer = Mid(cadena, posiciónCaracter + Trim(Len(carácter)), Len(cadena) - Len(cadenaIzq))
    cadenaTotal = cadenaTotal & cadenaIzq
    Select Case carácter
        Case "("
            carácterCorregido = " abre paréntesis, "
        Case ")"
            carácterCorregido = " cierra paréntesis, "
        Case "_"
            carácterCorregido = " sobre " 'para fracciones
        Case "$"
            carácterCorregido = " pesos, "
        Case " 1 "
            carácterCorregido = " uno, "
        Case "-"
            carácterCorregido = " menos "
        Case "*"
            carácterCorregido = " multiplicado por "
        Case "/"
            carácterCorregido = " dividido "
        Case "{"
            carácterCorregido = " abre llave, "
        Case "}"
            carácterCorregido = " cierra llave, "
        Case "m "
            carácterCorregido = " eme. "
        Case "s "
            carácterCorregido = " ese. "
        Case "b "
            carácterCorregido = " be larga. "
        Case "v "
            carácterCorregido = " ve corta. "
        Case "y "
            carácterCorregido = " ih griega. "
        Case "b. "
            carácterCorregido = " be larga. "
        Case "v. "
            carácterCorregido = " ve corta. "
        Case "y. "
            carácterCorregido = " ih griega. "
        Case "l "
            carácterCorregido = " ele. "
        Case "h "
            carácterCorregido = " ache. "
        Case "p "
            carácterCorregido = " pe. "
        Case "n "
            carácterCorregido = " ene. "
        Case "m. "
            carácterCorregido = " eme. "
        Case "s. "
            carácterCorregido = " ese. "
        Case "l. "
            carácterCorregido = " ele. "
        Case "h. "
            carácterCorregido = " ache. "
        Case "p. "
            carácterCorregido = " pe. "
        Case "n. "
            carácterCorregido = " ene. "
        Case "g. "
            carácterCorregido = " je. "
        Case "u. "
            carácterCorregido = " uh. "
        Case "d. "
            carácterCorregido = " de. "
        Case "á. "
            carácterCorregido = " a con acento. "
        Case "é. "
            carácterCorregido = " e con acento. "
        Case "í. "
            carácterCorregido = " i con acento. "
        Case "ó. "
            carácterCorregido = " o con acento. "
        Case "ú. "
            carácterCorregido = " u con acento. "
        Case "ü. "
            carácterCorregido = " u con diéresis. "
        Case "à. "
            carácterCorregido = " a con acento grave. "
        Case "è. "
            carácterCorregido = " e con acento grave. "
        Case "ì. "
            carácterCorregido = " i con acento grave. "
        Case "ò. "
            carácterCorregido = " o con acento grave. "
        Case "ù. "
            carácterCorregido = " u con acento grave. "
        Case "â. "
            carácterCorregido = " a con acento circunflejo. "
        Case "ê. "
            carácterCorregido = " e con acento circunflejo. "
        Case "î. "
            carácterCorregido = " i con acento circunflejo. "
        Case "ô. "
            carácterCorregido = " o con acento circunflejo. "
        Case "û. "
            carácterCorregido = " u con acento circunflejo. "
    End Select
    cadenaTotal = cadenaTotal & carácterCorregido & cadenaDer
    corregirCadena = cadenaTotal
End Function


Public Function medioDelRenglón(cuadroRTF As RichTextBox) As Boolean
    Dim cont As Long, cadena As String ', swEmpezando As Boolean
    Dim líneaInicio As Long, líneaActual As Long ', TotalDeLíneasEnRTF As Long
    
    líneaInicio = cuadroRTF.GetLineFromChar(cuadroRTF.SelStart)
    líneaActual = líneaInicio
    cont = cuadroRTF.SelStart
    If cont <> Len(cuadroRTF.Text) Then
        While líneaActual = líneaInicio And cont <= Len(cuadroRTF.Text)
            If Mid(cuadroRTF.Text, cont + 1, 1) <> Chr(10) And Mid(cuadroRTF.Text, cont + 1, 1) <> Chr(13) Then
                cadena = cadena + Mid(cuadroRTF.Text, cont + 1, 1)
            End If
            cont = cont + 1
            líneaActual = cuadroRTF.GetLineFromChar(cont)
        Wend
        If Trim(cadena) = "" Then
            medioDelRenglón = False
        Else
            medioDelRenglón = True
        End If
    Else
        medioDelRenglón = False
    End If
    
End Function

'Public Function oraciónSiguiente(cuadroRTF As RichTextBox) As String
'    Dim cont As Long, cadena As String ', swEmpezando As Boolean
'    Dim líneaInicio As Long, líneaActual As Long ', TotalDeLíneasEnRTF As Long
'
'    líneaInicio = cuadroRTF.GetLineFromChar(cuadroRTF.SelStart)
'
'    cont = cuadroRTF.SelStart
'
'    If cont <> 0 Then 'si no se está ya al comienzo de la línea se busca el punto de inicio
'        líneaActual = líneaInicio
'        Do While líneaActual = líneaInicio
'            cont = cont - 1
'            líneaActual = cuadroRTF.GetLineFromChar(cont)
'            If cont = 0 Then Exit Do
'        Loop
'    End If
'
'    cont = cont + 1
'    'If líneaInicio = 0 Then swEmpezando = True
'    líneaActual = líneaInicio
'    While líneaActual = líneaInicio And cont <= Len(cuadroRTF.Text)
'        If Mid(cuadroRTF.Text, cont, 1) <> Chr(10) And Mid(cuadroRTF.Text, cont, 1) <> Chr(13) Then
'            cadena = cadena + Mid(cuadroRTF.Text, cont, 1)
'        End If
'        cont = cont + 1
'        líneaActual = cuadroRTF.GetLineFromChar(cont)
'    Wend
'    oraciónSiguiente = cadena
'End Function


Public Function corregirPalabra(palabra As String) As Boolean 'true palabra encontrada, false no encontrada
    'Dim palabra As String
    Dim num As Integer
    Dim palabraArchivo As String
    Dim archivo As Integer
        
    If palabra = "" Then 'si el parámetro es vacío, se considera q está correcto
        corregirPalabra = True
    Else
        If swAspellInstalado = False Then
            archivo = FreeFile
            
            If palabra = "" Or palabra = "," Or palabra = "." Or palabra = ":" Or palabra = "?" Then
                corregirPalabra = True
                Exit Function
            End If
            
            Open App.path + "\datos\palabras.txt" For Input As #archivo 'se abre la lista de palabras
            Do Until EOF(archivo) 'chequeamos si la palabra está en la lista
                Line Input #archivo, palabraArchivo
                    If palabraArchivo = palabra Then 'si encuentra una palabra igual
                        Close #archivo
                        corregirPalabra = True
                        Exit Function
                    End If
                    
        '            If Asc(LCase(Left(palabraArchivo, 1))) > Asc(LCase(Left(palabra, 1))) Then 'si ya pasó la primera letra de la palabra
        '                Close #archivo
        '                corregirPalabra = False
        '                Exit Function
        '            End If
            Loop
            corregirPalabra = False
            Close #archivo
        Else
            Call objPipe.Write_(palabra & vbCrLf)
            Call Sleep(200)
            If InStr(1, objPipe.Read, "*") Then 'si está el asterisco en lo que devuelve aspell es que la palabra es correcta
                corregirPalabra = True
            Else
                corregirPalabra = False
            End If
        End If
    End If
End Function

Public Function obtenerPalabra(cuadroRTF As RichTextBox) As String
    Dim cont As Long, cadena As String
    If cuadroRTF.Text <> "" Then
        cont = cuadroRTF.SelStart
        'If cont <> Len(cuadroRTF.Text) Then
        cont = cont - 1
        cadena = Mid(cuadroRTF.Text, cont, 1) 'si se está en el comienzo del cuadro de texto
        While Left(cadena, 1) <> " " And Left(cadena, 1) <> "," And _
        Left(cadena, 1) <> "." And Left(cadena, 1) <> ";" And _
        Left(cadena, 1) <> ":" And Left(cadena, 1) <> "?" And _
        Right(cadena, 1) <> Chr(13) And Left(cadena, 1) <> "!" And cont <> 1
            cont = cont - 1
            cadena = Mid(cuadroRTF.Text, cont, 1) + cadena
        Wend
    Else
        cadena = ""
    End If
    
    obtenerPalabra = Trim(cadena)
End Function


Public Sub llenarComboMaterias(Combo As ComboBox)
    Dim archivolibre As Byte, cadena As String
    On Error GoTo manejoError
    archivolibre = FreeFile 'se abren las materias
    Open App.path + "\datos\materias.txt" For Input As archivolibre
    While Not EOF(archivolibre)
        Line Input #archivolibre, cadena
        Combo.AddItem Trim(cadena) 'se añaden las materias al combo
    Wend
    Close #archivolibre
    Exit Sub
manejoError:
    If Err.Number = 52 Then
        Open App.path + "\datos\materias.txt" For Output As #archivolibre 'se abre el trabajo ya guardado
        Close #archivolibre
        Resume
    End If
    
    If Err.Number = 429 Then
'        MsgBox "soy el controlador de la función llenarComboMaterias", , "Para mi creador"
        frmMsgBox.cadenaAMostrar = "Soy el controlador de la función llenarComboMaterias"
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
    End If
End Sub


Public Sub regularVelocidadVoz()
    Dim aux As Integer, aux2 As Integer
    On Error GoTo manejoError
    If swSapi5 = True Then 'si se trabaja con sapi5
        If Not Voz Is Nothing Then Voz.Rate = velocidadVoz
    Else 'si se está trabajando con sapi4
        If Not vozSapi4 Is Nothing Then
            aux = Int(((vozSapi4.MaxSpeed - vozSapi4.MinSpeed) / 20) * velocidadVoz) 'se divide por las unidades del slider del form control
            aux2 = ((vozSapi4.MaxSpeed - vozSapi4.MinSpeed) / 2) + vozSapi4.MinSpeed + aux
            If aux2 <= vozSapi4.MinSpeed Then aux2 = vozSapi4.MinSpeed + 1
            If aux2 >= vozSapi4.MaxSpeed Then aux2 = vozSapi4.MaxSpeed - 1
            vozSapi4.Speed = aux2
        End If
    End If
    Exit Sub
manejoError:
    frmMsgBox.cadenaAMostrar = "Soy el controlador del la función regularVelocidadVoz. Error número: " + Str(Err.Number) + ", descripción: " + Err.Description
    frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
    frmMsgBox.Show 1
    Exit Sub
End Sub


'Public Function obtenerVersiónWindows() As String
'    Dim retvalue As Integer
'
'    osInfo.dwosversioninfosize = 148
'    osInfo.szcsdversion = Space$(128)
'    retvalue = GetVersionEx(osInfo)
'    With osInfo
'        Select Case .dwplatformid
'            Case 1
'                Select Case .dwminorversion
'            Case 0
'                obtenerVersiónWindows = "windows 95"
'            Case 10
'                obtenerVersiónWindows = "windows 98"
'            Case 90
'                obtenerVersiónWindows = "windows millennium"
'                End Select
'            Case 2
'                Select Case .dwmajorversion
'                    Case 3
'                        obtenerVersiónWindows = "windows nt 3.51"
'                    Case 4
'                        obtenerVersiónWindows = "windows nt 4.0"
'                    Case 5
'                        If .dwminorversion = 0 Then
'                            obtenerVersiónWindows = "windows 2000"
'                        Else
'                            obtenerVersiónWindows = "windows xp"
'                        End If
'                End Select
'            Case Else
'                obtenerVersiónWindows = "falló"
'        End Select
'    End With
''    leerDatosSO (App.Path + "\datos\datosSO.lle")
'End Function

'Public Sub leerDatosSO(dónde As String)
'    Dim miRegistro As osVersionInfo, sistemaRepetido As Boolean
'
'    sistemaRepetido = False
'    Open dónde For Random As #1 Len = Len(miRegistro)
'    Do While Not EOF(1)   ' Repite hasta el final del archivo.
'       Get #1, , miRegistro   ' Lee el registro siguiente.
'       If miRegistro.dwbuildnumber = osInfo.dwbuildnumber And _
'       miRegistro.dwmajorversion = osInfo.dwmajorversion And _
'       miRegistro.dwminorversion = osInfo.dwminorversion And _
'       miRegistro.dwosversioninfosize = osInfo.dwosversioninfosize And _
'       miRegistro.dwplatformid = osInfo.dwplatformid And _
'       miRegistro.szcsdversion = osInfo.szcsdversion Then
'            sistemaRepetido = True
'            Exit Do
'        End If
'    Loop
'    Close #1   ' Cierra el archivo.
'
'    If sistemaRepetido = False Then
'        MsgBox "acá iría el instalador porque es la primera vez que se abre en este sistema" 'shell instalador de acuerdo al sistema
'        Call guardarDatosSO(dónde, osInfo) 'se añade el sistema a la lista de los guardados como que ya se instaló
'    End If
'End Sub

'Public Sub guardarDatosSO(dónde As String, quéRegistro As osVersionInfo)
'    Dim archivolibre As Integer
'
'    archivolibre = FreeFile 'se abre el archivo para guardar los datos del sistema operativo
'    Open dónde For Random As #archivolibre Len = Len(quéRegistro)
'    Put #archivolibre, 1, quéRegistro
'    Close #archivolibre
'End Sub

Public Sub centrarFormulario(quéForm As Form)
    Dim centroXform As Single, centroYform As Single
    centroXform = (Screen.Width - quéForm.ScaleWidth) / 2
    centroYform = (Screen.Height - quéForm.ScaleHeight) / 2
    Call quéForm.Move(centroXform, centroYform)
End Sub


Public Sub controlarCaracteresEspeciales(teclaPulsada As Integer, caja As TextBox)
    If teclaPulsada = 34 Or teclaPulsada = Asc("|") Or teclaPulsada = Asc("\") Or teclaPulsada = Asc("/") _
    Or teclaPulsada = Asc("?") Or teclaPulsada = Asc("*") Or teclaPulsada = Asc(">") Or teclaPulsada = Asc("<") _
    Or teclaPulsada = Asc(":") Or teclaPulsada = Asc(",") Or teclaPulsada = Asc(";") Or teclaPulsada = Asc(".") _
    Or teclaPulsada = Asc("-") Or teclaPulsada = Asc("_") Then
        caja.Text = Left(caja.Text, Len(caja.Text) - 1)
        frmMsgBox.cadenaAMostrar = "No se pueden escribir los siguientes signos en el nombre: . , ; - \ : / < > ? * | " + Chr(34) + "."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        SendKeys "^{end}"
    End If
End Sub

'Public Function decodificarArchivo(archivo As String) As String
'    archivo = Right(archivo, Len(archivo) - cantPrefijo)
'    archivo = Left(archivo, Len(archivo) - 4)
'    decodificarArchivo = archivo
'    'archivo = Format(archivo, "Long Date")
'End Function

'Public Function conocerPalabraEnOración(cuadroRTF As RichTextBox) As String
'    Dim cont As Long, cadena As String, renglónActual As Long
'    cont = cuadroRTF.SelStart
'    If Len(cuadroRTF.Text) <> cont Then 'si no se está al final de la hoja
'        cont = cont + 1
'        cadena = Mid(cuadroRTF.Text, cont, 1) 'si se está en el comienzo del cuadro de texto
'        While Right(cadena, 1) <> " " And Right(cadena, 1) <> "," And _
'        Right(cadena, 1) <> "." And Right(cadena, 1) <> ";" And _
'        Right(cadena, 1) <> ":" And Right(cadena, 1) <> "?" And _
'        Right(cadena, 1) <> "!" And Right(cadena, 1) <> Chr(13) And _
'        Len(cuadroRTF.Text) <> cont
'            cont = cont + 1
'            cadena = cadena + Mid(cuadroRTF.Text, cont, 1)
'        Wend
'
'        If cadena + Mid(cuadroRTF.Text, cont, 1) <> Chr(13) Then
'            cadena = cadena + Mid(cuadroRTF.Text, cont, 1)
'        End If
'
'        If Len(cadena) = 1 And cadena = Chr(13) Then cadena = "aparte"
'        If cadena = "." Then cadena = "punto"
'        If cadena = "," Then cadena = "coma"
'        If cadena = ";" Then cadena = "punto y coma"
'        If cadena = ":" Then cadena = "dos puntos"
'        If cadena = Chr(34) Then cadena = "comillas"
''        if cadena="(" then cadena="abre paréntesis"
'
''        If cuadroRTF.SelStart <> 0 Then 'si no se está al principio de la hoja
'            conocerPalabraEnOración = cadena '"estás al final de tu hoja"
''        Else
''            If cadena = "." Or cadena = "," Or cadena = ";" Then
''                conocerPalabraEnOración = "estás al principio de la hoja, delante del signo. " + cadena
''            Else
''                conocerPalabraEnOración = "estás al principio de la hoja, delante de la palabra. " + cadena
''            End If
''        End If
''
''        If swLeerRenglones = True Then
''            If renglónActual <> renglónAnterior Then conocerPalabraEnOración = conocerPalabraEnOración & ". renglón " & Str(renglónActual + 1)
''        End If
'    Else
''        If cuadroRTF.Text <> "" Then
'            conocerPalabraEnOración = "final de la hoja"
''        Else
''            conocerPalabraEnOración = "no hay nada escrito, la hoja está vacía"
''        End If
'    End If
''    renglónAnterior = renglónActual
'End Function

Public Function SalirDelPrograma() As Boolean
    frmMsgBox.swMostrarCancelar = False
    frmMsgBox.cadenaAMostrar = "¿Realmente querés salir del programa?"
    frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
    frmMsgBox.Show 1
    If frmMsgBox.swResultadoMostrado = True Then
        SalirDelPrograma = True
    Else
        SalirDelPrograma = False
    End If
End Function

Public Sub chauPrograma()
    On Error Resume Next
    sonido = sndPlaySound(App.path + "\sonidos\fin.wav", SND_SYNC)
    Unload frmOculto
    Set Voz = Nothing
    Set vozSapi4 = Nothing
    If swCuadernoAbierto = True Then Unload frmCuaderno 'si está abierto el cuaderno se lo cierra
    If swLibroAbierto = True Then Unload frmLectorLibro 'si está abierto el lector de libros, se lo cierra
    If swActividadAbierta = True Then Unload frmLectorActividad 'si está abierto el lector de actividad, se lo cierra
    If frmReproductorMúsica.swEstoyAbierto = True Then Unload frmReproductorMúsica
    If objPipe.Running = True Then
        Call objPipe.Terminate
    End If
    KillProcess ("cmd.exe")
    KillProcess ("aspell.exe")
    Set objPipe = Nothing
    End
End Sub

'
'Public Sub instalar()
'    Dim versiónOS As String
'    versiónOS = obtenerVersiónWindows
'
'    Select Case versiónOS
'        Case "windows 95"
'            'call ejecutar (App.Path + "\Ejecutables\instalador98.exe")
'            MsgBox "Acá va el instalador del win 95 del programa porque es la primera vez que se abre"
'        Case "windows 98"
'            'call ejecutar ( App.Path + "\Ejecutables\instalador98.exe")
'            MsgBox "Acá va el instalador del win 98 del programa porque es la primera vez que se abre"
'        Case "windows millennium"
'            MsgBox "Acá va el instalador del millenium del programa porque es la primera vez que se abre"
'        Case "windows nt 3.51"
'            MsgBox "Acá va el instalador del win 3.51 del programa porque es la primera vez que se abre"
'        Case "windows nt 4.0"
'            MsgBox "Acá va el instalador del win nt del programa porque es la primera vez que se abre"
'        Case "windows 2000"
'            MsgBox "Acá va el instalador del win 2000 del programa porque es la primera vez que se abre"
'        Case "windows xp"
'            MsgBox "Acá va el instalador del xp del programa porque es la primera vez que se abre"
'        Case "falló"
'    End Select
'
'
''    On Error GoTo manejoErrorInstalar
'''    Dim miUbicación As String
'    Dim lectorRegistro, x
'    Set lectorRegistro = CreateObject("WScript.Shell")
''
''    Call RegFlash 'registrar flash
''    Call RegBotónTransp 'registrar los botones transparentes
''
''    'instalar sapi5 si no está instalada
''    x = lectorRegistro.regRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\MSMary\409")
''    If x <> "Microsoft Mary" Then
''        If versiónOS = "windows xp" Then
''            Shell App.Path + "\ejecutables\Sapi5 (para XP).msi", vbNormalFocus
''        Else
''            Shell App.Path + "\ejecutables\Sapi5 (para Windows 98 Me 2000).msi", vbNormalFocus
''        End If
''    End If
'
''    'arreglar sapi4
''    'instalar sapi4 si no está instalada
''    x = lectorRegistro.regRead("HKEY_LOCAL_MACHINE\SOFTWARE\Voice\TextToSpeech\Engine")
''    If x <> "Microsoft Mary" Then
''        Shell App.Path + "\ejecutables\Sapi4.exe", vbNormalFocus
''    End If
'
''    miUbicación = App.Path + "\Mochila_Virtual.exe"
'    'lectorRegistro.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\", miUbicación
'    lectorRegistro.RegWrite "HKEY_LOCAL_MACHINE\Software\ReyNegro-ReyBlanco\MochilaVirtual\", "1"
'    lectorRegistro.RegWrite "HKEY_LOCAL_MACHINE\Software\ReyNegro-ReyBlanco\MochilaVirtual\datos\", "0"
'    Set lectorRegistro = Nothing
''    Exit Sub
''manejoErrorInstalar:
''    MsgBox "soy el error de la función instalar. Mi número es " + Str(Err.Number) + ", y mi descripción es " + Err.Description
''    Resume Next
'End Sub

'Declare Function DLLSelfRegister Lib "VB6STKIT.DLL" (ByVal lpDllName As String) As Integer
'
'Public Function SelfRegisterDLL(NombreDll As String) As Boolean
'Dim liRet As Integer
'
'On Error Resume Next
'
'    liRet = DLLSelfRegister(NombreDll)
'
'    If liRet = 0 Then
'        SelfRegisterDLL = True
'    Else
'        SelfRegisterDLL = False
'    End If
'
'End Function

'Public Sub instalarSapis()
'    Dim versiónOS As String
'    versiónOS = obtenerVersiónWindows
'
'    Select Case versiónOS
'        Case "windows xp"
'            'call ejecutar( App.Path + "\ejecutables\Sapi5 (para XP).msi")
'        Case Else
'            'call ejecutar( App.Path + "\ejecutables\Sapi5 (para Windows 98 Me 2000).msi")
'    End Select
'
'    'Shell App.Path + "\ejecutables\Sapi4.exe", vbNormalFocus
'
'
'
'
''    Dim lectorRegistro, x
'''    miUbicación = App.Path + "\Mochila_Virtual.exe"
''    Set lectorRegistro = CreateObject("WScript.Shell")
''    x = lectorRegistro.regRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\MSMary")
''    If x <> "Microsoft Mary" Then
''        Shell App.Path + "\ejecutables\Sapi5.exe", vbNormalFocus
''    End If
''    Set lectorRegistro = Nothing
''    Exit Sub
''manejoErrorInstalar:
''    MsgBox "soy el error de la función instalar. Mi número es " + Err.Number + ", y mi descripción es " + Err.Description
''
'End Sub


Public Function PathCorto(archivo As String) As String
    Dim temp As String * 250, aux As Long  'Buffer
    
    PathCorto = String(255, 0)
    'Obtenemos el Path corto
    aux = GetShortPathName(archivo, temp, 164)
    'Sacamos los nulos al path
    PathCorto = Replace(temp, Chr(0), "")
End Function

Public Sub GuardarRecordatorio(quéRecordatorio As Recordatorio)
    Dim archivolibre As Byte, contador As Integer, auxRecordatorio As Recordatorio, j As Byte

    On Error GoTo manejoError
    archivolibre = FreeFile 'se abre el archivo para guardar los datos de las partidas
    Open App.path + "\recordatorios\" + Trim(Right(Format(quéRecordatorio.fecha, "dd/mm/yyyy"), 4)) + "\" + Trim(Str(Int(Mid(Format(quéRecordatorio.fecha, "dd/mm/yyyy"), 4, 2)))) + "\recordatorios.gui" For Random As archivolibre Len = Len(quéRecordatorio)
    contador = 0
    While Not EOF(archivolibre)
        contador = contador + 1
        Get archivolibre, contador, auxRecordatorio
    Wend
    quéRecordatorio.índiceEnArchivo = contador
    Put archivolibre, contador, quéRecordatorio
    Close archivolibre
    Exit Sub
manejoError:
    If Err.Number = 76 Then
        MkDir (App.path + "\recordatorios\" + Trim(Right(Format(quéRecordatorio.fecha, "dd/mm/yyyy"), 4)))
        For j = 1 To 12
            MkDir (App.path + "\recordatorios\" + Trim(Right(Format(quéRecordatorio.fecha, "dd/mm/yyyy"), 4)) + "\" + Trim(Str(j)))
        Next
        Resume Next
    End If
End Sub

Public Sub GuardarRecordatorioEnPosición(quéRecordatorio As Recordatorio, posición As Long)
    Dim archivolibre As Byte, auxRecordatorio As Recordatorio, j As Byte

    On Error GoTo manejoError
    archivolibre = FreeFile 'se abre el archivo para guardar los datos de las partidas
    Open App.path + "\recordatorios\" + Trim(Right(Format(quéRecordatorio.fecha, "dd/mm/yyyy"), 4)) + "\" + Trim(Str(Int(Mid(Format(quéRecordatorio.fecha, "dd/mm/yyyy"), 4, 2)))) + "\recordatorios.gui" For Random As archivolibre Len = Len(quéRecordatorio)
    quéRecordatorio.índiceEnArchivo = posición
    Put archivolibre, posición, quéRecordatorio
    Close archivolibre
    Exit Sub
manejoError:
    If Err.Number = 76 Then
        MkDir (App.path + "\recordatorios\" + Trim(Right(Format(quéRecordatorio.fecha, "dd/mm/yyyy"), 4)))
        For j = 1 To 12
            MkDir (App.path + "\recordatorios\" + Trim(Right(Format(quéRecordatorio.fecha, "dd/mm/yyyy"), 4)) + "\" + Trim(Str(j)))
        Next
        Resume Next
    End If
End Sub


Public Function sonidoForm(quéForm As Byte) As String
    Dim archivo As String
    
    'corregir que las variables se cargen desde la configuración
'    rutaMúsicaFormPrincipal = "principal.mid"
'    rutaMúsicaFormCuaderno = "cuaderno.mid"
'    rutaMúsicaFormActividad = "actividades.mid"
'    rutaMúsicaFormTareas = "tareas.mid"
'    rutaMúsicaFormLibros = "libros.mid"
'    rutaMúsicaFormAccesorios = "accesorios.mid"
    
    archivo = App.path + "\sonidos\formularios\"
    Select Case quéForm
        Case formularios.principal
            If Trim(usuario.rutaMúsicaFormPrincipal) <> "" And Trim(Left(usuario.rutaMúsicaFormPrincipal, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaMúsicaFormPrincipal)
            Else
                archivo = archivo + "principal.mid"
            End If
        Case formularios.cuaderno
            If Trim(usuario.rutaMúsicaFormCuaderno) <> "" And Trim(Left(usuario.rutaMúsicaFormCuaderno, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaMúsicaFormCuaderno)
            Else
                archivo = archivo + "cuaderno.mid"
            End If
        Case formularios.actividades
            If Trim(usuario.rutaMúsicaFormActividad) <> "" And Trim(Left(usuario.rutaMúsicaFormActividad, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaMúsicaFormActividad)
            Else
                archivo = archivo + "actividades.mid"
            End If
        Case formularios.tareasAnt
            If Trim(usuario.rutaMúsicaFormTareas) <> "" And Trim(Left(usuario.rutaMúsicaFormTareas, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaMúsicaFormTareas)
            Else
                archivo = archivo + "tareas.mid"
            End If
        Case formularios.libros
            If Trim(usuario.rutaMúsicaFormLibros) <> "" And Trim(Left(usuario.rutaMúsicaFormLibros, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaMúsicaFormLibros)
            Else
                archivo = archivo + "libros.mid"
            End If
        Case formularios.accesorios
            If Trim(usuario.rutaMúsicaFormAccesorios) <> "" And Trim(Left(usuario.rutaMúsicaFormAccesorios, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaMúsicaFormAccesorios)
            Else
                archivo = archivo + "accesorios.mid"
            End If
    End Select
    sonidoForm = archivo
End Function

Public Sub reproducirForm(quéForm As Byte)
    If swMúsicaDeFondo = True Then
        frmOculto.swContinuarReproducción = False
        frmOculto.media.Stop
        frmOculto.media.FileName = sonidoForm(quéForm)
        frmOculto.swContinuarReproducción = True
'        frmOculto.media.Play
    Else
        If frmOculto.media.PlayState = mpPlaying Then
            frmOculto.swContinuarReproducción = False
            frmOculto.media.Stop
        End If
   End If
End Sub

Public Sub cargarRecordatorios()
    Dim archivolibre As Byte, miRec As Recordatorio, mes As Byte, año As Integer ', día As Byte
    Dim contador As Integer, i As Integer
    On Error GoTo manejoError
    mes = Month(Date)
    año = Year(Date)
    archivolibre = FreeFile
    contador = 1 'se deja en blanco el primer recordatorioActivo. Si hay más de uno, es que hay recordatorios activos. Se evalúa en frmOculto
    Open App.path + "\recordatorios\" + Trim(Str(año)) + "\" + Trim(Str(mes)) + "\" + "recordatorios.gui" For Random As #archivolibre Len = Len(miRec)
    Do While Not EOF(archivolibre)   ' Repite hasta el final del archivo.
       Get #archivolibre, , miRec   ' Lee el registro siguiente.
       If Asc(Left(miRec.texto, 1)) <> 0 Then
            If miRec.yaAnunciado = False Then 'si no fue anunciado
                If Format(miRec.fecha, "dd/mm/yyyy") = Format(Date, "dd/mm/yyyy") Then 'si el recordatorio es de hoy
                    If Left(Format(miRec.hora, "HH:mm"), 2) <= Left(Format(Time, "HH:mm"), 2) Then 'si la hora es la actual o menor
                        'se controla que los minutos sean iguales o menores a los actuales
                        If Right(Format(miRec.hora, "HH:mm"), 2) <= Right(Format(Time, "HH:mm"), 2) Then
                            ReDim Preserve recordatoriosActivos(0 To contador)
                            recordatoriosActivos(contador) = miRec
                            contador = contador + 1
                        End If
                    End If
                Else
                    If Right(Format(miRec.fecha, "dd/mm/yyyy"), 4) < Right(Format(Date, "dd/mm/yyyy"), 4) Then   'si el año es menor al actual
                        ReDim Preserve recordatoriosActivos(0 To contador)
                        recordatoriosActivos(contador) = miRec
                        contador = contador + 1
                    Else
                        If Right(Format(miRec.fecha, "dd/mm/yyyy"), 4) = Right(Format(Date, "dd/mm/yyyy"), 4) Then  'si el año es igual al actual
                            If Mid(Format(miRec.fecha, "dd/mm/yyyy"), 4, 2) < Mid(Format(Date, "dd/mm/yyyy"), 4, 2) Then 'si el mes es menor al actual
                                ReDim Preserve recordatoriosActivos(0 To contador)
                                recordatoriosActivos(contador) = miRec
                                contador = contador + 1
                            Else
                                If Mid(Format(miRec.fecha, "dd/mm/yyyy"), 4, 2) = Mid(Format(Date, "dd/mm/yyyy"), 4, 2) Then 'si el mes es igual al actual
                                    If Left(Format(miRec.fecha, "dd/mm/yyyy"), 2) < Left(Format(Date, "dd/mm/yyyy"), 2) Then 'si el día es anterior al actual
                                        ReDim Preserve recordatoriosActivos(0 To contador)
                                        recordatoriosActivos(contador) = miRec
                                        contador = contador + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop
    Close #archivolibre
    Exit Sub
manejoError:
    Exit Sub
End Sub

'Public Sub contarFormularios(acción As Boolean)
'    If acción = True Then formulariosAbiertos = formulariosAbiertos + 1
'    If acción = False Then formulariosAbiertos = formulariosAbiertos - 1
'    If formulariosAbiertos = 1 Then End
'End Sub

Public Function mensajeSalir(quéMensaje As String) As Boolean
    frmMsgBox.swMostrarCancelar = False
    frmMsgBox.cadenaAMostrar = quéMensaje
    frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
    frmMsgBox.Show 1
    If frmMsgBox.swResultadoMostrado = True Then
        mensajeSalir = True
    Else
        mensajeSalir = False
    End If
End Function

Public Function existeCarpeta(ByVal rutaCarpeta As String) As Boolean 'si existe o no una carpeta o archivo
    Dim x As Integer
    On Error GoTo Fallo
    x = GetAttr(rutaCarpeta)
    existeCarpeta = True
    Exit Function
Fallo:
    existeCarpeta = False
End Function

Public Function leerRegistro(raíz As Long, clave As String, valor As String) As String
    Dim hClave As Long, longitud As Long, dato As String, ret As Long
    
    ret = RegOpenKeyEx(raíz, clave, 0, KEY_ALL_ACCESS, hClave)
    ret = RegQueryValueEx(hClave, valor, 0, 0, 0, longitud)
    dato = String(longitud, 0)
    ret = RegQueryValueEx(hClave, valor, 0&, REG_SZ, ByVal dato, longitud)
    ret = RegCloseKey(hClave)
    leerRegistro = Left(dato, longitud - 1)
End Function

'Private Sub espaciosDelDisco(disco As String, espacioLibre As Currency, espacioTotal As Currency, espacioOcupado As Currency)
'    Dim devolución As Long, SectoresporCluster As Long, BytesPorSector As Long, CantidadDeClustersLibres As Long, NúmeroTotalClusters As Long
'    Static swDemasiadoDiscoParaMí As Boolean 'para controlar que no salte error si tiene un disco muy grande y desborde las variables que cuentan los bytes
'
'    On Error GoTo manejoErrorEspacio:
'    If swDemasiadoDiscoParaMí = False Then
'        devolución = GetDiskFreeSpace(disco, SectoresporCluster, BytesPorSector, CantidadDeClustersLibres, NúmeroTotalClusters)
'        espacioLibre = CantidadDeClustersLibres * SectoresporCluster * BytesPorSector
'        espacioLibre = (espacioLibre / 1024) / 1024
'        espacioTotal = NúmeroTotalClusters * SectoresporCluster * BytesPorSector
'        espacioTotal = (espacioTotal / 1024) / 1024
'        espacioOcupado = espacioTotal - espacioLibre
'    End If
'    Exit Sub
'manejoErrorEspacio:
'    swDemasiadoDiscoParaMí = True
'    Exit Sub
'End Sub
'
'Public Sub chequearEspacioEnDisco(disco As String)
'    Dim libre As Currency, total As Currency, ocupado As Currency, swMostrarCuadro As Boolean
'    swMostrarCuadro = False
'    Call espaciosDelDisco(disco, libre, total, ocupado)
'    If libre < 20 And libre > 10 Then
'        frmMsgBox.cadenaAMostrar = "Se está acabando el espacio libre que queda en el disco en que está instalada la mochila. Considere liberar espacio a la brevedad."
'        swMostrarCuadro = True
'    ElseIf libre < 10 And libre >= 1 Then
'        frmMsgBox.cadenaAMostrar = "Queda muy poco espacio libre en el disco en que está instalada la mochila. Libere espacio o instale la mochila en otro disco."
'        swMostrarCuadro = True
'    ElseIf libre <= 1 Then
'        frmMsgBox.cadenaAMostrar = "Queda solamente 1 MB libre en el disco en que está instalada la mochila. Libere urgentemente espacio o instale la mochila en otro disco."
'        swMostrarCuadro = True
'    End If
'    If swMostrarCuadro = True Then
'        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
'        frmMsgBox.Show 1
'    End If
'End Sub

'Public Sub alwaysOnTop(formulario As Form, estado As Boolean)
'    Dim banderas As Long, ret As Long
'    'para que no cambie el tamaño ni la posición
'    banderas = SWP_NOMOVE Or SWP_NOSIZE
'    If estado Then
'        ret = SetWindowPos(formulario, HWND_TOPMOST, 0, 0, 0, 0, banderas)
'    Else
'        ret = SetWindowPos(formulario, HWND_NOTOPMOST, 0, 0, 0, 0, banderas)
'    End If
'End Sub

Public Sub ejecutar(aplicación As String)
    Dim handleProceso As Long
    Dim activa As Long
    Dim ret As Long

    handleProceso = OpenProcess(PROCESS_QUERY_INFORMATION, 0, Shell(aplicación, 1))
    Do
        ret = GetExitCodeProcess(handleProceso, activa)
        DoEvents
    Loop While activa = STILL_ACTIVE
End Sub

Public Function quéLetraSeApretó(númeroLetra As Integer) As String
    Dim auxString As String ', cadena As String

    Select Case UCase(Chr(númeroLetra))
        Case " "
            auxString = " espacio"
        Case "1"
            auxString = " uno"
        Case "´"
            auxString = " acento agudo"
        Case "¨"
            auxString = " diéresis"
        Case "`"
            auxString = " acento grave"
        Case "^"
            auxString = " acento circunflejo"
        Case "+"
            auxString = " más"
        Case "-"
            auxString = " menos"
        Case "_"
            auxString = " sobre "
        Case "*"
            auxString = " por"
        Case "/"
            auxString = " dividido"
        Case "="
            auxString = " igual"
        Case ","
            auxString = " coma"
        Case "."
            auxString = " punto"
        Case ";"
            auxString = " punto y coma"
        Case ":"
            auxString = " dos puntos"
        Case Chr(34) '"'"
            auxString = " comillas"
        Case "¡"
            auxString = " abre exclamación"
        Case "!"
            auxString = " cierra exclamación"
        Case "¿"
            auxString = " abre pregunta"
        Case "?"
            auxString = " cierra pregunta"
        Case "$"
            auxString = " signo pesos"
        Case "&"
            auxString = " anpersand"
        Case "\"
            auxString = " barra diagonal inversa"
        Case "º"
            auxString = " ordinal masculino"
        Case "ª"
            auxString = " ordinal femenino"
        Case "%"
            auxString = " porciento"
        Case "("
            auxString = " abre paréntesis"
        Case ")"
            auxString = " cierra paréntesis"
        Case "{"
            auxString = " abre llave"
        Case "}"
            auxString = " cierra llave"
        Case "Á"
            auxString = " a con acento"
        Case "É"
            auxString = " e con acento"
        Case "Í"
            auxString = " i con acento"
        Case "Ó"
            auxString = " o con acento"
        Case "Ú"
            auxString = " u con acento"
        Case "Ü"
            auxString = " u con diéresis"
        Case "B"
            auxString = " bé larga"
        Case "C"
            auxString = " cé"
        Case "D"
            auxString = " dé"
        Case "F"
            auxString = " éfe"
        Case "G"
            auxString = " gé"
        Case "H"
            auxString = " áche"
        Case "J"
            auxString = " jóta"
        Case "K"
            auxString = " ká"
        Case "L"
            auxString = " éle"
        Case "M"
            auxString = " éme"
        Case "N"
            auxString = " éne"
        Case "Ñ"
            auxString = " éñe"
        Case "P"
            auxString = " pé"
        Case "Q"
            auxString = " cú"
        Case "R"
            auxString = " érre"
        Case "S"
            auxString = " ése"
        Case "T"
            auxString = " té"
        Case "V"
            auxString = " vé corta"
        Case "W"
            auxString = " doble bé"
        Case "X"
            auxString = " équis"
        Case "Y"
            auxString = " i griega"
        Case "Z"
            auxString = " seta"
        Case "A"
            auxString = " ah"
        Case "E"
            auxString = " eh"
        Case "I"
            auxString = " ih"
        Case "O"
            auxString = " oh"
        Case "U"
            auxString = " uh"
        Case Else 'si es cualquier otro caracter
            auxString = Chr(númeroLetra)
    End Select
    
    'cadena = auxString
    If (númeroLetra >= 65 And númeroLetra <= 90) Then auxString = auxString + " mayúscula"
    
    If númeroLetra = 9 Then auxString = "avanzando hacia adelante un salto" 'si es un tab
    quéLetraSeApretó = auxString
End Function


Public Function traducirParaBorrar(letra As String) As String
    Dim auxString As String
    Select Case UCase(letra)
        Case " "
            auxString = " el espacio"
        Case "´"
            auxString = " el acento agudo "
        Case "¨"
            auxString = " la diéresis"
        Case "`"
            auxString = " el acento grave"
        Case "^"
            auxString = " el acento circunflejo"
        Case "&"
            auxString = " el ampersand"
        Case "+"
            auxString = " el más"
        Case "-"
            auxString = " el menos"
        Case "_"
            auxString = " el sobre"
        Case "*"
            auxString = " el por"
        Case "/"
            auxString = " el dividido"
        Case "="
            auxString = " el igual"
        Case ","
            auxString = " la coma"
        Case "."
            auxString = " el punto"
        Case ";"
            auxString = " el punto y coma"
        Case ":"
            auxString = " los dos puntos"
        Case Chr(34) '"'"
            auxString = " las comillas"
        Case "¡"
            auxString = " el abre exclamación"
        Case "!"
            auxString = " el cierra exclamación"
        Case "¿"
            auxString = " el abre pregunta"
        Case "?"
            auxString = " el cierra pregunta"
        Case "$"
            auxString = " el signo pesos"
        Case "%"
            auxString = " el porciento"
        Case "("
            auxString = " el abre paréntesis"
        Case ")"
            auxString = " el cierra paréntesis"
        Case "{"
            auxString = " el abre llave"
        Case "}"
            auxString = " el cierra llave"
        Case "Á"
            auxString = " la a con acento"
        Case "É"
            auxString = " la e con acento"
        Case "Í"
            auxString = " la i con acento"
        Case "Ó"
            auxString = " la o con acento"
        Case "Ú"
            auxString = " la u con acento"
        Case "B"
            auxString = " la bé larga"
        Case "C"
            auxString = " la cé"
        Case "D"
            auxString = " la dé"
        Case "F"
            auxString = " la éfe"
        Case "G"
            auxString = " la gé"
        Case "H"
            auxString = " la áche"
        Case "J"
            auxString = " la jóta"
        Case "K"
            auxString = " la ká"
        Case "L"
            auxString = " la éle"
        Case "M"
            auxString = " la éme"
        Case "N"
            auxString = " la éne"
        Case "Ñ"
            auxString = " la éñe"
        Case "P"
            auxString = " la pé"
        Case "Q"
            auxString = " la cú"
        Case "R"
            auxString = " la érre"
        Case "S"
            auxString = " la ése"
        Case "T"
            auxString = " la té"
        Case "V"
            auxString = " la vé corta"
        Case "W"
            auxString = " la doble bé"
        Case "X"
            auxString = " la équis"
        Case "Y"
            auxString = " la i griega"
        Case "Z"
            auxString = " la seta"
        Case "A"
            auxString = " la ah"
        Case "E"
            auxString = " la eh"
        Case "I"
            auxString = " la ih"
        Case "O"
            auxString = " la oh"
        Case "U"
            auxString = " la uh"
        Case "1"
            auxString = " el uno"
        Case "2"
            auxString = " el dos"
        Case "3"
            auxString = " el tres"
        Case "4"
            auxString = " el cuatro"
        Case "5"
            auxString = " el cinco"
        Case "6"
            auxString = " el seis"
        Case "7"
            auxString = " el siete"
        Case "8"
            auxString = " el ocho"
        Case "9"
            auxString = " el nueve"
        Case "0"
            auxString = " el cero"
        Case Else 'si es cualquier otro caracter
            auxString = " la " + letra
    End Select
    
    If (Asc(letra) >= 65 And Asc(letra) <= 90) Then auxString = auxString + " mayúscula"
    traducirParaBorrar = auxString
End Function


Sub evaluarSelección(cuadroRTF As RichTextBox, control As Boolean, Shift As Boolean, teclaQueSelecciona As Byte)    'As String
    Static LenSelecciónAnterior As Currency
    Dim estadoSelección As Byte, cadena As String
    
    If Len(cuadroRTF.SelText) < LenSelecciónAnterior Then estadoSelección = selección.disminuyó
    If Len(cuadroRTF.SelText) > LenSelecciónAnterior Then estadoSelección = selección.creció
    If Len(cuadroRTF.SelText) = LenSelecciónAnterior Then estadoSelección = selección.igual
    
    If estadoSelección <> selección.igual Then
        If cuadroRTF.SelText = "" And estadoSelección = selección.disminuyó Then cadena = "quitando la selección"
        If cuadroRTF.SelText = "" And LenSelecciónAnterior <> 0 And teclaQueSelecciona = tecla.borrar Then cadena = "borrando la selección"
        '++++++++++++++++++++++++++++++++++++++++++++
        'se selecciona con shift y control apretadas
        If control And Shift Then 'si se va seleccionando texto con control y shift
            If (teclaQueSelecciona = tecla.flechaDerecha Or teclaQueSelecciona = tecla.flechaIzquierda) Then   'seleccionado por palabras
                If cuadroRTF.Text <> "" Then
                    cadena = "texto seleccionado: " + cuadroRTF.SelText
                Else
                    cadena = "no se puede seleccionar porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.inicio Then
                If cuadroRTF.Text <> "" Then
                    If cuadroRTF.SelText <> "" Then
                        If estadoSelección = selección.creció Then cadena = "seleccionado todo el texto desde donde estabas hasta el principio de la hoja"
                        If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección desde donde estabas hasta el principio de la hoja"
                    Else
                        cadena = "se ha sacado la selección del texto"
                    End If
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.fin Then
                If cuadroRTF.Text <> "" Then
                    If cuadroRTF.SelText <> "" Then
                        If estadoSelección = selección.creció Then cadena = "seleccionado todo el texto desde donde estabas hasta el final de la hoja"
                        If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección desde donde estabas hasta el final de la hoja"
                        'cadena =  "seleccionado todo el texto desde donde estabas hasta el final de la hoja"
                    Else
                        cadena = "se ha dejado de seleccionar todo el texto"
                    End If
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.flechaArriba Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelección = selección.creció Then cadena = "seleccionado desde donde estabas hasta el principio del párrafo"
                    If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección desde donde estabas hasta el principio del párrafo"
                    'cadena =  "seleccionado desde donde estabas hasta el principio del párrafo"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.flechaAbajo Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelección = selección.creció Then cadena = "seleccionado desde donde estabas hasta el final del párrafo"
                    If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección desde donde estabas hasta el final del párrafo"
                    'cadena =  "seleccionado desde donde estabas hasta el final del párrafo"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.avancePágina Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelección = selección.creció Then cadena = "seleccionando varios renglones desde donde estabas hacia arriba en la hoja"
                    If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección varios renglones desde donde estabas hacia arriba en la hoja"
                    'cadena =  "seleccionando varios renglones desde donde estabas hacia arriba en la hoja"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.retrocesoPágina Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelección = selección.creció Then cadena = "seleccionando varios renglones desde donde estabas hacia abajo en la hoja"
                    If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección varios renglones desde donde estabas hacia abajo en la hoja"
                    'cadena =  "seleccionando varios renglones desde donde estabas hacia abajo en la hoja"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
        End If
        
        '++++++++++++++++++++++++++++++++++++++++++++
        'se selecciona solamente con shift
        If Shift And Not control Then 'si se va seleccionando texto con control y shift
            If (teclaQueSelecciona = tecla.flechaDerecha Or teclaQueSelecciona = tecla.flechaIzquierda) Then   'seleccionado por letras
                If cuadroRTF.Text <> "" Then
                    cadena = "texto seleccionado: "
                    If Left(cuadroRTF.SelText, 1) = " " Then cadena = cadena + " espacio "
                    If Left(cuadroRTF.SelText, 1) = Chr(13) Then cadena = cadena + " bajada de línea "
                    
                    cadena = cadena + cuadroRTF.SelText
                Else
                    cadena = "no se puede seleccionar porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.inicio Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelección = selección.creció Then cadena = "seleccionando el texto hasta el principio del renglón"
                    If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección, queda seleccionado: " + cuadroRTF.SelText
                    'cadena =  "seleccionando el texto hasta el principio del renglón"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.fin Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelección = selección.creció Then cadena = "seleccionando el texto hasta el final del renglón"
                    If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección, queda seleccionado: " + cuadroRTF.SelText
                    'cadena =  "seleccionando el texto hasta el final del renglón"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.flechaArriba Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelección = selección.creció Then cadena = "seleccionado desde donde estabas hasta la misma posición del renglón superior"
                    If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección desde donde estabas hasta la misma posición del renglón superior"
                    'cadena =  "seleccionado desde donde estabas hasta la misma posición del renglón superior"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.flechaAbajo Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelección = selección.creció Then cadena = "seleccionado desde donde estabas hasta la misma posición del renglón inferior"
                    If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección desde donde estabas hasta la misma posición del renglón inferior"
                    'cadena =  "seleccionado desde donde estabas hasta la misma posición del renglón inferior"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.avancePágina Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelección = selección.creció Then cadena = "seleccionando varios renglones desde donde estabas hacia arriba en la hoja"
                    If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección varios renglones desde donde estabas hacia arriba en la hoja"
                    'cadena =  "seleccionando varios renglones desde donde estabas hacia arriba en la hoja"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
            
            If teclaQueSelecciona = tecla.retrocesoPágina Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelección = selección.creció Then cadena = "seleccionando varios renglones desde donde estabas hacia abajo en la hoja"
                    If estadoSelección = selección.disminuyó Then cadena = "disminuyendo la selección varios renglones desde donde estabas hacia abajo en la hoja"
                    'cadena =  "seleccionando varios renglones desde donde estabas hacia abajo en la hoja"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja está vacía"
                End If
            End If
        End If
    
        
        If control And teclaQueSelecciona = tecla.a Then 'seleccionar todo el texto
            If cuadroRTF.Text <> "" Then
                cadena = "se seleccionó todo el texto de la hoja"
            Else
                cadena = "No se puede seleccionar porque la hoja está vacía"
            End If
        End If
        
        If cadena <> "" Then Decir cadena
    End If
    
    LenSelecciónAnterior = Len(cuadroRTF.SelText)
End Sub

Public Sub crearCarpetasPrograma()
    Dim j As Byte, archivolibre As Byte, cadena As String
    
    On Error GoTo manejoError
    archivolibre = FreeFile 'se abren las materias
    Open App.path + "\datos\materias.txt" For Input As archivolibre
    
    MkDir (App.path + "\trabajos")
    MkDir (App.path + "\Comunicaciones")
    MkDir (App.path + "\Recordatorios")
    MkDir (App.path + "\Recordatorios\" + Trim(Str(Year(Date))))
    For j = 1 To 12
        MkDir (App.path + "\Recordatorios\" + Trim(Str(Year(Date))) + "\" + Trim(Str(j)))
    Next
    For j = 1 To 12
        MkDir (App.path + "\Datos\" + Trim(Str(j)))
    Next
    While Not EOF(archivolibre)
        Line Input #archivolibre, cadena
        MkDir (App.path + "\trabajos\" + cadena) 'se crean las carpetas de cada materia
        For j = 1 To 12
            MkDir (App.path + "\trabajos\" + cadena + "\" + Trim(Str(j)))
            MkDir (App.path + "\trabajos\" + cadena + "\" + Trim(Str(j)) + "\datosHojas")
        Next
        MkDir (App.path + "\trabajos\" + cadena + "\actividades")
        MkDir (App.path + "\trabajos\" + cadena + "\soporte")
        MkDir (App.path + "\trabajos\" + cadena + "\evaluaciones") 'carpeta para poner evaluaciones falsas por si los papás quieren modificar una evaluación ya hecha ];)
        For j = 1 To 12
            MkDir (App.path + "\trabajos\" + cadena + "\actividades\" + Trim(Str(j)))
            MkDir (App.path + "\trabajos\" + cadena + "\actividades\" + Trim(Str(j)) + "\datosActividades")
            MkDir (App.path + "\trabajos\" + cadena + "\soporte\" + Trim(Str(j)))
            MkDir (App.path + "\trabajos\" + cadena + "\soporte\" + Trim(Str(j)) + "\datosSoporte")
            MkDir (App.path + "\trabajos\" + cadena + "\evaluaciones\" + Trim(Str(j)))
        Next
        MkDir (App.path + "\trabajos\" + cadena + "\libros")
    Wend
    Close #archivolibre
    
'    Dim lectorRegistro, x
'    Set lectorRegistro = CreateObject("WScript.Shell")
'    lectorRegistro.RegWrite "HKEY_LOCAL_MACHINE\Software\ReyNegro-ReyBlanco\MochilaVirtual\", "1"
'    lectorRegistro.RegWrite "HKEY_LOCAL_MACHINE\Software\ReyNegro-ReyBlanco\MochilaVirtual\datos\", "0"
'    Set lectorRegistro = Nothing
manejoError:
    Resume Next
End Sub


Public Function compararFechas(fecha1 As Date, fecha2 As Date) As Byte
    Dim día1 As Byte, día2 As Byte, mes1 As Byte, mes2 As Byte, año1 As Integer, año2 As Integer
    
    día1 = Left(Format(fecha1, "dd/mm/yyyy"), 2)
    día2 = Left(Format(fecha2, "dd/mm/yyyy"), 2)
    
    mes1 = Mid(Format(fecha1, "dd/mm/yyyy"), 4, 2)
    mes2 = Mid(Format(fecha2, "dd/mm/yyyy"), 4, 2)
    
    año1 = Right(Format(fecha1, "dd/mm/yyyy"), 4)
    año2 = Right(Format(fecha2, "dd/mm/yyyy"), 4)
    
    If año1 > año2 Then
        compararFechas = comparación.primeroMayor
        Exit Function
    End If
    
    If año1 < año2 Then
        compararFechas = comparación.primeroMenor
        Exit Function
    End If
    
    If año1 = año2 Then
        If mes1 > mes2 Then
            compararFechas = comparación.primeroMayor
            Exit Function
        End If
        
        If mes1 < mes2 Then
            compararFechas = comparación.primeroMenor
            Exit Function
        End If
        
        If mes1 = mes2 Then
            If día1 > día2 Then
                compararFechas = comparación.primeroMayor
                Exit Function
            End If
            
            If día1 < día2 Then
                compararFechas = comparación.primeroMenor
                Exit Function
            End If
            
            If día1 = día2 Then
                compararFechas = comparación.iguales
                Exit Function
            End If
        End If
    End If
End Function


Public Function compararHora(primerHora As Date, segundaHora As Date) As Byte
    Dim hora1 As Byte, hora2 As Byte, minutos1 As Byte, minutos2 As Byte
    
    hora1 = Left(Format(primerHora, "HH:mm"), 2)
    hora2 = Left(Format(segundaHora, "HH:mm"), 2)
    minutos1 = Right(Format(primerHora, "HH:mm"), 2)
    minutos2 = Right(Format(segundaHora, "HH:mm"), 2)
    
    If hora1 > hora2 Then
        compararHora = comparación.primeroMayor
        Exit Function
    End If
    
    If hora1 < hora2 Then
        compararHora = comparación.primeroMenor
        Exit Function
    End If
    
    If hora1 = hora2 Then
        If minutos1 > minutos2 Then
            compararHora = comparación.primeroMayor
            Exit Function
        End If
    
        If minutos1 < minutos2 Then
            compararHora = comparación.primeroMenor
            Exit Function
        End If
        
        If minutos1 = minutos2 Then
            compararHora = comparación.iguales
            Exit Function
        End If
    End If
End Function


' ----------------------------------------------------------------------------------------
' \\ --   Subrutina para cargar en forma dinámica el menú de opciones
' ----------------------------------------------------------------------------------------
Public Sub Cargar_Menu(El_SubMenu As Object, palabras() As String)
    Dim i As Integer
    
    ' -- Por si hay Submenu cargados, los descarga a todos
    For i = 1 To El_SubMenu.Count - 1
        Unload El_SubMenu(i)
    Next
    
    If UBound(palabras) <> 0 Then 'si se manda alguna palabra para el menú
        For i = 0 To UBound(palabras)
            ' -- Establecer el caption del primer SubMenu
            El_SubMenu(El_SubMenu.Count - 1).Caption = palabras(i)
            
            ' -- Crear otro menú dinamicamente mediante Load
            If i <> UBound(palabras) Then Load El_SubMenu(El_SubMenu.Count)
        Next
    Else
        El_SubMenu(El_SubMenu.Count - 1).Caption = "No sé qué palabra sugerirte"
    End If
End Sub


Public Sub Cargar_Menú_En_Lista(lista As ListBox, palabras() As String)
    Dim i As Byte
    
    lista.Clear
    For i = 0 To UBound(palabras)
        lista.AddItem palabras(i)
    Next
    lista.Visible = True
    lista.SetFocus
End Sub

Public Function buscarPalabraParaCorregir(cuadroRTF As RichTextBox) As String
    Dim cont As Long ', cadena As String,
    Dim cont2 As Long
    
    cont = cuadroRTF.SelStart
    If cont <> 0 Then
        'Do While Mid(cuadroRtf.Text, cont, 1) <> Chr(1) 'se busca el comienzo de la palabra
        Do While (Asc(Mid(cuadroRTF.Text, cont, 1)) >= 65 And Asc(Mid(cuadroRTF.Text, cont, 1)) <= 90) _
        Or (Asc(Mid(cuadroRTF.Text, cont, 1)) >= 97 And Asc(Mid(cuadroRTF.Text, cont, 1)) <= 122) _
        Or (Asc(Mid(cuadroRTF.Text, cont, 1)) >= 192 And Asc(Mid(cuadroRTF.Text, cont, 1)) <= 220) _
        Or (Asc(Mid(cuadroRTF.Text, cont, 1)) >= 224 And Asc(Mid(cuadroRTF.Text, cont, 1)) <= 252)
            cont = cont - 1
            If cont = 0 Then Exit Do
        Loop
    End If
    
    cont2 = cont + 1
    
    'se busca ve cuánto mide la palabra
    If cont2 < Len(cuadroRTF.Text) Then
        Do While (Asc(Mid(cuadroRTF.Text, cont2, 1)) >= 65 And Asc(Mid(cuadroRTF.Text, cont2, 1)) <= 90) _
        Or (Asc(Mid(cuadroRTF.Text, cont2, 1)) >= 97 And Asc(Mid(cuadroRTF.Text, cont2, 1)) <= 122) _
        Or (Asc(Mid(cuadroRTF.Text, cont2, 1)) >= 192 And Asc(Mid(cuadroRTF.Text, cont2, 1)) <= 220) _
        Or (Asc(Mid(cuadroRTF.Text, cont2, 1)) >= 224 And Asc(Mid(cuadroRTF.Text, cont2, 1)) <= 252)
            cont2 = cont2 + 1
            If cont2 >= Len(cuadroRTF.Text) Then Exit Do
        Loop
    End If
    
'    If cont = 0 Then
'        cont = 1
'    Else
        cont = cont + 1
'    End If
    
    If cont2 = Len(cuadroRTF.Text) Then ' cont = 1 And cont2 = Len(cuadroRtf.Text) Then 'si hay una sola palabra escrita o si es la última palabra
        buscarPalabraParaCorregir = Trim(Mid(cuadroRTF.Text, cont, cont2))
    Else
        If cont2 - cont >= 0 Then
            buscarPalabraParaCorregir = Trim(Mid(cuadroRTF.Text, cont, cont2 - cont))
        Else
            buscarPalabraParaCorregir = ""
        End If
    End If
End Function


Public Sub KillProcess(ByVal processName As String)
On Error GoTo ErrHandler
    Dim oWMI
    Dim ret
    Dim sService
    Dim oWMIServices
    Dim oWMIService
    Dim oServices
    Dim oService
    Dim servicename

    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")

    For Each oService In oServices
        servicename = _
            LCase(Trim(CStr(oService.Name) & ""))

        If InStr(1, servicename, _
            LCase(processName), vbTextCompare) > 0 Then
            ret = oService.Terminate
        End If
    Next

    Set oServices = Nothing
    Set oWMI = Nothing
    Exit Sub
ErrHandler:
    Err.Clear
End Sub


Public Function arreglarCadena(cadena As String) As String() 'para el corrector aspell
    Dim posición As Integer, cantidadDevolución As Integer, cadenaAux As String
    Dim temp() As String, contador As Integer, i As Integer ', tempAux(0 To 10) As String
    
    ReDim temp(0 To 0) 'para que no dé error si no hay palabras que coincidan
    posición = InStr(1, cadena, "*") 'se ve si la palabra fue correcta
    If posición = 0 Then
        posición = InStr(1, cadena, "&") 'se dejan sólo las palabras devueltas
        cadenaAux = Trim(Right(cadena, Len(cadena) - posición))
        
        posición = InStr(1, cadenaAux, " ") 'se busca el largo del array que devuelve aspell
        cadenaAux = Right(cadenaAux, Len(cadenaAux) - posición)
        posición = InStr(1, cadenaAux, " ")
        If IsNumeric(Left(cadenaAux, posición)) Then 'si hay alguna devolución para la palabra
            cantidadDevolución = Int(Left(cadenaAux, posición))
            ReDim Preserve temp(0 To cantidadDevolución - 1) 'se estira el array según la cantidad de palabras que devulve aspell
            
            If cantidadDevolución > 0 Then
                posición = InStr(1, cadenaAux, ":") 'se busca dejar sólo las palabras sugeridas
                cadenaAux = Trim(Right(cadenaAux, Len(cadenaAux) - posición))
            End If
            
            posición = InStr(1, cadenaAux, "ù") 'se quitan los caracteres ù
            Do While posición
                cadenaAux = Left(cadenaAux, posición - 1) + "é" + Right(cadenaAux, Len(cadenaAux) - posición)
                posición = InStr(1, cadenaAux, "ù")
            Loop
                    
            posición = InStr(1, cadenaAux, "ý") 'se quitan los caracteres ý
            Do While posición
                cadenaAux = Left(cadenaAux, posición - 1) + "í" + Right(cadenaAux, Len(cadenaAux) - posición)
                posición = InStr(1, cadenaAux, "ý")
            Loop
                   
           'Debug.Print cadenaAux
           
            posición = InStr(1, cadenaAux, "ø") 'se quitan los caracteres ý
            Do While posición
                cadenaAux = Left(cadenaAux, posición - 1) + "è" + Right(cadenaAux, Len(cadenaAux) - posición)
                posición = InStr(1, cadenaAux, "ø")
            Loop
            
            posición = InStr(1, cadenaAux, ",") 'se llena el array con las palabras devueltas
            contador = 0
            Do While posición
                If contador >= UBound(temp) Then Exit Do 'nos aseguramos que el contador no supere el límite del array así no da error
                temp(contador) = controlar_A_Acentuada(Trim(Left(cadenaAux, posición - 1)))
                cadenaAux = Right(cadenaAux, Len(cadenaAux) - posición)
                posición = InStr(1, cadenaAux, ",")
                contador = contador + 1
            Loop
            temp(contador) = controlar_A_Acentuada(Trim(Left(cadenaAux, Len(cadenaAux) - 4))) 'se carga la última palabra
        End If
    End If
    
'    If UBound(temp) > UBound(tempAux) Then
'        For i = 0 To UBound(tempAux)
'            tempAux(i) = temp(i)
'        Next
'        arreglarCadena = tempAux 'se devuelve el array
'    Else
        arreglarCadena = temp
'    End If
End Function


Public Function esSigno(cadena As String) As Boolean
    Dim i As Byte
    For i = 0 To 254
        If (i >= 65 And i <= 90) _
        Or (i >= 97 And i <= 122) _
        Or (i >= 192 And i <= 220) _
        Or (i >= 224 And i <= 252) Then
            If InStr(1, cadena, Chr(i)) Then 'con que haya una sola letra, se devuelve Falso
                esSigno = False
                Exit Function
            End If
        End If
    Next
    esSigno = True 'si no se devolvió falso, es que son sólo signos. Devuelve verdadero
End Function

Public Function separarEnLetras(palabra As String) As String
    Dim i As Integer, cadenaTemp As String
    
    For i = 1 To Len(palabra)
        cadenaTemp = cadenaTemp + Mid(palabra, i, 1) + ". "
    Next
    separarEnLetras = controlarDeletreo(cadenaTemp)
End Function

Public Function controlarDeletreo(cadena As String) As String
    Dim carácter(26) As String, posiciónCaracter As Long, cadenaFinal As String
    Dim i As Byte, swEntróAlFor As Boolean, swYaEmpezó As Boolean
    
    cadena = LCase(cadena)
    
    If cadena <> "" Then
        carácter(0) = "m. "
        carácter(1) = "s. "
        carácter(2) = "l. "
        carácter(3) = "h. "
        carácter(4) = "p. "
        carácter(5) = "n. "
        carácter(6) = "á. "
        carácter(7) = "é. "
        carácter(8) = "í. "
        carácter(9) = "ó. "
        carácter(10) = "ú. "
        carácter(11) = "à. "
        carácter(12) = "è. "
        carácter(13) = "ì. "
        carácter(14) = "ò. "
        carácter(15) = "ù. "
        carácter(16) = "â. "
        carácter(17) = "ê. "
        carácter(18) = "î. "
        carácter(19) = "ô. "
        carácter(20) = "û. "
        carácter(21) = "g. "
        carácter(22) = "u. "
        carácter(23) = "d. "
        carácter(24) = "b. "
        carácter(25) = "v. "
        carácter(26) = "y. "

        swEntróAlFor = False
        swYaEmpezó = False
        
        For i = 0 To UBound(carácter)
            If swYaEmpezó = False Then
                posiciónCaracter = InStr(1, cadena, carácter(i))
            Else
                posiciónCaracter = InStr(1, cadenaFinal, carácter(i))
            End If
            
            Do While posiciónCaracter <> 0
                If swEntróAlFor = False Then
                    cadenaFinal = corregirCadena(cadena, posiciónCaracter, carácter(i))
                Else
                    cadenaFinal = corregirCadena(cadenaFinal, posiciónCaracter, carácter(i))
                End If
                posiciónCaracter = InStr(1, cadenaFinal, carácter(i))
                swEntróAlFor = True
                swYaEmpezó = True
            Loop
        Next
        
        If cadenaFinal = "" Then cadenaFinal = cadena
        controlarDeletreo = cadenaFinal
    End If
End Function



Public Function controlar_A_Acentuada(devoluciónAspell As String) As String
    Dim Pos As Integer, palabraCambiada As String
    
    Pos = InStr(1, devoluciónAspell, "ñ") 'se quitan los caracteres ñ
    If Pos <> 0 Then 'si hay alguna ñ
        Do While Pos
            palabraCambiada = Left(devoluciónAspell, Pos - 1) + "á" + Right(devoluciónAspell, Len(devoluciónAspell) - Pos)
            Pos = InStr(Pos + 1, devoluciónAspell, "ñ")
                    
            'se ve si la palabra cambiada es correcta para dejar de cambiar ñ
            If corregirPalabra(palabraCambiada) Then Exit Do
        Loop
        
        'se ve si la palabra cambiada es correcta
        If corregirPalabra(palabraCambiada) Then 'si la palabra es correcta
            controlar_A_Acentuada = palabraCambiada
        Else
            controlar_A_Acentuada = devoluciónAspell
        End If
    Else 'si no hay ñ en la palabra, se la devuelve sin cambio
        controlar_A_Acentuada = devoluciónAspell
    End If
End Function


Public Sub corregirConPalabraSeleccionada(cuadroRTF As RichTextBox, palabra As String)
    Dim cont As Long ', cadena As String,
    Dim cont2 As Long
    
    cont = cuadroRTF.SelStart
    If cont <> 0 Then
        'Do While Mid(cuadroRtf.Text, cont, 1) <> Chr(1) 'se busca el comienzo de la palabra
        Do While (Asc(Mid(cuadroRTF.Text, cont, 1)) >= 65 And Asc(Mid(cuadroRTF.Text, cont, 1)) <= 90) _
        Or (Asc(Mid(cuadroRTF.Text, cont, 1)) >= 97 And Asc(Mid(cuadroRTF.Text, cont, 1)) <= 122) _
        Or (Asc(Mid(cuadroRTF.Text, cont, 1)) >= 192 And Asc(Mid(cuadroRTF.Text, cont, 1)) <= 220) _
        Or (Asc(Mid(cuadroRTF.Text, cont, 1)) >= 224 And Asc(Mid(cuadroRTF.Text, cont, 1)) <= 252)
            cont = cont - 1
            If cont = 0 Then Exit Do
        Loop
    End If
    
    cont2 = cont + 1
    
    'se busca ve cuánto mide la palabra
    If cont2 < Len(cuadroRTF.Text) Then
        Do While (Asc(Mid(cuadroRTF.Text, cont2, 1)) >= 65 And Asc(Mid(cuadroRTF.Text, cont2, 1)) <= 90) _
        Or (Asc(Mid(cuadroRTF.Text, cont2, 1)) >= 97 And Asc(Mid(cuadroRTF.Text, cont2, 1)) <= 122) _
        Or (Asc(Mid(cuadroRTF.Text, cont2, 1)) >= 192 And Asc(Mid(cuadroRTF.Text, cont2, 1)) <= 220) _
        Or (Asc(Mid(cuadroRTF.Text, cont2, 1)) >= 224 And Asc(Mid(cuadroRTF.Text, cont2, 1)) <= 252)
            cont2 = cont2 + 1
            If cont2 >= Len(cuadroRTF.Text) Then Exit Do
        Loop
    End If
    
    If cont2 = Len(cuadroRTF.Text) Then cont2 = cont2 + 1
    
    cuadroRTF.SelStart = cont ' Comenzamos desde la cantidad de caracteres menos 1
    cuadroRTF.SelLength = cont2 - cont - 1 ' Con un maximo de un caracter.
    cuadroRTF.SelText = palabra ' Borramos
End Sub


Public Function buscarEntrada(quéCadena As String, diccionarioElegido As String) As String
    Dim cadenaDiccionario As String
    Dim palabraEncontrada As Boolean
    Dim archivolibre As Byte 'el manejador del diccionario
    Dim palabraNuevaLínea As String
    Dim posiciónDosPuntos As Integer
    
    If existeCarpeta(App.path + "\diccionarios\" + diccionarioElegido) Then
        'si existe el diccionario
        palabraEncontrada = False
        buscarEntrada = ""
        
        archivolibre = FreeFile
        Open App.path + "\diccionarios\" + diccionarioElegido For Input As archivolibre
        Do While Not EOF(archivolibre)   ' Repite el bucle hasta el final del archivo.
            Line Input #archivolibre, cadenaDiccionario ' Lee el carácter en dos variables.
            If palabraEncontrada = False Then 'se ve si la palabra está en el diccionario
                posiciónDosPuntos = InStr(1, cadenaDiccionario, ":") - 1
                If posiciónDosPuntos > 0 Then
                    If LCase(Trim(Left(cadenaDiccionario, posiciónDosPuntos))) = LCase(Trim(quéCadena)) Then
                    'esto sirve para buscar la coincidencia letra por letra de lo que se escribe -> 'If LCase(Trim(Left(cadenaDiccionario, Len(quéCadena)))) = LCase(Trim(quéCadena)) Then
                        buscarEntrada = cadenaDiccionario
                        palabraEncontrada = True
                    End If
                End If
            Else 'si se encontró la palabra, se ve si la definición sigue en el próximo renglón
                palabraNuevaLínea = Left(cadenaDiccionario, InStr(1, cadenaDiccionario, " "))
                'chequear si está en mayúsculas, si es así, salir de la función, sinó, sumar la cadena a la ya obtenida
                If UCase(palabraNuevaLínea) = palabraNuevaLínea Then
                    Exit Function
                Else
                    buscarEntrada = buscarEntrada + " " + cadenaDiccionario
                End If
            End If
        Loop
        Close archivolibre
    Else
        buscarEntrada = ""
    End If
End Function


'Public Function matrizInicializada(quéMatriz() As Object) As Boolean 'para controlar que una matriz dinámica ya esté dimensionada
'    On Error GoTo error:
'    If UBound(quéMatriz) Then matrizInicializada = True
'    Exit Function
'error:
'    matrizInicializada = False
'End Function


Function Aplicar_ScrollBar(ListBox As ListBox) As Long
     Dim ret          As Long
     Dim i            As Integer
     Dim j            As Long
     Dim Ancho_Maximo As Long
     Dim Ancho_Texto  As Long
     Dim LBParent     As Object

     Set LBParent = ListBox.Parent
     Ancho_Maximo = -1
     j = -1
     For i = 0 To ListBox.ListCount - 1
        Ancho_Texto = LBParent.TextWidth(ListBox.List(i))
        If Ancho_Texto > Ancho_Maximo Then
            Ancho_Maximo = Ancho_Texto + (10 * Screen.TwipsPerPixelX)
            j = i
        End If
     Next
     Set LBParent = Nothing
     ' -- Establecer el Scroll
     ret = SendMessage(ListBox.hwnd, LB_SETHORIZONTALEXTENT, (Ancho_Maximo / Screen.TwipsPerPixelX), ByVal 0&)
     ' -- retornar item mas largo
     Aplicar_ScrollBar = j
End Function

