Attribute VB_Name = "funciones"
Option Explicit

Public Voz As SpVoice
Public vozSapi4 As DirectSS 'As TextToSpeech
Private Enum selecci�n
    creci�
    disminuy�
    igual
End Enum

'funciones para registrar los ocx
'Public Declare Function RegFlash Lib "Flash9f.ocx" Alias "DllRegisterServer" () As Long
'Public Declare Function UnRegFlash Lib "Flash9f.ocx" Alias "DllUnregisterServer" () As Long
'Public Declare Function RegBot�nTransp Lib "TransparentButton.ocx" Alias "DllRegisterServer" () As Long
'Public Declare Function UnRegBot�nTransp Lib "TransparentButton.ocx" Alias "DllUnregisterServer" () As Long

Public Function guardarOrdenCap�tulosDesdeMatriz(materia As String, libro As String, ParamArray listaCap�tulos()) As Boolean
    Dim cap�tulo As Variant, auxMatriz() As String, i As Integer, contador As Integer, archivolibre As Byte
    On Error GoTo error
    archivolibre = FreeFile
    Open App.path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCap�tulos" For Output As #archivolibre
    contador = 0
    For Each cap�tulo In listaCap�tulos
        For i = 0 To UBound(cap�tulo)
            ReDim Preserve auxMatriz(0 To contador)
            auxMatriz(contador) = cap�tulo(contador)
            contador = contador + 1
        Next
    Next cap�tulo
    
    For i = 0 To UBound(auxMatriz)
        Print #archivolibre, auxMatriz(i)
    Next i
    
    Close #archivolibre
    
    guardarOrdenCap�tulosDesdeMatriz = True
    Exit Function
error:
    guardarOrdenCap�tulosDesdeMatriz = False
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

Public Sub guardarOrdenCap�tulosLibro(listaCap�tulos As ListBox, materia As String, libro As String)
    Dim i As Integer
    Open App.path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCap�tulos" For Output As #1
    listaCap�tulos.Refresh
    For i = 0 To listaCap�tulos.ListCount - 1
        Print #1, listaCap�tulos.List(i)
    Next
    Close #1
End Sub


Public Sub guardarHistorial(listaMaterias As ListBox)
    Dim swArchivoRepetido As Boolean, nextline As String, i As Integer
    
    On Error GoTo manejoError
    'se guardan las materias en el historial
    For i = 0 To listaMaterias.ListCount - 1 'se chequea que cada materia no est� ya guardada
        Open App.path + "\datos\historialMaterias.txt" For Input As #1 'se abre el trabajo ya guardado para leerlo
        Do While Not EOF(1) 'chequeamos que no est� en la lista ya el registro del archivo a guardar
            Line Input #1, nextline
            If nextline = listaMaterias.List(i) Then
                swArchivoRepetido = True
                Exit Do
            End If
        Loop
        Close #1
        
        If swArchivoRepetido = False Then
            Open App.path + "\datos\historialMaterias.txt" For Append As #1 'se abre el historial para a�adir las materias
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
        For i = 0 To Combo1.ListCount - 1 'se activa la voz sapi5 que el usuario us� en la sesi�n anterior
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
'            MsgBox "Se va a instalar la voz del programa. Por favor, presione el bot�n " + Chr(34) + "S�" + Chr(34) + " en el cuadro que va a aparecer.", , "Informaci�n"
'            Call ejecutar(App.Path + "\ejecutables\TTS3000.exe")
'            swInstalarVoz = True
'        End If
    Else
        For i = 0 To Combo2.ListCount - 1 'se activa la voz sapi4 que el usuario us� en la sesi�n anterior
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
        MsgBox "El programa necesita tener instaladas Sapi4 y Sapi5. Una de ellas o ambas faltan. Por favor inst�lelas y reinicie el programa. Ambas sapi est�n en la carpeta " + Chr(34) + "Ejecutables" + Chr(34) + "  que viene con el programa. Si por cualquier eventualidad no estuviesen all�, se pueden descargar gratuitamente desde la p�gina de Microsoft. Tenga presente instalar la sapi5 que es propia de su Windows, ya sea sapi5 para Windows XP, � sapi5 para versiones anteriores. Sapi4 es id�ntica para cualquier versi�n de Windows.", , "No se encuentra una SAPI"
        End
    End If
    
    If Err.Number = 53 Then
        'On Error GoTo manejoerrorCancelar
'        aceptar =
        MsgBox "No se encuentra el instalador de la voz que viene con el programa. Inst�lelo usted manualmente desde la carpeta " + Chr(34) + "Ejecutables" + Chr(34) + "  que viene con el programa. Si por cualquier eventualidad no estuviese all�, se puede descargar gratuitamente desde la p�gina de Microsoft con el nombre de TTS3000.", , "Imposible instalar la voz del programa"
'        If aceptar = vbYes Then
'            frmControl.di�logo.CancelError = True
'            frmControl.di�logo.Filter = "Archivos Ejecutables (*.exe);*.rtf"
'            frmControl.di�logo.ShowOpen
'            If Err.Number = cdlCancel Then Exit Sub
'            Shell frmControl.di�logo.FileName
'        End If
        Exit Sub
    End If
    
'    If Err.Number = 429 Then MsgBox "soy el controlador del m�dulo funciones, error 429", , "Para mi creador"
    
'    MsgBox "soy el controlador del la funci�n llenarComboVoz. Error n�mero: " + Str(Err.Number) + ", descripci�n: " + Err.Description, , "Para mi creador"
    frmMsgBox.cadenaAMostrar = "Soy el controlador del la funci�n llenarComboVoz. Error n�mero: " + Str(Err.Number) + ", descripci�n: " + Err.Description
    frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
    frmMsgBox.Show 1
    Exit Sub
'    Exit Sub
'manejoerrorCancelar:
'    Exit Sub
End Sub

'Sub guardarTrabajo(NombreArchivo, textoAGuardar)
'    Dim nextLine As String 'aqu� se almacena el contenido del registro de los guardados para chequear que no se repita un nombre
'    Dim cadena As String 'para almacenar la lista de guardados cuando se trabaja con tareas viejas
'
'    On Error GoTo manejoError
''    If swNuevoContinuar = False Then 'si se trabaja en un archivo nuevo
'        Open App.Path + NombreArchivo For Output As #1   ' Abre el archivo para operaciones de salida.
''    Else 'si se est� trabajando con una tarea vieja
''        Open MiTrabajo For Output As #1 'se abre el trabajo ya guardado
''    End If
'    Print #1, textoAGuardar
'    Close #1
'
'    Open App.Path + "\trabajos\listadeGuardados.txt" For Input As #2
'    Do Until EOF(2) 'chequeamos que no est� en la lista ya el registro del archivo a guardar
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
'    Open App.Path + "\trabajos\listadeGuardados.txt" For Append As #2 'abrimos el registro de los archivos guardados para a�adir este que hemos guardado ahora
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

'Public Sub GuardarRTF(nombreRuta As String, cu�lRTF As RichTextBox)
'        cu�lRTF.SaveFile nombreRuta, rtfRTF
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
        .mostrarTodasLasTareas = swMostrarA�oEnTareas
        .mostrarTodasLasActividades = swMostrarA�oEnActividades
        .nombre = nombreUsuario
        .usuarioMujer = swUsuarioMujer
        '.leerSignoPuntuaci�n = swLeerSignosPuntuaci�n
        .imprimirDirecto = swImprimirDirecto
        .colorFondo = colorFondo
        .fuenteColor = colorFuente
        .fuenteNombre = NombreFuente
        .fuenteTama�o = tama�oFuente
        .velocidadVoz = velocidadVoz
        .swLeerRenglones = swLeerRenglones
        .swUsarCorrectorOrtogr�fico = swUsarCorrectorOrtogr�fico
        .nombreVozSapi4 = nombreSapi4
        .nombreVozSapi5 = nombreSapi5
        '.swInstalarVoz = swInstalarVoz
        .swM�sicaDeFondo = swM�sicaDeFondo
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

Public Sub Decir(ByVal qu� As String, Optional usarbanderasSpVoice As Boolean = True, Optional esperarSapi4 As Boolean)
    On Error Resume Next 'esto se lo agregu� por la compu de franco, que tira un error siempre
    If swHablarVoz = True Then 'variable general del programa
        If swSapi5 = True Then
            If usarbanderasSpVoice = True Then
                Voz.Speak qu�, SVSFPurgeBeforeSpeak Or SVSFlagsAsync ' Or SVSFNLPSpeakPunc
            Else
                Voz.Speak qu�, SVSFPurgeBeforeSpeak Or SVSFlagsAsync
            End If
        Else
            If esperarSapi4 = False Then
                vozSapi4.AudioReset
                vozSapi4.Speak qu�
            Else
                vozSapi4.Speak qu�
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
'    cadena = Mid(cuadroRTF.Text, cont, 1) 'si se est� en el comienzo del cuadro de texto
'    While Right(cadena, 1) <> " " And Right(cadena, 1) <> "," And _
'    Right(cadena, 1) <> "." And Right(cadena, 1) <> ";" And _
'    Right(cadena, 1) <> ":" And Right(cadena, 1) <> "?" And _
'    Right(cadena, 1) <> "!" And Len(cuadroRTF.Text) <> cont
'        cont = cont + 1
'        cadena = cadena + Mid(cuadroRTF.Text, cont, 1)
'    Wend
'
'    If Len(cuadroRTF.Text) <> cont Then 'si no se est� al final de la hoja
'        decirPalabra = cadena
'    Else
'        decirPalabra = cadena '"est�s al final de tu hoja"
'    End If
'    ponerPuntoInserci�nEn = cont
'End Function

Public Function decirPalabraSiguiente(cuadroRTF As RichTextBox) As String
    Dim cont As Long, cadena As String, rengl�nActual As Long
    cont = cuadroRTF.SelStart
    If swLeerRenglones = True Then rengl�nActual = cuadroRTF.GetLineFromChar(cuadroRTF.SelStart)
    If Len(cuadroRTF.Text) <> cont Then 'si no se est� al final de la hoja
        cont = cont + 1
        cadena = Mid(cuadroRTF.Text, cont, 1) 'si se est� en el comienzo del cuadro de texto
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
'        if cadena="(" then cadena="abre par�ntesis"
        
        If cuadroRTF.SelStart <> 0 Then 'si no se est� al principio de la hoja
            decirPalabraSiguiente = cadena
            If cuadroRTF.SelBold = True Then decirPalabraSiguiente = decirPalabraSiguiente + " en negrita"
            If IsNull(cuadroRTF.SelBold) Then decirPalabraSiguiente = decirPalabraSiguiente + " parte en negrita"
            If cuadroRTF.SelUnderline = True Then decirPalabraSiguiente = decirPalabraSiguiente + " subrayada"
            If IsNull(cuadroRTF.SelUnderline) Then decirPalabraSiguiente = decirPalabraSiguiente + " parte subrayada"
        Else
            If cadena = "." Or cadena = "," Or cadena = ";" Or cadena = "aparte" Then
                decirPalabraSiguiente = "est�s al principio de la hoja, delante del signo. " + cadena
            Else
                decirPalabraSiguiente = "est�s al principio de la hoja, delante de la palabra. " + cadena
                If cuadroRTF.SelBold = True Then decirPalabraSiguiente = decirPalabraSiguiente + " en negrita"
                If IsNull(cuadroRTF.SelBold) Then decirPalabraSiguiente = decirPalabraSiguiente + " parte en negrita"
                If cuadroRTF.SelUnderline = True Then decirPalabraSiguiente = decirPalabraSiguiente + " subrayada"
                If IsNull(cuadroRTF.SelUnderline) Then decirPalabraSiguiente = decirPalabraSiguiente + " parte subrayada"
            End If
        End If
        
        If swLeerRenglones = True Then
            If rengl�nActual <> rengl�nAnterior Then decirPalabraSiguiente = decirPalabraSiguiente & ". rengl�n " & Str(rengl�nActual + 1)
        End If
    Else
        If cuadroRTF.Text <> "" Then
            decirPalabraSiguiente = "llegaste al final de la hoja"
        Else
            decirPalabraSiguiente = "no hay nada escrito, la hoja est� vac�a"
        End If
    End If
    decirPalabraSiguiente = controlarCadena(decirPalabraSiguiente)
    rengl�nAnterior = rengl�nActual
End Function

Public Function decirPalabraAnterior(cuadroRTF As RichTextBox) As String
    Dim cont As Long, cadena As String
    If cuadroRTF.SelStart <> 0 Then 'si no se est� al principio de la carpeta
        If cuadroRTF.Text <> "" Then
            cont = cuadroRTF.SelStart
            'If cont <> Len(cuadroRtf.Text) Then cont = cont - 1
            cadena = Mid(cuadroRTF.Text, cont, 1) 'si se est� en el comienzo del cuadro de texto
            While Left(cadena, 1) <> " " And Left(cadena, 1) <> "," And _
            Left(cadena, 1) <> "." And Left(cadena, 1) <> ";" And _
            Left(cadena, 1) <> ":" And Left(cadena, 1) <> "?" And Left(cadena, 1) <> Chr(10) And _
            Left(cadena, 1) <> Chr(13) And Left(cadena, 1) <> "!" And cont <> 1
                cont = cont - 1
                cadena = Mid(cuadroRTF.Text, cont, 1) + cadena
            Wend
            
            'si es un n�mero con . que no lea una parte del n�mero sino todo �l
            If IsNumeric(Right(cadena, Len(cadena) - 1)) And Left(cadena, 1) = "." And cont <> 1 Then
                'se toma el siguiente n�mero
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
    Dim cont As Long, cadena As String, rengl�nActual As Long
    If swLeerRenglones = True Then rengl�nActual = cuadroRTF.GetLineFromChar(cuadroRTF.SelStart)
    cont = cuadroRTF.SelStart
    If Len(cuadroRTF.Text) <> cont Then
        cont = cont + 1
        cadena = Mid(cuadroRTF.Text, cont, 1) 'si se est� en el comienzo del cuadro de texto
        If cadena = " " Then
            decirLetraSiguiente = "espacio"
        ElseIf cadena = Chr(13) Then
            decirLetraSiguiente = "bajada de l�nea del rengl�n " + Str(cuadroRTF.GetLineFromChar(cuadroRTF.SelStart) + 1)
        ElseIf cadena = ":" Then
            decirLetraSiguiente = "dos puntos"
        Else
            decirLetraSiguiente = cadena '"est�s al final de tu hoja"
        End If
        
        If cuadroRTF.SelBold = True Then decirLetraSiguiente = decirLetraSiguiente + " en negrita"
        If cuadroRTF.SelUnderline = True Then decirLetraSiguiente = decirLetraSiguiente + " subrayada"

        If swLeerRenglones = True Then
            If rengl�nActual <> rengl�nAnterior Then decirLetraSiguiente = decirLetraSiguiente & ". rengl�n " & Str(rengl�nActual + 1)
        End If
    Else
        If cuadroRTF.Text <> "" Then
            decirLetraSiguiente = "llegaste al final de la hoja"
        Else
            decirLetraSiguiente = "no hay nada escrito, la hoja est� vac�a"
        End If
    End If
    decirLetraSiguiente = controlarCadena(decirLetraSiguiente)
    rengl�nAnterior = rengl�nActual
End Function

'Public Function decirLetraAnterior(cuadroRTF As RichTextBox) As String
'    Dim cont As Long, cadena As String
'    cont = cuadroRTF.SelStart
'    If cont <> Len(cuadroRTF.Text) Then cont = cont - 1
'    cadena = Mid(cuadroRTF.Text, cont, 1) 'si se est� en el comienzo del cuadro de texto
'
''    If Len(cuadroRTF.Text) <> cont Then 'si no se est� al final de la hoja
''        decirPalabraAnterior = cadena
''    Else
'        decirLetraAnterior = cadena '"est�s al final de tu hoja"
''    End If
'End Function
'

Public Function decirOraci�nSiguiente(cuadroRTF As RichTextBox) As String
    Dim cont As Long, cadena As String, swEmpezando As Boolean
    Dim l�neaInicio As Long, l�neaActual As Long, TotalDeL�neasEnRTF As Long
    
    l�neaInicio = cuadroRTF.GetLineFromChar(cuadroRTF.SelStart)
    TotalDeL�neasEnRTF = cuadroRTF.GetLineFromChar(Len(cuadroRTF.Text))
    
    cont = cuadroRTF.SelStart
    
    If cont <> 0 Then 'si no se est� ya al comienzo de la l�nea se busca el punto de inicio
        l�neaActual = l�neaInicio
        Do While l�neaActual = l�neaInicio
            cont = cont - 1
            l�neaActual = cuadroRTF.GetLineFromChar(cont)
            If cont = 0 Then Exit Do
        Loop
    End If
    
    cont = cont + 1
    If l�neaInicio = 0 Then swEmpezando = True
    l�neaActual = l�neaInicio
    While l�neaActual = l�neaInicio And cont <= Len(cuadroRTF.Text)
        If Mid(cuadroRTF.Text, cont, 1) <> Chr(10) And Mid(cuadroRTF.Text, cont, 1) <> Chr(13) Then
            cadena = cadena + Mid(cuadroRTF.Text, cont, 1)
        End If
        cont = cont + 1
        l�neaActual = cuadroRTF.GetLineFromChar(cont)
    Wend
            
    If TotalDeL�neasEnRTF <> l�neaInicio Then 'se eval�a si se est� en el �ltimo rengl�n
        If swEmpezando = False Then 'si no se est� al principio de la hoja
            If swLeerRenglones = True Then decirOraci�nSiguiente = "rengl�n " + Str(l�neaInicio + 1)
            If cadena = "" Then
                decirOraci�nSiguiente = decirOraci�nSiguiente + ". rengl�n en blanco"
            Else
                If swLeerRenglones = True Then
                    decirOraci�nSiguiente = decirOraci�nSiguiente + " dice." + cadena
                Else
                    decirOraci�nSiguiente = decirOraci�nSiguiente + cadena
                End If
            End If
        Else
            If swLeerRenglones = True Then decirOraci�nSiguiente = "est�s en el primer rengl�n de la hoja "
            If cadena = "" Then
                decirOraci�nSiguiente = decirOraci�nSiguiente + ". el rengl�n est� en blanco"
            Else
                If swLeerRenglones = True Then
                    decirOraci�nSiguiente = decirOraci�nSiguiente + ". el rengl�n dice. " + cadena
                Else
                    decirOraci�nSiguiente = decirOraci�nSiguiente + cadena
                End If
            End If
        End If
    Else
        If cuadroRTF.Text <> "" Then
            If swLeerRenglones = True Then decirOraci�nSiguiente = "llegaste a el �ltimo rengl�n de la hoja"
            If cadena = "" Then
                decirOraci�nSiguiente = decirOraci�nSiguiente + ". el rengl�n est� en blanco"
            Else
                If swLeerRenglones = True Then
                    decirOraci�nSiguiente = decirOraci�nSiguiente + ". el rengl�n dice. " + cadena
                Else
                    decirOraci�nSiguiente = decirOraci�nSiguiente + cadena
                End If
            End If
        Else
            decirOraci�nSiguiente = "no hay nada escrito, la hoja est� vac�a"
        End If
    End If
    
    decirOraci�nSiguiente = controlarCadena(decirOraci�nSiguiente)
End Function

Public Function controlarCadena(cadena As String) As String
    Dim car�cter(7) As String, posici�nCaracter As Long, cadenaFinal As String
    Dim i As Byte, swEntr�AlFor As Boolean, swYaEmpez� As Boolean
    
    If cadena <> "" Then
        car�cter(0) = "("
        car�cter(1) = ")"
        car�cter(2) = "-"
        car�cter(3) = "*"
        car�cter(4) = "/"
        car�cter(5) = "{"
        car�cter(6) = "}"
        car�cter(7) = " 1 "
        
        swEntr�AlFor = False
        
        For i = 0 To UBound(car�cter)
            If swYaEmpez� = False Then
                posici�nCaracter = InStr(1, cadena, car�cter(i))
            Else
                posici�nCaracter = InStr(1, cadenaFinal, car�cter(i))
            End If
            
            Do While posici�nCaracter <> 0
                If swEntr�AlFor = False Then
                    cadenaFinal = corregirCadena(cadena, posici�nCaracter, car�cter(i))
                Else
                    cadenaFinal = corregirCadena(cadenaFinal, posici�nCaracter, car�cter(i))
                End If
                posici�nCaracter = InStr(posici�nCaracter + 1, cadenaFinal, car�cter(i))
                swEntr�AlFor = True
                swYaEmpez� = True
            Loop
        Next
        
        If cadenaFinal = "" Then cadenaFinal = cadena
        
        car�cter(0) = " m "
        car�cter(1) = " s "
        car�cter(2) = " l "
        car�cter(3) = " h "
        car�cter(4) = " p "
        car�cter(5) = "$"
        car�cter(6) = "_"

        swEntr�AlFor = False
        swYaEmpez� = False
        
        For i = 0 To 6
'            If swYaEmpez� = False Then
'                posici�nCaracter = InStr(1, cadena, car�cter(i))
'            Else
                posici�nCaracter = InStr(1, cadenaFinal, car�cter(i))
'            End If
            
            Do While posici�nCaracter <> 0
'                If swEntr�AlFor = False Then
'                    cadenaFinal = corregirCadena(cadena, posici�nCaracter, car�cter(i))
'                Else
                    cadenaFinal = corregirCadena(cadenaFinal, posici�nCaracter, car�cter(i))
'                End If
                posici�nCaracter = InStr(posici�nCaracter + 1, cadenaFinal, car�cter(i))
'                swEntr�AlFor = True
'                swYaEmpez� = True
            Loop
        Next
        
'        If cadenaFinal = "" Then cadenaFinal = cadena
        controlarCadena = cadenaFinal
    End If
End Function

Public Function corregirCadena(cadena As String, posici�nCaracter As Long, car�cter As String) As String
    Dim cadenaIzq As String, cadenaDer As String, cadenaTotal As String
    Dim car�cterCorregido As String
    
    cadenaIzq = Left(cadena, posici�nCaracter - 1)
    cadenaDer = Mid(cadena, posici�nCaracter + Trim(Len(car�cter)), Len(cadena) - Len(cadenaIzq))
    cadenaTotal = cadenaTotal & cadenaIzq
    Select Case car�cter
        Case "("
            car�cterCorregido = " abre par�ntesis, "
        Case ")"
            car�cterCorregido = " cierra par�ntesis, "
        Case "_"
            car�cterCorregido = " sobre " 'para fracciones
        Case "$"
            car�cterCorregido = " pesos, "
        Case " 1 "
            car�cterCorregido = " uno, "
        Case "-"
            car�cterCorregido = " menos "
        Case "*"
            car�cterCorregido = " multiplicado por "
        Case "/"
            car�cterCorregido = " dividido "
        Case "{"
            car�cterCorregido = " abre llave, "
        Case "}"
            car�cterCorregido = " cierra llave, "
        Case "m "
            car�cterCorregido = " eme. "
        Case "s "
            car�cterCorregido = " ese. "
        Case "b "
            car�cterCorregido = " be larga. "
        Case "v "
            car�cterCorregido = " ve corta. "
        Case "y "
            car�cterCorregido = " ih griega. "
        Case "b. "
            car�cterCorregido = " be larga. "
        Case "v. "
            car�cterCorregido = " ve corta. "
        Case "y. "
            car�cterCorregido = " ih griega. "
        Case "l "
            car�cterCorregido = " ele. "
        Case "h "
            car�cterCorregido = " ache. "
        Case "p "
            car�cterCorregido = " pe. "
        Case "n "
            car�cterCorregido = " ene. "
        Case "m. "
            car�cterCorregido = " eme. "
        Case "s. "
            car�cterCorregido = " ese. "
        Case "l. "
            car�cterCorregido = " ele. "
        Case "h. "
            car�cterCorregido = " ache. "
        Case "p. "
            car�cterCorregido = " pe. "
        Case "n. "
            car�cterCorregido = " ene. "
        Case "g. "
            car�cterCorregido = " je. "
        Case "u. "
            car�cterCorregido = " uh. "
        Case "d. "
            car�cterCorregido = " de. "
        Case "�. "
            car�cterCorregido = " a con acento. "
        Case "�. "
            car�cterCorregido = " e con acento. "
        Case "�. "
            car�cterCorregido = " i con acento. "
        Case "�. "
            car�cterCorregido = " o con acento. "
        Case "�. "
            car�cterCorregido = " u con acento. "
        Case "�. "
            car�cterCorregido = " u con di�resis. "
        Case "�. "
            car�cterCorregido = " a con acento grave. "
        Case "�. "
            car�cterCorregido = " e con acento grave. "
        Case "�. "
            car�cterCorregido = " i con acento grave. "
        Case "�. "
            car�cterCorregido = " o con acento grave. "
        Case "�. "
            car�cterCorregido = " u con acento grave. "
        Case "�. "
            car�cterCorregido = " a con acento circunflejo. "
        Case "�. "
            car�cterCorregido = " e con acento circunflejo. "
        Case "�. "
            car�cterCorregido = " i con acento circunflejo. "
        Case "�. "
            car�cterCorregido = " o con acento circunflejo. "
        Case "�. "
            car�cterCorregido = " u con acento circunflejo. "
    End Select
    cadenaTotal = cadenaTotal & car�cterCorregido & cadenaDer
    corregirCadena = cadenaTotal
End Function


Public Function medioDelRengl�n(cuadroRTF As RichTextBox) As Boolean
    Dim cont As Long, cadena As String ', swEmpezando As Boolean
    Dim l�neaInicio As Long, l�neaActual As Long ', TotalDeL�neasEnRTF As Long
    
    l�neaInicio = cuadroRTF.GetLineFromChar(cuadroRTF.SelStart)
    l�neaActual = l�neaInicio
    cont = cuadroRTF.SelStart
    If cont <> Len(cuadroRTF.Text) Then
        While l�neaActual = l�neaInicio And cont <= Len(cuadroRTF.Text)
            If Mid(cuadroRTF.Text, cont + 1, 1) <> Chr(10) And Mid(cuadroRTF.Text, cont + 1, 1) <> Chr(13) Then
                cadena = cadena + Mid(cuadroRTF.Text, cont + 1, 1)
            End If
            cont = cont + 1
            l�neaActual = cuadroRTF.GetLineFromChar(cont)
        Wend
        If Trim(cadena) = "" Then
            medioDelRengl�n = False
        Else
            medioDelRengl�n = True
        End If
    Else
        medioDelRengl�n = False
    End If
    
End Function

'Public Function oraci�nSiguiente(cuadroRTF As RichTextBox) As String
'    Dim cont As Long, cadena As String ', swEmpezando As Boolean
'    Dim l�neaInicio As Long, l�neaActual As Long ', TotalDeL�neasEnRTF As Long
'
'    l�neaInicio = cuadroRTF.GetLineFromChar(cuadroRTF.SelStart)
'
'    cont = cuadroRTF.SelStart
'
'    If cont <> 0 Then 'si no se est� ya al comienzo de la l�nea se busca el punto de inicio
'        l�neaActual = l�neaInicio
'        Do While l�neaActual = l�neaInicio
'            cont = cont - 1
'            l�neaActual = cuadroRTF.GetLineFromChar(cont)
'            If cont = 0 Then Exit Do
'        Loop
'    End If
'
'    cont = cont + 1
'    'If l�neaInicio = 0 Then swEmpezando = True
'    l�neaActual = l�neaInicio
'    While l�neaActual = l�neaInicio And cont <= Len(cuadroRTF.Text)
'        If Mid(cuadroRTF.Text, cont, 1) <> Chr(10) And Mid(cuadroRTF.Text, cont, 1) <> Chr(13) Then
'            cadena = cadena + Mid(cuadroRTF.Text, cont, 1)
'        End If
'        cont = cont + 1
'        l�neaActual = cuadroRTF.GetLineFromChar(cont)
'    Wend
'    oraci�nSiguiente = cadena
'End Function


Public Function corregirPalabra(palabra As String) As Boolean 'true palabra encontrada, false no encontrada
    'Dim palabra As String
    Dim num As Integer
    Dim palabraArchivo As String
    Dim archivo As Integer
        
    If palabra = "" Then 'si el par�metro es vac�o, se considera q est� correcto
        corregirPalabra = True
    Else
        If swAspellInstalado = False Then
            archivo = FreeFile
            
            If palabra = "" Or palabra = "," Or palabra = "." Or palabra = ":" Or palabra = "?" Then
                corregirPalabra = True
                Exit Function
            End If
            
            Open App.path + "\datos\palabras.txt" For Input As #archivo 'se abre la lista de palabras
            Do Until EOF(archivo) 'chequeamos si la palabra est� en la lista
                Line Input #archivo, palabraArchivo
                    If palabraArchivo = palabra Then 'si encuentra una palabra igual
                        Close #archivo
                        corregirPalabra = True
                        Exit Function
                    End If
                    
        '            If Asc(LCase(Left(palabraArchivo, 1))) > Asc(LCase(Left(palabra, 1))) Then 'si ya pas� la primera letra de la palabra
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
            If InStr(1, objPipe.Read, "*") Then 'si est� el asterisco en lo que devuelve aspell es que la palabra es correcta
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
        cadena = Mid(cuadroRTF.Text, cont, 1) 'si se est� en el comienzo del cuadro de texto
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
        Combo.AddItem Trim(cadena) 'se a�aden las materias al combo
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
'        MsgBox "soy el controlador de la funci�n llenarComboMaterias", , "Para mi creador"
        frmMsgBox.cadenaAMostrar = "Soy el controlador de la funci�n llenarComboMaterias"
        frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
    End If
End Sub


Public Sub regularVelocidadVoz()
    Dim aux As Integer, aux2 As Integer
    On Error GoTo manejoError
    If swSapi5 = True Then 'si se trabaja con sapi5
        If Not Voz Is Nothing Then Voz.Rate = velocidadVoz
    Else 'si se est� trabajando con sapi4
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
    frmMsgBox.cadenaAMostrar = "Soy el controlador del la funci�n regularVelocidadVoz. Error n�mero: " + Str(Err.Number) + ", descripci�n: " + Err.Description
    frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
    frmMsgBox.Show 1
    Exit Sub
End Sub


'Public Function obtenerVersi�nWindows() As String
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
'                obtenerVersi�nWindows = "windows 95"
'            Case 10
'                obtenerVersi�nWindows = "windows 98"
'            Case 90
'                obtenerVersi�nWindows = "windows millennium"
'                End Select
'            Case 2
'                Select Case .dwmajorversion
'                    Case 3
'                        obtenerVersi�nWindows = "windows nt 3.51"
'                    Case 4
'                        obtenerVersi�nWindows = "windows nt 4.0"
'                    Case 5
'                        If .dwminorversion = 0 Then
'                            obtenerVersi�nWindows = "windows 2000"
'                        Else
'                            obtenerVersi�nWindows = "windows xp"
'                        End If
'                End Select
'            Case Else
'                obtenerVersi�nWindows = "fall�"
'        End Select
'    End With
''    leerDatosSO (App.Path + "\datos\datosSO.lle")
'End Function

'Public Sub leerDatosSO(d�nde As String)
'    Dim miRegistro As osVersionInfo, sistemaRepetido As Boolean
'
'    sistemaRepetido = False
'    Open d�nde For Random As #1 Len = Len(miRegistro)
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
'        MsgBox "ac� ir�a el instalador porque es la primera vez que se abre en este sistema" 'shell instalador de acuerdo al sistema
'        Call guardarDatosSO(d�nde, osInfo) 'se a�ade el sistema a la lista de los guardados como que ya se instal�
'    End If
'End Sub

'Public Sub guardarDatosSO(d�nde As String, qu�Registro As osVersionInfo)
'    Dim archivolibre As Integer
'
'    archivolibre = FreeFile 'se abre el archivo para guardar los datos del sistema operativo
'    Open d�nde For Random As #archivolibre Len = Len(qu�Registro)
'    Put #archivolibre, 1, qu�Registro
'    Close #archivolibre
'End Sub

Public Sub centrarFormulario(qu�Form As Form)
    Dim centroXform As Single, centroYform As Single
    centroXform = (Screen.Width - qu�Form.ScaleWidth) / 2
    centroYform = (Screen.Height - qu�Form.ScaleHeight) / 2
    Call qu�Form.Move(centroXform, centroYform)
End Sub


Public Sub controlarCaracteresEspeciales(teclaPulsada As Integer, caja As TextBox)
    If teclaPulsada = 34 Or teclaPulsada = Asc("|") Or teclaPulsada = Asc("\") Or teclaPulsada = Asc("/") _
    Or teclaPulsada = Asc("?") Or teclaPulsada = Asc("*") Or teclaPulsada = Asc(">") Or teclaPulsada = Asc("<") _
    Or teclaPulsada = Asc(":") Or teclaPulsada = Asc(",") Or teclaPulsada = Asc(";") Or teclaPulsada = Asc(".") _
    Or teclaPulsada = Asc("-") Or teclaPulsada = Asc("_") Then
        caja.Text = Left(caja.Text, Len(caja.Text) - 1)
        frmMsgBox.cadenaAMostrar = "No se pueden escribir los siguientes signos en el nombre: . , ; - \ : / < > ? * | " + Chr(34) + "."
        frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
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

'Public Function conocerPalabraEnOraci�n(cuadroRTF As RichTextBox) As String
'    Dim cont As Long, cadena As String, rengl�nActual As Long
'    cont = cuadroRTF.SelStart
'    If Len(cuadroRTF.Text) <> cont Then 'si no se est� al final de la hoja
'        cont = cont + 1
'        cadena = Mid(cuadroRTF.Text, cont, 1) 'si se est� en el comienzo del cuadro de texto
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
''        if cadena="(" then cadena="abre par�ntesis"
'
''        If cuadroRTF.SelStart <> 0 Then 'si no se est� al principio de la hoja
'            conocerPalabraEnOraci�n = cadena '"est�s al final de tu hoja"
''        Else
''            If cadena = "." Or cadena = "," Or cadena = ";" Then
''                conocerPalabraEnOraci�n = "est�s al principio de la hoja, delante del signo. " + cadena
''            Else
''                conocerPalabraEnOraci�n = "est�s al principio de la hoja, delante de la palabra. " + cadena
''            End If
''        End If
''
''        If swLeerRenglones = True Then
''            If rengl�nActual <> rengl�nAnterior Then conocerPalabraEnOraci�n = conocerPalabraEnOraci�n & ". rengl�n " & Str(rengl�nActual + 1)
''        End If
'    Else
''        If cuadroRTF.Text <> "" Then
'            conocerPalabraEnOraci�n = "final de la hoja"
''        Else
''            conocerPalabraEnOraci�n = "no hay nada escrito, la hoja est� vac�a"
''        End If
'    End If
''    rengl�nAnterior = rengl�nActual
'End Function

Public Function SalirDelPrograma() As Boolean
    frmMsgBox.swMostrarCancelar = False
    frmMsgBox.cadenaAMostrar = "�Realmente quer�s salir del programa?"
    frmMsgBox.swS�No�Aceptar = True 'se elige que sea cuadro s�-no
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
    If swCuadernoAbierto = True Then Unload frmCuaderno 'si est� abierto el cuaderno se lo cierra
    If swLibroAbierto = True Then Unload frmLectorLibro 'si est� abierto el lector de libros, se lo cierra
    If swActividadAbierta = True Then Unload frmLectorActividad 'si est� abierto el lector de actividad, se lo cierra
    If frmReproductorM�sica.swEstoyAbierto = True Then Unload frmReproductorM�sica
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
'    Dim versi�nOS As String
'    versi�nOS = obtenerVersi�nWindows
'
'    Select Case versi�nOS
'        Case "windows 95"
'            'call ejecutar (App.Path + "\Ejecutables\instalador98.exe")
'            MsgBox "Ac� va el instalador del win 95 del programa porque es la primera vez que se abre"
'        Case "windows 98"
'            'call ejecutar ( App.Path + "\Ejecutables\instalador98.exe")
'            MsgBox "Ac� va el instalador del win 98 del programa porque es la primera vez que se abre"
'        Case "windows millennium"
'            MsgBox "Ac� va el instalador del millenium del programa porque es la primera vez que se abre"
'        Case "windows nt 3.51"
'            MsgBox "Ac� va el instalador del win 3.51 del programa porque es la primera vez que se abre"
'        Case "windows nt 4.0"
'            MsgBox "Ac� va el instalador del win nt del programa porque es la primera vez que se abre"
'        Case "windows 2000"
'            MsgBox "Ac� va el instalador del win 2000 del programa porque es la primera vez que se abre"
'        Case "windows xp"
'            MsgBox "Ac� va el instalador del xp del programa porque es la primera vez que se abre"
'        Case "fall�"
'    End Select
'
'
''    On Error GoTo manejoErrorInstalar
'''    Dim miUbicaci�n As String
'    Dim lectorRegistro, x
'    Set lectorRegistro = CreateObject("WScript.Shell")
''
''    Call RegFlash 'registrar flash
''    Call RegBot�nTransp 'registrar los botones transparentes
''
''    'instalar sapi5 si no est� instalada
''    x = lectorRegistro.regRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\MSMary\409")
''    If x <> "Microsoft Mary" Then
''        If versi�nOS = "windows xp" Then
''            Shell App.Path + "\ejecutables\Sapi5 (para XP).msi", vbNormalFocus
''        Else
''            Shell App.Path + "\ejecutables\Sapi5 (para Windows 98 Me 2000).msi", vbNormalFocus
''        End If
''    End If
'
''    'arreglar sapi4
''    'instalar sapi4 si no est� instalada
''    x = lectorRegistro.regRead("HKEY_LOCAL_MACHINE\SOFTWARE\Voice\TextToSpeech\Engine")
''    If x <> "Microsoft Mary" Then
''        Shell App.Path + "\ejecutables\Sapi4.exe", vbNormalFocus
''    End If
'
''    miUbicaci�n = App.Path + "\Mochila_Virtual.exe"
'    'lectorRegistro.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\", miUbicaci�n
'    lectorRegistro.RegWrite "HKEY_LOCAL_MACHINE\Software\ReyNegro-ReyBlanco\MochilaVirtual\", "1"
'    lectorRegistro.RegWrite "HKEY_LOCAL_MACHINE\Software\ReyNegro-ReyBlanco\MochilaVirtual\datos\", "0"
'    Set lectorRegistro = Nothing
''    Exit Sub
''manejoErrorInstalar:
''    MsgBox "soy el error de la funci�n instalar. Mi n�mero es " + Str(Err.Number) + ", y mi descripci�n es " + Err.Description
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
'    Dim versi�nOS As String
'    versi�nOS = obtenerVersi�nWindows
'
'    Select Case versi�nOS
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
'''    miUbicaci�n = App.Path + "\Mochila_Virtual.exe"
''    Set lectorRegistro = CreateObject("WScript.Shell")
''    x = lectorRegistro.regRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\MSMary")
''    If x <> "Microsoft Mary" Then
''        Shell App.Path + "\ejecutables\Sapi5.exe", vbNormalFocus
''    End If
''    Set lectorRegistro = Nothing
''    Exit Sub
''manejoErrorInstalar:
''    MsgBox "soy el error de la funci�n instalar. Mi n�mero es " + Err.Number + ", y mi descripci�n es " + Err.Description
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

Public Sub GuardarRecordatorio(qu�Recordatorio As Recordatorio)
    Dim archivolibre As Byte, contador As Integer, auxRecordatorio As Recordatorio, j As Byte

    On Error GoTo manejoError
    archivolibre = FreeFile 'se abre el archivo para guardar los datos de las partidas
    Open App.path + "\recordatorios\" + Trim(Right(Format(qu�Recordatorio.fecha, "dd/mm/yyyy"), 4)) + "\" + Trim(Str(Int(Mid(Format(qu�Recordatorio.fecha, "dd/mm/yyyy"), 4, 2)))) + "\recordatorios.gui" For Random As archivolibre Len = Len(qu�Recordatorio)
    contador = 0
    While Not EOF(archivolibre)
        contador = contador + 1
        Get archivolibre, contador, auxRecordatorio
    Wend
    qu�Recordatorio.�ndiceEnArchivo = contador
    Put archivolibre, contador, qu�Recordatorio
    Close archivolibre
    Exit Sub
manejoError:
    If Err.Number = 76 Then
        MkDir (App.path + "\recordatorios\" + Trim(Right(Format(qu�Recordatorio.fecha, "dd/mm/yyyy"), 4)))
        For j = 1 To 12
            MkDir (App.path + "\recordatorios\" + Trim(Right(Format(qu�Recordatorio.fecha, "dd/mm/yyyy"), 4)) + "\" + Trim(Str(j)))
        Next
        Resume Next
    End If
End Sub

Public Sub GuardarRecordatorioEnPosici�n(qu�Recordatorio As Recordatorio, posici�n As Long)
    Dim archivolibre As Byte, auxRecordatorio As Recordatorio, j As Byte

    On Error GoTo manejoError
    archivolibre = FreeFile 'se abre el archivo para guardar los datos de las partidas
    Open App.path + "\recordatorios\" + Trim(Right(Format(qu�Recordatorio.fecha, "dd/mm/yyyy"), 4)) + "\" + Trim(Str(Int(Mid(Format(qu�Recordatorio.fecha, "dd/mm/yyyy"), 4, 2)))) + "\recordatorios.gui" For Random As archivolibre Len = Len(qu�Recordatorio)
    qu�Recordatorio.�ndiceEnArchivo = posici�n
    Put archivolibre, posici�n, qu�Recordatorio
    Close archivolibre
    Exit Sub
manejoError:
    If Err.Number = 76 Then
        MkDir (App.path + "\recordatorios\" + Trim(Right(Format(qu�Recordatorio.fecha, "dd/mm/yyyy"), 4)))
        For j = 1 To 12
            MkDir (App.path + "\recordatorios\" + Trim(Right(Format(qu�Recordatorio.fecha, "dd/mm/yyyy"), 4)) + "\" + Trim(Str(j)))
        Next
        Resume Next
    End If
End Sub


Public Function sonidoForm(qu�Form As Byte) As String
    Dim archivo As String
    
    'corregir que las variables se cargen desde la configuraci�n
'    rutaM�sicaFormPrincipal = "principal.mid"
'    rutaM�sicaFormCuaderno = "cuaderno.mid"
'    rutaM�sicaFormActividad = "actividades.mid"
'    rutaM�sicaFormTareas = "tareas.mid"
'    rutaM�sicaFormLibros = "libros.mid"
'    rutaM�sicaFormAccesorios = "accesorios.mid"
    
    archivo = App.path + "\sonidos\formularios\"
    Select Case qu�Form
        Case formularios.principal
            If Trim(usuario.rutaM�sicaFormPrincipal) <> "" And Trim(Left(usuario.rutaM�sicaFormPrincipal, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaM�sicaFormPrincipal)
            Else
                archivo = archivo + "principal.mid"
            End If
        Case formularios.cuaderno
            If Trim(usuario.rutaM�sicaFormCuaderno) <> "" And Trim(Left(usuario.rutaM�sicaFormCuaderno, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaM�sicaFormCuaderno)
            Else
                archivo = archivo + "cuaderno.mid"
            End If
        Case formularios.actividades
            If Trim(usuario.rutaM�sicaFormActividad) <> "" And Trim(Left(usuario.rutaM�sicaFormActividad, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaM�sicaFormActividad)
            Else
                archivo = archivo + "actividades.mid"
            End If
        Case formularios.tareasAnt
            If Trim(usuario.rutaM�sicaFormTareas) <> "" And Trim(Left(usuario.rutaM�sicaFormTareas, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaM�sicaFormTareas)
            Else
                archivo = archivo + "tareas.mid"
            End If
        Case formularios.libros
            If Trim(usuario.rutaM�sicaFormLibros) <> "" And Trim(Left(usuario.rutaM�sicaFormLibros, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaM�sicaFormLibros)
            Else
                archivo = archivo + "libros.mid"
            End If
        Case formularios.accesorios
            If Trim(usuario.rutaM�sicaFormAccesorios) <> "" And Trim(Left(usuario.rutaM�sicaFormAccesorios, 1)) <> Chr(0) Then
                archivo = archivo + Trim(usuario.rutaM�sicaFormAccesorios)
            Else
                archivo = archivo + "accesorios.mid"
            End If
    End Select
    sonidoForm = archivo
End Function

Public Sub reproducirForm(qu�Form As Byte)
    If swM�sicaDeFondo = True Then
        frmOculto.swContinuarReproducci�n = False
        frmOculto.media.Stop
        frmOculto.media.FileName = sonidoForm(qu�Form)
        frmOculto.swContinuarReproducci�n = True
'        frmOculto.media.Play
    Else
        If frmOculto.media.PlayState = mpPlaying Then
            frmOculto.swContinuarReproducci�n = False
            frmOculto.media.Stop
        End If
   End If
End Sub

Public Sub cargarRecordatorios()
    Dim archivolibre As Byte, miRec As Recordatorio, mes As Byte, a�o As Integer ', d�a As Byte
    Dim contador As Integer, i As Integer
    On Error GoTo manejoError
    mes = Month(Date)
    a�o = Year(Date)
    archivolibre = FreeFile
    contador = 1 'se deja en blanco el primer recordatorioActivo. Si hay m�s de uno, es que hay recordatorios activos. Se eval�a en frmOculto
    Open App.path + "\recordatorios\" + Trim(Str(a�o)) + "\" + Trim(Str(mes)) + "\" + "recordatorios.gui" For Random As #archivolibre Len = Len(miRec)
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
                    If Right(Format(miRec.fecha, "dd/mm/yyyy"), 4) < Right(Format(Date, "dd/mm/yyyy"), 4) Then   'si el a�o es menor al actual
                        ReDim Preserve recordatoriosActivos(0 To contador)
                        recordatoriosActivos(contador) = miRec
                        contador = contador + 1
                    Else
                        If Right(Format(miRec.fecha, "dd/mm/yyyy"), 4) = Right(Format(Date, "dd/mm/yyyy"), 4) Then  'si el a�o es igual al actual
                            If Mid(Format(miRec.fecha, "dd/mm/yyyy"), 4, 2) < Mid(Format(Date, "dd/mm/yyyy"), 4, 2) Then 'si el mes es menor al actual
                                ReDim Preserve recordatoriosActivos(0 To contador)
                                recordatoriosActivos(contador) = miRec
                                contador = contador + 1
                            Else
                                If Mid(Format(miRec.fecha, "dd/mm/yyyy"), 4, 2) = Mid(Format(Date, "dd/mm/yyyy"), 4, 2) Then 'si el mes es igual al actual
                                    If Left(Format(miRec.fecha, "dd/mm/yyyy"), 2) < Left(Format(Date, "dd/mm/yyyy"), 2) Then 'si el d�a es anterior al actual
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

'Public Sub contarFormularios(acci�n As Boolean)
'    If acci�n = True Then formulariosAbiertos = formulariosAbiertos + 1
'    If acci�n = False Then formulariosAbiertos = formulariosAbiertos - 1
'    If formulariosAbiertos = 1 Then End
'End Sub

Public Function mensajeSalir(qu�Mensaje As String) As Boolean
    frmMsgBox.swMostrarCancelar = False
    frmMsgBox.cadenaAMostrar = qu�Mensaje
    frmMsgBox.swS�No�Aceptar = True 'se elige que sea cuadro s�-no
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

Public Function leerRegistro(ra�z As Long, clave As String, valor As String) As String
    Dim hClave As Long, longitud As Long, dato As String, ret As Long
    
    ret = RegOpenKeyEx(ra�z, clave, 0, KEY_ALL_ACCESS, hClave)
    ret = RegQueryValueEx(hClave, valor, 0, 0, 0, longitud)
    dato = String(longitud, 0)
    ret = RegQueryValueEx(hClave, valor, 0&, REG_SZ, ByVal dato, longitud)
    ret = RegCloseKey(hClave)
    leerRegistro = Left(dato, longitud - 1)
End Function

'Private Sub espaciosDelDisco(disco As String, espacioLibre As Currency, espacioTotal As Currency, espacioOcupado As Currency)
'    Dim devoluci�n As Long, SectoresporCluster As Long, BytesPorSector As Long, CantidadDeClustersLibres As Long, N�meroTotalClusters As Long
'    Static swDemasiadoDiscoParaM� As Boolean 'para controlar que no salte error si tiene un disco muy grande y desborde las variables que cuentan los bytes
'
'    On Error GoTo manejoErrorEspacio:
'    If swDemasiadoDiscoParaM� = False Then
'        devoluci�n = GetDiskFreeSpace(disco, SectoresporCluster, BytesPorSector, CantidadDeClustersLibres, N�meroTotalClusters)
'        espacioLibre = CantidadDeClustersLibres * SectoresporCluster * BytesPorSector
'        espacioLibre = (espacioLibre / 1024) / 1024
'        espacioTotal = N�meroTotalClusters * SectoresporCluster * BytesPorSector
'        espacioTotal = (espacioTotal / 1024) / 1024
'        espacioOcupado = espacioTotal - espacioLibre
'    End If
'    Exit Sub
'manejoErrorEspacio:
'    swDemasiadoDiscoParaM� = True
'    Exit Sub
'End Sub
'
'Public Sub chequearEspacioEnDisco(disco As String)
'    Dim libre As Currency, total As Currency, ocupado As Currency, swMostrarCuadro As Boolean
'    swMostrarCuadro = False
'    Call espaciosDelDisco(disco, libre, total, ocupado)
'    If libre < 20 And libre > 10 Then
'        frmMsgBox.cadenaAMostrar = "Se est� acabando el espacio libre que queda en el disco en que est� instalada la mochila. Considere liberar espacio a la brevedad."
'        swMostrarCuadro = True
'    ElseIf libre < 10 And libre >= 1 Then
'        frmMsgBox.cadenaAMostrar = "Queda muy poco espacio libre en el disco en que est� instalada la mochila. Libere espacio o instale la mochila en otro disco."
'        swMostrarCuadro = True
'    ElseIf libre <= 1 Then
'        frmMsgBox.cadenaAMostrar = "Queda solamente 1 MB libre en el disco en que est� instalada la mochila. Libere urgentemente espacio o instale la mochila en otro disco."
'        swMostrarCuadro = True
'    End If
'    If swMostrarCuadro = True Then
'        frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
'        frmMsgBox.Show 1
'    End If
'End Sub

'Public Sub alwaysOnTop(formulario As Form, estado As Boolean)
'    Dim banderas As Long, ret As Long
'    'para que no cambie el tama�o ni la posici�n
'    banderas = SWP_NOMOVE Or SWP_NOSIZE
'    If estado Then
'        ret = SetWindowPos(formulario, HWND_TOPMOST, 0, 0, 0, 0, banderas)
'    Else
'        ret = SetWindowPos(formulario, HWND_NOTOPMOST, 0, 0, 0, 0, banderas)
'    End If
'End Sub

Public Sub ejecutar(aplicaci�n As String)
    Dim handleProceso As Long
    Dim activa As Long
    Dim ret As Long

    handleProceso = OpenProcess(PROCESS_QUERY_INFORMATION, 0, Shell(aplicaci�n, 1))
    Do
        ret = GetExitCodeProcess(handleProceso, activa)
        DoEvents
    Loop While activa = STILL_ACTIVE
End Sub

Public Function qu�LetraSeApret�(n�meroLetra As Integer) As String
    Dim auxString As String ', cadena As String

    Select Case UCase(Chr(n�meroLetra))
        Case " "
            auxString = " espacio"
        Case "1"
            auxString = " uno"
        Case "�"
            auxString = " acento agudo"
        Case "�"
            auxString = " di�resis"
        Case "`"
            auxString = " acento grave"
        Case "^"
            auxString = " acento circunflejo"
        Case "+"
            auxString = " m�s"
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
        Case "�"
            auxString = " abre exclamaci�n"
        Case "!"
            auxString = " cierra exclamaci�n"
        Case "�"
            auxString = " abre pregunta"
        Case "?"
            auxString = " cierra pregunta"
        Case "$"
            auxString = " signo pesos"
        Case "&"
            auxString = " anpersand"
        Case "\"
            auxString = " barra diagonal inversa"
        Case "�"
            auxString = " ordinal masculino"
        Case "�"
            auxString = " ordinal femenino"
        Case "%"
            auxString = " porciento"
        Case "("
            auxString = " abre par�ntesis"
        Case ")"
            auxString = " cierra par�ntesis"
        Case "{"
            auxString = " abre llave"
        Case "}"
            auxString = " cierra llave"
        Case "�"
            auxString = " a con acento"
        Case "�"
            auxString = " e con acento"
        Case "�"
            auxString = " i con acento"
        Case "�"
            auxString = " o con acento"
        Case "�"
            auxString = " u con acento"
        Case "�"
            auxString = " u con di�resis"
        Case "B"
            auxString = " b� larga"
        Case "C"
            auxString = " c�"
        Case "D"
            auxString = " d�"
        Case "F"
            auxString = " �fe"
        Case "G"
            auxString = " g�"
        Case "H"
            auxString = " �che"
        Case "J"
            auxString = " j�ta"
        Case "K"
            auxString = " k�"
        Case "L"
            auxString = " �le"
        Case "M"
            auxString = " �me"
        Case "N"
            auxString = " �ne"
        Case "�"
            auxString = " ��e"
        Case "P"
            auxString = " p�"
        Case "Q"
            auxString = " c�"
        Case "R"
            auxString = " �rre"
        Case "S"
            auxString = " �se"
        Case "T"
            auxString = " t�"
        Case "V"
            auxString = " v� corta"
        Case "W"
            auxString = " doble b�"
        Case "X"
            auxString = " �quis"
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
            auxString = Chr(n�meroLetra)
    End Select
    
    'cadena = auxString
    If (n�meroLetra >= 65 And n�meroLetra <= 90) Then auxString = auxString + " may�scula"
    
    If n�meroLetra = 9 Then auxString = "avanzando hacia adelante un salto" 'si es un tab
    qu�LetraSeApret� = auxString
End Function


Public Function traducirParaBorrar(letra As String) As String
    Dim auxString As String
    Select Case UCase(letra)
        Case " "
            auxString = " el espacio"
        Case "�"
            auxString = " el acento agudo "
        Case "�"
            auxString = " la di�resis"
        Case "`"
            auxString = " el acento grave"
        Case "^"
            auxString = " el acento circunflejo"
        Case "&"
            auxString = " el ampersand"
        Case "+"
            auxString = " el m�s"
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
        Case "�"
            auxString = " el abre exclamaci�n"
        Case "!"
            auxString = " el cierra exclamaci�n"
        Case "�"
            auxString = " el abre pregunta"
        Case "?"
            auxString = " el cierra pregunta"
        Case "$"
            auxString = " el signo pesos"
        Case "%"
            auxString = " el porciento"
        Case "("
            auxString = " el abre par�ntesis"
        Case ")"
            auxString = " el cierra par�ntesis"
        Case "{"
            auxString = " el abre llave"
        Case "}"
            auxString = " el cierra llave"
        Case "�"
            auxString = " la a con acento"
        Case "�"
            auxString = " la e con acento"
        Case "�"
            auxString = " la i con acento"
        Case "�"
            auxString = " la o con acento"
        Case "�"
            auxString = " la u con acento"
        Case "B"
            auxString = " la b� larga"
        Case "C"
            auxString = " la c�"
        Case "D"
            auxString = " la d�"
        Case "F"
            auxString = " la �fe"
        Case "G"
            auxString = " la g�"
        Case "H"
            auxString = " la �che"
        Case "J"
            auxString = " la j�ta"
        Case "K"
            auxString = " la k�"
        Case "L"
            auxString = " la �le"
        Case "M"
            auxString = " la �me"
        Case "N"
            auxString = " la �ne"
        Case "�"
            auxString = " la ��e"
        Case "P"
            auxString = " la p�"
        Case "Q"
            auxString = " la c�"
        Case "R"
            auxString = " la �rre"
        Case "S"
            auxString = " la �se"
        Case "T"
            auxString = " la t�"
        Case "V"
            auxString = " la v� corta"
        Case "W"
            auxString = " la doble b�"
        Case "X"
            auxString = " la �quis"
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
    
    If (Asc(letra) >= 65 And Asc(letra) <= 90) Then auxString = auxString + " may�scula"
    traducirParaBorrar = auxString
End Function


Sub evaluarSelecci�n(cuadroRTF As RichTextBox, control As Boolean, Shift As Boolean, teclaQueSelecciona As Byte)    'As String
    Static LenSelecci�nAnterior As Currency
    Dim estadoSelecci�n As Byte, cadena As String
    
    If Len(cuadroRTF.SelText) < LenSelecci�nAnterior Then estadoSelecci�n = selecci�n.disminuy�
    If Len(cuadroRTF.SelText) > LenSelecci�nAnterior Then estadoSelecci�n = selecci�n.creci�
    If Len(cuadroRTF.SelText) = LenSelecci�nAnterior Then estadoSelecci�n = selecci�n.igual
    
    If estadoSelecci�n <> selecci�n.igual Then
        If cuadroRTF.SelText = "" And estadoSelecci�n = selecci�n.disminuy� Then cadena = "quitando la selecci�n"
        If cuadroRTF.SelText = "" And LenSelecci�nAnterior <> 0 And teclaQueSelecciona = tecla.borrar Then cadena = "borrando la selecci�n"
        '++++++++++++++++++++++++++++++++++++++++++++
        'se selecciona con shift y control apretadas
        If control And Shift Then 'si se va seleccionando texto con control y shift
            If (teclaQueSelecciona = tecla.flechaDerecha Or teclaQueSelecciona = tecla.flechaIzquierda) Then   'seleccionado por palabras
                If cuadroRTF.Text <> "" Then
                    cadena = "texto seleccionado: " + cuadroRTF.SelText
                Else
                    cadena = "no se puede seleccionar porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.inicio Then
                If cuadroRTF.Text <> "" Then
                    If cuadroRTF.SelText <> "" Then
                        If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionado todo el texto desde donde estabas hasta el principio de la hoja"
                        If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n desde donde estabas hasta el principio de la hoja"
                    Else
                        cadena = "se ha sacado la selecci�n del texto"
                    End If
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.fin Then
                If cuadroRTF.Text <> "" Then
                    If cuadroRTF.SelText <> "" Then
                        If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionado todo el texto desde donde estabas hasta el final de la hoja"
                        If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n desde donde estabas hasta el final de la hoja"
                        'cadena =  "seleccionado todo el texto desde donde estabas hasta el final de la hoja"
                    Else
                        cadena = "se ha dejado de seleccionar todo el texto"
                    End If
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.flechaArriba Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionado desde donde estabas hasta el principio del p�rrafo"
                    If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n desde donde estabas hasta el principio del p�rrafo"
                    'cadena =  "seleccionado desde donde estabas hasta el principio del p�rrafo"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.flechaAbajo Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionado desde donde estabas hasta el final del p�rrafo"
                    If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n desde donde estabas hasta el final del p�rrafo"
                    'cadena =  "seleccionado desde donde estabas hasta el final del p�rrafo"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.avanceP�gina Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionando varios renglones desde donde estabas hacia arriba en la hoja"
                    If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n varios renglones desde donde estabas hacia arriba en la hoja"
                    'cadena =  "seleccionando varios renglones desde donde estabas hacia arriba en la hoja"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.retrocesoP�gina Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionando varios renglones desde donde estabas hacia abajo en la hoja"
                    If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n varios renglones desde donde estabas hacia abajo en la hoja"
                    'cadena =  "seleccionando varios renglones desde donde estabas hacia abajo en la hoja"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
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
                    If Left(cuadroRTF.SelText, 1) = Chr(13) Then cadena = cadena + " bajada de l�nea "
                    
                    cadena = cadena + cuadroRTF.SelText
                Else
                    cadena = "no se puede seleccionar porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.inicio Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionando el texto hasta el principio del rengl�n"
                    If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n, queda seleccionado: " + cuadroRTF.SelText
                    'cadena =  "seleccionando el texto hasta el principio del rengl�n"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.fin Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionando el texto hasta el final del rengl�n"
                    If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n, queda seleccionado: " + cuadroRTF.SelText
                    'cadena =  "seleccionando el texto hasta el final del rengl�n"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.flechaArriba Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionado desde donde estabas hasta la misma posici�n del rengl�n superior"
                    If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n desde donde estabas hasta la misma posici�n del rengl�n superior"
                    'cadena =  "seleccionado desde donde estabas hasta la misma posici�n del rengl�n superior"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.flechaAbajo Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionado desde donde estabas hasta la misma posici�n del rengl�n inferior"
                    If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n desde donde estabas hasta la misma posici�n del rengl�n inferior"
                    'cadena =  "seleccionado desde donde estabas hasta la misma posici�n del rengl�n inferior"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.avanceP�gina Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionando varios renglones desde donde estabas hacia arriba en la hoja"
                    If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n varios renglones desde donde estabas hacia arriba en la hoja"
                    'cadena =  "seleccionando varios renglones desde donde estabas hacia arriba en la hoja"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
            
            If teclaQueSelecciona = tecla.retrocesoP�gina Then
                If cuadroRTF.Text <> "" Then
                    If estadoSelecci�n = selecci�n.creci� Then cadena = "seleccionando varios renglones desde donde estabas hacia abajo en la hoja"
                    If estadoSelecci�n = selecci�n.disminuy� Then cadena = "disminuyendo la selecci�n varios renglones desde donde estabas hacia abajo en la hoja"
                    'cadena =  "seleccionando varios renglones desde donde estabas hacia abajo en la hoja"
                Else
                    cadena = "no se puede seleccionar porque porque la hoja est� vac�a"
                End If
            End If
        End If
    
        
        If control And teclaQueSelecciona = tecla.a Then 'seleccionar todo el texto
            If cuadroRTF.Text <> "" Then
                cadena = "se seleccion� todo el texto de la hoja"
            Else
                cadena = "No se puede seleccionar porque la hoja est� vac�a"
            End If
        End If
        
        If cadena <> "" Then Decir cadena
    End If
    
    LenSelecci�nAnterior = Len(cuadroRTF.SelText)
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
        MkDir (App.path + "\trabajos\" + cadena + "\evaluaciones") 'carpeta para poner evaluaciones falsas por si los pap�s quieren modificar una evaluaci�n ya hecha ];)
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
    Dim d�a1 As Byte, d�a2 As Byte, mes1 As Byte, mes2 As Byte, a�o1 As Integer, a�o2 As Integer
    
    d�a1 = Left(Format(fecha1, "dd/mm/yyyy"), 2)
    d�a2 = Left(Format(fecha2, "dd/mm/yyyy"), 2)
    
    mes1 = Mid(Format(fecha1, "dd/mm/yyyy"), 4, 2)
    mes2 = Mid(Format(fecha2, "dd/mm/yyyy"), 4, 2)
    
    a�o1 = Right(Format(fecha1, "dd/mm/yyyy"), 4)
    a�o2 = Right(Format(fecha2, "dd/mm/yyyy"), 4)
    
    If a�o1 > a�o2 Then
        compararFechas = comparaci�n.primeroMayor
        Exit Function
    End If
    
    If a�o1 < a�o2 Then
        compararFechas = comparaci�n.primeroMenor
        Exit Function
    End If
    
    If a�o1 = a�o2 Then
        If mes1 > mes2 Then
            compararFechas = comparaci�n.primeroMayor
            Exit Function
        End If
        
        If mes1 < mes2 Then
            compararFechas = comparaci�n.primeroMenor
            Exit Function
        End If
        
        If mes1 = mes2 Then
            If d�a1 > d�a2 Then
                compararFechas = comparaci�n.primeroMayor
                Exit Function
            End If
            
            If d�a1 < d�a2 Then
                compararFechas = comparaci�n.primeroMenor
                Exit Function
            End If
            
            If d�a1 = d�a2 Then
                compararFechas = comparaci�n.iguales
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
        compararHora = comparaci�n.primeroMayor
        Exit Function
    End If
    
    If hora1 < hora2 Then
        compararHora = comparaci�n.primeroMenor
        Exit Function
    End If
    
    If hora1 = hora2 Then
        If minutos1 > minutos2 Then
            compararHora = comparaci�n.primeroMayor
            Exit Function
        End If
    
        If minutos1 < minutos2 Then
            compararHora = comparaci�n.primeroMenor
            Exit Function
        End If
        
        If minutos1 = minutos2 Then
            compararHora = comparaci�n.iguales
            Exit Function
        End If
    End If
End Function


' ----------------------------------------------------------------------------------------
' \\ --   Subrutina para cargar en forma din�mica el men� de opciones
' ----------------------------------------------------------------------------------------
Public Sub Cargar_Menu(El_SubMenu As Object, palabras() As String)
    Dim i As Integer
    
    ' -- Por si hay Submenu cargados, los descarga a todos
    For i = 1 To El_SubMenu.Count - 1
        Unload El_SubMenu(i)
    Next
    
    If UBound(palabras) <> 0 Then 'si se manda alguna palabra para el men�
        For i = 0 To UBound(palabras)
            ' -- Establecer el caption del primer SubMenu
            El_SubMenu(El_SubMenu.Count - 1).Caption = palabras(i)
            
            ' -- Crear otro men� dinamicamente mediante Load
            If i <> UBound(palabras) Then Load El_SubMenu(El_SubMenu.Count)
        Next
    Else
        El_SubMenu(El_SubMenu.Count - 1).Caption = "No s� qu� palabra sugerirte"
    End If
End Sub


Public Sub Cargar_Men�_En_Lista(lista As ListBox, palabras() As String)
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
    
    'se busca ve cu�nto mide la palabra
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
    
    If cont2 = Len(cuadroRTF.Text) Then ' cont = 1 And cont2 = Len(cuadroRtf.Text) Then 'si hay una sola palabra escrita o si es la �ltima palabra
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
    Dim posici�n As Integer, cantidadDevoluci�n As Integer, cadenaAux As String
    Dim temp() As String, contador As Integer, i As Integer ', tempAux(0 To 10) As String
    
    ReDim temp(0 To 0) 'para que no d� error si no hay palabras que coincidan
    posici�n = InStr(1, cadena, "*") 'se ve si la palabra fue correcta
    If posici�n = 0 Then
        posici�n = InStr(1, cadena, "&") 'se dejan s�lo las palabras devueltas
        cadenaAux = Trim(Right(cadena, Len(cadena) - posici�n))
        
        posici�n = InStr(1, cadenaAux, " ") 'se busca el largo del array que devuelve aspell
        cadenaAux = Right(cadenaAux, Len(cadenaAux) - posici�n)
        posici�n = InStr(1, cadenaAux, " ")
        If IsNumeric(Left(cadenaAux, posici�n)) Then 'si hay alguna devoluci�n para la palabra
            cantidadDevoluci�n = Int(Left(cadenaAux, posici�n))
            ReDim Preserve temp(0 To cantidadDevoluci�n - 1) 'se estira el array seg�n la cantidad de palabras que devulve aspell
            
            If cantidadDevoluci�n > 0 Then
                posici�n = InStr(1, cadenaAux, ":") 'se busca dejar s�lo las palabras sugeridas
                cadenaAux = Trim(Right(cadenaAux, Len(cadenaAux) - posici�n))
            End If
            
            posici�n = InStr(1, cadenaAux, "�") 'se quitan los caracteres �
            Do While posici�n
                cadenaAux = Left(cadenaAux, posici�n - 1) + "�" + Right(cadenaAux, Len(cadenaAux) - posici�n)
                posici�n = InStr(1, cadenaAux, "�")
            Loop
                    
            posici�n = InStr(1, cadenaAux, "�") 'se quitan los caracteres �
            Do While posici�n
                cadenaAux = Left(cadenaAux, posici�n - 1) + "�" + Right(cadenaAux, Len(cadenaAux) - posici�n)
                posici�n = InStr(1, cadenaAux, "�")
            Loop
                   
           'Debug.Print cadenaAux
           
            posici�n = InStr(1, cadenaAux, "�") 'se quitan los caracteres �
            Do While posici�n
                cadenaAux = Left(cadenaAux, posici�n - 1) + "�" + Right(cadenaAux, Len(cadenaAux) - posici�n)
                posici�n = InStr(1, cadenaAux, "�")
            Loop
            
            posici�n = InStr(1, cadenaAux, ",") 'se llena el array con las palabras devueltas
            contador = 0
            Do While posici�n
                If contador >= UBound(temp) Then Exit Do 'nos aseguramos que el contador no supere el l�mite del array as� no da error
                temp(contador) = controlar_A_Acentuada(Trim(Left(cadenaAux, posici�n - 1)))
                cadenaAux = Right(cadenaAux, Len(cadenaAux) - posici�n)
                posici�n = InStr(1, cadenaAux, ",")
                contador = contador + 1
            Loop
            temp(contador) = controlar_A_Acentuada(Trim(Left(cadenaAux, Len(cadenaAux) - 4))) 'se carga la �ltima palabra
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
    esSigno = True 'si no se devolvi� falso, es que son s�lo signos. Devuelve verdadero
End Function

Public Function separarEnLetras(palabra As String) As String
    Dim i As Integer, cadenaTemp As String
    
    For i = 1 To Len(palabra)
        cadenaTemp = cadenaTemp + Mid(palabra, i, 1) + ". "
    Next
    separarEnLetras = controlarDeletreo(cadenaTemp)
End Function

Public Function controlarDeletreo(cadena As String) As String
    Dim car�cter(26) As String, posici�nCaracter As Long, cadenaFinal As String
    Dim i As Byte, swEntr�AlFor As Boolean, swYaEmpez� As Boolean
    
    cadena = LCase(cadena)
    
    If cadena <> "" Then
        car�cter(0) = "m. "
        car�cter(1) = "s. "
        car�cter(2) = "l. "
        car�cter(3) = "h. "
        car�cter(4) = "p. "
        car�cter(5) = "n. "
        car�cter(6) = "�. "
        car�cter(7) = "�. "
        car�cter(8) = "�. "
        car�cter(9) = "�. "
        car�cter(10) = "�. "
        car�cter(11) = "�. "
        car�cter(12) = "�. "
        car�cter(13) = "�. "
        car�cter(14) = "�. "
        car�cter(15) = "�. "
        car�cter(16) = "�. "
        car�cter(17) = "�. "
        car�cter(18) = "�. "
        car�cter(19) = "�. "
        car�cter(20) = "�. "
        car�cter(21) = "g. "
        car�cter(22) = "u. "
        car�cter(23) = "d. "
        car�cter(24) = "b. "
        car�cter(25) = "v. "
        car�cter(26) = "y. "

        swEntr�AlFor = False
        swYaEmpez� = False
        
        For i = 0 To UBound(car�cter)
            If swYaEmpez� = False Then
                posici�nCaracter = InStr(1, cadena, car�cter(i))
            Else
                posici�nCaracter = InStr(1, cadenaFinal, car�cter(i))
            End If
            
            Do While posici�nCaracter <> 0
                If swEntr�AlFor = False Then
                    cadenaFinal = corregirCadena(cadena, posici�nCaracter, car�cter(i))
                Else
                    cadenaFinal = corregirCadena(cadenaFinal, posici�nCaracter, car�cter(i))
                End If
                posici�nCaracter = InStr(1, cadenaFinal, car�cter(i))
                swEntr�AlFor = True
                swYaEmpez� = True
            Loop
        Next
        
        If cadenaFinal = "" Then cadenaFinal = cadena
        controlarDeletreo = cadenaFinal
    End If
End Function



Public Function controlar_A_Acentuada(devoluci�nAspell As String) As String
    Dim Pos As Integer, palabraCambiada As String
    
    Pos = InStr(1, devoluci�nAspell, "�") 'se quitan los caracteres �
    If Pos <> 0 Then 'si hay alguna �
        Do While Pos
            palabraCambiada = Left(devoluci�nAspell, Pos - 1) + "�" + Right(devoluci�nAspell, Len(devoluci�nAspell) - Pos)
            Pos = InStr(Pos + 1, devoluci�nAspell, "�")
                    
            'se ve si la palabra cambiada es correcta para dejar de cambiar �
            If corregirPalabra(palabraCambiada) Then Exit Do
        Loop
        
        'se ve si la palabra cambiada es correcta
        If corregirPalabra(palabraCambiada) Then 'si la palabra es correcta
            controlar_A_Acentuada = palabraCambiada
        Else
            controlar_A_Acentuada = devoluci�nAspell
        End If
    Else 'si no hay � en la palabra, se la devuelve sin cambio
        controlar_A_Acentuada = devoluci�nAspell
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
    
    'se busca ve cu�nto mide la palabra
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


Public Function buscarEntrada(qu�Cadena As String, diccionarioElegido As String) As String
    Dim cadenaDiccionario As String
    Dim palabraEncontrada As Boolean
    Dim archivolibre As Byte 'el manejador del diccionario
    Dim palabraNuevaL�nea As String
    Dim posici�nDosPuntos As Integer
    
    If existeCarpeta(App.path + "\diccionarios\" + diccionarioElegido) Then
        'si existe el diccionario
        palabraEncontrada = False
        buscarEntrada = ""
        
        archivolibre = FreeFile
        Open App.path + "\diccionarios\" + diccionarioElegido For Input As archivolibre
        Do While Not EOF(archivolibre)   ' Repite el bucle hasta el final del archivo.
            Line Input #archivolibre, cadenaDiccionario ' Lee el car�cter en dos variables.
            If palabraEncontrada = False Then 'se ve si la palabra est� en el diccionario
                posici�nDosPuntos = InStr(1, cadenaDiccionario, ":") - 1
                If posici�nDosPuntos > 0 Then
                    If LCase(Trim(Left(cadenaDiccionario, posici�nDosPuntos))) = LCase(Trim(qu�Cadena)) Then
                    'esto sirve para buscar la coincidencia letra por letra de lo que se escribe -> 'If LCase(Trim(Left(cadenaDiccionario, Len(qu�Cadena)))) = LCase(Trim(qu�Cadena)) Then
                        buscarEntrada = cadenaDiccionario
                        palabraEncontrada = True
                    End If
                End If
            Else 'si se encontr� la palabra, se ve si la definici�n sigue en el pr�ximo rengl�n
                palabraNuevaL�nea = Left(cadenaDiccionario, InStr(1, cadenaDiccionario, " "))
                'chequear si est� en may�sculas, si es as�, salir de la funci�n, sin�, sumar la cadena a la ya obtenida
                If UCase(palabraNuevaL�nea) = palabraNuevaL�nea Then
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


'Public Function matrizInicializada(qu�Matriz() As Object) As Boolean 'para controlar que una matriz din�mica ya est� dimensionada
'    On Error GoTo error:
'    If UBound(qu�Matriz) Then matrizInicializada = True
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

