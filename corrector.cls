VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "corrector_ortografía"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type palabra
    nombre As String
    repetición As Long
End Type

Private array_palabras() As palabra
Private swHayDiccionario As Boolean

Private Function controlarPalabra(palabras() As String) As String() 'revisa el diccionario para ver si la palabra está en la lista
    Dim i As Double, j As Double, temp() As palabra, contador As Double
    Dim arrayTemp() As String, array_devolución(0 To 5) As String
       
    ReDim temp(0 To 0) 'para que no dé error si no hay palabras que coincidan
    contador = 0
    
    For j = 0 To UBound(palabras)
        For i = 0 To UBound(array_palabras) '-1
            If palabras(j) = array_palabras(i).nombre Then
                ReDim Preserve temp(0 To contador) As palabra
                temp(contador) = array_palabras(i)
                contador = contador + 1
            End If
        Next
    Next
    
    temp = ordenar_array(temp)
    ReDim Preserve arrayTemp(0 To UBound(temp))
    
    For i = 0 To UBound(temp)
        arrayTemp(i) = temp(i).nombre
    Next
    
    If UBound(arrayTemp) > 5 Then 'se cargan sólamente los 5 primeros resultados
        For i = 0 To 5
            array_devolución(i) = arrayTemp(i)
        Next
        controlarPalabra = array_devolución
    Else
        controlarPalabra = arrayTemp
    End If
End Function

Public Function Controlar_Un_Error(palabra As String) As String()
    'Dim temp(0 To 0) As String
    If swHayDiccionario = True Then
        'temp(0) = palabra
        Controlar_Un_Error = controlarPalabra(unirArrays(inserción(palabra), transposición(palabra), ausencia(palabra), alteración(palabra)))
    Else
        MsgBox "No se ha especificado el archivo del diccionario en la propiedad Archivo_de_diccionario. No habrá corrector ortográfico.", , "Atención!"
    End If
End Function

Public Function Controlar_Dos_Errores(palabra As String) As String()
    Dim temp() As String, i As Double, temp2() As String, temp1() As String, j As Double, hastaDónde As Double
    If swHayDiccionario = True Then
        temp() = unirArrays(inserción(palabra), transposición(palabra), ausencia(palabra), alteración(palabra))
        hastaDónde = UBound(temp)
        
        For i = 0 To hastaDónde
            'temp1 = unirArrays(inserción(temp(i)), transposición(temp(i)), ausencia(temp(i)), alteración(temp(i)))
            temp1 = unir_2_Arrays(inserción(temp(i)), transposición(temp(i)))
            temp2 = unir_2_Arrays(temp, temp1)
            temp = temp2
        Next
        
        
        
        Controlar_Dos_Errores = controlarPalabra(temp2)
    Else
        MsgBox "No se ha especificado el archivo del diccionario en la propiedad Archivo_de_diccionario. No habrá corrector ortográfico.", , "Atención!"
    End If
End Function

Private Function inserción(palabra As String) As String() 'hay una letra de más
    Dim i As Long, j As Integer, resultados() As String, temp() As String
    Dim contador As Long, swEscribir As Boolean
    
    contador = 0
    'For i = 0 To UBound(palabra)
        For j = 0 To Len(palabra) - 1
            ReDim Preserve resultados(0 To contador)
            resultados(contador) = Mid(palabra, 1, j) + Mid(palabra, j + 2, Len(palabra))
            contador = contador + 1
        Next
    'Next
    
    'revisar que no haya palabras repetidas en el array
    contador = 0
    For i = 0 To UBound(resultados)
        swEscribir = True
        For j = i + 1 To UBound(resultados)
            If resultados(i) = resultados(j) Then
                swEscribir = False
                Exit For
            End If
        Next
        
        If swEscribir = True Then
            ReDim Preserve temp(0 To contador)
            temp(contador) = resultados(i)
            contador = contador + 1
        End If
    Next
    
    inserción = temp 'se devuelve el array
End Function

Private Function transposición(palabra As String) As String()
    Dim i As Long, j As Integer, resultados() As String, temp() As String
    Dim contador As Long, swEscribir As Boolean
    
    contador = 0
    'For i = 0 To UBound(palabra)
        For j = 0 To Len(palabra) - 1 '2
            ReDim Preserve resultados(0 To contador)
            resultados(contador) = Mid(palabra, 1, j) + Mid(palabra, j + 2, 1) + Mid(palabra, j + 1, 1) + Mid(palabra, j + 3, Len(palabra))
            contador = contador + 1
            'Debug.Print resultados(j)
        Next
    'Next
        
    'revisar que no haya palabras repetidas en el array
    contador = 0
    For i = 0 To UBound(resultados)
        swEscribir = True
        For j = i + 1 To UBound(resultados)
            If resultados(i) = resultados(j) Then
                swEscribir = False
                Exit For
            End If
        Next
        
        If swEscribir = True Then
            ReDim Preserve temp(0 To contador)
            temp(contador) = resultados(i)
            contador = contador + 1
        End If
    Next
        
    transposición = temp
End Function


Private Function ausencia(palabra As String) As String() 'falta una letra
    Dim i As Long, j As Integer, resultados() As String, temp() As String, z As Double
    Dim contador As Long, swEscribir As Boolean, alfabeto(0 To 33) As String
    
    j = 0
    For i = 97 To 122 'se carga el alfabeto
        alfabeto(j) = Chr(i)
        j = j + 1
    Next
    
    alfabeto(j) = Chr(225)
    j = j + 1
    alfabeto(j) = Chr(233)
    j = j + 1
    alfabeto(j) = Chr(237)
    j = j + 1
    alfabeto(j) = Chr(237)
    j = j + 1
    alfabeto(j) = Chr(241)
    j = j + 1
    alfabeto(j) = Chr(243)
    j = j + 1
    alfabeto(j) = Chr(250)
    j = j + 1
    alfabeto(j) = Chr(252)
    
    contador = 0
    'For z = 0 To UBound(palabra)
        For j = 0 To UBound(alfabeto)
            For i = 0 To Len(palabra)
                ReDim Preserve resultados(0 To contador)
                resultados(contador) = Mid(palabra, 1, i) + alfabeto(j) + Mid(palabra, i + 1, Len(palabra))
                contador = contador + 1
                'Debug.Print resultados(contador)
            Next
        Next
    'Next
    
    'revisar que no haya palabras repetidas en el array
    contador = 0
    For i = 0 To UBound(resultados)
        swEscribir = True
        For j = i + 1 To UBound(resultados)
            If resultados(i) = resultados(j) Then
                swEscribir = False
                Exit For
            End If
        Next

        If swEscribir = True Then
            ReDim Preserve temp(0 To contador)
            temp(contador) = resultados(i)
            contador = contador + 1
        End If
    Next
    
    ausencia = temp
End Function

Private Function alteración(palabra As String) As String() 'cambiar una letra por otra
    Dim i As Long, j As Integer, resultados() As String, temp() As String, z As Double
    Dim contador As Long, swEscribir As Boolean, alfabeto(0 To 33) As String
    
    j = 0
    For i = 97 To 122 'se carga el alfabeto
        alfabeto(j) = Chr(i)
        j = j + 1
    Next
    
    alfabeto(j) = Chr(225)
    j = j + 1
    alfabeto(j) = Chr(233)
    j = j + 1
    alfabeto(j) = Chr(237)
    j = j + 1
    alfabeto(j) = Chr(237)
    j = j + 1
    alfabeto(j) = Chr(241)
    j = j + 1
    alfabeto(j) = Chr(243)
    j = j + 1
    alfabeto(j) = Chr(250)
    j = j + 1
    alfabeto(j) = Chr(252)
    
    contador = 0
    'For z = 0 To UBound(palabra)
        For j = 0 To UBound(alfabeto)
            For i = 0 To Len(palabra) - 1
                ReDim Preserve resultados(0 To contador)
                resultados(contador) = Mid(palabra, 1, i) + alfabeto(j) + Mid(palabra, i + 2, Len(palabra))
                contador = contador + 1
                'Debug.Print resultados(contador - 1)
            Next
        Next
    'Next
    
    'revisar que no haya palabras repetidas en el array
    contador = 0
    For i = 0 To UBound(resultados)
        swEscribir = True
        For j = i + 1 To UBound(resultados)
            If resultados(i) = resultados(j) Then
                swEscribir = False
                Exit For
            End If
        Next

        If swEscribir = True Then
            ReDim Preserve temp(0 To contador)
            temp(contador) = resultados(i)
            contador = contador + 1
        End If
    Next
    
    alteración = temp
End Function

Private Function unirArrays(array1() As String, array2() As String, array3() As String, array4() As String) As String()
    Dim i As Double, temp() As String
    Dim hastaDónde As Double, desdedónde As Double
    
    
    ReDim temp(0 To UBound(array1)) 'se carga el array1 en el temp
    temp = array1
    
    desdedónde = UBound(array1) + 1 'el array1 + el array2
    hastaDónde = UBound(array1) + 1 + UBound(array2)
    ReDim Preserve temp(0 To hastaDónde)
    For i = desdedónde To hastaDónde 'sumamos el array2 al temp
        'Debug.Print temp(268)
        temp(i) = array2(i - desdedónde)
    Next
    
    desdedónde = UBound(temp) + 1 'el resultado de (array1+array2) + array3
    hastaDónde = UBound(temp) + 1 + UBound(array3)
    ReDim Preserve temp(0 To hastaDónde)
    For i = desdedónde To hastaDónde 'sumamos el array3 al temp
        temp(i) = array3(i - desdedónde)
    Next
    
    desdedónde = UBound(temp) + 1 'el resultado de (array1+array2+array3) + array4
    hastaDónde = UBound(temp) + 1 + UBound(array4)
    ReDim Preserve temp(0 To hastaDónde)
    For i = desdedónde To hastaDónde 'sumamos el array4 al temp
        temp(i) = array4(i - desdedónde)
    Next
    
    unirArrays = temp
End Function

Private Sub Class_Initialize()
    swHayDiccionario = False
End Sub

Public Sub Cargar_Diccionario(path As String)
    'revisar que controle que el archivo exista
    
    Dim archivolibre As Byte, cadena As String
    Dim posiciónEspacio As Integer, contador As Double
    
    If existeCarpeta(path) Then 'se controla que la ruta sea correcta
        archivolibre = FreeFile
        Open path For Input As archivolibre
        contador = 0
        While Not EOF(archivolibre)
            ReDim Preserve array_palabras(0 To contador)
            Line Input #archivolibre, cadena
            posiciónEspacio = InStr(cadena, " ")
            If posiciónEspacio <> 0 Then 'se controla el formato del diccionario
                array_palabras(contador).nombre = LCase(Trim(Left(cadena, posiciónEspacio)))
                array_palabras(contador).repetición = Int(Trim(Right(cadena, Len(cadena) - posiciónEspacio - 1)))
            
                contador = contador + 1
            Else
                MsgBox "El diccionario tiene un formato incorrecto en la línea: " + Str(contador) + ". Recuerde que el formato correcto es palabra, espacio, número (repeticiones). Resuelva el formato y vuelva a intentar.", , "Atención"
                swHayDiccionario = False
                Close archivolibre
                Exit Sub
            End If
        Wend
        
        Close archivolibre
        swHayDiccionario = True
    Else
        MsgBox "La ruta del diccionario no es correcta, por favor indique una adecuada.", , "Atención"
        swHayDiccionario = False
    End If
End Sub

Private Function unir_2_Arrays(array1() As String, array2() As String) As String()
    '**********************
    'si se quiere hacer que realmente corrija la palabra con dos errores de distancia
    'descomentar las líenas inferiores y comentar la que dice For i = 0 To UBound(array2)
    'dejé que corrija en dos errores parcialmente pues insume muchísmo tiempo
    '**********************
    Dim i As Double, temp() As String
    'Dim hastaDónde As Double, desdedónde As Double
    Dim contador As Double, j As Double, swRepetido As Boolean


    ReDim temp(0 To UBound(array1)) 'se carga el array1 en el temp
    temp = array1

    'desdedónde = UBound(array1) + 1 'el array1 + el array2
    'hastaDónde = UBound(array1) + 1 + UBound(array2)
    
    contador = UBound(array1) + 1
    'For i = desdedónde To hastaDónde 'sumamos el array2 al temp
    For i = 0 To UBound(array2)
        swRepetido = False
        For j = 0 To UBound(array1)
            If array2(i) = array1(j) Then
                swRepetido = True
                Exit For
            End If
        Next
        
        If swRepetido = False Then
            ReDim Preserve temp(0 To contador)
            temp(contador) = array2(i)
            contador = contador + 1
        End If
    Next
    
'    Dim temp() As String, i As Double, j As Double, contador As Double, swRepetido As Boolean
'
'    contador = 0
'    For i = 0 To UBound(array1)
'        swRepetido = False
'        For j = 0 To UBound(array2)
'            If array1(i) = array2(j) Then
'                swRepetido = True
'                Exit For
'            End If
'        Next
'
'        If swRepetido = False Then
'            ReDim Preserve temp(0 To contador)
'            temp(contador) = array1(i)
'            contador = contador + 1
'        End If
'    Next
'
'    For i = 0 To UBound(array2)
'        swRepetido = False
'        For j = 0 To UBound(array1)
'            If array2(i) = array1(j) Then
'                swRepetido = True
'                Exit For
'            End If
'        Next
'
'        If swRepetido = False Then
'            ReDim Preserve temp(0 To contador)
'            temp(contador) = array2(i)
'            contador = contador + 1
'        End If
'    Next

    unir_2_Arrays = temp
End Function

Private Function existeCarpeta(ByVal rutaCarpeta As String) As Boolean 'si existe o no una carpeta o archivo
    Dim x As Integer
    On Error GoTo Fallo
    x = GetAttr(rutaCarpeta)
    existeCarpeta = True
    Exit Function
Fallo:
    existeCarpeta = False
End Function

    
Private Function ordenar_array(array1() As palabra) As palabra()
    Dim j As Long, aux As palabra, cambio As Boolean, i As Long
    j = 0
    Do
        cambio = False
        j = j + 1
        For i = 0 To UBound(array1) - j
            If array1(i).repetición < array1(i + 1).repetición Then
                aux = array1(i)
                array1(i) = array1(i + 1)
                array1(i + 1) = aux
                cambio = True
            End If
        Next
    Loop While cambio = True
    
    ordenar_array = array1
End Function
