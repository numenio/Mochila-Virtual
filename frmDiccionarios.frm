VERSION 5.00
Begin VB.Form frmDiccionarioElegido 
   Caption         =   "Diccionario"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   Icon            =   "frmDiccionarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmDiccionarios.frx":6852
   ScaleHeight     =   6270
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo 
      Height          =   2520
      Left            =   1440
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Respuesta del diccionario:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba la palabra a buscar:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   840
      TabIndex        =   0
      Top             =   3840
      Width           =   4455
   End
End
Attribute VB_Name = "frmDiccionarioElegido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public diccionarioElegido As String
Public swEstoyAbierto As Boolean
Private swDiccionarioReci�nCargado As Boolean
Private largoTexto As Integer
Private cadenaBorrar As String

Private Sub Form_Activate()
    Combo.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el men� de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then Unload Me
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    If KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") _
    Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") _
    Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") _
    Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") _
    Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") Or KeyAscii = Asc("�") _
    Then
        KeyAscii = KeyAscii - 32
'        If Combo.ListCount > 0 Then
'            If Combo.ListCount > 1 Then
'                Decir qu�LetraSeApret�(Asc(LCase(Chr(KeyAscii)))) + ". quedan " + Str(Combo.ListCount) + "palabras. pod�s verlas con las flechas"
'            Else 'si s�lamente hay una palabra
'                Decir qu�LetraSeApret�(Asc(LCase(Chr(KeyAscii)))) + ". queda una palabra. pod�s verla con las flechas"
'            End If
'        Else 'si no quedan palabras en la lista
'            Decir qu�LetraSeApret�(Asc(LCase(Chr(KeyAscii)))) + ". no hay palabras en el diccionario que empiecen con " + Combo.Text
'        End If
    End If
    
    'leer la palabra al apretar espacio, punto, coma, etc.
    If KeyAscii = 32 Or KeyAscii = Asc(".") Or KeyAscii = Asc(",") Or KeyAscii = Asc(";") Or KeyAscii = Asc(":") _
    Or KeyAscii = Asc("-") Or KeyAscii = Asc("_") Then
        Decir Combo.Text
    End If
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    Call cargarPalabras
    swDiccionarioReci�nCargado = True
    Decir "abriendo el diccionario " + Left(Me.diccionarioElegido, InStr(1, Me.diccionarioElegido, ".") - 1) + ". Escrib� la palabra que busques y acept� con enter"
    Me.swEstoyAbierto = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If swCuadernoAbierto = True Then
        Decir "cerrando el diccionario, volviendo a tu carpeta"
    ElseIf frmLectorEvaluaciones.swEstoyAbierto = True Then
        Decir "cerrando el diccionario, volviendo a la evaluaci�n"
    Else
        frmAccesorios.Show
    End If
    Me.swEstoyAbierto = False
End Sub

Private Sub combo_DblClick()
    Label1.Caption = buscarEntrada(Combo.List(Combo.ListIndex), diccionarioElegido)
End Sub

Private Sub combo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cadenaEncontrada As String, shiftkey As Byte
    
    shiftkey = Shift And 7

    If KeyCode = vbKeyReturn Then 'si se da enter
        If Combo.Text <> "" Then 'si hay algo escrito
            If Combo.ListCount <> 0 Then 'si a�n quedan palabras en el combo luego de ir filtrando
                cadenaEncontrada = buscarEntrada(Combo.Text, diccionarioElegido)
                
                If cadenaEncontrada <> "" Then
                    Label1.Caption = cadenaEncontrada
                Else
                    Label1.Caption = "La palabra " + Trim(Combo.Text) + " est� mal escrita o no existe en el diccionario"
                End If
                Decir Label1.Caption
                Combo.Text = ""
                largoTexto = 0
                Call cargarPalabras
            Else
                Decir "la palabra " + LCase(Combo.Text) + " no est� en el diccionario"
            End If
        Else
            Decir "no has escrito ninguna palabra para buscar en el diccionario"
        End If
    End If
    
    If KeyCode = vbKeyBack Then 'si se est� borrando
        If Combo.Text <> "" Then 'si hay algo escrito
            If Len(Combo.Text) - 1 > largoTexto Then 'si hay m�s texto en el combo que el q escribi� el usuario, es decir que se usaron las flechas para ver las palabras de la lista
                cadenaBorrar = "borrando el texto que escrib� cuando usaste las flechas para completar " + LCase(Combo.Text) + ", queda escrito " + controlarCadena(LCase(Left(Combo.Text, largoTexto)))
            Else 'si el texto no es m�s largo
                If Right(Combo.Text, 1) = " " Then 'si se borra el espacio
                    cadenaBorrar = "borrando el espacio"
                Else 'si no es espacio
                    cadenaBorrar = "borrando la letra " + qu�LetraSeApret�(Asc(LCase(Right(Combo.Text, 1))))
                End If
                
                If Len(Combo.Text) - 1 > 0 Then 'si queda algo escrito
                    cadenaBorrar = cadenaBorrar + ". queda escrito "
                    cadenaBorrar = cadenaBorrar + controlarCadena(LCase(Left(Combo.Text, Len(Combo.Text) - 1)))
                Else 'si se borr� todo
                    cadenaBorrar = cadenaBorrar + ". borraste todo"
                End If
                largoTexto = largoTexto - 1
                If Combo.Text = "" Then largoTexto = 0
            End If
        Else 'si no hay nada escrito
            largoTexto = 0
            cadenaBorrar = "ya est� todo borrado"
            KeyCode = 0
        End If
    End If
    
    If KeyCode = vbKeyF1 Then 'repetir la palabra encontrada con F1
        If Label1.Caption = "" Then
            Decir "no puedo leerte la definici�n porque a�n no has buscado ninguna palabra"
        Else
            Decir "La definici�n dice: " + Label1.Caption
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copia el contenido al portapapeles
        If Label1.Caption <> "" Then
            Clipboard.Clear
            Clipboard.SetText (Label1.Caption)
            Decir "definici�n copiada"
        Else
            Decir "no puedo copiar la definici�n porque a�n no has buscado ninguna palabra"
        End If
        KeyCode = 0
        Shift = 0
        Exit Sub
    End If
    
    If KeyCode = vbKeyF10 Then
        Decir "pasando a tu carpeta, para volver al diccionario us� efe diez"
        frmCuaderno.RichTextBox1.SetFocus
        SendKeys ("%")
    End If
    
    If (KeyCode >= vbKey0 And KeyCode <= vbKeyZ) Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") _
    Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") _
    Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") _
    Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") _
    Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") _
    Or KeyCode = Asc(" ") Or KeyCode = Asc(".") Or KeyCode = Asc(",") Or KeyCode = Asc("'") Then
        largoTexto = Len(Combo.Text) + 1 'SE GUARDA EN UNA VARIABLE CU�NTO SE HA ESCRITO EN EL TEXTO
    End If
End Sub


Private Sub conjuntoPalabrasBorrando(Optional cadena As String = "")
    Dim archivolibre As Byte
    Dim cadenaDiccionario As String
    Dim palabra As String
    Dim posici�nDosPuntos As Integer
    Dim auxString As String

    auxString = Combo.Text
    Combo.Clear
    archivolibre = FreeFile
    Open App.path + "\diccionarios\" + diccionarioElegido For Input As archivolibre
    Do While Not EOF(archivolibre)   ' Repite el bucle hasta el final del archivo.
        Line Input #archivolibre, cadenaDiccionario ' Lee el car�cter en dos variables.
        posici�nDosPuntos = InStr(1, cadenaDiccionario, ":") - 1
        If posici�nDosPuntos > 0 Then
            palabra = Trim(Left(cadenaDiccionario, posici�nDosPuntos))
            If UCase(palabra) = palabra Then 'si est� en may�sculas, o sea que es palabra y no definici�n
                If cadena <> "" Then 'se van filtrando las palabras seg�n lo que escribe el usuario
                    If LCase(Trim(Left(palabra, Len(cadena)))) = LCase(Trim(cadena)) Then
                        Combo.AddItem palabra
                    End If
                Else
                    Combo.AddItem palabra
                End If
            End If
        End If
    Loop
    Close archivolibre
    
    If cadena = "" Then 'se carga si se acaba de cargar el diccionario completo, as� si ya est� todo borrado y se aprieta backspace, no se carga de nuevo todo el diccionario
        swDiccionarioReci�nCargado = True
    Else
        swDiccionarioReci�nCargado = False
    End If
    
    Combo.Text = auxString
    Combo.SelStart = Len(Combo.Text)
End Sub

'Private Sub cargarPalabras()
'    Dim archivolibre As Byte
'    Dim cadenaDiccionario As String
'    Dim palabra As String
'    Dim posici�nDosPuntos As Integer
'    Dim contador As Double
'
'    contador = 0
'    archivolibre = FreeFile
'    ReDim palabrasDiccionario(0 To 0)
'    Open App.path + "\diccionarios\" + diccionarioElegido For Input As archivolibre
'    Do While Not EOF(archivolibre)   ' Repite el bucle hasta el final del archivo.
'        Line Input #archivolibre, cadenaDiccionario ' Lee el car�cter en dos variables.
'        posici�nDosPuntos = InStr(1, cadenaDiccionario, ":") - 1
'        If posici�nDosPuntos > 0 Then
'            palabra = Trim(Left(cadenaDiccionario, posici�nDosPuntos))
'            If UCase(palabra) = palabra Then 'si est� en may�sculas, o sea que es palabra y no definici�n
'                ReDim Preserve palabrasDiccionario(0 To contador)
'                palabrasDiccionario(contador) = palabra
'                contador = contador + 1
'            End If
'        End If
'    Loop
'    Close archivolibre
'End Sub


Private Sub cargarPalabras()
    Dim archivolibre As Byte
    Dim cadenaDiccionario As String
    Dim palabra As String
    Dim posici�nDosPuntos As Integer
    Dim contador As Double
    
    Combo.Clear
    contador = 0
    archivolibre = FreeFile
    ReDim palabrasDiccionario(0 To 0)
    Open App.path + "\diccionarios\" + diccionarioElegido For Input As archivolibre
    Do While Not EOF(archivolibre)   ' Repite el bucle hasta el final del archivo.
        Line Input #archivolibre, cadenaDiccionario ' Lee el car�cter en dos variables.
        posici�nDosPuntos = InStr(1, cadenaDiccionario, ":") - 1
        If posici�nDosPuntos > 0 Then
            palabra = Trim(Left(cadenaDiccionario, posici�nDosPuntos))
            If UCase(palabra) = palabra Then 'si est� en may�sculas, o sea que es palabra y no definici�n
                Combo.AddItem palabra
            End If
        End If
    Loop
    Close archivolibre
    'Combo.Text = auxString
End Sub


Private Sub conjuntoPalabras(Optional cadena As String = "")
    Dim palabra() As String
    Dim i As Integer, auxString As String
    
    If Combo.ListCount > 0 Then
        ReDim palabra(0 To Combo.ListCount - 1)
        
        For i = 0 To Combo.ListCount - 1
            palabra(i) = Combo.List(i)
        Next
        
        auxString = Combo.Text
        
        Combo.Clear
    ' esto ser�a para usar una matriz para filtrar las palabras en lugar de usar los elementos del listbox
    '    For i = 0 To UBound(palabrasDiccionario)
    '        If LCase(Trim(Left(palabrasDiccionario(i), Len(cadena)))) = LCase(Trim(cadena)) Then
    '            combo.AddItem palabrasDiccionario(i)
    '        End If
    '    Next
        
        For i = 0 To UBound(palabra)
            If LCase(Trim(Left(palabra(i), Len(cadena)))) = LCase(Trim(cadena)) Then
                Combo.AddItem palabra(i)
            End If
        Next
        
        Combo.Text = auxString
        Combo.SelStart = Len(Combo.Text)
    End If
End Sub

Private Sub combo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then
        If swDiccionarioReci�nCargado = False Then Call conjuntoPalabrasBorrando(Combo.Text)
        If Combo.ListCount > 1 Then
            cadenaBorrar = cadenaBorrar + Str(Combo.ListCount) + "palabras. pod�s verlas con las flechas"
        ElseIf Combo.ListCount = 1 Then 'si s�lamente hay una palabra
            cadenaBorrar = cadenaBorrar + ". una palabra. pod�s verla con las flechas"
        ElseIf Combo.ListCount = 0 Then 'si ya no hay palabras en la lista
            cadenaBorrar = cadenaBorrar + ". no hay palabras en el diccionario que empiecen con " + LCase(Combo.Text)
        End If
        Decir cadenaBorrar
    End If
    
    If ((KeyCode >= vbKey0 And KeyCode <= vbKeyZ) Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") _
    Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") _
    Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") _
    Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") _
    Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�") Or KeyCode = Asc("�")) And _
    (Shift And 7) <> vbCtrlMask Then
        If Combo.ListCount > 0 Then
            Call conjuntoPalabras(Combo.Text)
            
            If Combo.ListCount > 1 Then
                Decir qu�LetraSeApret�(Asc(LCase(Chr(KeyCode)))) + ". " + LCase(Combo.Text) + ". quedan " + Str(Combo.ListCount) + "palabras. pod�s verlas con las flechas"
            ElseIf Combo.ListCount = 1 Then 'si s�lamente hay una palabra
                Decir qu�LetraSeApret�(Asc(LCase(Chr(KeyCode)))) + ". " + LCase(Combo.Text) + ". queda una palabra. pod�s verla con las flechas"
            ElseIf Combo.ListCount = 0 Then 'si ya no hay palabras en la lista
                Decir qu�LetraSeApret�(Asc(LCase(Chr(KeyCode)))) + ". no hay palabras en el diccionario que empiecen con " + LCase(Combo.Text)
            End If
        Else 'si no quedan palabras en la lista
            Decir qu�LetraSeApret�(Asc(LCase(Chr(KeyCode)))) + ". no hay palabras en el diccionario que empiecen con " + LCase(Combo.Text)
        End If
        
        'If Combo.ListCount <> 0 Then Call conjuntoPalabras(Combo.Text)
        swDiccionarioReci�nCargado = False
    End If
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If largoTexto > 0 Then
            If largoTexto <= Len(Combo.Text) Then
                Combo.SelStart = largoTexto
                Combo.SelLength = Len(Combo.Text) - largoTexto
            End If
        End If
        Decir LCase(Combo.List(Combo.ListIndex)) + ". " + separarEnLetras(LCase(Combo.List(Combo.ListIndex)))
    End If
End Sub
