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
Private swDiccionarioReciénCargado As Boolean
Private largoTexto As Integer
Private cadenaBorrar As String

Private Sub Form_Activate()
    Combo.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyEscape Then Unload Me
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    If KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = Asc("ñ") Or KeyAscii = Asc("á") Or KeyAscii = Asc("é") _
    Or KeyAscii = Asc("í") Or KeyAscii = Asc("ó") Or KeyAscii = Asc("ú") Or KeyAscii = Asc("ü") _
    Or KeyAscii = Asc("à") Or KeyAscii = Asc("è") Or KeyAscii = Asc("ì") Or KeyAscii = Asc("ò") Or KeyAscii = Asc("ù") _
    Or KeyAscii = Asc("â") Or KeyAscii = Asc("ê") Or KeyAscii = Asc("î") Or KeyAscii = Asc("ô") Or KeyAscii = Asc("û") _
    Or KeyAscii = Asc("ä") Or KeyAscii = Asc("ë") Or KeyAscii = Asc("ï") Or KeyAscii = Asc("ö") Or KeyAscii = Asc("ç") _
    Then
        KeyAscii = KeyAscii - 32
'        If Combo.ListCount > 0 Then
'            If Combo.ListCount > 1 Then
'                Decir quéLetraSeApretó(Asc(LCase(Chr(KeyAscii)))) + ". quedan " + Str(Combo.ListCount) + "palabras. podés verlas con las flechas"
'            Else 'si sólamente hay una palabra
'                Decir quéLetraSeApretó(Asc(LCase(Chr(KeyAscii)))) + ". queda una palabra. podés verla con las flechas"
'            End If
'        Else 'si no quedan palabras en la lista
'            Decir quéLetraSeApretó(Asc(LCase(Chr(KeyAscii)))) + ". no hay palabras en el diccionario que empiecen con " + Combo.Text
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
    swDiccionarioReciénCargado = True
    Decir "abriendo el diccionario " + Left(Me.diccionarioElegido, InStr(1, Me.diccionarioElegido, ".") - 1) + ". Escribí la palabra que busques y aceptá con enter"
    Me.swEstoyAbierto = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If swCuadernoAbierto = True Then
        Decir "cerrando el diccionario, volviendo a tu carpeta"
    ElseIf frmLectorEvaluaciones.swEstoyAbierto = True Then
        Decir "cerrando el diccionario, volviendo a la evaluación"
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
            If Combo.ListCount <> 0 Then 'si aún quedan palabras en el combo luego de ir filtrando
                cadenaEncontrada = buscarEntrada(Combo.Text, diccionarioElegido)
                
                If cadenaEncontrada <> "" Then
                    Label1.Caption = cadenaEncontrada
                Else
                    Label1.Caption = "La palabra " + Trim(Combo.Text) + " está mal escrita o no existe en el diccionario"
                End If
                Decir Label1.Caption
                Combo.Text = ""
                largoTexto = 0
                Call cargarPalabras
            Else
                Decir "la palabra " + LCase(Combo.Text) + " no está en el diccionario"
            End If
        Else
            Decir "no has escrito ninguna palabra para buscar en el diccionario"
        End If
    End If
    
    If KeyCode = vbKeyBack Then 'si se está borrando
        If Combo.Text <> "" Then 'si hay algo escrito
            If Len(Combo.Text) - 1 > largoTexto Then 'si hay más texto en el combo que el q escribió el usuario, es decir que se usaron las flechas para ver las palabras de la lista
                cadenaBorrar = "borrando el texto que escribí cuando usaste las flechas para completar " + LCase(Combo.Text) + ", queda escrito " + controlarCadena(LCase(Left(Combo.Text, largoTexto)))
            Else 'si el texto no es más largo
                If Right(Combo.Text, 1) = " " Then 'si se borra el espacio
                    cadenaBorrar = "borrando el espacio"
                Else 'si no es espacio
                    cadenaBorrar = "borrando la letra " + quéLetraSeApretó(Asc(LCase(Right(Combo.Text, 1))))
                End If
                
                If Len(Combo.Text) - 1 > 0 Then 'si queda algo escrito
                    cadenaBorrar = cadenaBorrar + ". queda escrito "
                    cadenaBorrar = cadenaBorrar + controlarCadena(LCase(Left(Combo.Text, Len(Combo.Text) - 1)))
                Else 'si se borró todo
                    cadenaBorrar = cadenaBorrar + ". borraste todo"
                End If
                largoTexto = largoTexto - 1
                If Combo.Text = "" Then largoTexto = 0
            End If
        Else 'si no hay nada escrito
            largoTexto = 0
            cadenaBorrar = "ya está todo borrado"
            KeyCode = 0
        End If
    End If
    
    If KeyCode = vbKeyF1 Then 'repetir la palabra encontrada con F1
        If Label1.Caption = "" Then
            Decir "no puedo leerte la definición porque aún no has buscado ninguna palabra"
        Else
            Decir "La definición dice: " + Label1.Caption
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copia el contenido al portapapeles
        If Label1.Caption <> "" Then
            Clipboard.Clear
            Clipboard.SetText (Label1.Caption)
            Decir "definición copiada"
        Else
            Decir "no puedo copiar la definición porque aún no has buscado ninguna palabra"
        End If
        KeyCode = 0
        Shift = 0
        Exit Sub
    End If
    
    If KeyCode = vbKeyF10 Then
        Decir "pasando a tu carpeta, para volver al diccionario usá efe diez"
        frmCuaderno.RichTextBox1.SetFocus
        SendKeys ("%")
    End If
    
    If (KeyCode >= vbKey0 And KeyCode <= vbKeyZ) Or KeyCode = Asc("Ñ") Or KeyCode = Asc("Á") Or KeyCode = Asc("É") _
    Or KeyCode = Asc("Í") Or KeyCode = Asc("Ó") Or KeyCode = Asc("Ú") Or KeyCode = Asc("Ü") _
    Or KeyCode = Asc("À") Or KeyCode = Asc("È") Or KeyCode = Asc("Ì") Or KeyCode = Asc("Ò") Or KeyCode = Asc("Ù") _
    Or KeyCode = Asc("Â") Or KeyCode = Asc("Ê") Or KeyCode = Asc("Î") Or KeyCode = Asc("Ô") Or KeyCode = Asc("Û") _
    Or KeyCode = Asc("Ä") Or KeyCode = Asc("Ë") Or KeyCode = Asc("Ï") Or KeyCode = Asc("Ö") Or KeyCode = Asc("Ç") _
    Or KeyCode = Asc(" ") Or KeyCode = Asc(".") Or KeyCode = Asc(",") Or KeyCode = Asc("'") Then
        largoTexto = Len(Combo.Text) + 1 'SE GUARDA EN UNA VARIABLE CUÁNTO SE HA ESCRITO EN EL TEXTO
    End If
End Sub


Private Sub conjuntoPalabrasBorrando(Optional cadena As String = "")
    Dim archivolibre As Byte
    Dim cadenaDiccionario As String
    Dim palabra As String
    Dim posiciónDosPuntos As Integer
    Dim auxString As String

    auxString = Combo.Text
    Combo.Clear
    archivolibre = FreeFile
    Open App.path + "\diccionarios\" + diccionarioElegido For Input As archivolibre
    Do While Not EOF(archivolibre)   ' Repite el bucle hasta el final del archivo.
        Line Input #archivolibre, cadenaDiccionario ' Lee el carácter en dos variables.
        posiciónDosPuntos = InStr(1, cadenaDiccionario, ":") - 1
        If posiciónDosPuntos > 0 Then
            palabra = Trim(Left(cadenaDiccionario, posiciónDosPuntos))
            If UCase(palabra) = palabra Then 'si está en mayúsculas, o sea que es palabra y no definición
                If cadena <> "" Then 'se van filtrando las palabras según lo que escribe el usuario
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
    
    If cadena = "" Then 'se carga si se acaba de cargar el diccionario completo, así si ya está todo borrado y se aprieta backspace, no se carga de nuevo todo el diccionario
        swDiccionarioReciénCargado = True
    Else
        swDiccionarioReciénCargado = False
    End If
    
    Combo.Text = auxString
    Combo.SelStart = Len(Combo.Text)
End Sub

'Private Sub cargarPalabras()
'    Dim archivolibre As Byte
'    Dim cadenaDiccionario As String
'    Dim palabra As String
'    Dim posiciónDosPuntos As Integer
'    Dim contador As Double
'
'    contador = 0
'    archivolibre = FreeFile
'    ReDim palabrasDiccionario(0 To 0)
'    Open App.path + "\diccionarios\" + diccionarioElegido For Input As archivolibre
'    Do While Not EOF(archivolibre)   ' Repite el bucle hasta el final del archivo.
'        Line Input #archivolibre, cadenaDiccionario ' Lee el carácter en dos variables.
'        posiciónDosPuntos = InStr(1, cadenaDiccionario, ":") - 1
'        If posiciónDosPuntos > 0 Then
'            palabra = Trim(Left(cadenaDiccionario, posiciónDosPuntos))
'            If UCase(palabra) = palabra Then 'si está en mayúsculas, o sea que es palabra y no definición
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
    Dim posiciónDosPuntos As Integer
    Dim contador As Double
    
    Combo.Clear
    contador = 0
    archivolibre = FreeFile
    ReDim palabrasDiccionario(0 To 0)
    Open App.path + "\diccionarios\" + diccionarioElegido For Input As archivolibre
    Do While Not EOF(archivolibre)   ' Repite el bucle hasta el final del archivo.
        Line Input #archivolibre, cadenaDiccionario ' Lee el carácter en dos variables.
        posiciónDosPuntos = InStr(1, cadenaDiccionario, ":") - 1
        If posiciónDosPuntos > 0 Then
            palabra = Trim(Left(cadenaDiccionario, posiciónDosPuntos))
            If UCase(palabra) = palabra Then 'si está en mayúsculas, o sea que es palabra y no definición
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
    ' esto sería para usar una matriz para filtrar las palabras en lugar de usar los elementos del listbox
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
        If swDiccionarioReciénCargado = False Then Call conjuntoPalabrasBorrando(Combo.Text)
        If Combo.ListCount > 1 Then
            cadenaBorrar = cadenaBorrar + Str(Combo.ListCount) + "palabras. podés verlas con las flechas"
        ElseIf Combo.ListCount = 1 Then 'si sólamente hay una palabra
            cadenaBorrar = cadenaBorrar + ". una palabra. podés verla con las flechas"
        ElseIf Combo.ListCount = 0 Then 'si ya no hay palabras en la lista
            cadenaBorrar = cadenaBorrar + ". no hay palabras en el diccionario que empiecen con " + LCase(Combo.Text)
        End If
        Decir cadenaBorrar
    End If
    
    If ((KeyCode >= vbKey0 And KeyCode <= vbKeyZ) Or KeyCode = Asc("Ñ") Or KeyCode = Asc("Á") Or KeyCode = Asc("É") _
    Or KeyCode = Asc("Í") Or KeyCode = Asc("Ó") Or KeyCode = Asc("Ú") Or KeyCode = Asc("Ü") _
    Or KeyCode = Asc("À") Or KeyCode = Asc("È") Or KeyCode = Asc("Ì") Or KeyCode = Asc("Ò") Or KeyCode = Asc("Ù") _
    Or KeyCode = Asc("Â") Or KeyCode = Asc("Ê") Or KeyCode = Asc("Î") Or KeyCode = Asc("Ô") Or KeyCode = Asc("Û") _
    Or KeyCode = Asc("Ä") Or KeyCode = Asc("Ë") Or KeyCode = Asc("Ï") Or KeyCode = Asc("Ö") Or KeyCode = Asc("Ç")) And _
    (Shift And 7) <> vbCtrlMask Then
        If Combo.ListCount > 0 Then
            Call conjuntoPalabras(Combo.Text)
            
            If Combo.ListCount > 1 Then
                Decir quéLetraSeApretó(Asc(LCase(Chr(KeyCode)))) + ". " + LCase(Combo.Text) + ". quedan " + Str(Combo.ListCount) + "palabras. podés verlas con las flechas"
            ElseIf Combo.ListCount = 1 Then 'si sólamente hay una palabra
                Decir quéLetraSeApretó(Asc(LCase(Chr(KeyCode)))) + ". " + LCase(Combo.Text) + ". queda una palabra. podés verla con las flechas"
            ElseIf Combo.ListCount = 0 Then 'si ya no hay palabras en la lista
                Decir quéLetraSeApretó(Asc(LCase(Chr(KeyCode)))) + ". no hay palabras en el diccionario que empiecen con " + LCase(Combo.Text)
            End If
        Else 'si no quedan palabras en la lista
            Decir quéLetraSeApretó(Asc(LCase(Chr(KeyCode)))) + ". no hay palabras en el diccionario que empiecen con " + LCase(Combo.Text)
        End If
        
        'If Combo.ListCount <> 0 Then Call conjuntoPalabras(Combo.Text)
        swDiccionarioReciénCargado = False
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
