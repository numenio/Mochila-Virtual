VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmTítuloHoja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Título de la hoja"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   Icon            =   "frmTítuloHoja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTítuloHoja.frx":08CA
   ScaleHeight     =   1710
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   609
      _Version        =   393217
      MaxLength       =   64
      TextRTF         =   $"frmTítuloHoja.frx":2922
   End
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   495
      Left            =   1943
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Caption         =   "Guardar el título de la hoja"
      EstiloDelBoton  =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escribir un título a la hoja para después encontrarla fácilmente:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   4815
   End
End
Attribute VB_Name = "frmTítuloHoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nombreArchivo As String
Dim controlPresionado As Boolean
Dim archivoGuardado As Boolean

Private Sub ButtonTransparent1_Click()
    Call richTextbox1_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift And 7 = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If Shift And 7 = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda del cuaderno
         frmAyuda.formulario = formularios.títuloHoja
         frmAyuda.Show
         Exit Sub
    End If
    
    If Shift And 7 = vbCtrlMask Then Decir ""
End Sub

Private Sub Form_Load()
    Call centrarFormulario(Me)
    Decir "Escribile un título a tu hoja para que más tarde puedas encontrarla más fácilmente"
    archivoGuardado = False
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If archivoGuardado = False Then 'si no ha guardado la hoja, se impide cerrar la ventana
        Decir "no has escrito ningún título a tu hoja, escribí uno y apretá enter"
        Cancel = 1
    End If
End Sub

Private Sub richTextbox1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim archivolibre As Byte, shiftkey As Integer
    
    shiftkey = Shift And 7
    If shiftkey = vbCtrlMask Then controlPresionado = True
    
    If KeyCode = vbKeyReturn Then
        If Trim(RichTextBox1.Text) <> "" Then
            archivolibre = FreeFile 'se abre el archivo para guardar los datos de las partidas
            Open nombreArchivo For Random As archivolibre Len = 64
            Put archivolibre, 1, RichTextBox1.Text
            Close archivolibre
            archivoGuardado = True
            Unload Me
            Exit Sub
        Else
            Decir "no has escrito un título. primero escribilo y después apretá enter"
        End If
    ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        Decir "está escrito " + RichTextBox1.Text
    ElseIf KeyCode = vbKeyF1 Then
        Decir Label1.Caption
    End If
    
    Dim letra As String, auxString As String
    If KeyCode = vbKeyDelete Then
        If RichTextBox1.Text <> "" Then 'si no está vacío
            If RichTextBox1.SelStart <> Len(RichTextBox1.Text) Then 'y no está al final de la hoja
                If RichTextBox1.SelText <> "" Then 'si hay algo seleccionado
                    Decir "borrando el texto seleccionado"
                    Exit Sub
                End If
                letra = Mid(RichTextBox1.Text, RichTextBox1.SelStart + 1, 1)
                If letra = " " Then
                    Decir "borrando a la derecha el espacio", False
                ElseIf letra = Chr(9) Then
                    Decir "borrando a la derecha un salto"
                Else
                    auxString = traducirParaBorrar(letra)
                    Decir "borrando a la derecha " + auxString
                End If
            Else
                Decir "imposible borrar, estás al final del título"
            End If
        Else
            Decir "no se puede borrar a la derecha porque el título está vacío"
        End If
    End If
    
    If KeyCode = vbKeyBack Then
        If RichTextBox1.Text = "" Then
            Decir "Ya está todo borrado"
        Else
            If RichTextBox1.SelStart = 0 Then
                Decir "imposible borrar porque estás al principio del título"
            Else
                If RichTextBox1.SelText = "" Then 'si no hay nada seleccionado
                    letra = Mid(RichTextBox1.Text, RichTextBox1.SelStart, 1)
                Else
                    Decir "borrando el texto seleccionado"
                    Exit Sub
                End If
    
                If letra = " " Then
                    Decir "borrando el espacio"
                ElseIf letra = Chr(9) Then
                    Decir "borrando un salto"
                Else
                    auxString = traducirParaBorrar(letra)
                    Decir "borrando " + auxString
                End If
            End If
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyEnd Then Decir "final del texto"
    If shiftkey = 0 And KeyCode = vbKeyHome Then Decir "principio del texto"
End Sub

Private Sub richTextbox1_KeyPress(KeyAscii As Integer)
    Dim cadena As String
    If Len(RichTextBox1.Text) < 64 Then
        If KeyAscii >= 32 And KeyAscii <= 255 And controlPresionado = False Then cadena = quéLetraSeApretó(KeyAscii)

        If KeyAscii = 9 Then cadena = "salto hacia adelante" 'tab
        If KeyAscii = 39 Then cadena = "apóstrofo"
        If KeyAscii = 123 Then cadena = "abre llave"
        If KeyAscii = 125 Then cadena = "cierra llave"
        If KeyAscii = 91 Then cadena = "abre corchete"
        If KeyAscii = 93 Then cadena = "cierra corchete"
        If KeyAscii = 64 Then cadena = "arroba"

        'leer la palabra al apretar espacio, punto, coma, etc.
        If KeyAscii = 32 Or KeyAscii = Asc(".") Or KeyAscii = Asc(",") Or KeyAscii = Asc(";") Or KeyAscii = Asc(":") _
        Or KeyAscii = Asc("-") Or KeyAscii = Asc("_") Then cadena = cadena + decirPalabraAnterior(RichTextBox1)

        If cadena <> "" Then
            If RichTextBox1.SelBold = True Then cadena = cadena + " en negrita"
            If RichTextBox1.SelUnderline = True Then cadena = cadena + " subrayada"
            Decir cadena
        End If
'        If KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace Then Decir Chr(KeyAscii)
    Else
        Decir "ya se escribieron las 64 letras que le podés poner al título de tu hoja"
    End If
    controlPresionado = False 'se resetea la variable
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    
    shiftkey = Shift And 7
            
    If shiftkey = vbCtrlMask And KeyCode = vbKeyLeft Then 'leer por palabras retrocediendo
        Decir decirPalabraSiguiente(RichTextBox1) 'cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyRight Then 'leer por palabras avanzando
        Decir decirPalabraSiguiente(RichTextBox1) 'cadena
    End If
        
    If shiftkey = 0 And KeyCode = vbKeyRight Then 'avanzar de a caracter
        Decir decirLetraSiguiente(RichTextBox1)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyLeft Then 'retroceder de a caracter
        If RichTextBox1.SelStart = 0 And RichTextBox1.Text <> "" Then
            Decir "Estás en el principio del texto, delante de la letra " + decirLetraSiguiente(RichTextBox1)
        Else
            Decir decirLetraSiguiente(RichTextBox1)
        End If
    End If
    
    Dim TeclaShift, TeclaControl
    TeclaShift = (Shift And vbShiftMask) > 0
    TeclaControl = (Shift And vbCtrlMask) > 0

    Dim teclaApretada As Byte, control As Boolean, shift2 As Boolean
    Select Case KeyCode
        Case vbKeyA
            teclaApretada = tecla.a
        Case vbKeyUp
            teclaApretada = tecla.flechaArriba
        Case vbKeyDown
            teclaApretada = tecla.flechaAbajo
        Case vbKeyLeft
            teclaApretada = tecla.flechaIzquierda
        Case vbKeyRight
            teclaApretada = tecla.flechaDerecha
        Case vbKeyPageUp
            teclaApretada = tecla.avancePágina
        Case vbKeyPageDown
            teclaApretada = tecla.retrocesoPágina
        Case vbKeyHome
            teclaApretada = tecla.inicio
        Case vbKeyEnd
            teclaApretada = tecla.fin
        Case vbKeyBack
            teclaApretada = tecla.borrar
        Case vbKeyDelete
            teclaApretada = tecla.borrar
    End Select

    If TeclaControl Then
        control = True
    Else
        control = False
    End If

    If TeclaShift Then
        shift2 = True
    Else
        shift2 = False
    End If

    Call evaluarSelección(RichTextBox1, control, shift2, teclaApretada) 'se ve si hay selección
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copiar
        If RichTextBox1.SelText <> "" Then
            Decir "se copió el texto seleccionado"
        Else
            Decir "No se puede copiar porque no hay texto seleccionado. para seleccionar, usar shift más las flechas"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyX Then 'cortar
        If RichTextBox1.SelText <> "" Then
            Decir "se cortó el texto seleccionado"
        Else
            Decir "No se puede cortar porque no hay texto seleccionado. para seleccionar, usar shift más las flechas"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyV Then 'pegar
        If Clipboard.GetText <> "" Then
            Decir "texto pegado: " + Clipboard.GetText
        Else
            Decir "No se puede pegar porque no hay texto copiado o cortado. para copiar, usar control más ce. para cortar, usar control más équis"
        End If
    End If
    controlPresionado = False 'se resetea la variable
End Sub

