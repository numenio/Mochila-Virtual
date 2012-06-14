VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmOculto 
   Caption         =   "Oculto"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   3360
   End
   Begin MediaPlayerCtl.MediaPlayer media 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmOculto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public swContinuarReproducción As Boolean
'Dim x

Private Sub Form_Load()
    On Error GoTo manejoErrorInicio
    
    If App.PrevInstance = True Then End 'si ya hay cargada una versión de la mochila no se pude cargar otra
    Me.Visible = False
    ReDim recordatoriosActivos(0 To 0)
'    Dim lectorRegistro
'    Set lectorRegistro = CreateObject("WScript.Shell")
'    x = lectorRegistro.regRead("HKEY_LOCAL_MACHINE\Software\ReyNegro-ReyBlanco\MochilaVirtual\")
    'If (x <> "1") Or (
    If Not existeCarpeta(App.path + "\Datos\Inicio") Then 'chequeamos la carpeta inicio, que está ahí solo para saber que ya se hicieron las carpetas
        Call crearCarpetasPrograma 'primer uso del programa, se crean las carpetas de la mochila
    End If
'    Set lectorRegistro = Nothing
    Call cargarRecordatorios
    frmSplash.Show
    Exit Sub
manejoErrorInicio:
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call contarFormularios(False)
End Sub

Private Sub media_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    If OldState = mpPlaying And NewState = mpStopped And swContinuarReproducción = True Then
        media.Play
    End If
End Sub

Private Sub Timer1_Timer()
    Static horaAnterior As Date
    Dim i As Integer, j As Integer
    
    If Right(Format(horaAnterior, "HH:mm"), 2) <> Right(Format(Time, "HH:mm"), 2) Then
        Call cargarRecordatorios 'si han cambiado los minutos
        horaAnterior = Format(Time, "HH:mm")
    End If
    
    If UBound(recordatoriosActivos) > 0 Then
        For i = 1 To UBound(recordatoriosActivos) 'se empieza en el recordatorio de índice 1 pues el de índice 0 se deja en blanco
            If recordatoriosActivos(i).yaAnunciado = False Then
                sonido = sndPlaySound(App.path + "\Sonidos\Recordatorios\" + usuario.rutaSonidosRecordatorios, SND_ASYNC)
                frmMsgBox.swMostrarCancelar = False
                If Format(recordatoriosActivos(i).fecha, "dd/mm/yyyy") <> Format(Date, "dd/mm/yyyy") Then frmMsgBox.cadenaAMostrar = "Recordatorio del día " + Format(recordatoriosActivos(i).fecha) + ". "
                frmMsgBox.cadenaAMostrar = frmMsgBox.cadenaAMostrar + "Recordatorio de la hora " + Trim(Left(Str(recordatoriosActivos(i).hora), 6)) + ": " + Trim(recordatoriosActivos(i).texto)
                frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
                If frmMsgBox.swEstoyAbierto = False Then
                    frmMsgBox.Show 1
                    recordatoriosActivos(i).yaAnunciado = True
                    Call GuardarRecordatorioEnPosición(recordatoriosActivos(i), recordatoriosActivos(i).índiceEnArchivo)
                End If
            End If
        Next
    End If
End Sub
