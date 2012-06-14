VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   3240
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Interval        =   5800
      Left            =   3240
      Top             =   2280
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _cx             =   13361
      _cy             =   8705
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   0   'False
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ret As Long ', x,
Dim Y
Dim lectorRegistro

Private Sub Form_Load()
    On Error Resume Next
    Call centrarFormulario(Me)
    Set lectorRegistro = CreateObject("WScript.Shell")
'    x = lectorRegistro.regRead("HKEY_LOCAL_MACHINE\Software\ReyNegro-ReyBlanco\MochilaVirtual\datos\")
    rutaDeAspell = lectorRegistro.regRead("HKEY_LOCAL_MACHINE\SOFTWARE\Aspell\Path")
    If Len(rutaDeAspell) > 0 Then 'si aspell está instalado
        swAspellInstalado = True
        '*************************
        'se ve si está instalado el diccionario de español
        Set lectorRegistro = CreateObject("WScript.Shell")
        Y = lectorRegistro.regRead("HKEY_LOCAL_MACHINE\SOFTWARE\Aspell-es\UninstallString")
        If Y <> "" Then 'si está instalado
            idiomaAspell = "es"
            Set objPipe = New clsPipe 'se instancia un pipe para comunicarse con aspell
        Else
            swAspellInstalado = False 'si no está instalado, se prefiere el corrector propio
        End If
    Else
        swAspellInstalado = False
    End If
    
        
    flash.Movie = App.path + "\sonidos\swf\cybertalk.swf"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sonido = sndStopSound(0, 0)
    'ret = mciExecute("Close all")
    Set lectorRegistro = Nothing
    frmOculto.media.Stop
End Sub

Private Sub Timer1_Timer()
    Static sw As Boolean, direcciónArchivo As String
    
    flash.Stop
    flash.Movie = App.path + "\sonidos\swf\intro.swf"
    direcciónArchivo = PathCorto(App.path + "\sonidos\inicio.wav")
    If sw = False Then
'        ret = mciExecute("OPEN " & direcciónArchivo)
'        ret = mciExecute("PLAY " & direcciónArchivo)
        frmOculto.swContinuarReproducción = False
        frmOculto.media.FileName = direcciónArchivo
        frmOculto.media.Play
    End If
    flash.Play
    If sonido <> 1 Then sonido = sndPlaySound(App.path + "\sonidos\reyes.wav", SND_ASYNC)
    If sw = True Then
        Timer2.Enabled = True
        Timer1.Enabled = False
    End If
    sw = True
End Sub

Private Sub Timer2_Timer()
    flash.Stop
    'If x = 1 Then
    If existeCarpeta(App.path + "\Datos\Inicio") Then
        frmOculto.media.Volume = -2500
        frmPrincipal.Show
    Else
        'lectorRegistro.RegWrite "HKEY_LOCAL_MACHINE\Software\ReyNegro-ReyBlanco\MochilaVirtual\datos\", "1"
        MkDir (App.path + "\Datos\Inicio") 'sólo para saber que ya se crearon las carpetas
        frmPrimerosDatos.Show
    End If
    Set lectorRegistro = Nothing
    Unload Me
End Sub
