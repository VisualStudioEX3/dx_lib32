VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_Sound - Efectos basicos en musica"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll3 
      Height          =   315
      LargeChange     =   50
      Left            =   1553
      Max             =   200
      Min             =   25
      SmallChange     =   5
      TabIndex        =   5
      Top             =   1733
      Value           =   100
      Width           =   2475
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   315
      LargeChange     =   10
      Left            =   1553
      Max             =   100
      Min             =   -100
      TabIndex        =   3
      Top             =   1373
      Width           =   2475
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      LargeChange     =   5
      Left            =   1553
      Max             =   100
      TabIndex        =   1
      Top             =   1013
      Value           =   100
      Width           =   2475
   End
   Begin VB.Label Label3 
      Caption         =   "Velocidad"
      Height          =   255
      Left            =   533
      TabIndex        =   4
      Top             =   1793
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Balance"
      Height          =   255
      Left            =   533
      TabIndex        =   2
      Top             =   1433
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Volumen"
      Height          =   255
      Left            =   533
      TabIndex        =   0
      Top             =   1073
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Audio As dx_Sound_Class ' Instancia del objeto de audio de dx_lib32.
Private Sample As Long ' Guarda el identificador de la muestra de musica.

Private Sub Form_Load()
    Set Audio = New dx_Sound_Class ' Creamos la instancia del objeto.
    Audio.Init Me.hWnd ' Inicializamos el motor de audio por defecto.
    Sample = Audio.MUSIC_Load(App.Path & "\sample.mp3") ' Cargamos en memoria la muestra de musica.
    Audio.MUSIC_Play Sample, Primary_Buffer, True ' Reproducimos la muestra en bucle en el canal primario de musica.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Audio.MUSIC_Unload Sample ' Descargamos de memoria la muestra de music.
    Audio.Terminate ' Terminamos la ejecucion de la clase de audio y liberamos los recursos utilizados.
    Set Audio = Nothing ' Destruimos la instancia del objeto de audio.
End Sub

Private Sub HScroll1_Change()
    Audio.MUSIC_SetVolume Primary_Buffer, HScroll1.value ' Modificamos el nivel de volumen del canal primario.
End Sub

Private Sub HScroll2_Change()
    Audio.MUSIC_SetPan Primary_Buffer, HScroll2.value ' Modificamos el nivel de balance del canal primario.
End Sub

Private Sub HScroll3_Change()
    Audio.MUSIC_SetRate Primary_Buffer, HScroll3.value ' Modificamos el nivel de velocidad del canal primario.
End Sub
