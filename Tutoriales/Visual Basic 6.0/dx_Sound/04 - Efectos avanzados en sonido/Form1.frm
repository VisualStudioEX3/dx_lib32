VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_Sound - Efectos avanzados en sonidos"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Efecto Ondas de Reverberencia"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2303
      Width           =   2595
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Efecto Gargle"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   2003
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Efecto Flanger"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1703
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Efecto eco"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1403
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Efecto distorsion"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1103
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Efecto compresor"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   803
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Efecto coro"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   503
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Audio As dx_Sound_Class ' Instancia del objeto de audio de dx_lib32.
Private Sample As Long ' Guarda el identificador de la muestra de sonido.
Private FX() As Sound_Effects

Private Sub Form_Load()
    Set Audio = New dx_Sound_Class ' Creamos la instancia del objeto.
    Audio.Init Me.hWnd, 64 ' Inicializamos el motor de audio con 64 canales para efectos de sonido.
    Sample = Audio.SOUND_Load(App.Path & "\sample.wav") ' Cargamos en memoria la muestra de sonido.
    Audio.SOUND_Play Sample, 0, True ' Reproducimos la muestra de sonido en bucle en el canal 0.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Audio.SOUND_Unload Sample ' Descargamos de memoria la muestra de sonido.
    Audio.Terminate ' Terminamos la ejecucion de la clase de audio y liberamos los recursos utilizados.
    Set Audio = Nothing ' Destruimos la instancia del objeto de audio.
End Sub

Private Sub Check1_Click(Index As Integer)
    ' Utilizamos la funcion auxiliar para generar el array con los valores de los efectos seleccionados:
    FX = Audio.SOUND_FX_MakeArrayEffects(CBool(Check1(0).Value), _
                                         CBool(Check1(1).Value), _
                                         CBool(Check1(2).Value), _
                                         CBool(Check1(3).Value), _
                                         CBool(Check1(4).Value), _
                                         CBool(Check1(5).Value), _
                                         CBool(Check1(6).Value) _
                                        )
                                        
    Audio.SOUND_FX_SetEffects 0, FX ' Aplicamos los efectos al canal.
End Sub
