VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_Sound - Reproduccion basica de musica"
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
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2513
      TabIndex        =   2
      Top             =   1313
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2033
      TabIndex        =   1
      Top             =   1313
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1553
      TabIndex        =   0
      Top             =   1313
      Width           =   495
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Audio.MUSIC_Unload Sample ' Descargamos de memoria la muestra de musica.
    Audio.Terminate ' Terminamos la ejecucion de la clase de audio y liberamos los recursos utilizados.
    Set Audio = Nothing ' Destruimos la instancia del objeto de audio.
End Sub

Private Sub Command1_Click()
    Audio.MUSIC_Play Sample, Primary_Buffer ' Reproducimos la muestra en el canal primario de musica.
End Sub

Private Sub Command2_Click()
    Audio.MUSIC_Pause Primary_Buffer ' Pausamos la reproduccion del canal. Si el canal esta en pausa se reanuda la reproduccion.
End Sub

Private Sub Command3_Click()
    Audio.MUSIC_Stop Primary_Buffer ' Detenemos la reproduccion del canal y lo dejamos libre para ser utilizado por otra muestra.
End Sub
