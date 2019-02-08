VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form tercm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JTE CONFIGURATION MANAGER V1.3"
   ClientHeight    =   10320
   ClientLeft      =   5670
   ClientTop       =   780
   ClientWidth     =   17925
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "jte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   17925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBT 
      Caption         =   "BT COM"
      Height          =   270
      Left            =   8880
      TabIndex        =   46
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Interval        =   5
      Left            =   11640
      Top             =   6960
   End
   Begin VB.Frame frame_cmd_diter 
      Caption         =   "DITER COMMANDS"
      Height          =   1092
      Left            =   8760
      TabIndex        =   43
      Top             =   6000
      Visible         =   0   'False
      Width           =   2652
      Begin VB.CommandButton cmd1_diter 
         Caption         =   "CMD1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Index           =   3
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   1092
      End
      Begin VB.CommandButton cmd2_diter 
         Caption         =   "CMD2"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Index           =   2
         Left            =   1440
         TabIndex        =   44
         Top             =   360
         Width           =   1092
      End
   End
   Begin VB.CommandButton tasto_diter 
      Caption         =   "START DITER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   10200
      TabIndex        =   42
      Top             =   2280
      Width           =   1212
   End
   Begin VB.CommandButton comando_stop_seriale 
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   14400
      TabIndex        =   41
      Top             =   0
      Width           =   2412
   End
   Begin VB.CommandButton tasto_debug 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   10920
      TabIndex        =   40
      Top             =   600
      Width           =   492
   End
   Begin VB.CommandButton tasto_abort 
      Caption         =   "ABORT"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   8760
      TabIndex        =   39
      Top             =   9000
      Visible         =   0   'False
      Width           =   2652
   End
   Begin VB.CommandButton comando_clear_seriale 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   11760
      TabIndex        =   38
      Top             =   0
      Width           =   2412
   End
   Begin VB.TextBox text_debug3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   14760
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox text_debug2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   15480
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox text_debug1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   11640
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   7680
      Visible         =   0   'False
      Width           =   5892
   End
   Begin VB.TextBox totale_righe 
      Height          =   396
      Left            =   13800
      TabIndex        =   34
      Top             =   8880
      Width           =   2892
   End
   Begin VB.TextBox totale_byte 
      Height          =   396
      Left            =   13800
      TabIndex        =   33
      Top             =   8280
      Width           =   2892
   End
   Begin RichTextLib.RichTextBox text_report 
      Height          =   6855
      Left            =   1320
      TabIndex        =   30
      Top             =   840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   12091
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"jte.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13680
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frame_update 
      Caption         =   "UPDATE FIRMWARE"
      Height          =   2532
      Left            =   8760
      TabIndex        =   22
      Top             =   3360
      Width           =   2652
      Begin VB.CommandButton tasto_recall_report 
         Caption         =   "RECALL REPORT"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   2412
      End
      Begin VB.ComboBox combo_device 
         Height          =   390
         ItemData        =   "jte.frx":01D1
         Left            =   120
         List            =   "jte.frx":01DB
         TabIndex        =   28
         Text            =   "sel.device"
         Top             =   840
         Width           =   2412
      End
      Begin VB.CommandButton tasto_update 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   2412
      End
      Begin VB.Label label_file_hex 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SELECT FILE"
         Height          =   372
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2412
      End
   End
   Begin VB.TextBox text_macchina 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "jte.frx":01F9
      Top             =   0
      Visible         =   0   'False
      Width           =   6216
   End
   Begin VB.CommandButton tasto_extra 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   13080
      TabIndex        =   20
      Top             =   6600
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   11640
      Top             =   6480
   End
   Begin VB.Frame frame_tabelle 
      Caption         =   "DATA TYPE"
      Height          =   1692
      Left            =   8760
      TabIndex        =   19
      Top             =   7200
      Visible         =   0   'False
      Width           =   2652
      Begin VB.ComboBox combo_tabelle 
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "jte.frx":0217
         Left            =   120
         List            =   "jte.frx":0219
         TabIndex        =   26
         Text            =   "tabelle"
         Top             =   360
         Visible         =   0   'False
         Width           =   2412
      End
      Begin VB.CommandButton tasto_selezione_tabella 
         Caption         =   "SELECT"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   2412
      End
   End
   Begin VB.TextBox seriale 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5772
      Left            =   11640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   600
      Width           =   6132
   End
   Begin VB.Frame frame_com 
      Caption         =   "PC COMM PORT"
      Height          =   1695
      Left            =   8760
      TabIndex        =   11
      Top             =   480
      Width           =   2052
      Begin VB.ComboBox cbmComPort 
         Height          =   390
         ItemData        =   "jte.frx":021B
         Left            =   120
         List            =   "jte.frx":025B
         TabIndex        =   48
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkTTL 
         Caption         =   "TTL"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Frame frame_punti 
      Caption         =   "Load Point"
      Height          =   1092
      Left            =   8760
      TabIndex        =   5
      Top             =   9000
      Visible         =   0   'False
      Width           =   2652
      Begin VB.OptionButton Option1 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   2040
         TabIndex        =   16
         Top             =   720
         Width           =   440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   15
         Top             =   720
         Width           =   440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   14
         Top             =   720
         Width           =   440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   13
         Top             =   720
         Width           =   440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   360
         Width           =   440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   440
      End
   End
   Begin VB.Frame frame_comandi 
      Caption         =   "RESET COMMANDS"
      Height          =   1092
      Left            =   8760
      TabIndex        =   3
      Top             =   6000
      Width           =   2652
      Begin VB.CommandButton comando_reset 
         Caption         =   "SOFT"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Index           =   1
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   1092
      End
      Begin VB.CommandButton comando_reset 
         Caption         =   "HARD"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1092
      End
   End
   Begin MSFlexGridLib.MSFlexGrid win 
      Height          =   9015
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   15901
      _Version        =   393216
      Rows            =   50
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton tasto_exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   8760
      TabIndex        =   1
      Top             =   8880
      Width           =   2652
   End
   Begin VB.CommandButton tasto_uart 
      Caption         =   "START DEBUG"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   8760
      TabIndex        =   0
      Top             =   2280
      Width           =   1212
   End
   Begin MSCommLib.MSComm uart 
      Left            =   12240
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   19200
   End
   Begin VB.Label Label1 
      Caption         =   "Totale righe"
      Height          =   255
      Index           =   1
      Left            =   11760
      TabIndex        =   32
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Totale bytes"
      Height          =   255
      Index           =   0
      Left            =   11760
      TabIndex        =   31
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Label label_timeout 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "timeout"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Shape puntatore 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   12360
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape cornice 
      BorderColor     =   &H000000FF&
      Height          =   375
      Left            =   8760
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "tercm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'str(numero)=stringa del numero
'str(3)=" 3"
'trim(str(3))="3"
'str(123)="123"
'val("3")=3
'val("123")=123
'asc("0")=48
'chr(48)="0"

' comandi PC->micro->
'     risposte micro->PC
' ## blocca trasmissione (buffer pieno)
' #$ sblocca trasmissione
' #: richiesta stringa di intestazione
'     0xf2, intestazione
'     0xf3, nomi tabelle
'     0xf0, nomi variabili tabella
'     0xf1, variabili tabella
' #. invia i dati della tabella corrente
' #R reset


' #P0..9 punti
' #T0..9 invia tabella 0..26

' #+01..99 incrementa
' #-01..99 decrementa
' #*01..99 doppio click
' #/01..99 doppio click





Option Explicit

Dim errore_download As Integer
Dim dimensione_file As Long
Dim errore_comando As Integer


Dim timer_errori As Integer
Dim fase_comandi, resetta_fase_comandi As Integer
Dim timer_comandi As Integer
Dim cicla As Integer
Dim centesimi, secondi, clock_1s As Integer
Dim sec1, sec2, tim1, tim2 As Integer
Dim timeout As Integer
Dim numero_tabelle, indice_tabella, tabella_selezionata As Integer
Dim tabella_con_punti(10) As Integer
Dim blocca_ricezione, sblocca_ricezione As Integer
Dim esegue_timer As Integer
Dim comando_interno As String
Dim comando_esterno(50) As String
Dim indice_comando_esterno As Integer
Dim nome_file_hex_con_path, nome_file_hex_senza_path As String
Dim path_file_hex As String
Dim timer_text_report As Integer
Dim micro As String
Dim timer_programmazione As Integer
Dim stato_seriale As Boolean
Dim salta_caricamento_punto As Boolean
Dim colonna_selezionata, bak_colonna_selezionata As Integer
Dim riga_selezionata, bak_riga_selezionata As Integer
Dim finestra_stringhe As Integer
Dim stop_stringhe As Boolean

Const max_secondi_programmazione = 120

Dim numero_variabili_totale, numero_variabili_parziale As Integer

Dim numero_byte_seriale, numero_stringhe_seriale

Dim aaa, stringa, car As String
Dim contatore, riga, colonna, ii, jj, kk As Integer
Dim posizione, verso As Integer
Dim portacom As Integer
Dim debug1, debug2, debug3 As Integer
Dim debug4, debug5, debug6, debug7, debug8, debug9, debug10 As Integer
Dim debugs1, debugs2, debugs3 As String

Dim focus As Integer

Dim processo_download As Long

Dim in_ricezione As Integer
Dim diter As Boolean
Dim fase_diter, timer_diter As Integer
Dim chk_diter As Integer
Dim comando_diter As String
Dim dit1, dit2, dit3 As Integer
Dim righe_diter As Integer
Dim dit As String


' variabili per isp
Dim rx_cmd, tx_cmd As String





'fine variabili


Sub comandi_diter()

  If timer_programmazione > 0 Then GoTo fine_comandi_diter
  
  If fase_diter = 1 Then
  'trasmetto il primo comando, richiesta nome
    trasmetti_comando_diter ("?h1")
  ElseIf fase_diter = 3 Then
  'trasmetto il secondo comando, richiesta versione
    trasmetti_comando_diter ("?h2")
  ElseIf fase_diter = 5 Then
  'trasmetto il terzo comando, richiesta seriale
    trasmetti_comando_diter ("?h3")
  ElseIf fase_diter = 7 Then
    trasmetti_comando_diter ("?h4")
  ElseIf fase_diter = 9 Then
    trasmetti_comando_diter ("?h5")
  ElseIf fase_diter = 11 Then
    trasmetti_comando_diter ("?h6")
  ElseIf fase_diter = 13 Then
    trasmetti_comando_diter ("?h7")
  ElseIf fase_diter = 15 Then
    trasmetti_comando_diter ("?s1")
  ElseIf fase_diter = 17 Then
    trasmetti_comando_diter ("?s2")
  ElseIf fase_diter = 19 Then
    trasmetti_comando_diter ("?s3")
  ElseIf fase_diter = 21 Then
    trasmetti_comando_diter ("?s4")
  ElseIf fase_diter = 23 Then
    trasmetti_comando_diter ("?s5")
  ElseIf fase_diter = 25 Then
    trasmetti_comando_diter ("?s6")
  ElseIf fase_diter = 27 Then
    trasmetti_comando_diter ("?m1")
  ElseIf fase_diter = 29 Then
  
    If righe_diter < 250 Then
      righe_diter = righe_diter + 1
    Else
      righe_diter = 0
      seriale.Text = ""
      totale_byte = ""
      totale_righe = ""
      numero_byte_seriale = 0
      numero_stringhe_seriale = 0
    End If
    
    trasmetti_comando_diter ("?p")
  ElseIf fase_diter < 50 Then
    If timer_diter < 5 Then
    'If timer_diter < 1 Then
      timer_diter = timer_diter + 1
    Else
      fase_diter = 29
    End If
  ElseIf fase_diter = 100 Then
  'fermato apposta
  Else
  'aspetto ricezione
    If timer_diter < 100 Then
      timer_diter = timer_diter + 1
    Else
      fase_diter = 1
    End If
  End If
fine_comandi_diter:
End Sub
      
      
Private Sub cbmComPort_Change()
    portacom = cbmComPort.ListIndex + 1
End Sub

Private Sub cbmComPort_Click()
    portacom = cbmComPort.ListIndex + 1

End Sub

Private Sub chkBT_Click()
    If chkBT.Value Then
        chkTTL.Enabled = True
    Else
        chkTTL.Value = False
        chkTTL.Enabled = False
    End If
    
End Sub

Private Sub cmd1_diter_Click(Index As Integer)
  'cambio data
  'trasmetti_comando_diter ("!h6:2018-03-15")

  'cambio ora
  'trasmetti_comando_diter ("!h7:09:03:15")

  'cambio numero seriale
  'trasmetti_comando_diter ("!h3:AA18AAA1122")
  
  'comando errato, parametro non modificabile
  'trasmetti_comando_diter ("!h4:AA18AAA1122")
  
  'azzera numero accensioni
  'trasmetti_comando_diter ("!s1:1")
  
  'imposta ore funzionamento in stick
  'trasmetti_comando_diter ("!s2:1")
  
  'richiesta mode[sessione]
  'trasmetti_comando_diter ("?m2")
  
  'richiesta cfg1
  'trasmetti_comando_diter ("?cfg1")
  
  'richiesta cfg1
  'trasmetti_comando_diter ("?cfg10")
  
  'richiesta set corrente stick
  'trasmetti_comando_diter ("?ms10")
  
  '
  trasmetti_comando_diter ("?p2")
  seriale.Text = seriale.Text + vbCrLf + "Tx: " + comando_diter
  
End Sub

Private Sub cmd2_diter_Click(Index As Integer)
  
  'modifica mode[sessione]
  'trasmetti_comando_diter ("!m2:1")
  
  'modifica configurazione
  'trasmetti_comando_diter ("!cfg10:OFF")
  'trasmetti_comando_diter ("!cfg10:ON")
  'trasmetti_comando_diter ("!cfg10:2018-05-12")
  
  'corrente stick
  trasmetti_comando_diter ("!ms10:123")
  
  seriale.Text = seriale.Text + vbCrLf + "Tx: " + comando_diter

End Sub
      
      
      
      
      
      
      
Sub ricevi_risposta_diter()
            
  'cerco e verifico il chk
  dit1 = 1
  While dit1 < Len(aaa)
    If Mid(aaa, dit1, 1) = ";" Then
      dit2 = dit1
    End If
    dit1 = dit1 + 1
  Wend
  dit = Mid(aaa, dit2 + 1, Len(aaa) - dit2)
  
  dit3 = 0
  dit1 = 1
  While dit1 <= dit2
    dit3 = dit3 + Asc(Mid(aaa, dit1, 1))
    dit1 = dit1 + 1
  Wend
              
  If dit3 = Val(dit) Then
  'chk esatto
    dit = Mid(aaa, 1, dit2 - 1)
    If Mid(dit, 1, 3) = "!h1" Then
      'nome macchina
      text_macchina.Text = Mid(dit, 5, Len(dit) - 4)
      text_macchina.Visible = True
    ElseIf Mid(dit, 1, 3) = "!h2" Then
      'versione macchina
      text_macchina.Text = text_macchina.Text + " ver:" + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!h3" Then
      'versione macchina
      text_macchina.Text = text_macchina.Text + " ser:" + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!h4" Then
      'tipo CPU
      text_macchina.Text = text_macchina.Text + " CPU:" + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!h5" Then
      'n. CPU
      text_macchina.Text = text_macchina.Text + " n.CPU:" + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!h6" Then
      'data
      seriale.Text = seriale.Text + vbCrLf + "Data: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!h7" Then
      'ora
      seriale.Text = seriale.Text + vbCrLf + "ora: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!s1" Then
      seriale.Text = seriale.Text + vbCrLf + "n.accensioni: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!s2" Then
      seriale.Text = seriale.Text + vbCrLf + "n.ore funzionamento in MMA: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!s3" Then
      seriale.Text = seriale.Text + vbCrLf + "n.ore funzionamento in TIG DC: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!s4" Then
      seriale.Text = seriale.Text + vbCrLf + "n.ore funzionamento in TIG AC: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!s5" Then
      seriale.Text = seriale.Text + vbCrLf + "n.ore funzionamento in MIG: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!s6" Then
      seriale.Text = seriale.Text + vbCrLf + "n.ore funzionamento in MIG PULSE: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!m1" Then
      seriale.Text = seriale.Text + vbCrLf + "Sessione: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "!m2" Then
      seriale.Text = seriale.Text + vbCrLf + "Mode: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 5) = "!cfg1" Then
      seriale.Text = seriale.Text + vbCrLf + "Config1: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 5) = "!cfg10" Then
      seriale.Text = seriale.Text + vbCrLf + "Config10: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 5) = "!cfg3" Then
      seriale.Text = seriale.Text + vbCrLf + "Config3: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 5) = "!cfg4" Then
      seriale.Text = seriale.Text + vbCrLf + "Config4: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 5) = "!cfg5" Then
      seriale.Text = seriale.Text + vbCrLf + "Config5: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 5) = "!cfg6" Then
      seriale.Text = seriale.Text + vbCrLf + "Config6: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 5) = "!cfg7" Then
      seriale.Text = seriale.Text + vbCrLf + "Config7: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 5) = "!cfg8" Then
      seriale.Text = seriale.Text + vbCrLf + "Config8: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 5) = "!cfg9" Then
      seriale.Text = seriale.Text + vbCrLf + "Config9: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 6) = "!cfg10" Then
      seriale.Text = seriale.Text + vbCrLf + "Config10: " + Mid(dit, 5, Len(dit) - 4)
    ElseIf Mid(dit, 1, 3) = "?e0" Then
      seriale.Text = seriale.Text + vbCrLf + "Ricevuto Chk errato: " + dit
    ElseIf Mid(dit, 1, 3) = "?e1" Then
      seriale.Text = seriale.Text + vbCrLf + "Ricevuto Comando errato: " + dit
    ElseIf Mid(dit, 1, 3) = "?e2" Then
      seriale.Text = seriale.Text + vbCrLf + "Ricevuto Parametro errato: " + dit
    ElseIf Mid(dit, 1, 3) = "?e3" Then
      seriale.Text = seriale.Text + vbCrLf + "Ricevuto Richiesta Modifica Parametro Non Modificabile: " + dit
    ElseIf Mid(dit, 1, 3) = "!p0" Then
      seriale.Text = seriale.Text + vbCrLf + "P.dati vuoto: " + dit
    ElseIf Mid(dit, 1, 3) = "!p1" Then
      seriale.Text = seriale.Text + vbCrLf + "P.dati in stand_by: " + dit
    ElseIf Mid(dit, 1, 3) = "!p2" Then
      seriale.Text = seriale.Text + vbCrLf + "P.dati nuovo set: " + dit
    ElseIf Mid(dit, 1, 3) = "!p3" Then
      seriale.Text = seriale.Text + vbCrLf + "P.dati in ciclo: " + Mid(dit, 1, 30)
    ElseIf Mid(dit, 1, 3) = "!p4" Then
      seriale.Text = seriale.Text + vbCrLf + "P.dati in saldatura: " + Mid(dit, 1, 30)
    ElseIf Mid(dit, 1, 3) = "!p5" Then
      seriale.Text = seriale.Text + vbCrLf + "P.dati in allarme: " + dit
    Else
      seriale.Text = seriale.Text + vbCrLf + "Risposta sconosciuta: " + aaa
    End If
      
    fase_diter = fase_diter + 1
  
  Else
    'chk errato
    seriale.Text = seriale.Text + vbCrLf + "chk errato: " + aaa
    seriale.SelStart = Len(seriale.Text)
  End If
    
  seriale.SelStart = Len(seriale.Text)
  in_ricezione = 100
End Sub
      
      
      

Sub trasmetti_comando_diter(comando)
    comando_diter = comando + ";"
    calcola_chk_diter
    comando_diter = Chr(254) + comando_diter + Trim(Str(chk_diter)) + Chr(255)
    uart.Output = comando_diter
    fase_diter = fase_diter + 1
    timer_diter = 0
End Sub




























Private Sub comando_stop_seriale_Click()
  If diter = True Then
    If fase_diter = 100 Then
      fase_diter = 15
      comando_stop_seriale.Caption = "STOP"
    Else
      fase_diter = 100
      comando_stop_seriale.Caption = "START"
    End If
  Else
  
    If stop_stringhe = True Then
      stop_stringhe = False
      comando_stop_seriale.Caption = "STOP"
    Else
      stop_stringhe = True
      comando_stop_seriale.Caption = "START"
    End If
  End If

End Sub


Private Sub tasto_abort_Click()
  timer_programmazione = 1000
End Sub

Private Sub comando_clear_seriale_Click()
  seriale.Text = ""
  totale_byte = ""
  totale_righe = ""
  numero_byte_seriale = 0
  numero_stringhe_seriale = 0
End Sub

Private Sub comando_reset_Click(Index As Integer)
  If Index = 0 Then
    If chkBT.Value = 1 Then
        invia_comando ("AT+STM32_RST_SEQUENCE=1" + vbCrLf)
        attendi_risposta ("+STM32_RST_SEQUENCE=1" + vbCrLf)
    Else
    
  ' devo eseguire un reset, lo faccio hardware
    uart.DTREnable = True
    centesimi = 0
    While centesimi < 100
      DoEvents
    Wend
    uart.DTREnable = False
    End If
    fase_comandi = 1
    win.Clear
    seriale.Text = ""
    numero_byte_seriale = 0
    numero_stringhe_seriale = 0
    frame_tabelle.Visible = False
    combo_tabelle.Clear
  ElseIf Index = 1 Then
    If (indice_comando_esterno < 40) Then indice_comando_esterno = indice_comando_esterno + 1
    comando_esterno(indice_comando_esterno) = "#R"
    centesimi = 0
    While centesimi < 10
      DoEvents
    Wend
  End If
    
  focus = Index + 1
  
End Sub



Private Sub Form_Load()
  On Error GoTo qui
  
  numero_variabili_totale = 0
  numero_variabili_parziale = 0
  numero_tabelle = 0
  indice_comando_esterno = 0
  focus = 0
  
  contatore = 0
  riga = 1
  colonna = 0
  win.Col = 0
  win.Row = 0
resize:
  tercm.Height = 10150
  tercm.Width = 11700
  win.ColWidth(0) = 5000
  win.ColWidth(1) = 1700
  win.ColWidth(2) = 30
  win.ColWidth(3) = 700
  win.ColWidth(4) = 700
  win.Width = win.ColWidth(0) + win.ColWidth(1) + win.ColWidth(2) + win.ColWidth(3) + win.ColWidth(4) + 360
  
  seriale.Height = 7572
  
  win.ColAlignment(0) = 1   'sinistra basso
  win.ColAlignment(1) = 4   'centro basso
  win.ColAlignment(2) = 4  'centro basso
  win.ColAlignment(3) = 4  'centro basso
  win.ColAlignment(4) = 4  'centro basso
  
  colonna = 1
  GoTo li

qui:
  tasto_uart.Caption = "ERR"
li:
  
  On Error GoTo qui2
  portacom = 1
  micro = "STM32F030F4"
  
  'apro il file coi dati di configurazione
  Open App.Path + "\config.txt" For Input As #1
  
  Input #1, riga
  If riga <> "" Then
    portacom = CInt(riga)
  Else:
    portacom = 1
  End If
  
  Input #1, riga
  If Len(riga) > 1 Then
    micro = riga
    combo_device.Text = micro
  Else
   combo_device.ListIndex = Val(riga)
  End If
  
  Input #1, nome_file_hex_con_path
  If nome_file_hex_con_path > "" Then
  
    'tolgo il path per visualizzare solo il nome del file
    Dim tt, ii As Integer
    nome_file_hex_senza_path = nome_file_hex_con_path
    ii = Len(nome_file_hex_senza_path)
    If ii > 0 Then
    
      tt = 1
      While tt > 0
        If Mid(nome_file_hex_senza_path, ii, 1) <> "\" Then
          If ii > 1 Then
            ii = ii - 1
          Else
            tt = 0
          End If
        Else
          ii = ii + 1
          nome_file_hex_senza_path = Mid(nome_file_hex_senza_path, ii, Len(nome_file_hex_senza_path) - ii + 1)
          nome_file_hex_senza_path = Left(nome_file_hex_senza_path, Len(nome_file_hex_senza_path) - 4)
          path_file_hex = Left(nome_file_hex_con_path, ii - 1)
          tt = 0
        End If
      Wend
    
    End If
  
    label_file_hex.Caption = nome_file_hex_senza_path

  End If
  
  Input #1, riga
  If Len(riga) > 0 Then
    If riga = "1" Then
      finestra_stringhe = 1
      tercm.Width = 18000
      tasto_debug.Caption = "<-"
    End If
  End If
  
  'tercm.StartUpPosition = 2
  Input #1, riga
  chkBT.Value = CInt(riga)
  Input #1, riga
  chkTTL.Value = CInt(riga)
  
  Close #1
  GoTo li2

qui2:


li2:
  
  cbmComPort.ListIndex = portacom - 1
  'Option2(portacom - 1).Value = True

End Sub


Private Sub label_file_hex_Click()
  CommonDialog1.Filter = "Text files|*.hex"
  CommonDialog1.ShowOpen
  
  Dim file_da_prog As String
  
  file_da_prog = CommonDialog1.FileName
  If Len(file_da_prog) > 4 Then
  
    nome_file_hex_con_path = file_da_prog
    
    'tolgo il path per visualizzare solo il nome del file
    Dim tt, ii As Integer
    nome_file_hex_senza_path = nome_file_hex_con_path
    ii = Len(nome_file_hex_senza_path)
    If ii > 0 Then
    
      tt = 1
      While tt > 0
        If Mid(nome_file_hex_senza_path, ii, 1) <> "\" Then
          If ii > 0 Then
            ii = ii - 1
          Else
            tt = 0
          End If
        Else
          ii = ii + 1
          nome_file_hex_senza_path = Mid(nome_file_hex_senza_path, ii, Len(nome_file_hex_senza_path) - ii + 1)
          nome_file_hex_senza_path = Left(nome_file_hex_senza_path, Len(nome_file_hex_senza_path) - 4)
          path_file_hex = Left(nome_file_hex_con_path, ii - 1)
          tt = 0
        End If
      Wend
    
    End If
    label_file_hex.Caption = nome_file_hex_senza_path
  End If
End Sub


'sono i punti da caricare
Private Sub Option1_Click(Index As Integer)
  If salta_caricamento_punto Then
    salta_caricamento_punto = False
  Else
    If (indice_comando_esterno < 40) Then indice_comando_esterno = indice_comando_esterno + 1
    comando_esterno(indice_comando_esterno) = "#T" + Chr(Asc("0") + Index)
  End If
End Sub

Private Sub Option2_Click(Index As Integer)
   portacom = Index + 1
End Sub


Private Sub tasto_debug_Click()
  If tercm.Width = 18000 Then
    tercm.Width = 11700
    tasto_debug.Caption = "->"
    finestra_stringhe = 0
  Else
    tercm.Width = 18000
    tasto_debug.Caption = "<-"
    finestra_stringhe = 1
  End If
  
  If focus = 0 Then
    tasto_uart.SetFocus
  ElseIf focus = 1 Then
    comando_reset(0).SetFocus
  ElseIf focus = 2 Then
    comando_reset(1).SetFocus
  Else
    focus = 0
    tasto_uart.SetFocus
  End If
  

End Sub

Private Sub tasto_diter_Click()
  
  On Error GoTo seriale_occupata
  If uart.PortOpen = False Then
    diter = 1
    fase_diter = 0
    tasto_uart.Enabled = False
    frame_comandi.Enabled = False
    frame_comandi.Visible = False
    frame_cmd_diter.Enabled = True
    frame_cmd_diter.Visible = True
    text_macchina.FontSize = 10
    apri_seriale
    tasto_diter.Caption = "STOP"
    cornice.Visible = True
    posizione = cornice.Left
    verso = 1
    puntatore.Left = posizione
    puntatore.Top = cornice.Top
    puntatore.Visible = True
'    tasto_extra.Visible = True
'    comando(0).Visible = False
  Else
    tasto_uart.Enabled = True
    frame_comandi.Enabled = True
    frame_comandi.Visible = True
    frame_cmd_diter.Enabled = False
    frame_cmd_diter.Visible = False
    text_macchina.FontSize = 16
    chiudi_seriale
    tasto_diter.Caption = "START DITER"
    cornice.Visible = False
    puntatore.Visible = False
'    tasto_extra.Visible = False
'    comando(0).Visible = True
  End If
  GoTo fine_tasto_uart
  
seriale_occupata:
  tasto_uart.Caption = "COMM ERROR"
  timer_errori = 30
fine_tasto_uart:

  focus = 0
  

End Sub

Private Sub tasto_extra_Click()
'stampo la riga

End Sub

Private Sub tasto_recall_report_Click()
  If text_report.Visible = True Then
    text_report.Visible = False
    timer_text_report = 0
  Else
    text_report.Visible = True
    timer_text_report = 150
  End If
End Sub

'cambio tabella
Private Sub tasto_selezione_tabella_Click()
  If (indice_comando_esterno < 40) Then indice_comando_esterno = indice_comando_esterno + 1
  comando_esterno(indice_comando_esterno) = "#T" + Trim(Str(combo_tabelle.ListIndex))
  tabella_selezionata = combo_tabelle.ListIndex
  resetta_fase_comandi = 3
End Sub

Private Sub tasto_uart_Click()
  On Error GoTo seriale_occupata
  If uart.PortOpen = False Then
    tasto_diter.Enabled = False
    diter = 0
    apri_seriale
    tasto_uart.Caption = "STOP"
    cornice.Visible = True
    posizione = cornice.Left
    verso = 1
    puntatore.Left = posizione
    puntatore.Top = cornice.Top
    puntatore.Visible = True
  tasto_update.Enabled = True
'    tasto_extra.Visible = True
'    comando(0).Visible = False
  Else
    tasto_diter.Enabled = True
    chiudi_seriale
    tasto_uart.Caption = "START DEBUG"
    cornice.Visible = False
    puntatore.Visible = False
'  tasto_update.Enabled = False
'    tasto_extra.Visible = False
'    comando(0).Visible = True
  End If
  GoTo fine_tasto_uart
  
seriale_occupata:
  tasto_uart.Caption = "COMM ERROR"
  timer_errori = 30
fine_tasto_uart:

  focus = 0
  
End Sub

Private Sub tasto_exit_Click()
  If uart.PortOpen = True Then
    chiudi_seriale
  End If
  
  
  On Error GoTo qui
  Close
  
  Open App.Path + "\config.txt" For Output As #1
    If (portacom >= 1 And portacom <= 20) Then
        Print #1, Str(portacom)
    Else
        Print #1, "1"
    End If
    
  
  If combo_device.ListIndex = -1 Then
    If combo_device.Text <> "" Then
      Print #1, combo_device.Text
    Else
      Print #1, Str(combo_device.ListIndex)
    End If
  Else
    Print #1, Str(combo_device.ListIndex)
  End If
  
  'salvo il nome file
  If nome_file_hex_con_path > "" Then Print #1, nome_file_hex_con_path
  
  If finestra_stringhe = 1 Then
    Print #1, "1"
  Else
    Print #1, "0"
  End If
  
  'Bluetooth e TTL
  Print #1, Str(chkBT.Value)
  Print #1, Str(chkTTL.Value)
  
  
  
  Close #1
qui:
  End
End Sub

Private Sub tasto_update_Click()
  If timer_programmazione > 0 Then
    timer_programmazione = 1000
  Else
      
    If uart.PortOpen = True Then
      stato_seriale = True
      chiudi_seriale
      cornice.Visible = False
      text_macchina.Visible = False
      puntatore.Visible = False
      'inserisco una pausa
      centesimi = 0
      While centesimi < 10
        DoEvents
      Wend
    Else
      stato_seriale = False
    End If
    
    On Error Resume Next
    
    timer_programmazione = 1
    
    
    'tasto_uart.Enabled = False
    frame_com.Enabled = False
    frame_punti.Enabled = False
    frame_tabelle.Enabled = False
    frame_update.Enabled = False
    tasto_update.Caption = "prog..."
    tasto_abort.Enabled = True
    tasto_abort.Visible = True
    
    
    
    If combo_device.ListIndex >= 0 Then
      micro = combo_device.Text
    End If
    
    'faccio delle verifiche!!!
    If micro = "" Then micro = "STM32F303VE"
    
    If path_file_hex = "" Or nome_file_hex_senza_path = "" Then
      MsgBox "File non specificato o inesistene", vbCritical + vbOKOnly, "Errore!"
    Else
          
      Dim comando_update As String
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      'versione con downloader esterno.
'      If micro = "STM32F303VE" Then
'         comando_update = App.Path + "\prog\prog.bat " + " " + path_file_hex + " " + nome_file_hex_senza_path + " " + App.Path + "\"
'      ElseIf micro = "STM32F030F4" Then
'         comando_update = App.Path + "\prog\prog.bat " + " " + path_file_hex + " " + nome_file_hex_senza_path + " " + App.Path + "\"
'      Else
'         comando_update = App.Path + "\fm.bat " + Trim(Str(portacom)) + " " + micro + " " + path_file_hex + " " + nome_file_hex_senza_path + " " + App.Path + "\"
'      End If
'
'        processo_download = Shell(comando_update, vbNormalFocus)
'
'      If processo_download = 0 Then
'        MsgBox "Si è verificato un errore nell'apertura", vbCritical + vbOKOnly, "Errore!"
'        timer_programmazione = 1000
'      End If
      '...........fino a qui
      
      
      
    
      'versione da fare...
      text_report.Text = ""
      timer_text_report = 1000
      text_report.Visible = True
      programma_micro
      
      
      label_timeout.Visible = False
      tasto_update.Caption = "UPDATE"
      timer_programmazione = 0

      comando_reset_Click (0)

      tasto_uart.Enabled = True
      frame_com.Enabled = True
      frame_punti.Enabled = True
      frame_tabelle.Enabled = True
      frame_update.Enabled = True
      tasto_abort.Enabled = False
      tasto_abort.Visible = False

      tercm.SetFocus
      
    
    End If
    
    centesimi = 0
    While centesimi < 100
      DoEvents
    Wend
    
    text_report.Visible = False
    'aspetto la fine del download, ma metto un timeout
  End If

End Sub


Private Sub combo_device_LostFocus()
  micro = combo_device.Text
End Sub


Private Sub text_report_Click()
  text_report.Visible = False
End Sub


'Private Sub Timer2_Timer()
'
'  ms = ms + 1
'
'End Sub

'timer 10ms
'Private Sub Timer_Timer()

'End Sub
'timer 100ms
Private Sub Timer1_Timer()
  'ogni 100ms
  
  Dim ll As Integer
  Dim giri_al_secondo As Integer
  
  On Error GoTo fine_timer1
  
  giri_al_secondo = 20
  
  'questa roba serve per adattare il valore del timer un modo da avere un dato il + possibile giusto
    
  tim1 = tim1 + 1
  
  sec1 = Second(Now)
  If sec1 = sec2 Then
    clock_1s = 0
  Else
    clock_1s = 1
    
    If secondi < 20000 Then
      secondi = secondi + 1
    Else
      secondi = 10000
    End If
    
    sec2 = sec1

debug1 = tim1
debug2 = Timer1.Interval
        
    tim2 = Timer1.Interval
    
    If tim1 < giri_al_secondo Then
      tim1 = giri_al_secondo - tim1
      tim1 = (tim1 * 100) / giri_al_secondo
      tim2 = tim2 - (tim1 * tim2) / 100
      If tim2 < 5 Then tim2 = 5
    ElseIf tim1 > giri_al_secondo Then
      tim1 = tim1 - giri_al_secondo
      tim1 = (tim1 * 100) / giri_al_secondo
      tim2 = tim2 + (tim1 * tim2) / 100
      If tim2 > 250 Then tim2 = 250
    End If
      
    Timer1.Interval = tim2

debug3 = Timer1.Interval
    
    tim1 = 0
  
  End If
  
  
  'gestione barra di visualizzazione scambio dati
  If in_ricezione > 0 Then
    in_ricezione = in_ricezione - 1
  
    If verso = 1 Then
      If posizione < cornice.Left + cornice.Width - puntatore.Width Then
        posizione = posizione + 10
      Else
        posizione = posizione - 10
        verso = 0
      End If
    Else
      If posizione > cornice.Left Then
        posizione = posizione - 10
      Else
        posizione = posizione + 10
        verso = 1
      End If
    End If

  Else
  
  End If
  puntatore.Left = posizione

'text_debug1.Text = debug1
'text_debug2.Text = debug2
'text_debug3.Text = debug3
  
  
  
'
'  If timer_programmazione > 0 Then
'    If clock_1s Then
'      timer_programmazione = timer_programmazione + 1
'      label_timeout.Visible = True
'      If timer_programmazione < 295 Then
'        label_timeout.Caption = "Time:" + Str(timer_programmazione)
'      ElseIf timer_programmazione < max_secondi_programmazione Then
'        label_timeout.Caption = "RETRY!!!"
'      End If
'    End If
'
'
    
'    If timer_programmazione < 5 Then
'    ElseIf Len(Dir(path_file_hex + "report.txt")) > 0 Or timer_programmazione > 300 Then
'      label_timeout.Visible = False
'      tasto_update.Caption = "UPDATE"
'      timer_programmazione = 0
'
'      comando_reset_Click (0)
'
'      tasto_uart.Enabled = True
'      frame_com.Enabled = True
'      frame_punti.Enabled = True
'      frame_tabelle.Enabled = True
'      frame_update.Enabled = True
'      tasto_abort.Enabled = False
'      tasto_abort.Visible = False
'
'      tercm.SetFocus
'
      'visualizzo il risultato
'      If timer_programmazione > 300 Then GoTo nofile
'      On Error GoTo nofile
'
'      Dim ss As String
'      Open path_file_hex + "report.txt" For Input As #1
'      text_report.Text = ""
'      While Not EOF(1)
'
'        Input #1, ss
'
'        text_report.Text = text_report.Text + ss
'        text_report.Text = text_report.Text + vbCrLf
'      Wend
'      Close #1
'
'
'      'cerco le stringhe da evidenziare
'      If InStr(text_report.Text, "Erase complete") Then
'        text_report.SelStart = InStr(text_report.Text, "Erase complete") - 1
'        text_report.SelLength = Len("Erase complete")
'        text_report.SelBold = True
'        text_report.SelColor = vbGreen
'      ElseIf InStr(text_report.Text, "Erase failed") Then
'        text_report.SelStart = InStr(text_report.Text, "Erase failed") - 1
'        text_report.SelLength = Len("Erase failed")
'        text_report.SelBold = True
'        text_report.SelColor = vbRed
'      End If
'
'      If InStr(text_report.Text, "Hex file programming complete") Then
'        text_report.SelStart = InStr(text_report.Text, "Hex file programming complete") - 1
'        text_report.SelLength = Len("Hex file programming complete")
'        text_report.SelBold = True
'        text_report.SelColor = vbGreen
'      ElseIf InStr(text_report.Text, "Hex file programming failed") Then
'        text_report.SelStart = InStr(text_report.Text, "Hex file programming failed") - 1
'        text_report.SelLength = Len("Hex file programming failed")
'        text_report.SelBold = True
'        text_report.SelColor = vbRed
'      End If
'
'      If InStr(text_report.Text, "Connection failed") Then
'        text_report.SelStart = InStr(text_report.Text, "Connection failed") - 1
'        text_report.SelLength = Len("Connection failed")
'        text_report.SelBold = True
'        text_report.SelColor = vbRed
'      End If
'
'
'      text_report.Visible = True
'      timer_text_report = 10
'      tasto_recall_report.Enabled = True
'nofile:
'      Close
'      'pulisco lo schermo
'      win.Clear
'      seriale.Text = ""
'      numero_byte_seriale = 0
'      numero_stringhe_seriale = 0
'      frame_tabelle.Visible = False
'      combo_tabelle.Clear
            
'      If timer_programmazione > 0 Then
'      ElseIf stato_seriale = True Then
'        If diter = True Then
'          tasto_diter_Click
'        Else
'          tasto_uart_Click
'        End If
'        fase_comandi = 0
'        timer_comandi = 0
'      End If
    
    
'      ' devo eseguire un reset, lo faccio hardware
'    uart.DTREnable = True
'    centesimi = 0
'    While centesimi < 10
'      DoEvents
'    Wend
'    uart.DTREnable = False
'    fase_comandi = 1
'    win.Clear
'    seriale.Text = ""
'    numero_byte_seriale = 0
'    numero_stringhe_seriale = 0
'    frame_tabelle.Visible = False
'    combo_tabelle.Clear
'
'
'
'
'    End If
'  End If
  
  



  If tabella_con_punti(tabella_selezionata) > 0 Then
    If frame_punti.Visible = False Then
      frame_punti.Visible = True
      salta_caricamento_punto = True
      Option1(0) = 1
    End If
    If Option1(0) = 0 And Option1(1) = 0 And Option1(2) = 0 And Option1(3) = 0 And Option1(4) = 0 And Option1(5) = 0 And Option1(6) = 0 And Option1(7) = 0 And Option1(8) = 0 And Option1(9) = 0 Then Option1(0) = 1
  Else
    frame_punti.Visible = False
  End If

  
  
  If timer_errori > 1 Then
    timer_errori = timer_errori - 1
  ElseIf timer_errori = 1 Then
    tasto_uart.Caption = "START"
    timer_errori = 0
  End If



  If timer_text_report > 0 Then
    timer_text_report = timer_text_report - 1
  Else
    timer_text_report = 100
    text_report.Visible = False
  End If

  
  
  If uart.PortOpen = True Then
    frame_comandi.Enabled = True
    comando_reset(0).Enabled = True
    comando_reset(1).Enabled = True
    frame_com.Enabled = False
  Else
    frame_comandi.Enabled = False
    comando_reset(0).Enabled = False
    comando_reset(1).Enabled = False
    frame_com.Enabled = True
  End If

  
  
'  If label_file_hex.Caption = "SELECT FILE" Then
'    tasto_update.Enabled = False
'  Else
'    tasto_update.Enabled = True
'  End If





  'gestione comandi trasmissione e timeout ricezione risposte
  If uart.PortOpen = True Then
    
    If diter = True Then
      If fase_diter = 0 Then
        
        text_macchina.Visible = False
        label_timeout.Visible = False
        win.Clear
        seriale.Text = ""
        numero_byte_seriale = 0
        numero_stringhe_seriale = 0
        frame_tabelle.Visible = False
        combo_tabelle.Clear
        stringa = ""
        fase_diter = 1
      End If
    
    ElseIf timer_programmazione > 0 Then

    Else
    
      'gestione delle fasi di trasmissione
      If fase_comandi = 0 Then
        'richiesta versione
        
        text_macchina.Visible = False
        label_timeout.Visible = False
        'win.Clear
        'seriale.Text = ""
        numero_byte_seriale = 0
        numero_stringhe_seriale = 0
        frame_tabelle.Visible = False
        combo_tabelle.Clear
        stringa = ""
        
        If timer_comandi < 5 Then
          timer_comandi = timer_comandi + 1
        Else
          timer_comandi = 0
          fase_comandi = 1
          comando_interno = "#:"
        End If
      
      ElseIf fase_comandi = 1 Then
        'aspetto ricezione stringa macchina, altrimenti ripeto comando
        
        If timer_comandi < 20 Then
          timer_comandi = timer_comandi + 1
        Else
          timer_comandi = 0
          fase_comandi = 0
        End If
      
      ElseIf fase_comandi = 2 Then
        'se sono qui significa che ho ricevuto risposta al comando precedente
        
        comando_interno = "#T0"
        fase_comandi = 3
        timer_comandi = 0
      
      ElseIf fase_comandi = 3 Then
        'aspetto ricezione nomi variabili, altrimenti ripeto comando
        
        If in_ricezione Then
          timer_comandi = 0
        ElseIf timer_comandi < 20 Then
          timer_comandi = timer_comandi + 1
        Else
          timer_comandi = 0
          fase_comandi = 0
        End If
      
      ElseIf fase_comandi = 4 Then
        'se sono qui significa che ho ricevuto risposta al comando precedente
  'debug2 = centesimi - debug1
        comando_interno = "#."
        fase_comandi = 5
        timer_comandi = 0
      
      ElseIf fase_comandi = 5 Then
        'aspetto ricezione valori variabili, altrimenti ripeto comando
        If in_ricezione Then
          timer_comandi = 0
        ElseIf timer_comandi < 20 Then
          timer_comandi = timer_comandi + 1
        Else
          timer_comandi = 0
          fase_comandi = 0
        End If
      
      Else
        fase_comandi = 0
        timer_comandi = 0
      End If
    
    End If
  Else
    fase_comandi = 0
    timer_comandi = 0
    If text_macchina.Visible = True Then
      If label_timeout.Visible = False Then
        timeout = 10
        label_timeout.Caption = "Timeout: " + Trim(Str(timeout))
        label_timeout.Visible = True
      ElseIf clock_1s > 0 Then
        If timeout > 0 Then
          timeout = timeout - 1
          label_timeout.Caption = "Timeout: " + Trim(Str(timeout))
        Else
          text_macchina.Visible = False
          label_timeout.Visible = False
'          win.Clear
'          seriale.Text = ""
          numero_byte_seriale = 0
          numero_stringhe_seriale = 0
          frame_tabelle.Visible = False
          combo_tabelle.Clear
        End If
      End If
    End If
  End If


  'gestione della ricezione

  If timer_programmazione > 0 Then
    stringa = ""
  ElseIf uart.PortOpen = True Then
    aaa = uart.Input
    If Len(aaa) > 0 Then
      stringa = stringa + aaa
    End If
  End If
  
  
  ll = Len(stringa)
  'se sono troppo veloce a spedire e la stringa diventa troppo
  'lunga taglio via la parte piu' vecchia
  If ll > 5000 Then
    stringa = Right(stringa, 1000)
    ll = 900
  ElseIf ll > 2000 Then
    blocca_ricezione = 1
    sblocca_ricezione = 0
  ElseIf blocca_ricezione Then
    blocca_ricezione = 0
    sblocca_ricezione = 1
  End If
  
        
        
  ' qui trasmetto cio' che è in attesa, i comandi esterni hanno la precedenza
  If timer_programmazione > 0 Then
  ElseIf uart.PortOpen = True Then
    If blocca_ricezione Then
      uart.Output = "##"
    ElseIf sblocca_ricezione Then
      uart.Output = "#$"
      sblocca_ricezione = 0
    Else
    
      
      If diter = True Then
        comandi_diter
        
        'qui la gestione della trasmissione dei comandi alla scheda in emulazione diter
        
      
      
      
      
      
      Else
            
        If comando_esterno(indice_comando_esterno) > "" Then
          'aspetto la fine della ricezione del pacchetto precedente
          If comando_interno = "#." Then
            uart.Output = comando_esterno(indice_comando_esterno)
            comando_esterno(indice_comando_esterno) = ""
            If indice_comando_esterno > 0 Then indice_comando_esterno = indice_comando_esterno - 1
            If resetta_fase_comandi > 0 Then
              fase_comandi = resetta_fase_comandi
              resetta_fase_comandi = 0
              comando_interno = ""
            End If
          End If
        End If
        
        If comando_interno > "" Then
          uart.Output = comando_interno
          comando_interno = ""
  'debug3 = centesimi - debug1
        End If
      End If
    End If
  End If
        
        
        
  ll = Len(stringa)
        
  If diter = True Then
    'qqq
  End If
            
            
  If ll = 0 Then
  Else
    'analizzo e sistemo la stringa
    kk = 0
    While kk < ll
    
      'innanzitutto cerco un carattere di inizio stringa (240, 241, 242, 243, 244)
      ii = Asc(Left(stringa, 1))
      If ii = 240 Or ii = 241 Or ii = 242 Or ii = 243 Or ii = 244 Or ii = 254 Then
        'ho trovato un inizio stringa
          kk = Len(stringa) + 1
      Else
        'se il primo carattere non e' un inizio stringa lo elimino
        stringa = Right(stringa, ll - 1)
        ll = Len(stringa)
      End If
    Wend
    
    'adesso la stringa inizia con un carattere di inizio stringa oppure è rimasta una stringa vuota
    
    
    ll = Len(stringa)
    If ll > 0 Then
      cicla = 1
    Else
      cicla = 0
    End If
    
    
    While cicla > 0
      
      cicla = 0
      
      'prendo il primo carattere e vedo cosa è o sta arrivando
      jj = Asc(Left(stringa, 1))
  
  
      If jj = 254 And ll > 2 Then
      'stringa diter
        
        'cerco 0xff di fine stringa
        
        ii = 2
        aaa = ""
        kk = ll + 1
        While ii <= ll
          car = Mid(stringa, ii, 1)
          If car = Chr(255) Then
            'ho trovato la fine della stringa, la analizzo

            ricevi_risposta_diter
            
            stringa = Right(stringa, ll - ii)
          
          Else
            aaa = aaa + car
          End If
          ii = ii + 1
        Wend
      
          
          
          
          
          
          
          
          indice_tabella = 0
          numero_variabili_totale = 0
          numero_variabili_parziale = 0
          
        
        
        
        
        
        
        
        
        
        
        
        
      'nomi delle variabili
      ElseIf jj = 240 And ll > 1 Then
        
        If Asc(Mid(stringa, 2, 1)) = 255 Then
          'ho trovato l'ultima stringa
  
          If fase_comandi = 3 Then fase_comandi = 4
          stringa = Right(stringa, ll - 2)
          ll = Len(stringa)
          If ll > 0 Then cicla = 1
      
        ElseIf ll > 2 Then
        
          'ho trovato l'inizio di una stringa che contiene il nome della variabile, ne cerco la fine
          'estraggo la posizione
            
          riga = Asc(Mid(stringa, 2, 1))
          'se la posizione è la prima pulisco tutto
          If riga = 0 Then
            win.Clear
            seriale.Text = ""
            numero_byte_seriale = 0
            numero_stringhe_seriale = 0
          
          End If
            
          'faccio un test sulle righe
          If riga > win.Rows - 1 Then riga = win.Rows - 1
            
          colonna = 0
          ii = 3
          aaa = ""
          kk = ll + 1
            
          While ii < ll
            car = Mid(stringa, ii, 1)
            If car = Chr(255) Then
              'ho trovato la fine della stringa, la stampo
              win.Col = 0
              win.Row = riga
              If Right(aaa, 1) = "0" Then
              ElseIf Right(aaa, 1) = "1" Then
                win.Col = 3
                win.Row = riga
                win.Text = "-"
                win.Col = 4
                win.Text = "+"
              ElseIf Right(aaa, 1) = "2" Then
                win.Col = 3
                win.Row = riga
                win.Text = "OFF"
                win.Col = 4
                win.Text = "ON"
              End If
              win.Col = 0
              aaa = Left(aaa, Len(aaa) - 1)
              win.Text = aaa
              If riga = 0 Then
                win.CellFontBold = True
              End If
              stringa = Right(stringa, ll - ii)
              ll = Len(stringa)
              If ll > 0 Then cicla = 1
              ii = ll + 1
              kk = 0
              numero_variabili_totale = riga + 1
              numero_variabili_parziale = 0
              'aggiorna_puntatore
              in_ricezione = 100
                
            Else
              aaa = aaa + car
            End If
            ii = ii + 1
          Wend
        End If
      
      'valori delle variabili
      ElseIf jj = 241 And ll > 1 Then
        
        If Asc(Mid(stringa, 2, 1)) = 255 Then
          'ho trovato l'ultima stringa
          If fase_comandi = 5 Then
            fase_comandi = 4
          End If
          stringa = Right(stringa, ll - 2)
          ll = Len(stringa)
          If ll > 0 Then cicla = 1
        
        ElseIf ll > 2 Then
          'ho trovato l'inizio di una variabile (in formato stringa), ne cerco la fine
          riga = Asc(Mid(stringa, 2, 1))
          If riga > win.Rows Then
            If riga = 100 Then
              tasto_update.Enabled = False
            ElseIf riga = 101 Then
              tasto_update.Enabled = False
            ElseIf riga = 102 Then
              tasto_update.Enabled = True
            Else
              riga = win.Rows
            End If
          End If
          
          If riga < 100 Then
            bak_colonna_selezionata = win.ColSel
            bak_riga_selezionata = win.RowSel
          End If
          
          ii = 3
          aaa = ""
          kk = ll + 1
          While ii < ll
            car = Mid(stringa, ii, 1)
            If car = Chr(255) Then
              'ho trovato la fine della stringa, la stampo
              If riga < 100 Then
                win.Col = 1
                win.Row = riga
                If aaa = "" Then aaa = "0"
                win.Text = aaa
                win.CellFontBold = True
              Else
                riga = win.Rows
              End If
                
              stringa = Right(stringa, ll - ii)
              ll = Len(stringa)
              If ll > 0 Then cicla = 1
              ii = ll + 1
              kk = 0
              in_ricezione = 100
              If riga < 100 Then
                numero_variabili_parziale = riga + 1
              End If
              If numero_variabili_parziale = numero_variabili_totale Then
                If fase_comandi = 5 Then
                  fase_comandi = 4
  'debug1 = centesimi
                End If
                numero_variabili_parziale = 0
              End If
              
              If bak_colonna_selezionata < win.Cols Then win.Col = bak_colonna_selezionata
              If bak_riga_selezionata < win.Rows Then win.Row = bak_riga_selezionata
            Else
              aaa = aaa + car
            End If
            ii = ii + 1
          Wend
        End If
        
      ElseIf jj = 242 And ll > 2 Then
        'ho trovato l'inizio di una stringa che contiene il nome della macchina, ne cerco la fine
        ii = 2
        aaa = ""
        kk = ll + 1
        While ii <= ll
          car = Mid(stringa, ii, 1)
          If car = Chr(255) Then
            'ho trovato la fine della stringa, la stampo
            text_macchina.Text = aaa
            text_macchina.Visible = True
            in_ricezione = 100
            stringa = Right(stringa, ll - ii)
            If fase_comandi = 1 Then fase_comandi = 2
          Else
            aaa = aaa + car
          End If
          
          indice_tabella = 0
          numero_variabili_totale = 0
          numero_variabili_parziale = 0
          
          ii = ii + 1
        Wend
      
      ElseIf jj = 243 And ll > 2 Then
        'ho trovato l'inizio di una stringa che contiene i nomi delle tabelle
        ii = 2
        aaa = ""
        kk = ll + 1
        While ii <= ll
          car = Mid(stringa, ii, 1)
          If car = Chr(255) Then
            Dim indice As Integer
                      
            'ho trovato la fine della stringa, la gestisco
            'estraggo l'indice
            car = Left(aaa, 1)
            'tolgo l'indice
            aaa = Right(aaa, Len(aaa) - 1)
            indice = Asc(car)
            'estraggo la stringa
            If indice >= indice_tabella Then
              indice_tabella = indice_tabella + 1
              frame_tabelle.Visible = True
              combo_tabelle.AddItem Left(aaa, Len(aaa) - 1)
              combo_tabelle.Visible = True
              If indice = 0 Then
                combo_tabelle.ListIndex = 0
                tabella_selezionata = 0
              End If
            End If
            'vedo se la tabella è associata alla gestione dei punti
            car = Right(aaa, 1)
            If Asc(car) = 1 Then
              tabella_con_punti(indice) = 1
            Else
              tabella_con_punti(indice) = 0
            End If
            
            stringa = Right(stringa, ll - ii)
            timer_comandi = 0
            in_ricezione = 100
            
            ll = Len(stringa)
            If ll > 0 Then cicla = 1
            ii = ll + 1
            
          Else
            aaa = aaa + car
          End If
          
          ii = ii + 1
        Wend
      
      ElseIf jj = 244 And ll > 2 Then
        'ho trovato l'inizio di una stringa che contiene una stringa da visualizzare
        'sulla finestra delle stringhe
        
        
        timer_comandi = 0
        in_ricezione = 100
        
        ii = 2
        aaa = ""
        kk = ll + 1
        While ii <= ll
          car = Mid(stringa, ii, 1)
          If car = Chr(255) Then
            
            numero_byte_seriale = numero_byte_seriale + Len(aaa)
            numero_stringhe_seriale = numero_stringhe_seriale + 1
            
            'ho trovato la fine della stringa, la stampo
            If stop_stringhe = True Then
                    
            Else
              seriale.Text = seriale.Text + vbCrLf + aaa
              seriale.SelStart = Len(seriale.Text)
            End If
            
            timer_comandi = 0
            in_ricezione = 100
            stringa = Right(stringa, ll - ii)
          Else
            aaa = aaa + car
          End If
          
          ii = ii + 1
        Wend
          
          
          
          
          
        totale_byte.Text = numero_byte_seriale
        totale_righe.Text = numero_stringhe_seriale
              
      Else
        'fine delle ricerche
        kk = ll + 1
      End If
      
    Wend

  End If
  
fine_timer1:
  
'debug1 = fase_comandi
'debug2 = timer_comandi
'debug3 = debug2 - debug1
'deb1.Text = debug1
'deb2.Text = debug2
'deb3.Text = debug3
  
End Sub
Private Sub apri_seriale_per_isp()
  If uart.PortOpen = False Then
    uart.CommPort = portacom
    uart.Settings = "115200,e,8,1"
    
    uart.DTREnable = False
    
    
    
    'If micro = "STM32F303VE" Then
     ' uart.RTSEnable = False
    'Else
    '  uart.RTSEnable = True
    'End If
    uart.EOFEnable = False
    uart.Handshaking = comNone
    uart.InBufferSize = 1024
    uart.InputMode = comInputModeText
    uart.InputLen = 0
    uart.RThreshold = 100
    uart.SThreshold = 100
    uart.PortOpen = True
  
    'Modifiche effettuate per poter gestire l'interfaccia BT seriale
    If chkBT.Value = 1 Then
        If chkTTL.Value = 1 Then
            invia_comando ("AT+STYPE=T" + vbCrLf)
            attendi_risposta ("+STYPE=T" + vbCrLf)
        Else
            invia_comando ("AT+STYPE=R" + vbCrLf)
            attendi_risposta ("+STYPE=R" + vbCrLf)
        End If
        
        'setta il baud rate
        invia_comando ("AT+BRATE=115200" + vbCrLf)
        attendi_risposta ("+BRATE=115200" + vbCrLf)
        'setta i data bits
        invia_comando ("AT+DBITS=8" + vbCrLf)
        attendi_risposta ("+DBITS=8" + vbCrLf)
        'setta gli stop bits
        invia_comando ("AT+SBITS=1" + vbCrLf)
        attendi_risposta ("+SBITS=1" + vbCrLf)
        'setta la parità
        invia_comando ("AT+PARITY=E" + vbCrLf)
        attendi_risposta ("+PARITY=E" + vbCrLf)
    End If
  
  End If
End Sub


Private Sub apri_seriale()
  If uart.PortOpen = False Then
    uart.CommPort = portacom
    uart.Settings = "19200,n,8,1"
    uart.DTREnable = False
    'If micro = "STM32F303VE" Then

    uart.DTREnable = False

    'Else
    '  uart.RTSEnable = False
    'End If
    uart.EOFEnable = False
    uart.Handshaking = comNone
    uart.InBufferSize = 1024
    uart.InputMode = comInputModeText
    uart.InputLen = 0
    uart.RThreshold = 100
    uart.SThreshold = 100
    uart.PortOpen = True
    
    
        'Modifiche effettuate per poter gestire l'interfaccia BT seriale
    If chkBT.Value = 1 Then
        
        'Apre la porta seriale
        invia_comando ("AT+BYPASS=0" + vbCrLf)
        attendi_risposta ("+BYPASS=0" + vbCrLf)
        
        'Alza il pin di reset per essere sicuro che non resetti la scheda. Si arrangia il convertitore a decidere se il pin è effettivamente
        'a livello logico 1 o 0 in base al fatto che la porta sia una RS232 o una TTL
        invia_comando ("AT+RESET=1" + vbCrLf)
        attendi_risposta ("+RESET=1" + vbCrLf)
        
        
        If chkTTL.Value = 1 Then
            invia_comando ("AT+STYPE=T" + vbCrLf)
            attendi_risposta ("+STYPE=T" + vbCrLf)
        Else
            invia_comando ("AT+STYPE=R" + vbCrLf)
            attendi_risposta ("+STYPE=R" + vbCrLf)
        End If
        
        'setta il baud rate
        invia_comando ("AT+BRATE=19200" + vbCrLf)
        attendi_risposta ("+BRATE=19200" + vbCrLf)
        'setta i data bits
        invia_comando ("AT+DBITS=8" + vbCrLf)
        attendi_risposta ("+DBITS=8" + vbCrLf)
        'setta gli stop bits
        invia_comando ("AT+SBITS=1" + vbCrLf)
        attendi_risposta ("+SBITS=1" + vbCrLf)
        'setta la parità
        invia_comando ("AT+PARITY=N" + vbCrLf)
        attendi_risposta ("+PARITY=N" + vbCrLf)
        
        'Apre la porta seriale
        invia_comando ("AT+BYPASS=1" + vbCrLf)
        attendi_risposta ("+BYPASS=1" + vbCrLf)
        
        
    End If
  End If
End Sub


Private Sub chiudi_seriale()
  If uart.PortOpen = True Then
    uart.PortOpen = False
  End If
End Sub



Private Sub Timer2_Timer()
  'incremento contatore
  centesimi = centesimi + 10
  If centesimi > 20000 Then centesimi = 10000
End Sub

Private Sub win_Click()

  colonna_selezionata = win.ColSel
  riga_selezionata = win.RowSel

'text_debug1.Visible = True
'debug4 = debug4 + 1
'text_debug1.Text = Str(debug4) + "," + Str(colonna_selezionata) + "," + Str(riga_selezionata)

  If uart.PortOpen = True And fase_comandi > 3 Then
    If colonna_selezionata = 3 Then
      If (indice_comando_esterno < 40) Then indice_comando_esterno = indice_comando_esterno + 1
      If riga_selezionata > 9 Then
        comando_esterno(indice_comando_esterno) = "#-" + Trim(Str(riga_selezionata))
      Else
        comando_esterno(indice_comando_esterno) = "#-0" + Trim(Str(riga_selezionata))
      End If
      bak_colonna_selezionata = colonna_selezionata
      bak_riga_selezionata = riga_selezionata
    ElseIf colonna_selezionata = 4 Then
      If (indice_comando_esterno < 40) Then indice_comando_esterno = indice_comando_esterno + 1
      If riga_selezionata > 9 Then
        comando_esterno(indice_comando_esterno) = "#+" + Trim(Str(riga_selezionata))
      Else
        comando_esterno(indice_comando_esterno) = "#+0" + Trim(Str(riga_selezionata))
      End If
      bak_colonna_selezionata = colonna_selezionata
      bak_riga_selezionata = riga_selezionata
    Else
'debug1 = debug1 + 1
'debug2 = colonna_selezionata
    End If
  End If
  
  If focus = 0 Then
    tasto_uart.SetFocus
  ElseIf focus = 1 Then
    comando_reset(0).SetFocus
  ElseIf focus = 2 Then
    comando_reset(1).SetFocus
  Else
    focus = 0
    tasto_uart.SetFocus
  End If

End Sub

Private Sub win_DblClick()
  colonna_selezionata = win.ColSel
  riga_selezionata = win.RowSel
  
  If uart.PortOpen = True And fase_comandi > 3 Then
    If colonna_selezionata = 3 Then
      If (indice_comando_esterno < 40) Then indice_comando_esterno = indice_comando_esterno + 1
      If riga_selezionata > 9 Then
        comando_esterno(indice_comando_esterno) = "#/" + Trim(Str(riga_selezionata))
      Else
        comando_esterno(indice_comando_esterno) = "#/0" + Trim(Str(riga_selezionata))
      End If
      bak_colonna_selezionata = colonna_selezionata
      bak_riga_selezionata = riga_selezionata
    ElseIf colonna_selezionata = 4 Then
      If (indice_comando_esterno < 40) Then indice_comando_esterno = indice_comando_esterno + 1
      If riga_selezionata > 9 Then
        comando_esterno(indice_comando_esterno) = "#*" + Trim(Str(riga_selezionata))
      Else
        comando_esterno(indice_comando_esterno) = "#*0" + Trim(Str(riga_selezionata))
      End If
      bak_colonna_selezionata = colonna_selezionata
      bak_riga_selezionata = riga_selezionata
    End If
  End If

  If focus = 0 Then
    tasto_uart.SetFocus
  ElseIf focus = 1 Then
    comando_reset(0).SetFocus
  ElseIf focus = 2 Then
    comando_reset(1).SetFocus
  Else
    focus = 0
    tasto_uart.SetFocus
  End If

End Sub


Private Sub calcola_chk_diter()
  Dim ii As Integer
  
  chk_diter = 0
  
  ii = 1
  While ii <= Len(comando_diter)
    chk_diter = chk_diter + Asc(Mid(comando_diter, ii, 1))
    ii = ii + 1
  Wend
End Sub


Private Sub programma_micro()

Dim ttt As Integer

ttt = 0
  
  errore_download = 0


  tasto_update.Caption = "Download..."

  'inizio preparando il file
  
  'come prima cosa estraggo dal file hex il solo codice
  estrai_codice_hex
  If errore_download Then GoTo err_update

  
  'inizio aprendo la porta
  apri_seriale_per_isp
  
ttt = 1
  
  resetta_micro
  
  metti_il_micro_in_isp_mode
  If errore_download Then GoTo err_update
  
ttt = 2
  sincronizza_micro
  If errore_download Then GoTo err_update
ttt = 3
  sproteggi_read_flash
  If errore_download Then GoTo err_update
ttt = 4
  
  sincronizza_micro
  If errore_download Then GoTo err_update
ttt = 5

  sproteggi_write_flash
  If errore_download Then GoTo err_update
ttt = 6
  
  sincronizza_micro
  If errore_download Then GoTo err_update
ttt = 7
  
  cancella_flash
  If errore_download Then GoTo err_update
ttt = 8
  
  scarica_programma
  If errore_download Then GoTo err_update
  
  proteggi_read_flash
  If errore_download Then GoTo err_update
  
  resetta_micro
  
  chiudi_seriale


  If stato_seriale = True Then
    apri_seriale
  End If




  tasto_update.Caption = "UPDATE OK!"
  
  text_report.Text = text_report.Text + "UPDATE OK!"
  text_report.SelStart = InStr(text_report.Text, "UPDATE OK!") - 1
  text_report.SelLength = Len("UPDATE OK!")
  text_report.SelBold = True
  text_report.SelColor = vbGreen
  
  timer_programmazione = 995
  GoTo fine_tasto_update

err_update:
  timer_programmazione = 995
  
  On Error GoTo fine_tasto_update
  If InStr(text_report.Text, "ERROR") > 0 Then
    text_report.SelStart = InStr(text_report.Text, "ERROR") - 1
    text_report.SelLength = Len("ERROR")
    text_report.SelBold = True
    text_report.SelColor = vbRed
  End If
  text_report.Text = text_report.Text + "ttt=" + ttt
  
  GoTo fine_tasto_update
    
fine_tasto_update:
  


End Sub









Private Sub estrai_codice_hex()

Dim conta_riga As Long
Dim indirizzo_hex As Long
Dim dim_riga As Integer
Dim adr_riga As Long
Dim tipo_riga As Integer
Dim chk, chk_riga As Integer
Dim extra_address As Long
Dim riga_in, riga_out As String

  text_report.Text = text_report.Text + "Extract code from file...." + vbLf
  
  'apro il file
  Open nome_file_hex_con_path For Input As #1
  Open path_file_hex + "file_solo_codice.txt" For Output As #2
  conta_riga = 0
  indirizzo_hex = 0
  extra_address = 0
  riga_out = ""
  
loop_hex:
  If EOF(1) Then
    GoTo end_loop_hex
  End If
  
  Input #1, riga_in
  
'If conta_riga = 2049 Then
'  tasto_update.Caption = conta_riga
'End If
  
  If Len(riga_in) < 8 Then
    GoTo errore_riga_hex
  End If
  If Mid(riga_in, 1, 1) <> ":" Then
    GoTo errore_riga_hex
  End If
  dim_riga = dec(Mid(riga_in, 2, 2))
  adr_riga = Val(dec(Mid(riga_in, 4, 4))) + extra_address

'If adr_riga = 11456 Then
'  nop
'End If

  tipo_riga = Val(Mid(riga_in, 8, 2))
  conta_riga = conta_riga + 1
  




'struttura di un file hex, spiegazione dei byte per ogni riga:
'1: carattere di start, deve essere ":"
'2: indica il numero di byte della riga
'3,4: indirizzo
'5: record type
' 0-dati
' 1-eof
' 2-extended segment address
' 3-start segment address
' 4-extended linear address
' 5-start linear address
'6....: dati
'ultimo: checksum, complemento a 2 della somma di tutti i byte dopo il :

  If tipo_riga = 1 Then
    GoTo tipo_riga_1
  ElseIf tipo_riga = 2 Then
    GoTo tipo_riga_2
  ElseIf tipo_riga = 3 Then
    GoTo tipo_riga_3
  ElseIf tipo_riga = 4 Then
    GoTo tipo_riga_4
  ElseIf tipo_riga = 5 Then
    GoTo tipo_riga_5
  End If
  
tipo_riga_0:  'riga che contiene dati da trasmettere
  'verifico la lunghezza della riga
  If dim_riga > 16 Then
    GoTo errore_riga_hex
  End If
  If Len(riga_in) <> 11 + dim_riga * 2 Then
    GoTo errore_riga_hex
  End If
  'lunghezza ok, verifico chk
  
  'calcolo il chk della riga
  chk = 0
  For ii = 2 To Len(riga_in) - 3 Step 2
    chk = chk + dec(Mid(riga_in, ii, 2))
    If chk > 255 Then chk = chk - 256
  Next
  
  'lo complemento
  If chk > 0 Then
    chk = 256 - chk
  End If
    
  'estraggo il chk della riga
  chk_riga = dec(Right(riga_in, 2))
  'lo confronto
  If chk <> chk_riga Then
    GoTo errore_riga_hex
  End If
  
  'la riga è giusta, estraggo i dati
  riga_out = riga_out + Mid(riga_in, 10, dim_riga * 2)
  indirizzo_hex = indirizzo_hex + dim_riga
  
  
'  While dim_riga < 16
'    riga = riga + "FF"
'    dim_riga = dim_riga + 1
'    indirizzo_hex = indirizzo_hex + 1
'  Wend
  
  If Len(riga_out) >= 32 Then
    Print #2, Left(riga_out, 32)
    riga_out = Right(riga_out, Len(riga_out) - 32)
  End If
  GoTo loop_hex
  
tipo_riga_1:  'riga di fine file
  
  'completo con FF fino ad arrivare a multiplo di 256
  While indirizzo_hex Mod (4096) > 0
    riga_out = riga_out + "FF"
    If Len(riga_out) >= 32 Then
      Print #2, Left(riga_out, 32)
      riga_out = Right(riga_out, Len(riga_out) - 32)
    End If
    indirizzo_hex = indirizzo_hex + 1
  Wend
  
  
  If Len(riga_out) > 0 Then GoTo errore_riga_hex
  dimensione_file = indirizzo_hex
  text_report.Text = text_report.Text + "file is " + Trim(Str(dimensione_file)) + " byte long." + vbLf

  Print #2, ""

  If indirizzo_hex > 512000 Then
    errore_download = True
    text_report.Text = text_report.Text + "ERROR: file is too long for device!" + vbLf
  End If
  
  GoTo end_loop_hex

tipo_riga_2:
tipo_riga_3:
  GoTo loop_hex
tipo_riga_4:  'riga che contiene indirizzo esteso del file hex dul byte di dato
  'per micro ST
  GoTo loop_hex
  
  extra_address = dec(Mid(riga_in, 12, 2)) * 65536
  adr_riga = adr_riga + extra_address
  If indirizzo_hex > adr_riga Then
    GoTo errore_riga_hex
  ElseIf indirizzo_hex < adr_riga Then
    While indirizzo_hex < adr_riga
      riga_out = riga_out + "FF"
      indirizzo_hex = indirizzo_hex + 1
      If Len(riga_out) >= 32 Then
        Print #2, Left(riga_out, 32)
        riga_out = Right(riga_out, Len(riga_out) - 32)
      End If
    Wend
  End If
    
  indirizzo_hex = adr_riga
tipo_riga_5:  'riga iniziale
  GoTo loop_hex
  
errore_riga_hex:
    
  text_report.Text = text_report.Text + "File is not good!" + vbLf
  errore_download = True
  tasto_update.Caption = "ERROR"

end_loop_hex:

  Close #1
  Close #2
  

End Sub



Private Sub prepara_file_download()


End Sub





Private Sub resetta_micro()

   
  text_report.Text = text_report.Text + "Reset MicroPropcessor..." + vbLf
    

tasto_update.Caption = "ISP MODE"

    If chkBT.Value = 1 Then
        invia_comando ("AT+STM32_RST_SEQUENCE=1" + vbCrLf)
        attendi_risposta ("+STM32_RST_SEQUENCE=1" + vbCrLf)
        
    Else
    
  'metto il micro in run
  uart.RTSEnable = True
  centesimi = 0
  While centesimi < 50
    DoEvents
  Wend
  
  'metto il micro in reset
  uart.DTREnable = True
  centesimi = 0
  While centesimi < 50
    DoEvents
  Wend
  
  'tolgo il reset
  uart.DTREnable = False
  centesimi = 0
  While centesimi < 50
    DoEvents
  Wend
End If
End Sub



Private Sub metti_il_micro_in_isp_mode()
  
  tasto_update.Caption = "ISP MODE"
  
  If chkBT.Value = 1 Then
    invia_comando ("AT+BYPASS=0" + vbCrLf)
    attendi_risposta ("+BYPASS=0" + vbCrLf)
    invia_comando ("AT+STM32_BOOT_SEQUENCE=1" + vbCrLf)
    attendi_risposta ("+STM32_BOOT_SEQUENCE=1" + vbCrLf)
    invia_comando ("AT+BYPASS=1" + vbCrLf)
    attendi_risposta ("+BYPASS=1" + vbCrLf)
    'attesa
    centesimi = 0
    While centesimi < 50
        DoEvents
    Wend
  
  Else
  'metto il micro in ISP
  uart.RTSEnable = False
  centesimi = 0
  While centesimi < 50
    DoEvents
  Wend
  
  'metto il micro in reset
  uart.DTREnable = True
  centesimi = 0
  While centesimi < 50
    DoEvents
  Wend
  
  'tolgo il reset
  uart.DTREnable = False
  centesimi = 0
  While centesimi < 50
    DoEvents
  Wend
End If


End Sub

'devo inviare 0x7f e mi aspetto 0x79
Private Sub sincronizza_micro()
  Dim rx As String
  Dim prove As Integer
  
  On Error GoTo err_sincro
  
  prove = 0

  text_report.Text = text_report.Text + "Connect to MicroProcessor..." + vbLf

  'invio il 0x3f di sincronizzazione
  tasto_update.Caption = "SYNCRO"

  'svuoto buffer di ricezione per sicurezza
  rx = uart.Input

invia_punto_interrogativo:
  invia_comando (Chr(127))
  'attendo ACK
  attendi_risposta (Chr(121))
  
  If errore_comando = True Then
    If prove < 10 Then
      prove = prove + 1
      GoTo invia_punto_interrogativo
    Else
      GoTo err_sincro
    End If
  End If
  
  
  
  'invio get id command
  invia_comando (Chr(2) + Chr(253))
  attendi_risposta (Chr(121) + Chr(1) + Chr(4) + Chr(70) + Chr(121))
  
  If errore_comando = True Then
    GoTo err_sincro
  Else
   errore_comando = False
  End If
  
'  If Len(rx_cmd) <> 5 Then
'    errore_comando = True
'  ElseIf Asc(Mid(rx_cmd, 1, 1)) <> 121 Then
'    errore_comando = True
'  ElseIf Asc(Mid(rx_cmd, 2, 1)) <> 1 Then
'    errore_comando = True
'  ElseIf Asc(Mid(rx_cmd, 3, 1)) <> 4 Then
'    errore_comando = True
'  ElseIf Asc(Mid(rx_cmd, 4, 1)) <> 70 Then
'    errore_comando = True
'  ElseIf Asc(Mid(rx_cmd, 5, 1)) <> 121 Then
'    errore_comando = True
'  Else
'    errore_comando = False
'  End If
  
  GoTo fine_sincro
  
err_sincro:
  text_report.Text = text_report.Text + "ERROR: fail to connect!" + vbLf
  errore_download = True
  
fine_sincro:
  End Sub
  
Private Sub sproteggi_write_flash()

  On Error GoTo err_sproteggi_wr_flash

  text_report.Text = text_report.Text + "Prepare flash..." + vbLf

  tasto_update.Caption = "PREPARE"
  'sproteggi
  invia_comando (Chr(115) + Chr(140))
  attendi_risposta (Chr(121))
  If errore_comando = True Then
    If Len(rx_cmd) <> 1 Then
      GoTo err_sproteggi_wr_flash
    ElseIf Asc(rx_cmd) <> 31 Then
      GoTo err_sproteggi_wr_flash
    End If
  Else
    rx_cmd = ""
    attendi_risposta (Chr(121))
    If errore_comando = True Then GoTo err_sproteggi_wr_flash
  End If

  centesimi = 0
  While centesimi < 50
    timer_programmazione = 1
    DoEvents
  Wend

  GoTo fine_sproteggi_wr_flash
  
err_sproteggi_wr_flash:
  errore_download = True
  tasto_update.Caption = "ERR.PREPARE"
  text_report.Text = text_report.Text + "ERROR: can't sprotect device!"

fine_sproteggi_wr_flash:

End Sub

Private Sub sproteggi_read_flash()

  invia_comando (Chr(146) + Chr(109))
  attendi_risposta (Chr(121))
  If errore_comando = True Then
    If Len(rx_cmd) <> 1 Then
      GoTo err_sproteggi_rd_flash
    ElseIf Asc(rx_cmd) <> 31 Then
      GoTo err_sproteggi_rd_flash
    End If
  Else
    rx_cmd = ""
    attendi_risposta (Chr(121))
    If errore_comando = True Then GoTo err_sproteggi_rd_flash
  End If

  centesimi = 0
  While centesimi < 50
    timer_programmazione = 1
    DoEvents
  Wend
  
  GoTo fine_sproteggi_rd_flash
  
err_sproteggi_rd_flash:
  errore_download = True
  tasto_update.Caption = "ERR.PREPARE"
  text_report.Text = text_report.Text + "ERROR: can't sprotect device!"

fine_sproteggi_rd_flash:

End Sub
 

Private Sub proteggi_read_flash()

  invia_comando (Chr(130) + Chr(125))
  attendi_risposta (Chr(121))
  If errore_comando = True Then
'    If Len(rx_cmd) <> 1 Then
      GoTo err_proteggi_rd_flash
'    ElseIf Asc(rx_cmd) <> 31 Then
'      GoTo err_proteggi_rd_flash
'    End If
'  Else
'    rx_cmd = ""
'    attendi_risposta (Chr(121))
'    If errore_comando = True Then GoTo err_proteggi_rd_flash
  End If

  
  
'  centesimi = 0
'  While centesimi < 300
'    timer_programmazione = 1
'    DoEvents
'  Wend
  
  GoTo fine_proteggi_rd_flash
  
err_proteggi_rd_flash:
  errore_download = True
  tasto_update.Caption = "ERR.PROTECT"
  text_report.Text = text_report.Text + "ERROR: can't protect device!"

fine_proteggi_rd_flash:

End Sub
 





Private Sub cancella_flash()

  On Error GoTo err_delete_flash

  errore_comando = False
  
  text_report.Text = text_report.Text + "Delete old version..." + vbLf

  tasto_update.Caption = "ERASE"
  'erase
  invia_comando (Chr(68) + Chr(187))
  attendi_risposta (Chr(121))
  If errore_comando = True Then GoTo err_delete_flash

  invia_comando (Chr(255) + Chr(255) + Chr(0))


  attendi_risposta (Chr(121))
  If errore_comando = True Then GoTo err_delete_flash


  GoTo fine_delete_flash

err_delete_flash:
  errore_download = True
  tasto_update.Caption = "ERR.DELE"
  text_report.Text = text_report.Text + "ERROR: can't erase device!"
fine_delete_flash:

End Sub


Private Sub check_blank()
'non esiste un comando di blanck check quindi devo verificare tutta la flash
'  On Error GoTo err_check_blank
'
'  text_report.Text = text_report.Text + "Blank verify..."
'
'  tasto_update.Caption = "CHECK BLANK"
'  cmd = "I 1 " + Trim(Str(settori_device))
'  invia_comando (cmd + vbCr + vbLf)
'  attendi_risposta (cmd + vbCr + vbLf + "0" + vbCr + vbLf)
'  If errore_comando = True Then GoTo err_check_blank
'
'  text_report.Text = text_report.Text + "OK" + vbLf
'  GoTo fine_check_blank
'
'err_check_blank:
'  text_report.Text = text_report.Text + "ERROR: device is not blank!"
'  errore_download = True
'  tasto_update.Caption = "ERR.BLANK"

fine_check_blank:


End Sub



Private Sub leggi_id()

'  text_report.Text = text_report.Text + "ID Verify..."
'
'  'leggo il part ID
'  tasto_update.Caption = "READ ID"
'  invia_comando ("J" + vbCr + vbLf)
'  attendi_risposta ("J" + vbCr + vbLf + "0" + vbCr + vbLf + id_device + vbCr + vbLf)
'
'  If errore_comando = True Then
'    errore_download = True
'    text_report.Text = text_report.Text + "ERROR" + vbLf
'  Else
'    text_report.Text = text_report.Text + "OK" + vbLf
'  End If

End Sub





Private Sub unlock_micro()

'  'unlock
'  text_report.Text = text_report.Text + "Unlock MicroProcessor" + vbLf
'  tasto_update.Caption = "UNLOCK"
'  invia_comando ("U 23130" + vbCr + vbLf)
'  attendi_risposta ("U 23130" + vbCr + vbLf + "0" + vbCr + vbLf)
'
'  If errore_comando = True Then errore_download = True

End Sub
  


Private Sub leggi_flash()
  
'  On Error GoTo err_read_flash
'
'
'  Dim conta_righe_rx As Integer
'  Dim byte_da_leggere As Integer
'
'  text_report.Text = text_report.Text + "Read Flash..." + vbLf
'
'  Open path_file_hex + "file_read_flash.uuc" For Output As #1
'
'  'leggo la flash e la salvo in un file
'  cmd = "R 256 4"
'  invia_comando (cmd + vbCr + vbLf)
'  attendi_risposta (cmd + vbCr + vbLf + "0" + vbCr + vbLf)
'  If errore_comando = True Then GoTo err_read_flash
'
'  'pulisco la riga
'  rx_cmd = Right(rx_cmd, Len(rx_cmd) - Len(cmd + vbCr + vbLf + "0" + vbCr + vbLf))
'
'  'adesso mi arrivano i dati
'  conta_righe_rx = 0
'  byte_da_leggere = Int(256 / 3)
'  If Int(256 / 3) < 256 / 3 Then byte_da_leggere = byte_da_leggere + 1
'  byte_da_leggere = byte_da_leggere * 4
'  byte_da_leggere = byte_da_leggere + Int(byte_da_leggere / 60)
'  If Int(256 / 3) < 256 / 3 Then byte_da_leggere = byte_da_leggere + 1
'
'
'
'loop_ricevi_righe:
'  centesimi = 0
'  While centesimi < 100
'    DoEvents
'    aaa = uart.Input
'    If Len(aaa) > 0 Then
'      rx_cmd = rx_cmd + aaa
'      If Len(rx_cmd) >= byte_da_leggere Then centesimi = 1000
'    End If
'  Wend
'
'  Print #1, rx_cmd
'  byte_da_leggere = byte_da_leggere - Len(rx_cmd)
'
'  If byte_da_leggere > 0 Then
'    invia_comando ("OK" + vbCr + vbLf)
'    GoTo loop_ricevi_righe
'  End If
'
'  GoTo fine_read_flash
'
'err_read_flash:
'  errore_download = True
'
'fine_read_flash:
'
'  Close #1
'
End Sub


Private Sub scarica_programma()
  Dim riga, riga_tx As String
  Dim tentativi As Integer
  Dim chk As Byte
  Dim adr(5) As Byte
  Dim perc As Long

'  Dim flash, flash2 As Long
  

  On Error GoTo err_write_flash

  'è tutto pronto, devo solo inviare le righe

'  'cambio frequenza per andare + veloce
'  cmd = "B 38400 1"
'  invia_comando (cmd + vbCr + vbLf)
'  attendi_risposta (cmd + vbCr + vbLf + "0" + vbCr + vbLf)
'  If errore_comando = True Then GoTo err_write_flash
'
'  uart.Settings = "38400,n,8,1"
'
'
'
inizio_download_file:

  Open path_file_hex + "file_solo_codice.txt" For Input As #1

  text_report.Text = text_report.Text + "Download new version..." + vbLf


'per scrivere devo trasferire le righe in ram partendo dall'indirizzo giusto
'scrivo 256 byte alla volta

'le righe sono tutte pronte per la trasmissione
'ogni riga contiene 16 byte
'16 righe contengono 256 byte
'invio 256 byte alla volta


  If EOF(1) Then
    tasto_update.Caption = "ERR.FILE"
    GoTo err_download
  End If

  On Error GoTo err_download

  'flash = 134217728
  adr(0) = 0
  adr(1) = 0
  adr(2) = 0
  adr(3) = 8
  
  chk = 0

'  'metto echo off
'  cmd = "A 0"
'  invia_comando (cmd + vbCr + vbLf)
'  attendi_risposta (cmd + vbCr + vbLf + "0" + vbCr + vbLf)
'
loop_write_flash:


  If timer_programmazione > 500 Then GoTo err_download

  tentativi = 0

  riga_tx = ""

  'carico 16 righe
loop_carica_righe:
  Line Input #1, riga
  If EOF(1) Then
    GoTo ultime_righe
  Else
    riga_tx = riga_tx + riga
    If Len(riga_tx) < 256 * 2 Then
      GoTo loop_carica_righe
    End If
  End If

  'converto in binario
  riga = ""
  
  
  Dim nn As Integer
  
  chk = 255
  For nn = 0 To (Len(riga_tx) / 2) - 1
    car = Chr(dec(Mid(riga_tx, 1 + nn * 2, 2)))
    riga = riga + car
    chk = chk Xor Asc(car)
  Next


trasmetti_riga:
  
'If flash > 134250752 Then
'  flash2 = flash
'End If
  
  invia_comando (Chr(49) + Chr(206))
  attendi_risposta (Chr(121))
  If errore_comando = True Then GoTo err_download
    
  'trasmetto indirizzo + chk
'  flash2 = flash
'  adr(0) = flash2 Mod 256
'  flash2 = flash2 / 256
'  adr(1) = flash2 Mod 256
'  flash2 = flash2 / 256
'  adr(2) = flash2 Mod 256
'  flash2 = flash2 / 256
'  adr(3) = flash2 Mod 256
  adr(4) = adr(0) Xor adr(1) Xor adr(2) Xor adr(3)
  invia_comando (Chr(adr(3)) + Chr(adr(2)) + Chr(adr(1)) + Chr(adr(0)) + Chr(adr(4)))
  attendi_risposta (Chr(121))
  If errore_comando = True Then GoTo err_download
  
  'invio la riga
 
  
  invia_comando (Chr(255) + riga + Chr(chk))
  attendi_risposta (Chr(121))
  If errore_comando = True Then GoTo err_download
  
  'rileggo la riga per sicurezza
  
  invia_comando (Chr(17) + Chr(238))
  attendi_risposta (Chr(121))
  If errore_comando = True Then GoTo err_download
    
  'trasmetto indirizzo + chk
'  flash2 = flash
'  adr(0) = flash2 Mod 256
'  flash2 = flash2 / 256
'  adr(1) = flash2 Mod 256
'  flash2 = flash2 / 256
'  adr(2) = flash2 Mod 256
'  flash2 = flash2 / 256
'  adr(3) = flash2 Mod 256
  adr(4) = adr(0) Xor adr(1) Xor adr(2) Xor adr(3)
  invia_comando (Chr(adr(3)) + Chr(adr(2)) + Chr(adr(1)) + Chr(adr(0)) + Chr(adr(4)))
  attendi_risposta (Chr(121))
  If errore_comando = True Then GoTo err_download
  
  'trsmetto numero di byte da leggere (-1)
  invia_comando (Chr(255) + Chr(0))
  
  'devo leggere la riga
  attendi_risposta (Chr(121) + riga)
  
  If errore_comando = True Then
    If Len(rx_cmd) < 257 Then
      GoTo err_download
    Else
      If Mid(rx_cmd, 2, 256) <> riga Then
        GoTo err_download
      End If
    End If
  End If
  
  
  
  
  'flash = flash + 256
  If adr(1) < 255 Then
    adr(1) = adr(1) + 1
  Else
    adr(1) = 0
    If adr(2) < 255 Then
      adr(2) = adr(2) + 1
    Else
      adr(2) = 0
      adr(3) = adr(3) + 1
    End If
  End If
    
  
Dim percentuale As Long
  
  perc = adr(2)
  perc = perc * 256
  perc = perc + adr(1)
  perc = perc * 256
  perc = perc + adr(0)
  perc = perc * 100
  perc = perc / dimensione_file
  perc = Int(perc)
  tasto_update.Caption = Str(Trim(perc)) + "%"
'  tasto_update.Caption = Str(Trim(perc)) + "/" + Str(Trim(dimensione_file))
    
  
'  percentuale = flash - 134217728
'  percentuale = percentuale * 100
'  percentuale = percentuale / dimensione_file
'  percentuale = Int(percentuale)
'  tasto_update.Caption = Str(Trim(percentuale)) + "% " + Str(flash)

  GoTo loop_write_flash


ultime_righe:


  Close #1
  GoTo fine_write_flash

err_download:
  text_report.Text = text_report.Text + "ERROR: can't write device!"
  errore_download = True
  tasto_update.Caption = "ERR.WRITE"
  Close #1

err_write_flash:

  text_report.Text = text_report.Text + "ERROR: can't write device!"
  errore_download = True
  tasto_update.Caption = "ERR.WRITE"


fine_write_flash:

  'ripristino la velocità della seriale
'  uart.Settings = "19200,n,8,1"

End Sub

























'converte una stringa esadecimale in un numer decimale
Private Function dec(numero_hex As String) As Double
  
  Dim numero_dec As Double
  Dim len_numero As Integer
  
  numero_dec = 0
  For len_numero = 1 To Len(numero_hex)
    numero_dec = numero_dec * 16
  
    If Mid(numero_hex, len_numero, 1) = "0" Then
      numero_dec = numero_dec + 0
    ElseIf Mid(numero_hex, len_numero, 1) = "1" Then
      numero_dec = numero_dec + 1
    ElseIf Mid(numero_hex, len_numero, 1) = "2" Then
      numero_dec = numero_dec + 2
    ElseIf Mid(numero_hex, len_numero, 1) = "3" Then
      numero_dec = numero_dec + 3
    ElseIf Mid(numero_hex, len_numero, 1) = "4" Then
      numero_dec = numero_dec + 4
    ElseIf Mid(numero_hex, len_numero, 1) = "5" Then
      numero_dec = numero_dec + 5
    ElseIf Mid(numero_hex, len_numero, 1) = "6" Then
      numero_dec = numero_dec + 6
    ElseIf Mid(numero_hex, len_numero, 1) = "7" Then
      numero_dec = numero_dec + 7
    ElseIf Mid(numero_hex, len_numero, 1) = "8" Then
      numero_dec = numero_dec + 8
    ElseIf Mid(numero_hex, len_numero, 1) = "9" Then
      numero_dec = numero_dec + 9
    ElseIf Mid(numero_hex, len_numero, 1) = "A" Then
      numero_dec = numero_dec + 10
    ElseIf Mid(numero_hex, len_numero, 1) = "B" Then
      numero_dec = numero_dec + 11
    ElseIf Mid(numero_hex, len_numero, 1) = "C" Then
      numero_dec = numero_dec + 12
    ElseIf Mid(numero_hex, len_numero, 1) = "D" Then
      numero_dec = numero_dec + 13
    ElseIf Mid(numero_hex, len_numero, 1) = "E" Then
      numero_dec = numero_dec + 14
    ElseIf Mid(numero_hex, len_numero, 1) = "F" Then
      numero_dec = numero_dec + 15
    End If
  Next
  
  dec = numero_dec
  
End Function

'converte un numero decimale in una stringa
Private Function hex(ByVal numero_dec As Double) As String
  
  Dim numero_hex As String
  Dim cifra_hex As Integer
  
  numero_hex = ""
  While numero_dec
    cifra_hex = (numero_dec - Int(numero_dec / 16) * 16)
    If cifra_hex = 0 Then
      numero_hex = "0" + numero_hex
    ElseIf cifra_hex = 1 Then
      numero_hex = "1" + numero_hex
    ElseIf cifra_hex = 2 Then
      numero_hex = "2" + numero_hex
    ElseIf cifra_hex = 3 Then
      numero_hex = "3" + numero_hex
    ElseIf cifra_hex = 4 Then
      numero_hex = "4" + numero_hex
    ElseIf cifra_hex = 5 Then
      numero_hex = "5" + numero_hex
    ElseIf cifra_hex = 6 Then
      numero_hex = "6" + numero_hex
    ElseIf cifra_hex = 7 Then
      numero_hex = "7" + numero_hex
    ElseIf cifra_hex = 8 Then
      numero_hex = "8" + numero_hex
    ElseIf cifra_hex = 9 Then
      numero_hex = "9" + numero_hex
    ElseIf cifra_hex = 10 Then
      numero_hex = "A" + numero_hex
    ElseIf cifra_hex = 11 Then
      numero_hex = "B" + numero_hex
    ElseIf cifra_hex = 12 Then
      numero_hex = "C" + numero_hex
    ElseIf cifra_hex = 13 Then
      numero_hex = "D" + numero_hex
    ElseIf cifra_hex = 14 Then
      numero_hex = "E" + numero_hex
    ElseIf cifra_hex = 15 Then
      numero_hex = "F" + numero_hex
    End If
      
    numero_dec = Int(numero_dec / 16)
  Wend
  
  If numero_hex = "" Then numero_hex = "0"
  
  hex = numero_hex
  
End Function
'converte un numero decimale in una stringa
Private Function hex_ff(ByVal numero_dec As Double) As String
  
  Dim numero_hex As String
  Dim cifra_hex As Integer
  
  numero_hex = ""
  While numero_dec
    cifra_hex = (numero_dec - Int(numero_dec / 16) * 16)
    If cifra_hex = 0 Then
      numero_hex = "0" + numero_hex
    ElseIf cifra_hex = 1 Then
      numero_hex = "1" + numero_hex
    ElseIf cifra_hex = 2 Then
      numero_hex = "2" + numero_hex
    ElseIf cifra_hex = 3 Then
      numero_hex = "3" + numero_hex
    ElseIf cifra_hex = 4 Then
      numero_hex = "4" + numero_hex
    ElseIf cifra_hex = 5 Then
      numero_hex = "5" + numero_hex
    ElseIf cifra_hex = 6 Then
      numero_hex = "6" + numero_hex
    ElseIf cifra_hex = 7 Then
      numero_hex = "7" + numero_hex
    ElseIf cifra_hex = 8 Then
      numero_hex = "8" + numero_hex
    ElseIf cifra_hex = 9 Then
      numero_hex = "9" + numero_hex
    ElseIf cifra_hex = 10 Then
      numero_hex = "A" + numero_hex
    ElseIf cifra_hex = 11 Then
      numero_hex = "B" + numero_hex
    ElseIf cifra_hex = 12 Then
      numero_hex = "C" + numero_hex
    ElseIf cifra_hex = 13 Then
      numero_hex = "D" + numero_hex
    ElseIf cifra_hex = 14 Then
      numero_hex = "E" + numero_hex
    ElseIf cifra_hex = 15 Then
      numero_hex = "F" + numero_hex
    End If
      
    numero_dec = Int(numero_dec / 16)
  Wend
  
  If numero_hex = "" Then numero_hex = "0"
  If Len(numero_hex) < 2 Then numero_hex = "0" + numero_hex
  
  hex_ff = numero_hex
  
End Function





Private Sub invia_comando(comando As String)
  timer_text_report = 150

  rx_cmd = ""
  tx_cmd = comando
  uart.Output = tx_cmd
  End Sub


Private Sub attendi_risposta(stringa As String)

  Dim xx As Integer

  On Error GoTo errore_attesa_risposta
  
  'devo epurare le stringhe dei carattere CR e LF
'  xx = 1
'  While xx <= Len(stringa)
'    If Asc(Mid(stringa, xx, 1)) = 13 Then
'      stringa = Left(stringa, xx - 1) + Right(stringa, Len(stringa) - xx)
'    ElseIf Asc(Mid(stringa, xx, 1)) = 10 Then
'      stringa = Left(stringa, xx - 1) + Right(stringa, Len(stringa) - xx)
'    Else
'      xx = xx + 1
'    End If
'  Wend

  centesimi = 0
  While centesimi < 300
    timer_programmazione = 1
    aaa = uart.Input
    
    If Len(aaa) > 0 Then
'      'devo epurare le stringhe dei carattere CR e LF
'      xx = 1
'      While xx <= Len(aaa)
'        If Asc(Mid(aaa, xx, 1)) = 13 Then
'          aaa = Left(aaa, xx - 1) + Right(aaa, Len(aaa) - xx)
'        ElseIf Asc(Mid(aaa, xx, 1)) = 10 Then
'          aaa = Left(aaa, xx - 1) + Right(aaa, Len(aaa) - xx)
'        Else
'          xx = xx + 1
'        End If
'      Wend
      rx_cmd = rx_cmd + aaa
      If Len(rx_cmd) >= Len(stringa) Then centesimi = 1000
    End If
    DoEvents
  Wend

 If Len(rx_cmd) = 0 Then
    errore_comando = True
 ElseIf Mid(rx_cmd, 1, Len(stringa)) = Mid(stringa, 1, Len(stringa)) Then
    errore_comando = False
 Else
  'trovo il carattere di differenza,
  If Len(rx_cmd) > Len(stringa) Then
    errore_comando = True
  Else
    For xx = 1 To Len(rx_cmd)
      If Mid(rx_cmd, xx, 1) <> Mid(stringa, xx, 1) Then
        errore_comando = True
        xx = 1000
      End If
    Next
  End If
 End If
 
 GoTo fine_attesa_risposta

errore_attesa_risposta:
  errore_comando = True

fine_attesa_risposta:

End Sub



Private Sub nop()

End Sub

