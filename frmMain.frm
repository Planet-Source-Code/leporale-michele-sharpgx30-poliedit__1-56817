VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Editor melodie polifoniche per Sharp GX-30"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.ImageList ImagePrimoMenu 
      Left            =   4680
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B84
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstVociSuoneria 
      Height          =   450
      ItemData        =   "frmMain.frx":10F8
      Left            =   5040
      List            =   "frmMain.frx":10FA
      TabIndex        =   47
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSuggerimento 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1575
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   46
      Text            =   "frmMain.frx":10FC
      Top             =   2400
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   4560
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      Caption         =   "Velocità melodia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   3600
      TabIndex        =   40
      Top             =   4080
      Width           =   2055
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   41
         Text            =   "frmMain.frx":1185
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Selezionare la scala!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1935
      Left            =   8280
      TabIndex        =   33
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton TipoScala 
         Caption         =   "Scala - -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   38
         Top             =   1500
         Width           =   2600
      End
      Begin VB.OptionButton TipoScala 
         Caption         =   "Scala -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   37
         Top             =   1200
         Width           =   2600
      End
      Begin VB.OptionButton TipoScala 
         Caption         =   "Scala + +"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   300
         Width           =   2600
      End
      Begin VB.OptionButton TipoScala 
         Caption         =   "Scala +"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   2600
      End
      Begin VB.OptionButton TipoScala 
         Caption         =   "Scala Base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   900
         Value           =   -1  'True
         Width           =   2600
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Aggiunte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1400
      Left            =   3800
      TabIndex        =   28
      Top             =   2600
      Width           =   1400
      Begin VB.CommandButton aggiunta 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   3
         Left            =   700
         TabIndex        =   32
         ToolTipText     =   "Nota in terzina"
         Top             =   850
         Width           =   500
      End
      Begin VB.CommandButton aggiunta 
         Caption         =   "Lega"
         Height          =   500
         Index           =   2
         Left            =   200
         TabIndex        =   31
         ToolTipText     =   "Lega la nota alla successiva"
         Top             =   850
         Width           =   500
      End
      Begin VB.CommandButton aggiunta 
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   700
         TabIndex        =   30
         ToolTipText     =   "Nota con il punto"
         Top             =   350
         Width           =   500
      End
      Begin VB.CommandButton aggiunta 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   200
         TabIndex        =   29
         ToolTipText     =   "Nota con il Diesis (aumenta di un semitono)"
         Top             =   350
         Width           =   500
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Composizione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   4100
      Left            =   5760
      TabIndex        =   27
      Top             =   4000
      Width           =   5895
      Begin VB.Frame Frame8 
         Caption         =   "Contatore note:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   800
         Left            =   3840
         TabIndex        =   44
         ToolTipText     =   "Conta sia le note che le pause!"
         Top             =   200
         Width           =   1935
         Begin VB.Label lblContaNote 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   495
            Left            =   240
            TabIndex        =   45
            ToolTipText     =   "Conta sia le note che le pause!"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtNumeroVoci 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   43
         Text            =   "1"
         Top             =   630
         Width           =   1215
      End
      Begin VB.TextBox txtComposizione 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2880
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label lblCancellaUltima 
         Alignment       =   2  'Center
         Caption         =   "Cancella l'ultima nota!"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   615
         Left            =   2360
         TabIndex        =   48
         Top             =   360
         Width           =   1400
      End
      Begin VB.Label Label1 
         Caption         =   "Selezione voce (max 8-16-32):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         TabIndex        =   42
         Top             =   400
         Width           =   2415
      End
   End
   Begin VB.CommandButton nota 
      Caption         =   "SOL+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   7000
      TabIndex        =   25
      Top             =   500
      Width           =   630
   End
   Begin VB.CommandButton nota 
      Caption         =   "FA+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   6400
      TabIndex        =   24
      Top             =   650
      Width           =   500
   End
   Begin VB.CommandButton nota 
      Caption         =   "MI+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   5800
      TabIndex        =   23
      Top             =   800
      Width           =   500
   End
   Begin VB.CommandButton nota 
      Caption         =   "RE+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   5200
      TabIndex        =   22
      Top             =   950
      Width           =   520
   End
   Begin VB.CommandButton nota 
      Caption         =   "DO+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   4590
      TabIndex        =   21
      Top             =   1100
      Width           =   540
   End
   Begin VB.Frame Frame2 
      Caption         =   "Velocità note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1300
      Left            =   200
      TabIndex        =   15
      Top             =   2600
      Width           =   3255
      Begin VB.CommandButton velocita 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   9
         Left            =   200
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Semibreve"
         Top             =   500
         Width           =   500
      End
      Begin VB.CommandButton velocita 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   8
         Left            =   800
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Minima"
         Top             =   500
         Width           =   500
      End
      Begin VB.CommandButton velocita 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   7
         Left            =   1400
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Semiminima"
         Top             =   500
         Width           =   500
      End
      Begin VB.CommandButton velocita 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   6
         Left            =   2000
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Croma"
         Top             =   500
         Width           =   500
      End
      Begin VB.CommandButton velocita 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   5
         Left            =   2600
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Semicroma"
         Top             =   500
         Width           =   520
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pausa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1300
      Left            =   200
      TabIndex        =   9
      Top             =   3900
      Width           =   3255
      Begin VB.CommandButton pausa 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   4
         Left            =   2600
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Semicroma"
         Top             =   500
         Width           =   500
      End
      Begin VB.CommandButton pausa 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   3
         Left            =   2000
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Croma"
         Top             =   500
         Width           =   500
      End
      Begin VB.CommandButton pausa 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   1400
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Semiminima"
         Top             =   500
         Width           =   500
      End
      Begin VB.CommandButton pausa 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   800
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Minima"
         Top             =   500
         Width           =   500
      End
      Begin VB.CommandButton pausa 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   200
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Semibreve"
         Top             =   500
         Width           =   500
      End
   End
   Begin VB.CommandButton nota 
      Caption         =   "SI'"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   4000
      TabIndex        =   8
      ToolTipText     =   "B"
      Top             =   1250
      Width           =   500
   End
   Begin VB.CommandButton nota 
      Caption         =   "LA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   3400
      TabIndex        =   7
      ToolTipText     =   "A"
      Top             =   1400
      Width           =   500
   End
   Begin VB.CommandButton nota 
      Caption         =   "SOL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   2780
      TabIndex        =   6
      ToolTipText     =   "G"
      Top             =   1550
      Width           =   540
   End
   Begin VB.CommandButton nota 
      Caption         =   "FA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   2200
      TabIndex        =   5
      ToolTipText     =   "F"
      Top             =   1700
      Width           =   500
   End
   Begin VB.CommandButton nota 
      Caption         =   "MI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1600
      TabIndex        =   4
      ToolTipText     =   "E"
      Top             =   1850
      Width           =   500
   End
   Begin VB.CommandButton nota 
      Caption         =   "RE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1000
      TabIndex        =   3
      ToolTipText     =   "D"
      Top             =   2000
      Width           =   500
   End
   Begin VB.CommandButton nota 
      Caption         =   "DO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   400
      TabIndex        =   2
      ToolTipText     =   "C"
      Top             =   2150
      Width           =   500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Titolo Brano"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   200
      TabIndex        =   0
      Top             =   6900
      Width           =   5175
      Begin VB.TextBox txtTitoloBrano 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   120
         MaxLength       =   24
         TabIndex        =   1
         Top             =   600
         Width           =   4935
      End
   End
   Begin VB.Label lblScala 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scala Base"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   720
      TabIndex        =   26
      Top             =   120
      Width           =   2895
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   4540
      X2              =   4540
      Y1              =   240
      Y2              =   2640
   End
   Begin VB.Image Image1 
      Height          =   1515
      Left            =   200
      Picture         =   "frmMain.frx":11E1
      Top             =   5400
      Width           =   4245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      Index           =   5
      X1              =   200
      X2              =   1050
      Y1              =   2300
      Y2              =   2300
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   200
      X2              =   8000
      Y1              =   2000
      Y2              =   2000
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   200
      X2              =   8000
      Y1              =   1700
      Y2              =   1700
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   200
      X2              =   8000
      Y1              =   1400
      Y2              =   1400
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   200
      X2              =   8000
      Y1              =   1100
      Y2              =   1100
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   200
      X2              =   8000
      Y1              =   800
      Y2              =   800
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNuovo 
         Caption         =   "Nuova Melodia"
      End
      Begin VB.Menu mnuSalva 
         Caption         =   "Salva Melodia"
      End
      Begin VB.Menu mnuCarica 
         Caption         =   "Carica Melodia"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dichiarazione per l'uso di icone nel menu:
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Private Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Fine dichiarazione
Dim PrimaNota As Boolean
Dim ContaNote As Integer
Dim LegaturaPossibile As Boolean
Dim UltimaVoce As Integer

Private Sub Form_Load()
 'Per visualizzare le icone nel menu:
 Dim i%
 Dim hMenu, hSubMenu, menuID, x
 hMenu = GetMenu(hWnd)
 hSubMenu = GetSubMenu(hMenu, 0)
 'Primo Menu:
 For i = 1 To 3
  menuID = GetMenuItemID(hSubMenu, i - 1)
  x = SetMenuItemBitmaps(hMenu, menuID, &H4, ImagePrimoMenu.ListImages(i).Picture, ImagePrimoMenu.ListImages(i).Picture)
 Next
 PrimaNota = True
 Frame5.Visible = False
 Frame2.Visible = False
 velocita(7).BackColor = &HFF00&
 UltimaVoce = txtNumeroVoci
End Sub

Private Sub mnuCarica_Click()
 On Error GoTo errore
 Dim i As Integer
 Dim strTxtNomeBrano As String
 Dim strTxtComposizione(31) As String
 With cmnDialog
  .DefaultExt = "SGX"
  .DialogTitle = "Scegli la suoneria da caricare!"
  .FileName = Empty
  .Filter = "Sharp GX30(*.SGX)"
  .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNNoChangeDir + cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNExtensionDifferent
  .InitDir = App.Path & "\Suonerie"
  .ShowOpen
  'Apertura File Esistente
  Open .FileName For Input As 1
   Input #1, strTxtNomeBrano
   For i = 0 To 31
    Input #1, strTxtComposizione(i)
    lstVociSuoneria.AddItem strTxtComposizione(i)
   Next i
  Close 1
  txtTitoloBrano = strTxtNomeBrano
  txtComposizione = lstVociSuoneria.List(txtNumeroVoci - 1)
 End With
 i = Len(txtComposizione) - Len(Replace(UCase(txtComposizione), UCase(" "), ""))
 lblContaNote = i + 1
errore:
 If Err.Number = 1 Then Exit Sub
End Sub

Private Sub mnuNuovo_Click()
 PrimaNota = True
 lstVociSuoneria.Clear
 txtNumeroVoci = 1
 txtComposizione = ""
 txtTitoloBrano = ""
 TipoScala(0).Value = True
 lblContaNote = "0"
 velocita(7).BackColor = &HFF00&
End Sub

Private Sub mnuSalva_Click()
  On Error GoTo errore
 Dim i As Integer
 lstVociSuoneria.List(txtNumeroVoci - 1) = txtComposizione
 With cmnDialog
  .DefaultExt = "SGX"
  .DialogTitle = "Salva la suoneria per il tuo Sharp GX30!!!"
  .FileName = txtTitoloBrano
  .Filter = "Sharp GX30(*.SGX)"
  .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNNoChangeDir + cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNExtensionDifferent
  .InitDir = App.Path & "\Suonerie"
  .ShowSave
  'Creazione Nuovo File
  Open .FileName For Output As 1
   Write #1, txtTitoloBrano
   For i = 0 To 31
    Write #1, lstVociSuoneria.List(i)
   Next i
  Close 1
 End With
errore:
 If Err.Number = 1 Then Exit Sub
End Sub

Private Sub nota_click(Index As Integer)
 Dim i As Integer
 UltimaVoce = txtNumeroVoci
 Frame5.Visible = True
 Frame2.Visible = True
 lblCancellaUltima.Enabled = False
 If PrimaNota = False Then
  txtComposizione = txtComposizione & " "
 End If
 PrimaNota = False
 lblContaNote = lblContaNote + 1
 lblCancellaUltima.Enabled = True
 LegaturaPossibile = True
 For i = 0 To 3
  aggiunta(i).Visible = True
 Next i
 For i = 5 To 9
  velocita(i).Enabled = True
  velocita(i).BackColor = &HC0C0C0
 Next i
 velocita(7).BackColor = &HFF00&
 If (Index = 2) Or (Index = 6) Then
  aggiunta(0).Visible = False
  Else:  aggiunta(0).Visible = True
 End If
 If TipoScala(0) = True Then
  Select Case Index
   Case 0
    txtComposizione = txtComposizione & "1"
   Case 1
    txtComposizione = txtComposizione & "2"
   Case 2
    txtComposizione = txtComposizione & "3"
   Case 3
    txtComposizione = txtComposizione & "4"
   Case 4
    txtComposizione = txtComposizione & "5"
   Case 5
    txtComposizione = txtComposizione & "6"
   Case 6
    txtComposizione = txtComposizione & "7"
   Case 7
    txtComposizione = txtComposizione & "11"
   Case 8
    txtComposizione = txtComposizione & "22"
   Case 9
    txtComposizione = txtComposizione & "33"
   Case 10
    txtComposizione = txtComposizione & "44"
   Case 11
    txtComposizione = txtComposizione & "55"
  End Select
 End If
 If TipoScala(1) = True Then
  Select Case Index
   Case 0
    txtComposizione = txtComposizione & "11"
   Case 1
    txtComposizione = txtComposizione & "22"
   Case 2
    txtComposizione = txtComposizione & "33"
   Case 3
    txtComposizione = txtComposizione & "44"
   Case 4
    txtComposizione = txtComposizione & "55"
   Case 5
    txtComposizione = txtComposizione & "66"
   Case 6
    txtComposizione = txtComposizione & "77"
  End Select
 End If
 If TipoScala(2) = True Then
  Select Case Index
   Case 0
    txtComposizione = txtComposizione & "111"
   Case 1
    txtComposizione = txtComposizione & "222"
   Case 2
    txtComposizione = txtComposizione & "333"
   Case 3
    txtComposizione = txtComposizione & "444"
  End Select
 End If
 If TipoScala(3) = True Then
  Select Case Index
   Case 0
    txtComposizione = txtComposizione & "1111"
   Case 1
    txtComposizione = txtComposizione & "2222"
   Case 2
    txtComposizione = txtComposizione & "3333"
   Case 3
    txtComposizione = txtComposizione & "4444"
   Case 4
    txtComposizione = txtComposizione & "555"
   Case 5
    txtComposizione = txtComposizione & "6666"
   Case 6
    txtComposizione = txtComposizione & "7777"
  End Select
 End If
 If TipoScala(4) = True Then
  Select Case Index
   Case 5
    txtComposizione = txtComposizione & "666"
   Case 6
    txtComposizione = txtComposizione & "777"
  End Select
 End If
End Sub

Private Sub nota_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 If LegaturaPossibile = True Then
  If Button = 2 Then
   txtComposizione = txtComposizione & "8"
   aggiunta(2).Visible = False
  End If
  LegaturaPossibile = False
 End If
End Sub

Private Sub TipoScala_Click(Index As Integer)
 lblScala = TipoScala(Index).Caption
 Dim i As Integer
 If Index = 0 Then
  For i = 0 To 11
   nota(i).Visible = True
  Next i
 End If
 If Index <> 0 Then
  For i = 7 To 11
   nota(i).Visible = False
  Next i
 End If
 If Index = 1 Then
  For i = 0 To 6
   nota(i).Visible = True
  Next i
 End If
 If Index = 2 Then
  For i = 4 To 6
   nota(i).Visible = False
  Next i
  For i = 0 To 3
   nota(i).Visible = True
  Next i
 End If
 If Index = 3 Then
  For i = 0 To 6
   nota(i).Visible = True
  Next i
 End If
 If Index = 4 Then
  For i = 0 To 4
   nota(i).Visible = False
  Next i
  For i = 5 To 6
   nota(i).Visible = True
  Next i
 End If
End Sub

Private Sub txtNumeroVoci_Change()
 On Error GoTo errore
 Dim temp As Integer
 temp = Val(txtNumeroVoci)
 If IsNumeric(temp) Then
  If temp < 1 Or temp > 32 Then
   txtNumeroVoci = 1
  End If
 End If
 lstVociSuoneria.List(UltimaVoce - 1) = txtComposizione
 txtComposizione = lstVociSuoneria.List(txtNumeroVoci - 1)
 UltimaVoce = txtNumeroVoci
 temp = Len(txtComposizione) - Len(Replace(UCase(txtComposizione), UCase(" "), ""))
 If txtComposizione <> "" Then
  lblContaNote = temp + 1
  Else: lblContaNote = 0
  PrimaNota = True
 End If
errore:
 If Err.Number > 0 Then txtNumeroVoci = 1
End Sub

Private Sub velocita_Click(Index As Integer)
 Dim i As Integer
 For i = 5 To 9
  velocita(i).BackColor = &HC0C0C0
  velocita(i).Enabled = False
 Next i
 velocita(Index).BackColor = &HFF00&
 Select Case Index
  Case 5
   txtComposizione = txtComposizione & "**"
  Case 6
   txtComposizione = txtComposizione & "*"
  Case 7
   txtComposizione = txtComposizione & ""
  Case 8
   txtComposizione = txtComposizione & "#"
  Case 9
   txtComposizione = txtComposizione & "##"
 End Select
End Sub

Private Sub aggiunta_Click(Index As Integer)
 Select Case Index
  Case 0
   txtComposizione = txtComposizione & "^"
   aggiunta(0).Visible = False
  Case 1
   txtComposizione = txtComposizione & "9"
   aggiunta(3).Visible = False
   aggiunta(1).Visible = False
  Case 2
   txtComposizione = txtComposizione & "8"
   aggiunta(2).Visible = False
  Case 3
   txtComposizione = txtComposizione & "99"
   aggiunta(1).Visible = False
   aggiunta(3).Visible = False
  End Select
End Sub

Private Sub pausa_click(Index As Integer)
 Dim i As Integer
 UltimaVoce = txtNumeroVoci
 Frame5.Visible = False
 Frame2.Visible = False
 lblCancellaUltima.Enabled = False
 If PrimaNota = False Then
  txtComposizione = txtComposizione & " "
 End If
 PrimaNota = False
 lblContaNote = lblContaNote + 1
 lblCancellaUltima.Enabled = True
 For i = 0 To 4
  pausa(i).BackColor = &HC0C0C0
 Next i
 pausa(Index).BackColor = &HC0&
 Select Case Index
  Case 0
   txtComposizione = txtComposizione & "0##"
  Case 1
   txtComposizione = txtComposizione & "0#"
  Case 2
   txtComposizione = txtComposizione & "0"
  Case 3
   txtComposizione = txtComposizione & "0*"
  Case 4
   txtComposizione = txtComposizione & "0**"
 End Select
End Sub

Private Sub lblCancellaUltima_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 lblCancellaUltima.FontBold = True
End Sub

Private Sub lblCancellaUltima_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim i As Integer
 Dim n As Integer
 Dim z As String
 i = Len(txtComposizione)
 For n = i To 0 Step -1
  If n = 0 Then
   lblContaNote = lblContaNote - 1
   lblCancellaUltima.Enabled = False
   PrimaNota = True
   Exit Sub
  End If
  z = txtComposizione
  If (Right(z, 1) <> " ") Then
   Mid(z, n, 1) = vbNullChar
   txtComposizione = z
   Else: Mid(z, n, 1) = vbNullChar
         txtComposizione = z
         lblContaNote = lblContaNote - 1
         Exit Sub
  End If
 Next n
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 lblCancellaUltima.FontBold = False
End Sub

