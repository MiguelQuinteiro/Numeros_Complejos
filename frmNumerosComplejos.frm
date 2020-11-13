VERSION 5.00
Begin VB.Form frmNumerosComplejos 
   BackColor       =   &H8000000D&
   Caption         =   "Números Complejos"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   51
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtBuscaZ1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   50
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtBuscaZ2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   49
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtTrigonometricaZ2 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   48
      Top             =   3480
      Width           =   4215
   End
   Begin VB.TextBox txtPolarZ2 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   47
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txtBinomicaZ2 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   46
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtTrigonometricaZ1 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   42
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox txtPolarZ1 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   41
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtBinomicaZ1 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   40
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtCuadranteZ2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   39
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtCuadranteZ1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   38
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtAnguloRadianZ1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   37
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtAnguloRadianZ2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   36
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtAnguloZ2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   33
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtAnguloZ1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   32
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Calcular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9480
      TabIndex        =   31
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtModuloZ1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   28
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtModuloZ2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   27
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtOpuestoD 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   23
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtOpuestoC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   22
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtOpuestoB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   21
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtOpuestoA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   20
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtConjugadoD 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtConjugadoC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtConjugadoB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtConjugadoA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtDivision 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox txtMultiplicacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox txtResta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox txtSuma 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox txtD 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Text            =   "0.5"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Text            =   "-0.86602540378443864676372317075294"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Text            =   "0.5"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "0.86602540378443864676372317075294"
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "Trigonométrica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   45
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "Polar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   44
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "Binómica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   43
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "Ángulo Z2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   35
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "Ángulo Z1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   34
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Módulo Z1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   30
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Módulo Z2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   29
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Operaciones de Suma, Resta, Multiplicación y División :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   4200
      Width           =   9135
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Opuesto Z2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   25
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Opuesto Z1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   24
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Conjugado Z2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   19
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Conjugado Z1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   18
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "División"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   13
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Multiplicación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Resta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Suma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Número Complejo Z2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Número Complejo Z1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmNumerosComplejos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Declaración de tipo de datos
Private Type Complejo
  X As Double     ' Componente real
  Y As Double     ' Componente imaginaria
  CX As Double    ' Componente real conjugado
  CY As Double    ' Componente imaginaria conjugado
  OX As Double    ' Componente real opuesto
  OY As Double    ' Componente imaginaria opuesto
  R As Double     ' Módulo (norma, magnitud)
  An As Double    ' Ángulo (argumento)
  Cu As String    ' Cuadrante
End Type

' Declaración de Variables
Dim miPi As Double
Dim z1 As Complejo
Dim z2 As Complejo
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double

Dim Mod1 As String
Dim Mod2 As String
Dim Ang1 As String
Dim Ang2 As String

Dim miSuma As Complejo
Dim miResta As Complejo
Dim miMultiplicacion As Complejo
Dim miDivision As Complejo

Private Sub cmdBuscar_Click()
' Para Z1
  If txtBuscaZ1.Text <> "" Then
    txtA.Text = CambiaComa(Cos(Val(txtBuscaZ1.Text) * miPi / 180))
  End If
  If txtBuscaZ1.Text <> "" Then
    txtB.Text = CambiaComa(Sin(Val(txtBuscaZ1.Text) * miPi / 180))
  End If
  ' Para Z2
  If txtBuscaZ2.Text <> "" Then
    txtC.Text = CambiaComa(Cos(Val(txtBuscaZ2.Text) * miPi / 180))
  End If
  If txtBuscaZ1.Text <> "" Then
    txtD.Text = CambiaComa(Sin(Val(txtBuscaZ2.Text) * miPi / 180))
  End If
End Sub

' Al Cargar el formulario
Private Sub Form_Load()
  miPi = 3.14159265358979

End Sub

' Botón Calcular
Private Sub cmdCalcular_Click()
' Captura de datos desde el formulario
'Val (Replace(Datos.Text, ",", "."))
  a = Val(Replace(txtA.Text, ",", "."))
  b = Val(Replace(txtB.Text, ",", "."))
  c = Val(Replace(txtC.Text, ",", "."))
  d = Val(Replace(txtD.Text, ",", "."))

  ' Cálculos
  ' Para Z1
  z1.X = a
  z1.Y = b
  z1.R = cModulo(z1.X, z1.Y)
  z1.An = cAngulo(z1.X, z1.Y)
  z1.Cu = cCuadrante(z1.X, z1.Y)
  z1.CX = z1.X
  z1.CY = -z1.Y
  z1.OX = -z1.X
  z1.OY = -z1.Y
  ' Para Z2
  z2.X = c
  z2.Y = d
  z2.R = cModulo(z2.X, z2.Y)
  z2.An = cAngulo(z2.X, z2.Y)
  z2.Cu = cCuadrante(z2.X, z2.Y)
  z2.CX = z2.X
  z2.CY = -z2.Y
  z2.OX = -z2.X
  z2.OY = -z2.Y

  ' Operaciones con numeros complejos
  ' SUMA
  miSuma = SumaComplejo(z1, z2)
  miResta = RestaComplejo(z1, z2)
  miMultiplicacion = MultiplicaComplejo(z1, z2)
  miDivision = DivideComplejo(z1, z2)

  ' Inicializa
  Mod1 = Format(z1.R, "##,##0.0000")
  Ang1 = Format(z1.An, "##,##0.0000 º")
  Mod2 = Format(z2.R, "##,##0.0000")
  Ang2 = Format(z2.An, "##,##0.0000 º")

  ' Muestra los resultados
  ' Para Z1
  txtModuloZ1.Text = Mod1
  txtAnguloZ1.Text = Ang1
  txtAnguloRadianZ1.Text = Format(z1.An * miPi / 180, "##,##0.0000 Rad")
  txtCuadranteZ1.Text = z1.Cu
  txtConjugadoA.Text = z1.CX
  txtConjugadoB.Text = z1.CY
  txtOpuestoA.Text = z1.OX
  txtOpuestoB.Text = z1.OY
  ' Para Z2
  txtModuloZ2.Text = Mod2
  txtAnguloZ2.Text = Ang2
  txtAnguloRadianZ2.Text = Format(z2.An * miPi / 180, "##,##0.0000 Rad")
  txtCuadranteZ2.Text = z2.Cu
  txtConjugadoC.Text = z2.CX
  txtConjugadoD.Text = z2.CY
  txtOpuestoC.Text = z2.OX
  txtOpuestoD.Text = z2.OY

  ' Represantaciones de números complejos
  ' Binómica
  If z1.Y >= 0 Then
    txtBinomicaZ1.Text = " = " & Format(z1.X, "##,##0.000") & " + " & Format(z1.Y, "##,##0.000") & " i "
  Else
    txtBinomicaZ1.Text = " = " & Format(z1.X, "##,##0.000") & " - " & Format(Abs(z1.Y), "##,##0.000") & " i "
  End If
  If z2.Y >= 0 Then
    txtBinomicaZ2.Text = " = " & Format(z2.X, "##,##0.000") & " + " & Format(z2.Y, "##,##0.000") & " i "
  Else
    txtBinomicaZ2.Text = " = " & Format(z2.X, "##,##0.000") & " - " & Format(Abs(z2.Y), "##,##0.000") & " i "
  End If
  ' Polar
  txtPolarZ1.Text = " R= " & Mod1 & " -- @= " & Ang1
  txtPolarZ2.Text = " R= " & Mod2 & " -- @= " & Ang2
  ' Trigonométrica
  txtTrigonometricaZ1.Text = " " & Mod1 & " ( Cos" & Ang1 & " + i Sen" & Ang1 & " )"
  txtTrigonometricaZ2.Text = " " & Mod2 & " ( Cos" & Ang2 & " + i Sen" & Ang2 & " )"

  ' La Suma
  If miSuma.Y >= 0 Then
    txtSuma.Text = " = " & Format(miSuma.X, "##,##0.000") & " + " & Format(miSuma.Y, "##,##0.000") & " i "
  Else
    txtSuma.Text = " = " & Format(miSuma.X, "##,##0.000") & " - " & Format(Abs(miSuma.Y), "##,##0.000") & " i "
  End If
  ' La Resta
  If miResta.Y >= 0 Then
    txtResta.Text = " = " & Format(miResta.X, "##,##0.000") & " + " & Format(miResta.Y, "##,##0.000") & " i "
  Else
    txtResta.Text = " = " & Format(miResta.X, "##,##0.000") & " - " & Format(Abs(miResta.Y), "##,##0.000") & " i "
  End If
  ' La Multiplicacion
  If miMultiplicacion.Y >= 0 Then
    txtMultiplicacion.Text = " = " & Format(miMultiplicacion.X, "##,##0.000") & " + " & Format(miMultiplicacion.Y, "##,##0.000") & " i "
  Else
    txtMultiplicacion.Text = " = " & Format(miMultiplicacion.X, "##,##0.000") & " - " & Format(Abs(miMultiplicacion.Y), "##,##0.000") & " i "
  End If

  ' La Division
  If miDivision.Y >= 0 Then
    txtDivision.Text = " = " & Format(miDivision.X, "##,##0.000") & " + " & Format(miDivision.Y, "##,##0.000") & " i "
  Else
    txtDivision.Text = " = " & Format(miDivision.X, "##,##0.000") & " - " & Abs(Format(miDivision.Y, "##,##0.000")) & " i "
  End If


End Sub


'********************************************************************************
' FUNCIONES PARA TRABAJAR CON NUMEROS COMPLEJOS
'********************************************************************************

' Calcula el módulo del número complejo
Public Function cModulo(ByVal a As Double, ByVal b As Double) As Double
' Validación para datos igual a 0
  If a = 0 Then
    a = 0.000001
  End If
  If b = 0 Then
    b = 0.00000000001
  End If
  ' Cálculo del módulo
  cModulo = Sqr((a ^ 2) + (b ^ 2))
End Function

' Calcula el ángulo del número complejo
Public Function cAngulo(ByVal a As Double, ByVal b As Double) As Double
' Validación para datos igual a 0
  If a = 0 Then
    a = 0.000001
  End If
  If b = 0 Then
    b = 0.00000000001
  End If
  ' Cálculo del ángulo
  ' Primer Cuadrante      (entre 0º y 90º)
  If a > 0 And b > 0 Then
    cAngulo = (Atn(Abs(b) / Abs(a)) * 180 / miPi)
  End If
  ' Segundo Cuadrante     (entre 90º y 180º)
  If a < 0 And b > 0 Then
    cAngulo = 180 - (Atn(Abs(b) / Abs(a)) * 180 / miPi)
  End If
  ' Tercer Cuadrante      (entre 180º y 270º)
  If a < 0 And b < 0 Then
    cAngulo = 180 + (Atn(Abs(b) / Abs(a)) * 180 / miPi)
  End If
  ' Cuarto Cuadrante      (entre 270º y 360º)
  If a > 0 And b < 0 Then
    cAngulo = 360 - (Atn(Abs(b) / Abs(a)) * 180 / miPi)
  End If
End Function

' Calcula el cuadrante del número complejo
Public Function cCuadrante(ByVal a As Double, ByVal b As Double) As String
' Cálculo del ángulo
' Primer Cuadrante      (entre 0º y 90º)
  If a > 0 And b > 0 Then
    cCuadrante = " I.- Cuadrante"
  End If
  ' Segundo Cuadrante     (entre 90º y 180º)
  If a < 0 And b > 0 Then
    cCuadrante = " II.- Cuadrante"
  End If
  ' Tercer Cuadrante      (entre 180º y 270º)
  If a < 0 And b < 0 Then
    cCuadrante = " III.- Cuadrante"
  End If
  ' Cuarto Cuadrante      (entre 270º y 360º)
  If a > 0 And b < 0 Then
    cCuadrante = " IV.- Cuadrante"
  End If
End Function

' Suma de Complejos
Private Function SumaComplejo(ByRef z1 As Complejo, ByRef z2 As Complejo) As Complejo
  SumaComplejo.X = z1.X + z2.X
  SumaComplejo.Y = z1.Y + z2.Y

  SumaComplejo.R = cModulo(z1.X, z1.Y)
  SumaComplejo.An = cAngulo(z1.X, z1.Y)
  SumaComplejo.Cu = cCuadrante(z1.X, z1.Y)
  SumaComplejo.CX = z1.X
  SumaComplejo.CY = -z1.Y
  SumaComplejo.OX = -z1.X
  SumaComplejo.OY = -z1.Y
End Function

' Resta de Complejos
Private Function RestaComplejo(ByRef z1 As Complejo, ByRef z2 As Complejo) As Complejo
  RestaComplejo.X = z1.X - z2.X
  RestaComplejo.Y = z1.Y - z2.Y

  RestaComplejo.R = cModulo(z1.X, z1.Y)
  RestaComplejo.An = cAngulo(z1.X, z1.Y)
  RestaComplejo.Cu = cCuadrante(z1.X, z1.Y)
  RestaComplejo.CX = z1.X
  RestaComplejo.CY = -z1.Y
  RestaComplejo.OX = -z1.X
  RestaComplejo.OY = -z1.Y
End Function

' Multiplicacion de Complejos
Private Function MultiplicaComplejo(ByRef z1 As Complejo, ByRef z2 As Complejo) As Complejo
  MultiplicaComplejo.X = (z1.X * z2.X) + (z1.Y * z2.Y * (-1))
  MultiplicaComplejo.Y = (z1.X * z2.Y) + (z1.Y * z2.X)

  MultiplicaComplejo.R = cModulo(z1.X, z1.Y)
  MultiplicaComplejo.An = cAngulo(z1.X, z1.Y)
  MultiplicaComplejo.Cu = cCuadrante(z1.X, z1.Y)
  MultiplicaComplejo.CX = z1.X
  MultiplicaComplejo.CY = -z1.Y
  MultiplicaComplejo.OX = -z1.X
  MultiplicaComplejo.OY = -z1.Y
End Function

' Division de Complejos
Private Function DivideComplejo(ByRef z1 As Complejo, ByRef z2 As Complejo) As Complejo
  DivideComplejo.X = (z1.X * z2.CX) + (z1.Y * z2.CY * (-1))
  DivideComplejo.Y = (z1.X * z2.CY) + (z1.Y * z2.CX)

  DivideComplejo.X = DivideComplejo.X / ((z2.X ^ 2) - (z2.Y ^ 2) * (-1))
  DivideComplejo.Y = DivideComplejo.Y / ((z2.X ^ 2) - (z2.Y ^ 2) * (-1))

  DivideComplejo.R = cModulo(z1.X, z1.Y)
  DivideComplejo.An = cAngulo(z1.X, z1.Y)
  DivideComplejo.Cu = cCuadrante(z1.X, z1.Y)
  DivideComplejo.CX = z1.X
  DivideComplejo.CY = -z1.Y
  DivideComplejo.OX = -z1.X
  DivideComplejo.OY = -z1.Y
End Function

' Ajusta el punto por la coma
Private Function CambiaComa(ByRef n As Double) As String
  Dim i As Integer
  CambiaComa = ""
  For i = 1 To Len(n)
    If Mid(Str(n), i, 1) = "," Then
      CambiaComa = CambiaComa + "."
    Else
      CambiaComa = CambiaComa + Mid(Str(n), i, 1)
    End If
  Next i
End Function

