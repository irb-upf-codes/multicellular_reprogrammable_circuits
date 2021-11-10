VERSION 5.00
Begin VB.Form frmKVDiagram 
   AutoRedraw      =   -1  'True
   Caption         =   "CDesign - Multicellular Ciruits Designer"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   ScaleHeight     =   628
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1054
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   3255
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   46
      Top             =   5880
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   840
      TabIndex        =   44
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFC0C0&
      Height          =   4875
      Left            =   3240
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   813
      TabIndex        =   35
      Top             =   360
      Width           =   12255
      Begin VB.HScrollBar HS 
         Height          =   255
         LargeChange     =   10
         Left            =   0
         Max             =   100
         TabIndex        =   43
         Top             =   4560
         Width           =   12255
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   240
         ScaleHeight     =   305
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   321
         TabIndex        =   36
         Top             =   0
         Width           =   4815
         Begin VB.CommandButton Reprog 
            Caption         =   "Command1"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   45
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label TNOT 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NOT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   990
            TabIndex        =   42
            Top             =   2790
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label T4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id(d)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   41
            Top             =   1740
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label T3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id(c)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   40
            Top             =   1740
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label T2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id(b)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   39
            Top             =   1020
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label T1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id(a)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   38
            Top             =   1020
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Shape NOT 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   0
            Left            =   840
            Shape           =   3  'Circle
            Top             =   2460
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Shape C4 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            Height          =   1815
            Index           =   0
            Left            =   1320
            Shape           =   3  'Circle
            Top             =   1020
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Shape C3 
            BackColor       =   &H0080FF80&
            BackStyle       =   1  'Opaque
            Height          =   1815
            Index           =   0
            Left            =   480
            Shape           =   3  'Circle
            Top             =   1020
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Shape C2 
            BackColor       =   &H008080FF&
            BackStyle       =   1  'Opaque
            Height          =   1815
            Index           =   0
            Left            =   1320
            Shape           =   3  'Circle
            Top             =   300
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Shape C1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            Height          =   1815
            Index           =   0
            Left            =   480
            Shape           =   3  'Circle
            Top             =   300
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OUTPUT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   660
            TabIndex        =   37
            Top             =   3750
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Shape output 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000005&
            Height          =   615
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   3540
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Shape Conector 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000005&
            Height          =   255
            Index           =   0
            Left            =   930
            Top             =   3300
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Shape Camara 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000005&
            Height          =   2775
            Index           =   0
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   540
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Shape fondo 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   4455
            Left            =   0
            Top             =   45
            Visible         =   0   'False
            Width           =   2295
         End
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Simplification"
      Height          =   255
      Left            =   480
      TabIndex        =   34
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   -4680
      TabIndex        =   33
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Generate"
      Height          =   375
      Left            =   840
      TabIndex        =   32
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2280
      TabIndex        =   31
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Salida 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Entrada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   44
      Left            =   -15000
      TabIndex        =   18
      Text            =   "0"
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   43
      Left            =   -15000
      TabIndex        =   17
      Text            =   "0"
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   42
      Left            =   -15000
      TabIndex        =   16
      Text            =   "0"
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   41
      Left            =   -15000
      TabIndex        =   15
      Text            =   "0"
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   34
      Left            =   -15000
      TabIndex        =   14
      Text            =   "0"
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   33
      Left            =   -15000
      TabIndex        =   13
      Text            =   "0"
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   32
      Left            =   -15000
      TabIndex        =   12
      Text            =   "0"
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   31
      Left            =   -15000
      TabIndex        =   11
      Text            =   "0"
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   24
      Left            =   -15000
      TabIndex        =   10
      Text            =   "0"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   23
      Left            =   -15000
      TabIndex        =   9
      Text            =   "0"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   22
      Left            =   -15000
      TabIndex        =   8
      Text            =   "0"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   21
      Left            =   -15000
      TabIndex        =   7
      Text            =   "0"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   14
      Left            =   -15000
      TabIndex        =   6
      Text            =   "0"
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   13
      Left            =   -15000
      TabIndex        =   5
      Text            =   "0"
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   12
      Left            =   -15000
      TabIndex        =   4
      Text            =   "0"
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   11
      Left            =   -15000
      TabIndex        =   3
      Text            =   "0"
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSimplify 
      Caption         =   "Simplify"
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   9360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Chamber Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   47
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Num. Inputs:"
      Height          =   195
      Left            =   480
      TabIndex        =   30
      Top             =   480
      Width           =   900
   End
   Begin VB.Shape box2R2C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   1455
      Index           =   0
      Left            =   -9720
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape box1R1C 
      BackColor       =   &H8000000D&
      BorderWidth     =   2
      Height          =   735
      Index           =   0
      Left            =   -9720
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape box2R1C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1455
      Index           =   0
      Left            =   -14985
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape box1R2C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   735
      Index           =   0
      Left            =   -9720
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape box4R1C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Height          =   2895
      Index           =   0
      Left            =   -9720
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape box1R4C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H0080FF80&
      BorderWidth     =   2
      Height          =   735
      Index           =   0
      Left            =   -9720
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Shape box4R2C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2895
      Index           =   0
      Left            =   -9720
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape box2R4C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1455
      Index           =   0
      Left            =   -9720
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   4
      Left            =   9480
      TabIndex        =   28
      Top             =   8520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   3
      Left            =   9480
      TabIndex        =   27
      Top             =   7800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   2
      Left            =   9480
      TabIndex        =   26
      Top             =   7080
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   1
      Left            =   9480
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      Height          =   195
      Index           =   4
      Left            =   12000
      TabIndex        =   24
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   3
      Left            =   11490
      TabIndex        =   23
      Top             =   6000
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   2
      Left            =   10770
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   1
      Left            =   9930
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   0
      Left            =   9120
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblC 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   0
      Left            =   9480
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   -648
      X2              =   -616
      Y1              =   416
      Y2              =   384
   End
End
Attribute VB_Name = "frmKVDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit


Const vbWhite = &HFFFFFF
Const vbLightYellow = &H80FFFF
Const vbYellow = &HFFFF&
Const vbLightOrange = &H80C0FF
Const vbOrange = &H80FF&
Const vbLightGreen = &H80FF80
' vbGreen is a standard color
Const vbLightCyan = &HFFFF80
Const vbCyan = &HFFFF00
Const vbLightBlue = &HFF8080
' vbBlue is a standard color
Const vbLightPurple = &HFF80FF
Const vbPurple = &HFF00FF
Const vbMagenta = &HFF00FF
Dim Tested(1 To 4, 1 To 4) As String * 1

Private Sub SetValue(C1 As Byte, C2 As Byte)
 Tested(C1, C2) = "1"
End Sub

Public Function F() As String
 
 Dim C1 As Byte, C2 As Byte
 Dim C1a As Byte, C2a As Byte
 Dim Test As String, strTested As String, Bool As Boolean
 Dim Same As Boolean
 Dim Obj As Object
 ' Set tested to false
 For C1 = 1 To 4
  For C2 = 1 To 4
   Tested(C1, C2) = "0"
   txtKV(10 * C1 + C2).BackColor = vbWhite
   If txtKV(10 * C1 + C2) = "0" Then Tested(C1, C2) = "1"
  Next C2
 Next C1
 ' Unload all the gridlines
 For C1 = 1 To 4
  Select Case C1
   Case 1: Set Obj = box2R4C ' Mask8_Rows
   Case 2: Set Obj = box4R2C ' Mask8_Columns
   Case 3: Set Obj = box1R4C ' Mask4_Rows
   Case 4: Set Obj = box4R1C ' Mask4_Columns
   Case 5: Set Obj = box2R2C ' Mask4_Cubes
   Case 6: Set Obj = box1R2C ' Mask2_Rows
   Case 7: Set Obj = box2R1C ' Mask2_Columns
   Case 8: Set Obj = box1R1C ' Mask1
  End Select
  For C2 = 1 To Obj.Count - 1
   Unload Obj(C2)
  Next C2
 Next C1
Mask16:
 Test = txtKV(11) + txtKV(12) + txtKV(13) + txtKV(14)
 Test = Test + txtKV(21) + txtKV(22) + txtKV(23) + txtKV(24)
 Test = Test + txtKV(31) + txtKV(32) + txtKV(33) + txtKV(34)
 Test = Test + txtKV(41) + txtKV(42) + txtKV(43) + txtKV(44)
 If (InStr(Test, "0")) = 0 Then
  F = "1"
  GoTo EndFunction
 End If
Mask8_Rows:
 For C1 = 1 To 4
  If C1 < 4 Then
   C1a = C1 + 1
  Else
   C1a = 1
  End If
  Test = txtKV(C1 * 10 + 1) + txtKV(10 * C1 + 2) + txtKV(10 * C1 + 3) + txtKV(10 * C1 + 4)
  strTested = Tested(C1, 1) + Tested(C1, 2) + Tested(C1, 3) + Tested(C1, 4)
  Test = Test + txtKV(C1a * 10 + 1) + txtKV(10 * C1a + 2) + txtKV(10 * C1a + 3) + txtKV(10 * C1a + 4)
  strTested = strTested + Tested(C1a, 1) + Tested(C1a, 2) + Tested(C1a, 3) + Tested(C1a, 4)
  Bool = (InStr(strTested, "0") <> 0) And (InStr(Test, "0") = 0)
  If Bool Then
   Call SetValue(C1, 1): Call SetValue(C1a, 1)
   Call SetValue(C1, 2): Call SetValue(C1a, 2)
   Call SetValue(C1, 3): Call SetValue(C1a, 3)
   Call SetValue(C1, 4): Call SetValue(C1a, 4)
   Set Obj = box2R4C
   Load Obj(Obj.Count)
   With Obj(Obj.Count - 1)
    .Top = txtKV(10 * C1 + 1).Top - 120
    .Left = txtKV(10 * C1 + 1).Left - 120
    .Visible = True
   End With
   Same = Not ((CBool(Val(Left$(lblR(C1), 1)))) Xor (CBool(Val(Left$(lblR(C1a), 1)))))
   If Same Then
    If F <> Empty Then F = F + " + "
    F = F + Left$(lblR(0), 1)
    If Left$(lblR(C1), 1) = "0" Then F = F + "' "
   End If
   Same = Not ((CBool(Val(Right$(lblR(C1), 1)))) Xor (CBool(Val(Right$(lblR(C1a), 1)))))
   If Same Then
    If F <> Empty Then F = F + " + "
    F = F + Right$(lblR(0), 1)
    If Right$(lblR(C1), 1) = "0" Then F = F + "' "
   End If
  End If
 Next C1
Mask8_Columns:
 For C1 = 1 To 4
  If C1 < 4 Then
   C1a = C1 + 1
  Else
   C1a = 1
  End If
  Test = txtKV(10 + C1) + txtKV(20 + C1) + txtKV(30 + C1) + txtKV(40 + C1)
  strTested = Tested(1, C1) + Tested(2, C1) + Tested(3, C1) + Tested(4, C1)
  Test = Test + txtKV(10 + C1a) + txtKV(20 + C1a) + txtKV(30 + C1a) + txtKV(40 + C1a)
  strTested = strTested + Tested(1, C1a) + Tested(2, C1a) + Tested(3, C1a) + Tested(4, C1a)
  Bool = InStr(strTested, "0") <> 0 And (InStr(Test, "0")) = 0
  If Bool Then
   Call SetValue(1, C1): Call SetValue(1, C1a)
   Call SetValue(2, C1): Call SetValue(2, C1a)
   Call SetValue(3, C1): Call SetValue(3, C1a)
   Call SetValue(4, C1): Call SetValue(4, C1a)
   Set Obj = box4R2C
   Load Obj(Obj.Count)
   With Obj(Obj.Count - 1)
    .Top = txtKV(10 + C1).Top - 120
    .Left = txtKV(10 + C1).Left - 120
    .Visible = False 'True
   End With
   Same = Not ((CBool(Val(Left$(lblC(C1), 1)))) Xor (CBool(Val(Left$(lblC(C1a), 1)))))
   If Same Then
    If F <> Empty Then F = F + " + "
    F = F + Left$(lblC(0), 1)
    If Left$(lblC(C1), 1) = "0" Then F = F + "' "
   End If
   Same = Not ((CBool(Val(Right$(lblC(C1), 1)))) Xor (CBool(Val(Right$(lblC(C1a), 1)))))
   If Same Then
    If F <> Empty Then F = F + " + "
    F = F + Right$(lblC(0), 1)
    If Right$(lblC(C1), 1) = "0" Then F = F + "' "
   End If
  End If
 Next C1
Mask4_Rows:
 For C1 = 1 To 4
  Test = txtKV(10 * C1 + 1) + txtKV(10 * C1 + 2) + txtKV(10 * C1 + 3) + txtKV(10 * C1 + 4)
  strTested = Tested(C1, 1) + Tested(C1, 2) + Tested(C1, 3) + Tested(C1, 4)
  Bool = InStr(strTested, "0") <> 0 And InStr(Test, "0") = 0
  If Bool Then
   Call SetValue(C1, 1): Call SetValue(C1, 2)
   Call SetValue(C1, 3): Call SetValue(C1, 4)
   Set Obj = box1R4C
   Load Obj(Obj.Count)
   With Obj(Obj.Count - 1)
    .Top = txtKV(10 * C1 + 1).Top - 120
    .Left = txtKV(10 * C1 + 1).Left - 120
    .Visible = False 'True
   End With
   If F <> Empty Then F = F + " + "
    F = F + Left$(lblR(0), 1)
   If Left$(lblR(C1), 1) = "0" Then F = F + "' "
    F = F + Right$(lblR(0), 1)
   If Right$(lblR(C1), 1) = "0" Then F = F + "' "
  End If
 Next C1
Mask4_columns:
 For C1 = 1 To 4
  Test = txtKV(10 + C1) + txtKV(20 + C1) + txtKV(30 + C1) + txtKV(40 + C1)
  strTested = Tested(1, C1) + Tested(2, C1) + Tested(3, C1) + Tested(4, C1)
  Bool = InStr(strTested, "0") <> 0 And InStr(Test, "0") = 0
  If Bool Then
   Call SetValue(1, C1): Call SetValue(2, C1)
   Call SetValue(3, C1): Call SetValue(4, C1)
   Set Obj = box4R1C
   Load Obj(Obj.Count)
   With Obj(Obj.Count - 1)
    .Top = txtKV(10 + C1).Top - 120
    .Left = txtKV(10 + C1).Left - 120
    .Visible = False 'True
   End With
   If F <> Empty Then F = F + " + "
    F = F + Left$(lblC(0), 1)
   If Left$(lblC(C1), 1) = "0" Then F = F + "' "
    F = F + Right$(lblC(0), 1)
   If Right$(lblC(C1), 1) = "0" Then F = F + "' "
  End If
 Next C1
Mask4_Cubes:
 For C1 = 1 To 4
  If C1 < 4 Then
   C1a = C1 + 1
  Else
   C1a = 1
  End If
  For C2 = 1 To 4
   If C2 < 4 Then
    C2a = C2 + 1
   Else
    C2a = 1
   End If
   Test = txtKV(10 * C1 + C2) + txtKV(10 * C1 + C2a)
   strTested = Tested(C1, C2) + Tested(C1, C2a)
   Test = Test + txtKV(10 * C1a + C2) + txtKV(10 * C1a + C2a)
   strTested = strTested + Tested(C1a, C2) + Tested(C1a, C2a)
   Bool = InStr(strTested, "0") <> 0 And InStr(Test, "0") = 0
   If Bool Then
    Call SetValue(C1, C2): Call SetValue(C1a, C2)
    Call SetValue(C1, C2a): Call SetValue(C1a, C2a)
    Set Obj = box2R2C
    Load Obj(Obj.Count)
    With Obj(Obj.Count - 1)
     .Top = txtKV(10 * C1 + C2).Top - 120
     .Left = txtKV(10 * C1 + C2).Left - 120
     .Visible = False 'True
    End With
    If F <> Empty Then F = F + " + "
    Same = Not ((CBool(Val(Left$(lblR(C1), 1)))) Xor (CBool(Val(Left$(lblR(C1a), 1)))))
    If Same Then
     F = F + Left$(lblR(0), 1)
     If Left$(lblR(C1), 1) = "0" Then F = F + "' "
    End If
    Same = Not ((CBool(Val(Right$(lblR(C1), 1)))) Xor (CBool(Val(Right$(lblR(C1a), 1)))))
    If Same Then
     F = F + Right$(lblR(0), 1)
     If Right$(lblR(C1), 1) = "0" Then F = F + "' "
    End If
    Same = Not ((CBool(Val(Left$(lblC(C2), 1)))) Xor (CBool(Val(Left$(lblC(C2a), 1)))))
    If Same Then
     F = F + Left$(lblC(0), 1)
     If Left$(lblC(C2), 1) = "0" Then F = F + "' "
    End If
    Same = Not ((CBool(Val(Right$(lblC(C2), 1)))) Xor (CBool(Val(Right$(lblC(C2a), 1)))))
    If Same Then
     F = F + Right$(lblC(0), 1)
     If Right$(lblC(C2), 1) = "0" Then F = F + "' "
    End If
   End If
  Next C2
 Next C1
Mask2_Rows:
 For C1 = 1 To 4
  For C2 = 1 To 4
   If C2 < 4 Then
    C2a = C2 + 1
   Else
    C2a = 1
   End If
   Test = txtKV(10 * C1 + C2) + txtKV(10 * C1 + C2a)
   strTested = Tested(C1, C2) + Tested(C1, C2a)
   Bool = InStr(strTested, "0") <> 0 And InStr(Test, "0") = 0
   Bool = Bool And Not ((InStr("01", Left$(Test, 1)) = 0 And Tested(C1, C2a) = "1") Or (InStr("01", Right$(Test, 1)) = 0 And Tested(C1, C2) = "1"))
   If Bool Then
    Call SetValue(C1, C2): Call SetValue(C1, C2a)
    Set Obj = box1R2C
    Load Obj(Obj.Count)
    With Obj(Obj.Count - 1)
     .Top = txtKV(10 * C1 + C2).Top - 120
     .Left = txtKV(10 * C1 + C2).Left - 120
     .Visible = False 'True
    End With
    If F <> Empty Then F = F + " + "
    F = F + Left$(lblR(0), 1)
    If Left$(lblR(C1), 1) = "0" Then F = F + "' "
    F = F + Right$(lblR(0), 1)
    If Right$(lblR(C1), 1) = "0" Then F = F + "' "
    
    Same = Not ((CBool(Val(Left$(lblC(C2), 1)))) Xor (CBool(Val(Left$(lblC(C2a), 1)))))
    If Same Then
     F = F + Left$(lblC(0), 1)
     If Left$(lblC(C2), 1) = "0" Then F = F + "' "
    End If
    Same = Not ((CBool(Val(Right$(lblC(C2), 1)))) Xor (CBool(Val(Right$(lblC(C2a), 1)))))
    If Same Then
     F = F + Right$(lblC(0), 1)
     If Right$(lblC(C2), 1) = "0" Then F = F + "' "
    End If
   End If
  Next C2
 Next C1
Mask2_Columns:
 For C1 = 1 To 4
  If C1 < 4 Then
   C1a = C1 + 1
  Else
   C1a = 1
  End If
  For C2 = 1 To 4
   Test = txtKV(10 * C1 + C2) + txtKV(10 * C1a + C2)
   strTested = Tested(C1, C2) + Tested(C1a, C2)
   Bool = InStr(strTested, "0") <> 0 And InStr(Test, "0") = 0
   Bool = Bool And Not ((InStr("01", Left$(Test, 1)) = 0 And Tested(C1a, C2) = "1") Or (InStr("01", Right$(Test, 1)) = 0 And Tested(C1, C2) = "1"))
   If Bool Then
    Call SetValue(C1, C2)
    Call SetValue(C1a, C2)
    Set Obj = box2R1C
    Load Obj(Obj.Count)
    With Obj(Obj.Count - 1)
     .Top = txtKV(10 * C1 + C2).Top - 120
     .Left = txtKV(10 * C1 + C2).Left - 120
     .Visible = False 'True
    End With
    If F <> Empty Then F = F + " + "
    Same = Not ((CBool(Val(Left$(lblR(C1), 1)))) Xor (CBool(Val(Left$(lblR(C1a), 1)))))
    If Same Then
     F = F + Left$(lblR(0), 1)
     If Left$(lblR(C1), 1) = "0" Then F = F + "' "
    End If
    Same = Not ((CBool(Val(Right$(lblR(C1), 1)))) Xor (CBool(Val(Right$(lblR(C1a), 1)))))
    If Same Then
     F = F + Right$(lblR(0), 1)
     If Right$(lblR(C1), 1) = "0" Then F = F + "' "
    End If
    F = F + Left$(lblC(0), 1)
    If Left$(lblC(C2), 1) = "0" Then F = F + "' "
    F = F + Right$(lblC(0), 1)
    If Right$(lblC(C2), 1) = "0" Then F = F + "' "
   End If
  Next C2
 Next C1
Mask1:
 For C1 = 1 To 4
  For C2 = 1 To 4
   If Tested(C1, C2) = "0" And txtKV(10 * C1 + C2) = "1" Then
    Call SetValue(C1, C2)
    Set Obj = box1R1C
    Load Obj(Obj.Count)
    With Obj(Obj.Count - 1)
     .Top = txtKV(10 * C1 + C2).Top - 120
     .Left = txtKV(10 * C1 + C2).Left - 120
     .Visible = False 'True
    End With
    If F <> Empty Then F = F + " + "
    F = F + Left$(lblR(0), 1)
    If Left$(lblR(C1), 1) = "0" Then F = F + "' "
    F = F + Right$(lblR(0), 1)
    If Right$(lblR(C1), 1) = "0" Then F = F + "' "
    F = F + Left$(lblC(0), 1)
    If Left$(lblC(C2), 1) = "0" Then F = F + "' "
    F = F + Right$(lblC(0), 1)
    If Right$(lblC(C2), 1) = "0" Then F = F + "' "
   Else
    Tested(C1, C2) = "1"
   End If
  Next C2
 Next C1
EndFunction:
End Function

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()


End Sub


Private Sub Command3_Click()
'Crea la tabla de verdad

NInputs = Val(Me.Text1.Text)
If NInputs < 2 Or NInputs > 4 Then
    MsgBox "Wrong number of inputs (>1 <5).", 16
    Me.Text1.Text = ""
    Me.Text1.SetFocus
    Exit Sub
End If

Call LimpiaTabla

NumCombInput = 2 ^ NInputs

For r = 1 To NumCombInput
    Load Me.Entrada(r)
    Load Me.Salida(r)
    Me.Entrada(r).Left = Me.Entrada(0).Left
    Me.Entrada(r).Top = Me.Entrada(r - 1).Top + 25
    Me.Entrada(r).Visible = True
    Me.Salida(r).Left = Me.Salida(0).Left
    Me.Salida(r).Top = Me.Salida(r - 1).Top + 25
    Me.Salida(r).Visible = True
Next r

Me.Command4.Left = Me.Entrada(0).Left
Me.Command4.Top = Me.Entrada(NumCombInput).Top + 30
Me.Command4.Visible = True

Me.Command5.Left = Me.Entrada(0).Left
Me.Command5.Top = Me.Entrada(0).Top - 3 '- Me.Command5.Height
Me.Command5.Visible = True

'LLena las cajas input
For r = 0 To NumCombInput - 1
    Call PasaBinario(r, NInputs, StrBin)
    Me.Entrada(r + 1).Text = StrBin
Next r

Me.Salida(1).SetFocus
End Sub

Private Sub Command4_Click()
Me.HS.Value = 0
StrBin = ""

NInputs = Val(Me.Text1.Text)
Largo = 2 ^ NInputs
NumCombInput = 2 ^ NInputs

Me.Label4.Visible = True
Me.Text3.Visible = True

If Me.Check1.Value <> 1 Then
    func = ""
    For r = 1 To NumCombInput '- 1
    'MsgBox Me.Salida(r).Text
        If Val(Me.Salida(r).Text) = 1 Then
            Comb = Trim(Me.Entrada(r).Text)
            For t = 1 To Len(Comb)
                If Mid(Comb, t, 1) = "1" Then
                    func = func + Chr(64 + t)
                Else
                    func = func + Chr(64 + t) + "'"
                End If
            Next t
            func = func + "+"
        End If
    Next r
    If Right(func, 1) = "+" Then
        func = Left(func, Len(func) - 1)
        Me.Text2.Text = func
        GoTo 6
    End If
End If



'Inicializa mapa
Me.txtKV(11).Text = "0"
Me.txtKV(12).Text = "0"
Me.txtKV(13).Text = "0"
Me.txtKV(14).Text = "0"

Me.txtKV(21).Text = "0"
Me.txtKV(22).Text = "0"
Me.txtKV(23).Text = "0"
Me.txtKV(24).Text = "0"

Me.txtKV(31).Text = "0"
Me.txtKV(32).Text = "0"
Me.txtKV(33).Text = "0"
Me.txtKV(34).Text = "0"

Me.txtKV(41).Text = "0"
Me.txtKV(42).Text = "0"
Me.txtKV(43).Text = "0"
Me.txtKV(44).Text = "0"

For r = 1 To NumCombInput
    If Trim(Me.Salida(r).Text) = "" Then
        MsgBox "Truth table incomplete", 16
        Exit Sub
    End If
    If Val(Me.Salida(r).Text) = 1 Then
        'Si es 1 decide en que casilla colocarlo
        Call PasaBinario(r - 1, NInputs, StrEntrada)
        If NInputs = 2 Then
            'A 2 inputs solo se consideran las variables B y D, las otras se ignoran
            bit1 = Val(Left(StrEntrada, 1))
            bit2 = Val(Right(StrEntrada, 1))
            If bit1 = 0 Then
              a1 = 1
            Else
              a1 = 2
            End If
            If bit2 = 0 Then
                a2 = 1
            Else
                a2 = 2
            End If
            pose = a2 * 10 + a1
            Me.txtKV(pose).Text = "1"
            

        End If
        If NInputs = 3 Then
            'En este caso se descarta la C
            If r - 1 = 0 Then pose = 11
            If r - 1 = 1 Then pose = 21
            If r - 1 = 2 Then pose = 12
            If r - 1 = 3 Then pose = 22
            
            If r - 1 = 4 Then pose = 14
            If r - 1 = 5 Then pose = 24
            If r - 1 = 6 Then pose = 13
            If r - 1 = 7 Then pose = 23
            
            Me.txtKV(pose).Text = "1"
            
            

        End If
        If NInputs = 4 Then
            If r - 1 = 0 Then pose = 11
            If r - 1 = 1 Then pose = 21
            If r - 1 = 2 Then pose = 12
            If r - 1 = 3 Then pose = 22
            
            If r - 1 = 4 Then pose = 14
            If r - 1 = 5 Then pose = 24
            If r - 1 = 6 Then pose = 13
            If r - 1 = 7 Then pose = 23
            
            If r - 1 = 8 Then pose = 31
            If r - 1 = 9 Then pose = 32
            If r - 1 = 10 Then pose = 34
            If r - 1 = 11 Then pose = 33
            
            If r - 1 = 12 Then pose = 41
            If r - 1 = 13 Then pose = 42
            If r - 1 = 14 Then pose = 44
            If r - 1 = 15 Then pose = 43
            
            
            Me.txtKV(pose).Text = "1"
            
        End If
    End If
Next r

Me.Text2.Text = F

If NInputs = 2 Then
    'Filtramos la A y la C
    funcion = Me.Text2.Text
    For p = 1 To Len(funcion)
        If Mid(funcion, p, 1) = "C" Then
            funcion = Left(funcion, p - 1) & Right(funcion, Len(funcion) - p)
            'Mira si hay tilde
            If Mid(funcion, p, 1) = "'" Then
                funcion = Left(funcion, p - 1) & Right(funcion, Len(funcion) - p)
            End If
        End If
        If Mid(funcion, p, 1) = "A" Then
            funcion = Left(funcion, p - 1) & Right(funcion, Len(funcion) - p)
            'Mira si hay tilde
            If Mid(funcion, p, 1) = "'" Then
                funcion = Left(funcion, p - 1) & Right(funcion, Len(funcion) - p)
            End If
        End If
    Next p
    'Reemplaza la B por la A
    For r = 1 To Len(funcion)
        If Mid(funcion, r, 1) = "B" Then
            funcion = Left(funcion, r - 1) & "A" & Right(funcion, Len(funcion) - r)
        End If
    Next r
    'Reemplaza la D por la B
    For r = 1 To Len(funcion)
        If Mid(funcion, r, 1) = "D" Then
            funcion = Left(funcion, r - 1) & "B" & Right(funcion, Len(funcion) - r)
        End If
    Next r
    Me.Text2.Text = funcion
End If

If NInputs = 3 Then
    'Filtramos  la C
    funcion = Me.Text2.Text
    For p = 1 To Len(funcion)
        If Mid(funcion, p, 1) = "C" Then
            funcion = Left(funcion, p - 1) & Right(funcion, Len(funcion) - p)
            'Mira si hay tilde
            If Mid(funcion, p, 1) = "'" Then
                funcion = Left(funcion, p - 1) & Right(funcion, Len(funcion) - p)
            End If
        End If
    Next p
    funcion = UCase(funcion)
    'Reemplaza la D por la C
    For r = 1 To Len(funcion)
        If Mid(funcion, r, 1) = "D" Then
            funcion = Left(funcion, r - 1) & "C" & Right(funcion, Len(funcion) - r)
        End If
    Next r
    Me.Text2.Text = funcion
End If

6 'Determinal el numero de camaras y las carga en la pantalla
paso = Me.Camara(0).Width + 20
funcion = Me.Text2.Text

Nc = 1
For r = 1 To Len(funcion)
    If Mid(funcion, r, 1) = "+" Then Nc = Nc + 1
Next r

descargaCamaras

For r = 1 To Nc - 1
    Load Me.Camara(r): Me.Camara(r).Top = Me.Camara(0).Top: Me.Camara(r).Left = Me.Camara(0).Left + paso * r
    Load Me.C1(r): Me.C1(r).Top = Me.C1(0).Top: Me.C1(r).Left = Me.C1(0).Left + paso * r: Me.C1(r).ZOrder 0
    Load Me.C2(r): Me.C2(r).Top = Me.C2(0).Top: Me.C2(r).Left = Me.C2(0).Left + paso * r: Me.C2(r).ZOrder 0
    Load Me.C3(r): Me.C3(r).Top = Me.C3(0).Top: Me.C3(r).Left = Me.C3(0).Left + paso * r: Me.C3(r).ZOrder 0
    Load Me.C4(r): Me.C4(r).Top = Me.C4(0).Top: Me.C4(r).Left = Me.C4(0).Left + paso * r: Me.C4(r).ZOrder 0
    Load Me.T1(r): Me.T1(r).Top = Me.T1(0).Top: Me.T1(r).Left = Me.T1(0).Left + paso * r: Me.T1(r).ZOrder 0
    Load Me.T2(r): Me.T2(r).Top = Me.T2(0).Top: Me.T2(r).Left = Me.T2(0).Left + paso * r: Me.T2(r).ZOrder 0
    Load Me.T3(r): Me.T3(r).Top = Me.T3(0).Top: Me.T3(r).Left = Me.T3(0).Left + paso * r: Me.T3(r).ZOrder 0
    Load Me.T4(r): Me.T4(r).Top = Me.T4(0).Top: Me.T4(r).Left = Me.T4(0).Left + paso * r: Me.T4(r).ZOrder 0
    Load Me.NOT(r): Me.NOT(r).Top = Me.NOT(0).Top: Me.NOT(r).Left = Me.NOT(0).Left + paso * r: Me.NOT(r).ZOrder 0
    Load Me.TNOT(r): Me.TNOT(r).Top = Me.TNOT(0).Top: Me.TNOT(r).Left = Me.TNOT(0).Left + paso * r: Me.TNOT(r).ZOrder 0
    Load Me.Conector(r): Me.Conector(r).Top = Me.Conector(0).Top: Me.Conector(r).Left = Me.Conector(0).Left + paso * r: Me.Conector(r).ZOrder 0
    Load Me.Reprog(r): Me.Reprog(r).Top = Me.Reprog(0).Top: Me.Reprog(r).Left = Me.Reprog(0).Left + paso * r: Me.Reprog(r).ZOrder 0
Next r


For r = 0 To Nc - 1
    frmKVDiagram.Camara(r).Visible = True
    frmKVDiagram.C1(r).Visible = False
    frmKVDiagram.C2(r).Visible = False
    frmKVDiagram.C3(r).Visible = False
    frmKVDiagram.C4(r).Visible = False
    frmKVDiagram.T1(r).Visible = False
    frmKVDiagram.T2(r).Visible = False
    frmKVDiagram.T3(r).Visible = False
    frmKVDiagram.T4(r).Visible = False
    frmKVDiagram.NOT(r).Visible = True
    frmKVDiagram.TNOT(r).Visible = True
    frmKVDiagram.Conector(r).Visible = True
    frmKVDiagram.Reprog(r).Caption = "Chamber #" & Trim(Str(r + 1))
    frmKVDiagram.Reprog(r).Visible = True
    
Next r

frmKVDiagram.output.Visible = True
frmKVDiagram.output.Width = Me.Camara(Nc - 1).Left + Me.Camara(Nc - 1).Width - Me.Camara(0).Left
Me.Label3.Left = Me.output.Left + Me.output.Width / 2 - Me.Label3.Width / 2
Me.Label3.ZOrder 0
Me.Label3.Visible = True
Me.fondo.Width = Me.Camara(Nc - 1).Left + Me.Camara(Nc - 1).Width - Me.Camara(0).Left + 30
Me.Picture3.Width = Me.Camara(Nc - 1).Left + Me.Camara(Nc - 1).Width - Me.Camara(0).Left + 30
Me.fondo.Visible = True
Me.fondo.ZOrder 1


'Asigna textos y elimina celulas no necesarias

funcion = funcion & "+"

cont = 0
Do
    funcion = Trim(funcion)
    pp = InStr(1, funcion, "+")
    If pp = 0 Then Exit Sub
    cam = Trim(Left(funcion, pp - 1))
    funcion = Trim(Right(funcion, Len(funcion) - pp))
    Do
        If Len(cam) = 0 Then Exit Do
        letra = Left(cam, 1)
        cam = Trim(Right(cam, Len(cam) - 1))
        'mira si es not
        no = 0
        If Left(cam, 1) = "'" Then
            no = 1
            cam = Trim(Right(cam, Len(cam) - 1))
        End If
        'Escribe
        If letra = "A" Then
            frmKVDiagram.C1(cont).Visible = True
            frmKVDiagram.T1(cont).Visible = True
            If no = 0 Then
                frmKVDiagram.T1(cont).Caption = "NOT" & Chr(13) & Chr(10) & "(A)"
            Else
                frmKVDiagram.T1(cont).Caption = "ID" & Chr(13) & Chr(10) & "(A)"
            End If
        End If
    
        If letra = "B" Then
            frmKVDiagram.C2(cont).Visible = True
            frmKVDiagram.T2(cont).Visible = True
            If no = 0 Then
                frmKVDiagram.T2(cont).Caption = "NOT" & Chr(13) & Chr(10) & "(B)"
            Else
                frmKVDiagram.T2(cont).Caption = "ID" & Chr(13) & Chr(10) & "(B)"
            End If
        End If
        
        If letra = "C" Then
            frmKVDiagram.C3(cont).Visible = True
            frmKVDiagram.T3(cont).Visible = True
            If no = 0 Then
                frmKVDiagram.T3(cont).Caption = "NOT" & Chr(13) & Chr(10) & "(C)"
            Else
                frmKVDiagram.T3(cont).Caption = "ID" & Chr(13) & Chr(10) & "(C)"
            End If
        End If
        
        If letra = "D" Then
            frmKVDiagram.C4(cont).Visible = True
            frmKVDiagram.T4(cont).Visible = True
            If no = 0 Then
                frmKVDiagram.T4(cont).Caption = "NOT" & Chr(13) & Chr(10) & "(D)"
            Else
                frmKVDiagram.T4(cont).Caption = "ID" & Chr(13) & Chr(10) & "(D)"
            End If
        End If
    
    
    Loop
    cont = cont + 1
Loop


End Sub

Private Sub Command5_Click()
On Error GoTo 1
i = 0
Do
    Me.Salida(i).Text = ""
    i = i + 1
Loop
1 End Sub

Private Sub Form_Resize()

Me.Picture2.Width = Me.ScaleWidth - Me.Picture2.Left - 30
Me.HS.Width = Me.Picture2.Width - 3

End Sub



Private Sub HS_Change()
valor = Me.HS.Value

If Me.Picture3.Width <= Me.Picture2.Width Then
    Me.HS.Value = 0
Else
    Me.Picture3.Left = -(valor / 100) * (Me.Picture3.ScaleWidth - Me.Picture2.ScaleWidth)
End If
End Sub


Private Sub Reprog_Click(Index As Integer)

frase = Me.Reprog(Index).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
frase = frase & "External Inputs: {"
If Me.T1(Index).Visible = True Then frase = frase & "A "
If Me.T2(Index).Visible = True Then frase = frase & "B "
If Me.T3(Index).Visible = True Then frase = frase & "C "
If Me.T4(Index).Visible = True Then frase = frase & "D "
frase = Trim(frase) & "}"
frase = frase & Chr(13) & Chr(10)
frase = frase & Chr(13) & Chr(10)
frase = frase & "Input Layer Composition" & Chr(13) & Chr(10)
frase = frase & "===================" & Chr(13) & Chr(10)

Tex1 = UCase(Me.T1(Index).Caption): Tx1 = "":
For i = 1 To Len(Tex1):
    a = Mid(Tex1, i, 1):
    If Asc(a) = 13 Then a = " ":
    If Asc(a) = 10 Then a = " ":
    Tx1 = Tx1 & a:
Next i
Tex2 = UCase(Me.T2(Index).Caption): Tx2 = "":
For i = 1 To Len(Tex2):
    a = Mid(Tex2, i, 1):
    If Asc(a) = 13 Then a = " ":
    If Asc(a) = 10 Then a = " ":
    Tx2 = Tx2 & a:
Next i
Tex3 = UCase(Me.T3(Index).Caption): Tx3 = "":
    For i = 1 To Len(Tex3):
        a = Mid(Tex3, i, 1):
        If Asc(a) = 13 Then a = " ":
        If Asc(a) = 10 Then a = " ":
        Tx3 = Tx3 & a:
    Next i
Tex4 = UCase(Me.T4(Index).Caption): Tx4 = "":
For i = 1 To Len(Tex4):
    a = Mid(Tex4, i, 1):
    If Asc(a) = 13 Then a = " ":
    If Asc(a) = 10 Then a = " ":
    Tx4 = Tx4 & a:
Next i
c = 0
If Me.T1(Index).Visible = True Then c = c + 1: frase = frase & Space$(5) & "Cell #" & Trim(Str(c)) & ":  " & Tx1 & Chr(13) & Chr(10)
If Me.T2(Index).Visible = True Then c = c + 1: frase = frase & Space$(5) & "Cell #" & Trim(Str(c)) & ":  " & Tx2 & Chr(13) & Chr(10)
If Me.T3(Index).Visible = True Then c = c + 1: frase = frase & Space$(5) & "Cell #" & Trim(Str(c)) & ":  " & Tx3 & Chr(13) & Chr(10)
If Me.T4(Index).Visible = True Then c = c + 1: frase = frase & Space$(5) & "Cell #" & Trim(Str(c)) & ":  " & Tx4 & Chr(13) & Chr(10)

frase = frase & Chr(13) & Chr(10)
frase = frase & Chr(13) & Chr(10)
frase = frase & "Reprogramers" & Chr(13) & Chr(10)
frase = frase & "===========" & Chr(13) & Chr(10)

sw = 0
If Me.T1(Index).Visible = True Then
    pp = InStr(1, UCase(Me.T1(Index).Caption), "NOT")
    If pp > 0 Then
        frase = frase & Space$(0) & "Reprog. Cell #1" & Chr(13) & Chr(10): sw = 1
    End If
End If
If Me.T2(Index).Visible = True Then
    pp = InStr(1, UCase(Me.T2(Index).Caption), "NOT")
    If pp > 0 Then
        frase = frase & Space$(0) & "Reprog. Cell #2" & Chr(13) & Chr(10): sw = 1
    End If
End If
If Me.T3(Index).Visible = True Then
    pp = InStr(1, UCase(Me.T3(Index).Caption), "NOT")
    If pp > 0 Then
        frase = frase & Space$(0) & "Reprog. Cell #3" & Chr(13) & Chr(10): sw = 1
    End If
End If
If Me.T4(Index).Visible = True Then
    pp = InStr(1, UCase(Me.T4(Index).Caption), "NOT")
    If pp > 0 Then
        frase = frase & Space$(0) & "Reprog. Cell #4" & Chr(13) & Chr(10): sw = 1
    End If
End If

If sw = 0 Then
    frase = frase & "Reprogrammers Not Required"
End If
Me.Text3.Text = frase
End Sub

Private Sub Salida_KeyPress(Index As Integer, KeyAscii As Integer)
Nc = 2 ^ Val(Me.Text1.Text)
If KeyAscii = 13 Then
    NextIndex = Index + 1
    If NextIndex > Nc Then NextIndex = 1
    Me.Salida(NextIndex).SetFocus
End If

If KeyAscii <> 48 And KeyAscii <> 49 Then Me.Salida(Index).Text = "": Exit Sub
    NextIndex = Index + 1
    If NextIndex > Nc Then NextIndex = 1
    Me.Salida(NextIndex).SetFocus
End Sub


Private Sub Salida_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode <> 48 And KeyCode <> 49 Then Me.Salida(Index).Text = "": Exit Sub
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Call Command3_Click
End Sub


