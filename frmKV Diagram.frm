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
   Begin VB.CheckBox Check1 
      Caption         =   "Simplification"
      Height          =   255
      Left            =   480
      TabIndex        =   41
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4680
      TabIndex        =   35
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Generate"
      Height          =   375
      Left            =   840
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2280
      TabIndex        =   33
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
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Procesar Info"
      Height          =   615
      Left            =   6480
      TabIndex        =   30
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Funciones hasta 3-Inputs"
      Height          =   615
      Left            =   13200
      TabIndex        =   29
      Top             =   6120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   44
      Left            =   12000
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
      Left            =   11280
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
      Left            =   10560
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
      Left            =   9840
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
      Left            =   12000
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
      Left            =   11280
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
      Left            =   10560
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
      Left            =   9840
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
      Left            =   12000
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
      Left            =   11280
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
      Left            =   10560
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
      Left            =   9840
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
      Left            =   12000
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
      Left            =   11280
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
      Left            =   10560
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
      Left            =   9840
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
      Left            =   3720
      TabIndex        =   43
      Top             =   4560
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Shape Conector 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   3960
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape output 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   615
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Canninical Function"
      Height          =   210
      Left            =   3240
      TabIndex        =   42
      Top             =   600
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   3960
      TabIndex        =   40
      Top             =   3480
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Shape NOT 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   0
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
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
      Left            =   4320
      TabIndex        =   39
      Top             =   2520
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape C4 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   0
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
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
      Left            =   3600
      TabIndex        =   38
      Top             =   2520
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Shape C3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   0
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
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
      Left            =   4320
      TabIndex        =   37
      Top             =   1800
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape C2 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   0
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
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
      Left            =   3600
      TabIndex        =   36
      Top             =   1800
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape C1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   0
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Camara 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   2775
      Index           =   0
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Num. Inputs:"
      Height          =   195
      Left            =   480
      TabIndex        =   32
      Top             =   480
      Width           =   900
   End
   Begin VB.Shape box2R2C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   1455
      Index           =   0
      Left            =   9720
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
      Left            =   9720
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
      Left            =   9720
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
      Left            =   9720
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
      Left            =   9720
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
      Left            =   9720
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
      Left            =   9720
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
      Left            =   9720
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "10"
      Height          =   195
      Index           =   4
      Left            =   9480
      TabIndex        =   28
      Top             =   8520
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "11"
      Height          =   195
      Index           =   3
      Left            =   9480
      TabIndex        =   27
      Top             =   7800
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "01"
      Height          =   195
      Index           =   2
      Left            =   9480
      TabIndex        =   26
      Top             =   7080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "00"
      Height          =   195
      Index           =   1
      Left            =   9480
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      Caption         =   "10"
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
      Caption         =   "11"
      Height          =   195
      Index           =   3
      Left            =   11280
      TabIndex        =   23
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "01"
      Height          =   195
      Index           =   2
      Left            =   10560
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "00"
      Height          =   195
      Index           =   1
      Left            =   9720
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "CD"
      Height          =   195
      Index           =   0
      Left            =   9120
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblC 
      AutoSize        =   -1  'True
      Caption         =   "AB"
      Height          =   195
      Index           =   0
      Left            =   9480
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   648
      X2              =   616
      Y1              =   416
      Y2              =   384
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   3000
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmKVDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit


' ***********************************************
' ** This program solves a 4x4 KV-diagram      **
' ** It even looks for random values indicated **
' ** by 'X' or another non-numeric symbol.     **
' ** Input: values in the KV-diagram           **
' ** Ouput: the simpelest equation possible.   **
' ***********************************************

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

Private Sub cmdSimplify_Click()



 MsgBox F
End Sub

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
Dim Strings() As String

mm = FreeFile
Open "c:\circuitos_3.txt" For Output As mm
Close #mm

NInp = 3
Largo = 2 ^ NInp
For Num = 0 To 2 ^ Largo - 1     'Num en binario representa una combinacion de outputs
    Me.Caption = "Num=" & Num: DoEvents
    Call PasaBinario(Num, Largo, StrBin)
    For t = 1 To Largo
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
        'Coloca este string en el mapa de karnaugh
        Me.txtKV(11).Text = Mid(StrBin, Largo, 1): Me.txtKV(21).Text = Mid(StrBin, Largo, 1)
        Me.txtKV(12).Text = Mid(StrBin, Largo - 1, 1): Me.txtKV(22).Text = Mid(StrBin, Largo - 1, 1)
        Me.txtKV(13).Text = Mid(StrBin, Largo - 3, 1): Me.txtKV(23).Text = Mid(StrBin, Largo - 3, 1)
        Me.txtKV(14).Text = Mid(StrBin, Largo - 2, 1): Me.txtKV(24).Text = Mid(StrBin, Largo - 2, 1)
        
        Me.txtKV(31).Text = Mid(StrBin, Largo - 4, 1): Me.txtKV(41).Text = Mid(StrBin, Largo - 4, 1)
        Me.txtKV(32).Text = Mid(StrBin, Largo - 5, 1): Me.txtKV(42).Text = Mid(StrBin, Largo - 5, 1)
        Me.txtKV(33).Text = Mid(StrBin, Largo - 7, 1): Me.txtKV(43).Text = Mid(StrBin, Largo - 7, 1)
        Me.txtKV(34).Text = Mid(StrBin, Largo - 6, 1): Me.txtKV(44).Text = Mid(StrBin, Largo - 6, 1)
    Next t
    'Presenta la funcion canonica simplificada
    funcion = Trim(F) & " "
    'MsgBox Funcion
    If Trim(funcion) <> "" Then
        '**************************************************************************************************
        '*                                  MODELO CON LOGICA INVERSA                                     *
        '**************************************************************************************************
        '1º Determina el numero de camaras/wires contando el numero de signos +
        NumCam = 1
        For t = 1 To Len(funcion)
            If Mid(funcion, t, 1) = "+" Then NumCam = NumCam + 1
        Next t
        '2º Deterina el numero total de celulas que intervienen en el circuito
        NumTotCells = 0
        For t = 1 To Len(funcion)
            If Mid(funcion, t, 1) <> "+" And Mid(funcion, t, 1) <> " " And Mid(funcion, t, 1) <> "'" Then NumTotCells = NumTotCells + 1
        Next t
        '3º Deterina los diferentes tipos celulars que participan
        IdCells = ""
        NotCells = ""
        For t = 1 To Len(funcion)
            If Mid(funcion, t, 1) <> "+" And Mid(funcion, t, 1) <> " " And Mid(funcion, t, 1) <> "'" Then
                If Mid(funcion, t + 1, 1) = "'" Then
                    'Si entra aquí es una puerta NOT
                    'Mira si previamente ya existe
                    If InStr(1, NotCells, "#" & Mid(funcion, t, 1) & "$") = 0 Then
                        NotCells = NotCells & "#" & Mid(funcion, t, 1) & "$"
                    End If
                Else
                    'Si entra aquí es una puerta Id
                    'Mira si previamente ya existe
                    If InStr(1, IdCells, "#" & Mid(funcion, t, 1) & "$") = 0 Then
                        IdCells = IdCells & "#" & Mid(funcion, t, 1) & "$"
                    End If
                End If
            End If
        Next t
        'Hace recuento de puertas NOT
        NumNOTs = 0
        For t = 1 To Len(NotCells)
            If Mid(NotCells, t, 1) = "#" Then NumNOTs = NumNOTs + 1
        Next t
        'Hace recuento de puertas Id
        NumIDs = 0
        For t = 1 To Len(IdCells)
            If Mid(IdCells, t, 1) = "#" Then NumIDs = NumIDs + 1
        Next t
        
        '**************************************************************************************************
        '*                                  MODELO CON LOGICA DIRECTA (Nature)                            *
        '**************************************************************************************************
        '1º Determina el numero de wires
        NumWires = 0
        For t = 1 To Len(funcion)
            If Mid(funcion, t, 1) <> "+" And Mid(funcion, t, 1) <> " " And Mid(funcion, t, 1) <> "'" Then NumWires = NumWires + 1
            If Mid(funcion, t, 1) = "+" Then NumWires = NumWires - 1
        Next t
        NumWires = NumWires - 1
        
        '2º Determina el numero de celulas diferentes
        ReDim Strings(100) As String
        Ft = Trim(funcion) & "+"
        Cv = 0
        Do
            pp = InStr(1, Ft, "+")
            If pp = 0 Then Exit Do
            Info = Trim(Left(Ft, pp - 1)) & " "
            Ft = Trim(Right(Ft, Len(Ft) - pp))
            Cv = Cv + 1
            For w = 1 To Len(Info) - 1
                If Mid(Info, w, 1) <> " " And Mid(Info, w, 1) <> "'" And Mid(Info, w, 1) <> "+" Then
                    If Mid(Info, w + 1, 1) = "'" Then
                        Dato = UCase(Mid(Info, w, 1))
                    Else
                        Dato = LCase(Mid(Info, w, 1))
                    End If
                    Strings(Cv) = Strings(Cv) & Dato
                End If
            Next w
        Loop
        
        For t = 1 To Cv
            Strings(t) = Trim(Strings(t))
        Next t
        'Mira si hay celulas repetidas. Eso significa que detectan la misma señal externa (la misma letra) y la letra anterior debe ser la misma o bien ser la primera
        NumTotCells2 = NumTotCells
        For t = 97 + NInp To 97 Step -1
            GateId = Chr(t)
            For v = 1 To Cv
                pp = InStr(1, Strings(v), GateId)
                If pp > 1 And pp < Len(Strings(v)) Then
                    'Si está presente crea la secuencia de localización
                    SecGates = Mid(Strings(v), pp - 1, 1) & Mid(Strings(v), pp, 1)
                    'Para esta secuencia mira si existe en el restro de cadenas
                    For y = Cv + 1 To Cv
                        pp = InStr(1, Strings(y), SecGates)
                        If pp > 0 And y <> v Then
                            'Si entra aqui hay dos pares de gates iguales, por tanto podemos reducir un wire y un tipo celular (al anterior)
                            NumWires = NumWires - 1
                            NumIDs = NumID - 1
                        End If
                    Next y
                ElseIf pp = 1 And pp < Len(Strings(v)) Then
                    NumWires = NumWires - 1
                    NumIDs = NumIDs - 1
                End If
            Next v
        Next t
        For t = 97 + NInp To 97 Step -1
            GateNOT = UCase(Chr(t))
            For v = 1 To Cv
                pp = InStr(1, Strings(v), GateNOT)
                If pp > 1 And pp < Len(Strings(v)) Then
                    'Si está presente crea la secuencia de localización
                    SecGates = Mid(Strings(v), pp - 1, 1) & Mid(Strings(v), pp, 1)
                    'Para esta secuencia mira si existe en el restro de cadenas
                    For y = v + 1 To Cv
                        pp = InStr(1, Strings(y), SecGates)
                        If pp > 0 And y <> v Then
                            'Si entra aqui hay dos pares de gates iguales, por tanto podemos reducir un wire y un tipo celular (al anterior)
                            NumWires = NumWires - 1
                            NumTotCells2 = NumTotCells2 - 1
                        End If
                    Next y
                ElseIf pp = 1 And pp < Len(Strings(v)) Then
                    For y = v + 1 To Cv
                        pp = InStr(1, Strings(y), GateNOT)
                        If pp = 1 And y <> v Then
                            'Si entra aqui hay dos pares de gates iguales, por tanto podemos reducir un wire y un tipo celular (al anterior)
                            NumWires = NumWires - 1
                            NumTotCells2 = NumTotCells2 - 1
                        End If
                    Next y
                End If
            Next v
        Next t
        'Si una gate esta presente en todos los strings el nuemero de wires puede reducirse
        For t = 97 + NInp To 97 Step -1
            GateId = Chr(t)
            For y = 1 To Cv
                pp = InStr(1, Strings(y), GateId)
                If pp = 0 Then GoTo 1
            Next y
        Next t
        NumWires = NumWires - Cv + 1
        
1       For t = 97 + NInp To 97 Step -1
            GateNOT = UCase(Chr(t))
            For y = 1 To Cv
                pp = InStr(1, Strings(y), GateNOT)
                If pp = 0 Then GoTo 2
            Next y
        Next t
        NumWires = NumWires - Cv + 1
        
2       'Graba los resultados
        mm = FreeFile
        Open "c:\circuitos_3.txt" For Append As mm
            Print #mm, Num, NumCam, NumTotCells + NumCam, NumNOTs, NumIDs, NumTotCells2, NumWires
        Close #mm
    End If
    

Next Num

MsgBox "END"
End Sub

Private Sub Command2_Click()

Dim HistoCamsLI(1000)
Dim HistoCellsLI(1000)
Dim HistoIdLI(1000)
Dim HistoNOTLI(1000)

Dim HistoWiresLD(1000)
Dim HistoCellsLD(1000)

mm = FreeFile
Open "c:\circuitos_3.txt" For Input As mm
    While Not EOF(mm)
        Input #mm, Num, NumCam, NumTotCells, NumNOTs, NumIDs, NumTotCells2, NumWires
        HistoCamsLI(NumCam) = HistoCamsLI(NumCam) + 1
        HistoCellsLI(NumTotCells) = HistoCellsLI(NumTotCells) + 1
        HistoIdLI(NumIDs) = HistoIdLI(NumIDs) + 1
        HistoNOTLI(NumNOTs) = HistoNOTLI(NumNOTs) + 1
        HistoWiresLD(NumWires) = HistoWiresLD(NumWires) + 1
        HistoCellsLD(NumTotCells2) = HistoCellsLD(NumTotCells2) + 1
    Wend
Close #mm

'Graba Ficheros
mm = FreeFile
Open "c:\Histo_Cells_LI.txt" For Output As mm
    Acum = 0
    For t = 1 To 1000
        Acum = Acum + HistoCellsLI(t)
        Print #mm, t, Acum
    Next t
Close #mm

mm = FreeFile
Open "c:\Histo_Cams_LI.txt" For Output As mm
    Acum = 0
    For t = 1 To 1000
        Acum = Acum + HistoCamsLI(t)
        Print #mm, t, Acum
    Next t
Close #mm

mm = FreeFile
Open "c:\Histo_Id_LI.txt" For Output As mm
    Acum = 0
    For t = 1 To 1000
        Acum = Acum + HistoIdLI(t)
        Print #mm, t, Acum
    Next t
Close #mm

mm = FreeFile
Open "c:\Histo_NOT_LI.txt" For Output As mm
    Acum = 0
    For t = 1 To 1000
        Acum = Acum + HistoNOTLI(t)
        Print #mm, t, Acum
    Next t
Close #mm

mm = FreeFile
Open "c:\Histo_Wires_LD.txt" For Output As mm
    Acum = 0
    For t = 1 To 1000
        Acum = Acum + HistoWiresLD(t)
        Print #mm, t, Acum
    Next t
Close #mm

mm = FreeFile
Open "c:\Histo_Cells_LD.txt" For Output As mm
    Acum = 0
    For t = 1 To 1000
        Acum = Acum + HistoCellsLD(t)
        Print #mm, t, Acum
    Next t
Close #mm

MsgBox "END"

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

'LLena las cajas input
For r = 0 To NumCombInput - 1
    Call PasaBinario(r, NInputs, StrBin)
    Me.Entrada(r + 1).Text = StrBin
Next r

Me.Salida(1).SetFocus
End Sub

Private Sub Command4_Click()
StrBin = ""

NInputs = Val(Me.Text1.Text)
Largo = 2 ^ NInputs
NumCombInput = 2 ^ NInputs

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
Next r

frmKVDiagram.output.Visible = True
frmKVDiagram.output.Width = Me.Camara(Nc - 1).Left + Me.Camara(Nc - 1).Width - Me.Camara(0).Left
Me.Label3.Left = Me.output.Left + Me.output.Width / 2 - Me.Label3.Width / 2
Me.Label3.ZOrder 0
Me.Label3.Visible = True
Me.fondo.Width = Me.Camara(Nc - 1).Left + Me.Camara(Nc - 1).Width - Me.Camara(0).Left + 30
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


