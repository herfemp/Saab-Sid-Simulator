VERSION 5.00
Object = "{927716E6-2465-43E0-B6B8-7BF453E1A2ED}#1.0#0"; "DisplayMatOcx.ocx"
Begin VB.Form Saab_SID_Simulator_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saab SID Simulator"
   ClientHeight    =   6405
   ClientLeft      =   855
   ClientTop       =   1545
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Active Interface"
      Height          =   255
      Left            =   4920
      TabIndex        =   32
      Top             =   4080
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Custom USB Interface"
      Height          =   255
      Left            =   5160
      TabIndex        =   31
      Top             =   5400
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Custom Serial Interface"
      Height          =   255
      Left            =   5160
      TabIndex        =   30
      Top             =   5040
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "National Instruments (PCI-CAN, PCMCIA-CAN, etc.)"
      Height          =   495
      Left            =   5160
      TabIndex        =   29
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Timer ClockTimer 
      Interval        =   1000
      Left            =   7560
      Top             =   0
   End
   Begin VB.CheckBox CanBus_Listen_Option 
      Caption         =   "Can-Bus Listen"
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton CanBus_Transmit_Button 
      Caption         =   "Send"
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Clear_Button 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Display_Button 
      Caption         =   "Display"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox DisplayLine2 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   12
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox DisplayLine1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   12
      TabIndex        =   8
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton Exit_Button 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin DisplayMatOcx.DisplayMatrix DisplayMatrix2 
      Height          =   465
      Left            =   3360
      TabIndex        =   1
      Top             =   960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   820
      MtxForeColor    =   7
      MtxBackColor    =   0
      MtxBackForeColor=   0
      MtxFont         =   0
      MtxDotHeight    =   3
      MtxDotWidth     =   3
      MtxCaption      =   "123456789012"
   End
   Begin DisplayMatOcx.DisplayMatrix DisplayMatrix1 
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      MtxForeColor    =   7
      MtxBackColor    =   0
      MtxBackForeColor=   0
      MtxFont         =   0
      MtxDotHeight    =   3
      MtxDotWidth     =   3
      MtxCaption      =   "X*X*X*X*X*X*"
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "CAN Bus Interface Configuration"
      Height          =   255
      Left            =   4680
      TabIndex        =   28
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Line Line13 
      X1              =   4440
      X2              =   4440
      Y1              =   3720
      Y2              =   6240
   End
   Begin VB.Line Line12 
      X1              =   4440
      X2              =   4440
      Y1              =   2820
      Y2              =   2740
   End
   Begin VB.Line Line11 
      X1              =   2640
      X2              =   2640
      Y1              =   2820
      Y2              =   2740
   End
   Begin VB.Line Line10 
      X1              =   6360
      X2              =   6360
      Y1              =   2820
      Y2              =   2740
   End
   Begin VB.Line Line9 
      X1              =   4560
      X2              =   4560
      Y1              =   2820
      Y2              =   2740
   End
   Begin VB.Line Line8 
      X1              =   5880
      X2              =   6360
      Y1              =   2740
      Y2              =   2740
   End
   Begin VB.Line Line7 
      X1              =   4560
      X2              =   5040
      Y1              =   2740
      Y2              =   2740
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Real SID"
      Height          =   255
      Left            =   5160
      TabIndex        =   27
      Top             =   2640
      Width           =   735
   End
   Begin VB.Line Line6 
      X1              =   3960
      X2              =   4440
      Y1              =   2740
      Y2              =   2740
   End
   Begin VB.Line Line5 
      X1              =   2640
      X2              =   3120
      Y1              =   2740
      Y2              =   2740
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Simulator"
      Height          =   255
      Left            =   3200
      TabIndex        =   26
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label CAN_Msg_6 
      BackStyle       =   0  'Transparent
      Caption         =   "00 96 02 44 52 20 20 20"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   5880
      Width           =   3735
   End
   Begin VB.Label CAN_Msg_5 
      BackStyle       =   0  'Transparent
      Caption         =   "01 96 02 4E 53 50 4F 4E"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   24
      Top             =   5520
      Width           =   3735
   End
   Begin VB.Label CAN_Msg_4 
      BackStyle       =   0  'Transparent
      Caption         =   "02 96 02 32 20 54 52 41"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   5160
      Width           =   3735
   End
   Begin VB.Label CAN_Msg_3 
      BackStyle       =   0  'Transparent
      Caption         =   "03 96 01 45 59 20 20 20"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Label CAN_Msg_2 
      BackStyle       =   0  'Transparent
      Caption         =   "04 96 01 4F 54 45 20 4B"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Label CAN_Msg_1 
      BackStyle       =   0  'Transparent
      Caption         =   "45 96 01 32 20 52 45 4D"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   4080
      Width           =   3735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CAN Message Construction (ID 0x32F)"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7800
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3240
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   2280
      Shape           =   1  'Square
      Top             =   600
      Width           =   255
   End
   Begin VB.Label AMPM_Label 
      BackStyle       =   0  'Transparent
      Caption         =   "AM"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "DIST ARRIV ALARM SPD W"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   1680
      TabIndex        =   16
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TEMP   DTE    FUEL  SPD"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   2520
      TabIndex        =   15
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Time_Label 
      BackStyle       =   0  'Transparent
      Caption         =   "11:57"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   480
      TabIndex        =   13
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Text to Display"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   7800
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NIGHT PANEL"
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
      Left            =   6360
      TabIndex        =   6
      Top             =   1950
      Width           =   1455
   End
   Begin VB.Line Line3 
      X1              =   4560
      X2              =   4560
      Y1              =   1800
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2400
      Y1              =   1800
      Y2              =   2280
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   1875
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   1875
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SET"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   600
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   255
   End
   Begin VB.Shape Shape4 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Saab_SID_Simulator_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Clear_Button_Click()
    
    DisplayMatrix1.MtxCaption = ""
    DisplayMatrix2.MtxCaption = ""
    DisplayLine1.Text = ""
    DisplayLine2.Text = ""
    
    CAN_Msg_1.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_2.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_3.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_4.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_5.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_6.Caption = "XX XX XX XX XX XX XX XX"

End Sub

Private Sub ClockTimer_Timer()
    Time_Label.Caption = Format$(Now, "hh:mm")
    AMPM_Label.Caption = Format$(Now, "AMPM")
End Sub

Private Sub Display_Button_Click()
    DisplayMatrix1.MtxCaption = DisplayLine1.Text
    DisplayMatrix2.MtxCaption = DisplayLine2.Text
    
    If Len(DisplayLine1.Text) < 12 Then
        While Len(DisplayLine1.Text) < 12
            DisplayLine1.Text = DisplayLine1.Text & " "
        Wend
    End If
    
    If Len(DisplayLine2.Text) > 0 And Len(DisplayLine2.Text) < 12 Then
        While Len(DisplayLine2.Text) < 12
            DisplayLine2.Text = DisplayLine2.Text & " "
        Wend
    End If
    
    CAN_Msg_1.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_2.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_3.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_4.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_5.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_6.Caption = "XX XX XX XX XX XX XX XX"
    
    If Len(DisplayLine2.Text) = 0 Then
        LineCount = 1
    Else
        LineCount = 2
    End If
    
    If LineCount = 1 Then
        CAN_Msg_1.Caption = "42 96 01 "
    Else
        CAN_Msg_1.Caption = "45 96 01 "
    End If
    
    CAN_Msg_1.Caption = CAN_Msg_1.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 1, 1))) & " "
    CAN_Msg_1.Caption = CAN_Msg_1.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 2, 1))) & " "
    CAN_Msg_1.Caption = CAN_Msg_1.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 3, 1))) & " "
    CAN_Msg_1.Caption = CAN_Msg_1.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 4, 1))) & " "
    CAN_Msg_1.Caption = CAN_Msg_1.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 5, 1)))
    
    If LineCount = 1 Then
        CAN_Msg_2.Caption = "01 96 01 "
    Else
        CAN_Msg_2.Caption = "04 96 01 "
    End If
    
    CAN_Msg_2.Caption = CAN_Msg_2.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 6, 1))) & " "
    CAN_Msg_2.Caption = CAN_Msg_2.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 7, 1))) & " "
    CAN_Msg_2.Caption = CAN_Msg_2.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 8, 1))) & " "
    CAN_Msg_2.Caption = CAN_Msg_2.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 9, 1))) & " "
    CAN_Msg_2.Caption = CAN_Msg_2.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 10, 1)))
    
    If LineCount = 1 Then
        CAN_Msg_3.Caption = "00 96 01 "
    Else
        CAN_Msg_3.Caption = "03 96 01 "
    End If
    
    CAN_Msg_3.Caption = CAN_Msg_3.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 11, 1))) & " "
    CAN_Msg_3.Caption = CAN_Msg_3.Caption & Hex(Asc(Mid$(DisplayLine1.Text, 12, 1))) & " "
    CAN_Msg_3.Caption = CAN_Msg_3.Caption & "20 "
    CAN_Msg_3.Caption = CAN_Msg_3.Caption & "20 "
    CAN_Msg_3.Caption = CAN_Msg_3.Caption & "20"
    
    If LineCount = 1 Then
        CAN_Msg_4.Caption = ""
    Else
        CAN_Msg_4.Caption = "02 96 02 "
        CAN_Msg_4.Caption = CAN_Msg_4.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 1, 1))) & " "
        CAN_Msg_4.Caption = CAN_Msg_4.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 2, 1))) & " "
        CAN_Msg_4.Caption = CAN_Msg_4.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 3, 1))) & " "
        CAN_Msg_4.Caption = CAN_Msg_4.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 4, 1))) & " "
        CAN_Msg_4.Caption = CAN_Msg_4.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 5, 1)))
        
    End If

    If LineCount = 1 Then
        CAN_Msg_5.Caption = ""
    Else
        CAN_Msg_5.Caption = "01 96 02 "
        CAN_Msg_5.Caption = CAN_Msg_5.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 6, 1))) & " "
        CAN_Msg_5.Caption = CAN_Msg_5.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 7, 1))) & " "
        CAN_Msg_5.Caption = CAN_Msg_5.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 8, 1))) & " "
        CAN_Msg_5.Caption = CAN_Msg_5.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 9, 1))) & " "
        CAN_Msg_5.Caption = CAN_Msg_5.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 10, 1)))
        
    End If

    If LineCount = 1 Then
        CAN_Msg_6.Caption = ""
    Else
        CAN_Msg_6.Caption = "00 96 02 "
        CAN_Msg_6.Caption = CAN_Msg_6.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 11, 1))) & " "
        CAN_Msg_6.Caption = CAN_Msg_6.Caption & Hex(Asc(Mid$(DisplayLine2.Text, 12, 1))) & " "
        CAN_Msg_6.Caption = CAN_Msg_6.Caption & "20 "
        CAN_Msg_6.Caption = CAN_Msg_6.Caption & "20 "
        CAN_Msg_6.Caption = CAN_Msg_6.Caption & "20"
        
    End If
    
End Sub

Private Sub Exit_Button_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DisplayLine1.Text = ""
    DisplayLine2.Text = ""
    
    CAN_Msg_1.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_2.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_3.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_4.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_5.Caption = "XX XX XX XX XX XX XX XX"
    CAN_Msg_6.Caption = "XX XX XX XX XX XX XX XX"

End Sub
