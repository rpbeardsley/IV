VERSION 5.00
Begin VB.Form Conversion 
   Caption         =   "File conversion to ScanXP type .dat"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Text            =   "150"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3120
      TabIndex        =   6
      Text            =   "0.0856"
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "-0.0788"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
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
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "Scan length (in ps) in the linear region given above"
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
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "ms"
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
      Left            =   4080
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "and"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Take data between"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the full file path of the file to be converted :"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   4335
   End
End
Attribute VB_Name = "Conversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Y() As Double
Dim X() As Double

Private Sub Command1_Click()
    
    'Read in the data
    ReadFile (Text1.Text)
    
    
    'Convert the ms X axis from the scope into delay stage mm (2x real mm) for the FT
    'software to understand
    
    'Translation constant for X axis zeroing
    TranCon = Abs(X(0))
    
    For b = LBound(X) To UBound(X)
        'zero X axis
        X(b) = X(b) + TranCon
    Next
    
    'Calculate the correction factor to convert the full sweep to the stipulated
    'scan time and give it in delay stage mm
    CorFac = (Val(Text4.Text) / X(UBound(X))) / 6.6666666
    
    'Create the file to write the new data to
    Outfname = Text1.Text & "_CONVERTED.dat"
    Open Outfname For Output As #1
    
    'Write the new data
    For a = LBound(X) To UBound(X)
        
        'Multiply each X array element by the correction factor to give mm
        X(a) = X(a) * CorFac
        
        Print #1, "Nanostepper position:" & vbCrLf & X(a)
        
        i = 0
        Do Until i = 10
            Print #1, Format(Y(a), "0.00000E-00") + ", " + Format(Y(a), "0.00000E-00")
            i = i + 1
        Loop
        
        Print #1, vbCrLf
        
    Next a
    
    Close #1

End Sub


Private Sub ReadFile(fname)
    
    'Open the read file
    Open fname For Input As #1
    
    'Read in the data
    cnt = 0
    While Not EOF(1)
        Input #1, cx
        Input #1, cy
        
        'Only take data in the stipulated linear region of the trace
        If Val(Text3.Text) > cx And cx > Val(Text2.Text) Then
        
            ReDim Preserve X(0 To cnt)
            ReDim Preserve Y(0 To cnt)
        
            X(cnt) = cx
            Y(cnt) = cy
            
            cnt = cnt + 1
            
        End If
                
    Wend

    Close #1
        
End Sub

