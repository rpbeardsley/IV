VERSION 5.00
Begin VB.Form Conversion 
   Caption         =   "File conversion to strip header"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3960
      Width           =   6015
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   3360
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2895
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
      Left            =   2040
      TabIndex        =   0
      Top             =   4680
      Width           =   2175
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
     
    'Go throught the list of files
    firstfile = True
    For a = 0 To File1.ListCount - 1
        If File1.Selected(a) = True Then
            If firstfile = True Then
                pth = File1.Path
                If Right(pth, 1) <> "\" Then pth = pth + "\"
                pth = pth + File1.List(a)
                fname = File1.List(a)
                ReadFile pth, fname 'Read and strip the header
                'firstfile = False
            End If
        End If

        Close #1
        
    Next a
    
End Sub


Private Sub ReadFile(FPath, Name)

    Dim data As String
    Dim cx As Double
    Dim cy As Double
    
    Open FPath For Input As #1
    
    cnt = 0
    While Not EOF(1)
        
        Line Input #1, data
        
        If cnt < 20 Then
            cx = Val(Parse(data, 1))
            cy = Val(Parse(data, 2))
        Else
            Input #1, cx
            Input #1, cy
        End If
        
        
        ReDim Preserve X(0 To cnt)
        ReDim Preserve Y(0 To cnt)
        
        X(cnt) = cx
        Y(cnt) = cy
            
        cnt = cnt + 1
                  
    Wend

    Close #1
    
    'Output the result
    Outfname = Text1.Text + "\" + Replace$(Name, ".txt", ".dat")   'change the file type
    Open Outfname For Output As #1
    For v = LBound(X) To UBound(X)
        Print #1, Format(X(v), "00.00000") + "    " + Format(Y(v), "00.000000000")
    Next v
    Close #1
    
End Sub

Public Function Parse(DatStr As String, Counter As Integer) As String

    On Error GoTo PErr
    
    Xpos = InStr(DatStr, Chr(46)) - 2
    Ypos = InStrRev(DatStr, Chr(46)) - 2
    
    If Counter = 1 Then
        Parse = Mid(DatStr, Xpos, 8)
    ElseIf Counter = 2 Then
        Parse = Mid(DatStr, Ypos)
   End If
    
    Exit Function
PErr:
    
    Parse = "Search string does not exist"
    
End Function

Private Sub Dir1_Change()

On Error Resume Next

File1.Path = Dir1.Path
Text1.Text = Dir1.Path

End Sub

Private Sub Drive1_Change()

On Error Resume Next

Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()

Text1.Text = Dir1.Path

End Sub

Private Sub Text1_LostFocus()

    On Error Resume Next
    
    Drive1.Drive = Text1.Text
    Dir1.Path = Text1.Text

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    On Error Resume Next
    
    Drive1.Drive = Text1.Text
    Dir1.Path = Text1.Text
End If

End Sub
