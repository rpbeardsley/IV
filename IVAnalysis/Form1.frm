VERSION 5.00
Begin VB.Form IVTraceAverage 
   Caption         =   "Average IV Traces"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Save Lockin X and Y"
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   5400
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Display"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   5520
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7815
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   3480
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "IVTraceAverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Yr() As Double
Dim Yi() As Double
Dim N() As Long
Dim X() As Double

Dim meanX() As Double
Dim meanYr() As Double
Dim meanYi() As Double

Private Sub Command1_Click()

    'Clear arrays
    ReDim meanX(0 To 0)
    ReDim meanYi(0 To 0)
    ReDim meanYr(0 To 0)
    ReDim Yr(0 To 0)
    ReDim Yi(0 To 0)
    ReDim X(0 To 0)
    ReDim N(0 To 0)

    firstfile = True
    For a = 0 To File1.ListCount - 1
        If File1.Selected(a) = True Then
        
            If firstfile = True Then
                pth = File1.Path
                If Right(pth, 1) <> "\" Then pth = pth + "\"
                pth = pth + File1.List(a)
                LoadFirstFile pth
                firstfile = False
            Else
                pth = File1.Path
                If Right(pth, 1) <> "\" Then pth = pth + "\"
                pth = pth + File1.List(a)
                LoadFile pth, X
            End If
        
        End If
    Next a
    
    'Now the files have been summed into the arrays we
    'just need to divide the arrays by the number of files
    'to get the mean
    cnt = 0
    For a = LBound(X) To UBound(X)
        If N(a) <> 0 Then
            ReDim Preserve meanX(0 To cnt)
            ReDim Preserve meanYi(0 To cnt)
            ReDim Preserve meanYr(0 To cnt)
            meanX(a) = X(a)
            meanYi(a) = Yi(a) / N(a)
            meanYr(a) = Yr(a) / N(a)
            cnt = cnt + 1
        End If
    Next a
    
    MsgBox "Done!", vbOKOnly Or vbInformation, "Process Complete"

End Sub

Private Sub Command2_Click()

    Open Text2.Text For Output As #1

    For a = LBound(meanX) To UBound(meanX)
        If Check1.Value = 1 Then
            Print #1, Format(meanX(a), "0.00000E-00") + ", " + Format(meanYr(a), "0.00000E-00") + ", " + Format(meanYi(a), "0.00000E-00")
        Else
            Print #1, Format(meanX(a), "0.00000E-00") + ", " + Format(meanYr(a), "0.00000E-00")
        End If
    Next a

    Close #1

End Sub

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

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    On Error Resume Next
    
    Drive1.Drive = Text1.Text
    Dir1.Path = Text1.Text
End If

End Sub

Private Sub Text1_LostFocus()

    On Error Resume Next
    
    Drive1.Drive = Text1.Text
    Dir1.Path = Text1.Text

End Sub

Private Sub LoadFirstFile(fname)

    Open fname For Input As #1
    
    cnt = 0
    While Not EOF(1)
    
        'Read points
        Input #1, cx
        Input #1, cyr
        Input #1, cyi
        
        ReDim Preserve X(0 To cnt)
        ReDim Preserve Yi(0 To cnt)
        ReDim Preserve Yr(0 To cnt)
        ReDim Preserve N(0 To cnt)
        
        X(cnt) = cx
        Yi(cnt) = cyi
        Yr(cnt) = cyr
        N(cnt) = 1
        
        cnt = cnt + 1
        
    Wend

    Close #1

End Sub

Private Sub LoadFile(fname, xaxis)

'Loads file and interpolates its contents onto the specified
'x-axis

    Open fname For Input As #1
    
    firstpoint = True
    While Not EOF(1)
    
        'Read points
        Input #1, cx
        Input #1, cyr
        Input #1, cyi
        
        If firstpoint = True Then
            prevx = cx
            prevyr = cyr
            prevyi = cyi
            Input #1, cx
            Input #1, cyr
            Input #1, cyi
            firstpoint = False
        End If
        
        dyr = cyr - prevyr
        dyi = cyi - prevyi
        dx = cx - prevx
        
        For b = LBound(xaxis) To UBound(xaxis)
            If (prevx <= xaxis(b) And cx > xaxis(b)) Or (prevx >= xaxis(b) And cx < xaxis(b)) Then
                Yr(b) = Yr(b) + prevyr + (xaxis(b) - prevx) * dyr / dx
                Yi(b) = Yi(b) + prevyi + (xaxis(b) - prevx) * dyr / dx
                N(b) = N(b) + 1
                GoTo ext
            End If
        Next b
ext:

        prevyr = cyr
        prevyi = cyi
        prevx = cx
    
    Wend
    
    Close #1

End Sub
