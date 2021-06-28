VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Earthquake!"
   ClientHeight    =   3930
   ClientLeft      =   1065
   ClientTop       =   3465
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   10290
   Begin VB.CommandButton Command1 
      Caption         =   "Download Latest List"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   3120
      Width           =   4455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"earthquake.frx":0000
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ms - Surface Wave"
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mb - Body Wave"
      Height          =   255
      Left            =   7080
      TabIndex        =   15
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M1 - Richter Scale"
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Magnitude Key"
      Height          =   255
      Left            =   7080
      TabIndex        =   13
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "  DATE    UTC TIME    LAT.  LONG.  DEPTH  MAG. Q  LOCATION"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   10095
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnDownload 
         Caption         =   "Download Latest List"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FoundPos As Integer
Dim FoundPos2 As Integer
Dim FoundLine As Integer
Dim st As String
Dim st2 As String
Dim today As String
Dim s As String
Private Sub Command1_Click()
Call loadup
End Sub
Private Sub Form_Load()
' center the form
Form1.Top = (Screen.Height - Form1.Height) / 2
Form1.Left = (Screen.Width - Form1.Width) / 2
End Sub
Private Sub loadup()
Me.Caption = "Earthquake!   Loading..."
RichTextBox1.Text = Inet1.OpenURL("http://wwwneic.cr.usgs.gov/neis/bulletin/bulletin.html")
Me.Caption = "Earthquake!   Done!"
' Use Today's date to set up the search string
today = ">" & Right$(Date$, 2) & "/" & Left$(Date$, 2) & "/" & Mid$(Date$, 4, 2)
    
    For X = 0 To 9
    ' Find the text specified in the TextBox control.
    FoundPos = RichTextBox1.Find(today, FoundPos, , rtfWholeWord)

    If FoundPos <> -1 Then
        Me.Caption = "Earthquake! "
        ' Returns number of line containing found text.
        FoundLine = RichTextBox1.GetLineFromChar(FoundPos)
        
        ' Select 100 characters from the position where the date was found ...
        RichTextBox1.SelStart = FoundPos
        RichTextBox1.SelLength = 100
        
        ' Now, using the selected data, find the endpoint of the data ...
        FoundPos2 = RichTextBox1.Find("<", , , SelRTF)
        RichTextBox1.SelStart = FoundPos
        RichTextBox1.SelLength = (FoundPos2 - FoundPos)
        
        ' Truncate and clean-up the string ...
        st = Right$(RichTextBox1.SelRTF, FoundPos2 - FoundPos + 2)
        Label2(X).Caption = "  " & Left$(st, Len(st) - 3)
    Else
           If X = 0 Then
           Me.Caption = "Earthquake!   No Entries Found For Today."
           End If
        Exit Sub
    End If
        ' Move the pointer ahead for the next "find" ...
        FoundPos = FoundPos2
    Next X
    
End Sub
Private Sub Label1_Click()
Unload Me
End Sub
Private Sub mnDownload_Click()
Call loadup
End Sub

