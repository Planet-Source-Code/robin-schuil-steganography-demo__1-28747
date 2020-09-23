VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Steganography demo"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   5880
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "&Decode"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "&Encode"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Message"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   5655
      Begin VB.TextBox txtMessage 
         Height          =   975
         Left            =   1080
         MaxLength       =   160
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Steganography.frx":0000
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Text            =   "secret"
         Top             =   300
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Message:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #########################################################################################
' ## Steganography Demonstration                                                         ##
' ##                                                                                     ##
' ## This project is a demonstration of Steganography.                                   ##
' ## 'Steganography' is a name for hiding something in something else                    ##
' ## In this case we hide a message in a picture. The picture can be saved and no one    ##
' ## can see the message, unless they have this project and the correct password.        ##
' ##                                                                                     ##
' ## This version features a Steganography TAG at the beginning of each message          ##
' ## It was implemented to make sure the decoded message is a valid message.             ##
' ##                                                                                     ##
' ## For any questions or suggestions please contact beast@valleyalley.co.uk             ##
' #########################################################################################

Option Explicit

Private imgWidth As Long
Private imgHeight As Long
Private imgLoaded As Boolean
Private colPositions As Collection

Private Const Steganography_Tag = "TAG"

Private Sub EncodeByte(Value As Byte)

    On Error Resume Next

    Dim i As Integer
    Dim offset As Integer
    Dim bit As Byte
    
    Dim X As Long
    Dim Y As Long
    Dim C As Integer
    Dim strPosition As String
    
    Dim pixel As Long
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    
    offset = 1
    For i = 1 To 8
    
        ' Get random position and color channel
        Do
            X = Int(Rnd * imgWidth)
            Y = Int(Rnd * imgHeight)
            C = Int(Rnd * 100) Mod 3
            strPosition = "[" & C & "," & X & "," & Y & "]"
            colPositions.Add strPosition, strPosition
            If Err = 0 Then Exit Do
        Loop

        ' Get the Red-Green-Blue value of the pixel
        pixel = picImage.Point(X, Y)
        R = pixel And &HFF&
        G = (pixel And &HFF00&) \ &H100&
        B = (pixel And &HFF0000) \ &H10000
        
        ' Determine wether bit is 0 or 1
        If Value And offset Then bit = 1 Else bit = 0

        ' Add bit to the selected channel
        Select Case C
            Case 0
                R = (R And &HFE) Or bit
            Case 1
                G = (G And &HFE) Or bit
            Case 2
                B = (B And &HFE) Or bit
        End Select
        
        ' Update the pixel in the image
        picImage.PSet (X, Y), RGB(R, G, B)
        
        offset = offset * 2
    
    Next i

End Sub

Private Function DecodeByte() As Integer

    On Error Resume Next

    Dim Value As Integer
    Dim i As Integer
    Dim offset As Integer
    Dim bit As Byte
    
    Dim X As Long
    Dim Y As Long
    Dim C As Integer
    Dim strPosition As String
    
    Dim pixel As Long
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    
    offset = 1
    For i = 1 To 8
    
        ' Get random position and color channel
        Do
            X = Int(Rnd * imgWidth)
            Y = Int(Rnd * imgHeight)
            C = Int(Rnd * 100) Mod 3
            strPosition = "[" & C & "," & X & "," & Y & "]"
            colPositions.Add strPosition, strPosition
            If Err = 0 Then Exit Do
        Loop

        ' Get the Red-Green-Blue value of the pixel
        pixel = picImage.Point(X, Y)
        R = pixel And &HFF&
        G = (pixel And &HFF00&) \ &H100&
        B = (pixel And &HFF0000) \ &H10000

        ' Determine wether bit is 0 or 1
        Select Case C
            Case 0
                bit = (R And &H1)
            Case 1
                bit = (G And &H1)
            Case 2
                bit = (B And &H1)
        End Select

        ' Increase byte value
        If bit Then
            Value = Value Or offset
        End If

        offset = offset * 2
        
    Next i

    DecodeByte = Value
    
End Function

Private Function CalculateSeed(ByVal password As String) As Long

    ' Calculate a numeric seed based on the password
    ' You may use any method here, as long as the result is always
    ' the same for the same password.

    Dim Value As Long
    Dim ch As Long
    Dim shift1 As Long
    Dim shift2 As Long
    Dim i As Integer
    Dim str_len As Integer

    shift1 = 3
    shift2 = 17
    str_len = Len(password)
    
    For i = 1 To str_len
        ch = Asc(Mid$(password, i, 1))
        Value = Value Xor (ch * 2 ^ shift1)
        Value = Value Xor (ch * 2 ^ shift2)
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    
    CalculateSeed = Value

End Function

Private Sub cmdEncode_Click()

    Dim strMessage As String
    Dim i As Integer
    Dim message_length As Byte
    Dim seed As Long
    
    If imgLoaded = False Then
        MsgBox "Load an image first!"
        Exit Sub
    End If
        
    ' Initialize randomizer
    seed = CalculateSeed(CStr(txtPassword.Text))
    Rnd -1
    Randomize seed
    
    Set colPositions = New Collection
    
    ' Prepend the TAG to the message
    strMessage = Steganography_Tag & txtMessage.Text
    
    message_length = Len(strMessage)
    
    ' Store message length
    EncodeByte message_length
    
    ' Store message
    For i = 1 To message_length
        EncodeByte Asc(Mid(strMessage, i, 1))
    Next i

    While colPositions.Count
        colPositions.Remove 1
    Wend
    Set colPositions = Nothing

    ' Update the image
    picImage.Picture = picImage.Image

End Sub

Private Sub cmdDecode_Click()
    
    Dim strMessage As String
    Dim i As Integer
    Dim message_length As Integer
    Dim seed As Long
    
    If imgLoaded = False Then
        MsgBox "Load an image first!"
        Exit Sub
    End If
    
    ' Initialize randomizer
    seed = CalculateSeed(CStr(txtPassword.Text))
    Rnd -1
    Randomize seed
        
    Set colPositions = New Collection
        
    ' Read the message length
    message_length = DecodeByte
    
    For i = 1 To message_length
        strMessage = strMessage & Chr(DecodeByte)
    Next
    
    If Left(strMessage, 3) = Steganography_Tag Then
        txtMessage.Text = Mid(strMessage, 4)
    Else
        txtMessage.Text = ""
    End If
    
    While colPositions.Count
        colPositions.Remove 1
    Wend
    Set colPositions = Nothing
    
End Sub

Private Sub cmdLoad_Click()

    Dim sFilename As String
    
    ' Show dialog
    With CDialog
        .DialogTitle = "Open image"
        .Filter = "Windows Bitmap (*.bmp)|*.bmp|CompuServe Graphics Interchange (*.gif)|*.gif"
        .ShowOpen
        sFilename = .FileName
    End With
    
    ' Load image
    If sFilename <> "" Then
        picImage.Picture = LoadPicture(sFilename)
        imgWidth = picImage.Picture.Width
        imgHeight = picImage.Picture.Height
        If imgWidth > picImage.ScaleWidth Then imgWidth = picImage.ScaleWidth
        If imgHeight > picImage.ScaleHeight Then imgHeight = picImage.ScaleHeight
        imgLoaded = True
    End If
    
End Sub

Private Sub cmdQuit_Click()

    Dim reply As VbMsgBoxResult
    
    ' Confirm
    reply = MsgBox("Are you sure you want to quit?", vbYesNo Or vbQuestion, "Steganography")
    If reply = vbYes Then End

End Sub

Private Sub cmdSave_Click()

    Dim sFilename As String
    
    ' Show dialog
    With CDialog
        .DialogTitle = "Save image"
        .Filter = "Windows Bitmap (*.bmp)|*.bmp"
        .DefaultExt = ".bmp"
        .ShowSave
        sFilename = .FileName
    End With
    
    ' Save image
    If sFilename <> "" Then
        SavePicture picImage.Picture, sFilename
    End If

End Sub
