VERSION 5.00
Begin VB.Form Decompiler1 
   AutoRedraw      =   -1  'True
   Caption         =   "Decompiler1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "C:\file.exe"
      Top             =   0
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   3735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Decompiler1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program should not be able to ruin anything or change anything on
'your computer. If it does some how, I will not take any resposiblity in it.
'Computer Controller
Dim StringThing As String
Sub FormFinder(FileName As String, Lst)
Dim Countt As Long
     Countt = 1
Again:

     FChr$ = Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(5) + Chr(32)
     FChr1$ = Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(7) + Chr(32)
     FChr2$ = Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(6) + Chr(32)
     FChr3$ = Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(8) + Chr(32)
        
        If InStr(Countt, StringThing, FChr$, 1) Then
            A = InStr(Countt, StringThing, FChr$, 1) + 6
            If InStr(A, StringThing, Chr(1)) Then
            b = InStr(A, StringThing, Chr(1)) - 2
            For d = A To b
            For e = 0 To 31
            If Mid(StringThing, d, 1) = Chr(e) Then
                Countt = b
                GoTo Again
            End If
            Next
            For e = 127 To 255
            If Mid(StringThing, d, 1) = Chr(e) Then
                Countt = b
                GoTo Again
            End If
            Next
            Next
                If A > b Then
                Countt = A
                GoTo Again
                End If
                If Len(Mid(StringThing, A, b - A)) > 127 Then
                Countt = b
                GoTo Again
                End If
                Lst = Lst & vbCrLf & "Name - " & Mid(StringThing, A, b - A)
                If InStr(b + 1, StringThing, Chr(25) + Chr(1)) Then
                    C = InStr(b + 1, StringThing, Chr(25) + Chr(1))
                    Lst = Lst & "   Caption - " & Mid(StringThing, b + 5, C - b - 6)
                    Countt = C
                Else
                    Countt = b
                End If
            Else
            Countt = A
            End If
            DoEvents
            GoTo Again
        ElseIf InStr(Countt, StringThing, FChr1$, 1) Then
            A = InStr(Countt, StringThing, FChr1$, 1) + 6
            If InStr(A, StringThing, Chr(1)) Then
                b = InStr(A, StringThing, Chr(1)) - 2
            For d = A To b
            For e = 0 To 31
            If Mid(StringThing, d, 1) = Chr(e) Then
                Countt = b
                GoTo Again
            End If
            Next
            For e = 127 To 255
            If Mid(StringThing, d, 1) = Chr(e) Then
                Countt = b
                GoTo Again
            End If
            Next
            Next
                If A > b Then
                Countt = A
                GoTo Again
                End If
                If Len(Mid(StringThing, A, b - A)) > 127 Then
                Countt = b
                GoTo Again
                End If
                Lst = Lst & vbCrLf & "Name - " & Mid(StringThing, A, b - A)
                If InStr(b + 1, StringThing, Chr(25) + Chr(1)) Then
                    C = InStr(b + 1, StringThing, Chr(25) + Chr(1))
                    Lst = Lst & "   Caption - " & Mid(StringThing, b + 5, C - b - 6)
                    Countt = C
                Else
                    Countt = b
                End If
            Else
            Countt = A
            End If
            DoEvents
            GoTo Again
        ElseIf InStr(Countt, StringThing, FChr2$, 1) Then
            A = InStr(Countt, StringThing, FChr2$, 1) + 6
            If InStr(A, StringThing, Chr(1)) Then
                b = InStr(A, StringThing, Chr(1)) - 2
            For d = A To b
            For e = 0 To 31
            If Mid(StringThing, d, 1) = Chr(e) Then
                Countt = b
                GoTo Again
            End If
            Next
            For e = 127 To 255
            If Mid(StringThing, d, 1) = Chr(e) Then
                Countt = b
                GoTo Again
            End If
            Next
            Next
                If A > b Then
                Countt = A
                GoTo Again
                End If
                If Len(Mid(StringThing, A, b - A)) > 127 Then
                Countt = b
                GoTo Again
                End If
                Lst = Lst & vbCrLf & "Name - " & Mid(StringThing, A, b - A)
                If InStr(b + 1, StringThing, Chr(25) + Chr(1)) Then
                    C = InStr(b + 1, StringThing, Chr(25) + Chr(1))
                    Lst = Lst & "   Caption - " & Mid(StringThing, b + 5, C - b - 6)
                    Countt = C
                Else
                    Countt = b
                End If
            Else
            Countt = A
            End If
            DoEvents
            GoTo Again
        ElseIf InStr(Countt, StringThing, FChr3$, 1) Then
            A = InStr(Countt, StringThing, FChr3$, 1) + 6
            If InStr(A, StringThing, Chr(1)) Then
            b = InStr(A, StringThing, Chr(1)) - 2
            For d = A To b
            For e = 0 To 31
            If Mid(StringThing, d, 1) = Chr(e) Then
                Countt = b
                GoTo Again
            End If
            Next
            For e = 127 To 255
            If Mid(StringThing, d, 1) = Chr(e) Then
                Countt = b
                GoTo Again
            End If
            Next
            Next
                If A > b Then
                Countt = A
                GoTo Again
                End If
                If Len(Mid(StringThing, A, b - A)) > 127 Then
                Countt = b
                GoTo Again
                End If
                Lst = Lst & vbCrLf & "Name - " & Mid(StringThing, A, b - A)
                If InStr(b + 1, StringThing, Chr(25) + Chr(1)) Then
                    C = InStr(b + 1, StringThing, Chr(25) + Chr(1))
                    Lst = Lst & "   Caption - " & Mid(StringThing, b + 5, C - b - 6)
                    Countt = C
                Else
                    Countt = b
                End If
            Else
                Countt = A
            End If
            DoEvents
            GoTo Again
        Else
        End If
Caption = "Done"
End Sub
Sub ProjProp(FileName As String, Lst)
Caption = "Please Wait done with Loading"
FChr$ = " F i l e V e r s i o n"
    If InStr(1, StringThing, FChr$, 1) Then
        b = InStr(1, StringThing, FChr$, 1) + Len(FChr$) + 5
        C = InStr(b, StringThing, Chr(32) + Chr(32) + Chr(32), 1) - b
        Lst = Lst & vbCrLf & "File Version - " & RemoveSpaces(Mid(StringThing, b, C))
    Else
        Lst = Lst & vbCrLf & "File Version - None"
    End If
FChr$ = " F i l e D e s c r i p t i o n"
    If InStr(1, StringThing, FChr$, 1) Then
        b = InStr(1, StringThing, FChr$, 1) + Len(FChr$) + 5
        C = InStr(b, StringThing, Chr(32) + Chr(32) + Chr(32), 1) - b
        Lst = Lst & vbCrLf & "File Description - " & RemoveSpaces(Mid(StringThing, b, C))
    Else
        Lst = Lst & vbCrLf & "File Description - None"
    End If
FChr$ = " L e g a l C o p y r i g h t"
    If InStr(1, StringThing, FChr$, 1) Then
        b = InStr(1, StringThing, FChr$, 1) + Len(FChr$) + 3
        C = InStr(b, StringThing, Chr(1), 1) - b - 5
        Lst = Lst & vbCrLf & "Copyright - " & RemoveSpaces(Mid(StringThing, b, C))
    Else
        Lst = Lst & vbCrLf & "Copyright - None"
    End If
FChr$ = " C o m m e n t s"
    If InStr(1, StringThing, FChr$, 1) Then
        b = InStr(1, StringThing, FChr$, 1) + Len(FChr$) + 3
        C = InStr(b, StringThing, Chr(1), 1) - b - 5
        Lst = Lst & vbCrLf & "Comments - " & RemoveSpaces(Mid(StringThing, b, C))
    Else
        Lst = Lst & vbCrLf & "Comments - None"
    End If
FChr$ = " C o m p a n y N a m e"
    If InStr(1, StringThing, FChr$, 1) Then
        b = InStr(1, StringThing, FChr$, 1) + Len(FChr$) + 3
        C = InStr(b, StringThing, Chr(1), 1) - b - 5
        Lst = Lst & vbCrLf & "Company Name - " & RemoveSpaces(Mid(StringThing, b, C))
    Else
        Lst = Lst & vbCrLf & "Company Name - None"
    End If
FChr$ = " I n t e r n a l N a m e"
    If InStr(1, StringThing, FChr$, 1) Then
        b = InStr(1, StringThing, FChr$, 1) + Len(FChr$) + 3
        C = InStr(b, StringThing, Chr(1), 1) - b - 5
        Lst = Lst & vbCrLf & "Internal Name - " & RemoveSpaces(Mid(StringThing, b, C))
    Else
        Lst = Lst & vbCrLf & "Internal Name - None"
    End If
FChr$ = " O r i g i n a l F i l e n a m e"
    If InStr(1, StringThing, FChr$, 1) Then
        b = InStr(1, StringThing, FChr$, 1) + Len(FChr$) + 3
        C = InStr(b, StringThing, Chr(1), 1) - b - 5
        Lst = Lst & vbCrLf & "Original File Name - " & RemoveSpaces(Mid(StringThing, b, C))
    Else
        Lst = Lst & vbCrLf & "Original File Name - None"
    End If
FChr$ = " P r o d u c t N a m e"
    If InStr(1, StringThing, FChr$, 1) Then
        b = InStr(1, StringThing, FChr$, 1) + Len(FChr$) + 3
        C = InStr(b, StringThing, Chr(1), 1) - b - 5
        Lst = Lst & vbCrLf & "Product Name - " & RemoveSpaces(Mid(StringThing, b, C))
    Else
        Lst = Lst & vbCrLf & "Product Name - None"
    End If
FChr$ = " P r o d u c t V e r s i o n"
    If InStr(1, StringThing, FChr$, 1) Then
        b = InStr(1, StringThing, FChr$, 1) + Len(FChr$) + 3
        C = InStr(b, StringThing, Chr(1), 1) - b - 5
        Lst = Lst & vbCrLf & "Product Version - " & RemoveSpaces(Mid(StringThing, b, C))
    Else
        Lst = Lst & vbCrLf & "Product Version - None"
    End If
End Sub
Function RemoveSpaces(Str As String) As String
For A = 1 To Len(Str) Step 2
b = b & Mid(Str, A, 1)
Next
RemoveSpaces = b
End Function
Sub Loadexe(FileName$)
Caption = "Please wait"
Open FileName$ For Binary Access Read As #1 'Because this is an exe, you need to open it in binary
    For d = 1 To LOF(1) Step 1000 'Loop until end of file
    If d Mod 5 = 0 Then
    Caption = "Please Wait Percent done - " & Int(d / LOF(1) * 100) & "%"
    End If
        StringThing = String(1000, " ")
        Get #1, d, StringThing 'Get char. from specified position
        For b = 1 To 1000
            If Mid(StringThing, b, 1) = Chr(0) Then 'If the char. = nothen then
                Mid(StringThing, b, 1) = Chr(32) 'Change it to a space
            End If 'Question done
        Next
        A = A & StringThing 'record the char.
        DoEvents 'Don't freeze
    Next 'go back to For
Close 'Free mem
StringThing = A 'Save it for later
End Sub
Private Sub Command1_Click()
StringThing = ""
Loadexe Text2.Text
ProjProp Text2.Text, A
FormFinder Text2.Text, b
Text1.Text = A & b
End Sub

Private Sub Form_Resize()
Text1.Width = Me.Width - 90
Text1.Height = Me.Height - 750
Command1.Left = Me.Width - Command1.Width - 120
Text2.Width = Command1.Left - 10
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click
End Sub
