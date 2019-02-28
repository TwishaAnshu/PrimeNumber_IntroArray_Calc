VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17916
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   17916
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdORIMELIST 
      Caption         =   "ACTIVATE LIST"
      Height          =   372
      Left            =   14220
      TabIndex        =   33
      Top             =   4020
      Width           =   732
   End
   Begin VB.ListBox lstPNum 
      Height          =   816
      Left            =   15960
      TabIndex        =   32
      Top             =   4380
      Width           =   1272
   End
   Begin VB.ListBox lstPrime 
      Height          =   3696
      Left            =   12060
      TabIndex        =   29
      Top             =   1980
      Width           =   1752
   End
   Begin VB.ListBox lst2 
      Height          =   1392
      Left            =   3780
      TabIndex        =   23
      Top             =   4440
      Width           =   2712
   End
   Begin VB.TextBox txtDivisorsList2 
      Height          =   1032
      Left            =   8100
      TabIndex        =   20
      Text            =   "txtDivisorsList2"
      Top             =   5040
      Width           =   2712
   End
   Begin VB.ListBox lst1 
      Height          =   1584
      Left            =   4080
      TabIndex        =   19
      Top             =   1740
      Width           =   2412
   End
   Begin VB.TextBox txtDivisorsList 
      Height          =   1152
      Left            =   7980
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "Form1.frx":0000
      Top             =   1560
      Width           =   3552
   End
   Begin VB.TextBox txtNumber2 
      Height          =   552
      Left            =   960
      TabIndex        =   15
      Text            =   "txtNumber2"
      Top             =   6240
      Width           =   1032
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   492
      Left            =   11700
      TabIndex        =   13
      Top             =   6120
      Width           =   1032
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
      Height          =   432
      Left            =   11580
      TabIndex        =   12
      Top             =   6960
      Width           =   1212
   End
   Begin VB.TextBox txtNumber1 
      Height          =   552
      Left            =   1020
      TabIndex        =   0
      Text            =   "txtNumber1"
      Top             =   2040
      Width           =   1692
   End
   Begin VB.Label lblPerfectNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   972
      Left            =   16080
      TabIndex        =   31
      Top             =   3060
      Width           =   1392
   End
   Begin VB.Label Label111 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PERFECT NUMBER"
      Height          =   432
      Left            =   14220
      TabIndex        =   30
      Top             =   3000
      Width           =   1452
   End
   Begin VB.Label lblLCM 
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   14820
      TabIndex        =   28
      Top             =   2220
      Width           =   1032
   End
   Begin VB.Label lblSumOfDivisors2 
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   9420
      TabIndex        =   27
      Top             =   7080
      Width           =   972
   End
   Begin VB.Label lblNumberOfDivisors2 
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   9360
      TabIndex        =   26
      Top             =   6420
      Width           =   852
   End
   Begin VB.Label LABEL700 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUM "
      Height          =   432
      Left            =   8040
      TabIndex        =   25
      Top             =   6960
      Width           =   852
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# DIVISORS :"
      Height          =   372
      Left            =   7980
      TabIndex        =   24
      Top             =   6420
      Width           =   972
   End
   Begin VB.Label lblNotPrime2 
      Caption         =   "IS NOT PRIME"
      Height          =   312
      Left            =   2100
      TabIndex        =   22
      Top             =   7080
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label lblPrime2 
      Caption         =   "IT IS PRIME"
      Height          =   192
      Left            =   720
      TabIndex        =   21
      Top             =   6960
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label lblNotPrime 
      Caption         =   "IT IS NOT PRIME"
      Height          =   492
      Left            =   1080
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Label lblPrime 
      Caption         =   "IT IS A PRIME NUMBER!"
      Height          =   492
      Left            =   1200
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number2"
      Height          =   612
      Left            =   840
      TabIndex        =   14
      Top             =   4380
      Width           =   1212
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LCM"
      Height          =   312
      Left            =   14100
      TabIndex        =   11
      Top             =   2280
      Width           =   552
   End
   Begin VB.Label lblGCF 
      BorderStyle     =   1  'Fixed Single
      Height          =   372
      Left            =   14880
      TabIndex        =   10
      Top             =   1380
      Width           =   1872
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GCF"
      Height          =   372
      Left            =   13980
      TabIndex        =   9
      ToolTipText     =   "TO FIND THE GCF,FIRST ENTER THE SECOND NUMBER AND THEN THEN 1ST"
      Top             =   1440
      Width           =   612
   End
   Begin VB.Label lblSum 
      BorderStyle     =   1  'Fixed Single
      Height          =   612
      Left            =   10020
      TabIndex        =   8
      Top             =   3900
      Width           =   1212
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sum"
      Height          =   432
      Left            =   8220
      TabIndex        =   7
      Top             =   3660
      Width           =   1092
   End
   Begin VB.Label lblNumberOfDivisors 
      BorderStyle     =   1  'Fixed Single
      Height          =   492
      Left            =   9900
      TabIndex        =   6
      Top             =   3120
      Width           =   792
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# of Divisors"
      Height          =   372
      Left            =   8100
      TabIndex        =   5
      Top             =   3000
      Width           =   1272
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Divisors"
      Height          =   552
      Left            =   7920
      TabIndex        =   4
      Top             =   780
      Width           =   1692
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prime"
      Height          =   432
      Left            =   12060
      TabIndex        =   3
      Top             =   900
      Width           =   1392
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Divisors"
      Height          =   672
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   2172
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number 1"
      Height          =   612
      Left            =   1020
      TabIndex        =   1
      Top             =   720
      Width           =   1512
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Number, Sum, Divisor As Long
Dim DivisorList As String
Dim Number2, Sum2, Divisor2 As Long  ' declaring variable
Dim DivisorList2 As String
Dim GCF As Long
Dim I As Double
Dim CountPrime As Single
Dim LCM As Double


Private Sub cmdClear_Click()
txtNumber1.Text = ""
txtNumber2.Text = ""
txtDivisorsList = ""
lblPrime.Visible = False
lblNotPrime.Visible = False
lst1.Clear
lst2.Clear
txtDivisorsList = ""
lblNumberOfDivisors = ""
lblSumOfDivisors2 = ""
lblGCF = ""
lblLCM = ""

txtDivisorsList2 = ""
lblNumberOfDivisors2 = ""
lblSum = ""
lstPrime.Clear

End Sub

Private Sub cmdORIMELIST_Click()
If Number > Number2 Then



For CountPrime = 1 To Number
        CountPrime = 0
            For Divisor = 1 To Number
                If Number Mod Divisor = 0 Then
                CountPrime = CountPrime + 1
                End If
            Next Divisor
    If CountPrime = 2 Then
        lstPrime.AddItem (CountPrime)
    End If

    Next CountPrime
    lstPrime.AddItem (CountPrime)
 
ElseIf Number < Number2 Then
 
    For CountPrime = 1 To Number2
        CountPrime = 0
            For Divisor = 1 To Number2
                If Number2 Mod Divisor = 0 Then
                CountPrime = CountPrime + 1
                End If
            Next Divisor
    If CountPrime = 2 Then
        lstPrime.AddItem (CountPrime)
    End If

    Next CountPrime
    lstPrime.AddItem (CountPrime)
    
End If

End Sub

Private Sub cmdQuit_Click()
End ' to end program


End Sub




Private Sub txtNumber1_KeyPress(KeyAscii As Integer)
Dim Number, Sum, Divisor As Long  ' declaring variable
Dim DivisorList As String
Dim LCM As Long


If KeyAscii = 13 Then

    Number = Val(txtNumber1) ' read number
    Sum = 0 ' initializes sum of divisors
    DivisorList = ""
    For Divisor = 1 To Number ' loop of divisors
        If Number Mod Divisor = 0 Then ' test
        Sum = Sum + Divisor ' running total
        DivisorList = DivisorList & Str$(Divisor) + " "
        lst1.AddItem (Divisor)
        End If
        
        
    Next Divisor
    txtDivisorsList.Text = DivisorList 'display

    If Sum = Number + 1 Then ' test if prime
        lblPrime.Visible = True
        Else
        lblNotPrime.Visible = True
    End If
    
'GCFGCFGCFGCFGCFGCF

Number2 = Val(txtNumber2.Text)

    If Number <> 0 And Number2 <> 0 Then
For Divisor = 1 To Number
  If (Number Mod Divisor) = 0 And (Number2 Mod Divisor) = 0 Then
   GCF = Divisor
   
End If

Next Divisor
lblGCF.Caption = GCF
lblLCM.Caption = (Number * Number2) / GCF
End If

cmdClear.SetFocus ' shift focus to clear
End If
'''LCM = Val(lblLCM)
'''''LCM = (Number * Number2) / GCF

lblSum.Caption = Sum



      Dim tmpNum As Integer
        Const onespace As String = " "
        Const twospace As String = "  "
        Dim tmpWords() As String
        Dim tmpText As String
        
        tmpNum = 0
        tmpText = txtDivisorsList.Text
        Do Until tmpNum = 0
            tmpNum = InStr(tmpText, twospace)
            If tmpNum > 0 Then
                tmpText = Replace(tmpText, twospace, onespace)
            End If
        Loop
              
        tmpWords = Split(tmpText, onespace)
       lblNumberOfDivisors.Caption = UBound(tmpWords) / 2
       
       
       
        

If lblNotPrime.Visible = True Then
lblPerfectNumber.Caption = Number * 2

   ' perfec number = half the sum of all its positive dicisores
   
   lblPerfectNumber.Caption = (1 / 2) * Sum
   
   
   
       Number = Val(txtNumber1) ' read number
    Sum = 0 ' initializes sum of divisors
    DivisorList = ""
    For Divisor = 1 To Number ' loop of divisors
        If Number Mod Divisor = 0 Then ' test
        Sum = Sum + Divisor ' running total
        DivisorList = DivisorList & Str$(Divisor) + " "
        lst1.AddItem (Divisor)
        End If
        
        
    Next Divisor
    txtDivisorsList.Text = DivisorList 'display
   

End If



End Sub




Private Sub txtNumber2_KeyPress(KeyAscii As Integer)
Dim Number2, Sum2, Divisor2 As Long  ' declaring variable
Dim DivisorList2 As String
lblGCF = GCF

If KeyAscii = 13 Then
    Number2 = Val(txtNumber2) ' read number
    Sum2 = 0 ' initializes sum of divisors
    DivisorList2 = ""
    For Divisor2 = 1 To Number2 ' loop of divisors
        If Number2 Mod Divisor2 = 0 Then ' test
        Sum2 = Sum2 + Divisor2 ' running total
        DivisorList2 = DivisorList2 & Str$(Divisor2) + " "
        lst2.AddItem (Divisor2)
        End If
        
        
    Next Divisor2
    txtDivisorsList2.Text = DivisorList2 'display
 
    If Sum2 = Number2 + 1 Then ' test if prime
        lblPrime2.Visible = True
        Else
        lblNotPrime2.Visible = True
    End If
 

cmdClear.SetFocus ' shift focus to clear




End If
 lblSumOfDivisors2.Caption = Sum2
 
 
 Dim tmpNum As Integer
        Const onespace As String = " "
        Const twospace As String = "  "
        Dim tmpWords() As String
        Dim tmpText As String
        
        tmpNum = 0
        tmpText = txtDivisorsList2.Text
        Do Until tmpNum = 0
            tmpNum = InStr(tmpText, twospace)
            If tmpNum > 0 Then
                tmpText = Replace(tmpText, twospace, onespace)
            End If
        Loop
              
        tmpWords = Split(tmpText, onespace)
       lblNumberOfDivisors2.Caption = UBound(tmpWords) / 2



 





End Sub
'' wuwrheiojfe
