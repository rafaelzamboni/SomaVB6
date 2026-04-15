VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Soma"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSecondNumber 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton btnCalculate 
      Caption         =   "Calcular"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtFirstNumber 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   5640
      Picture         =   "Soma.frx":0000
      Top             =   3600
      Width           =   660
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Soma de dois números"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCalculate_Click()
    ' Verifica se os campos estão vazios
    If Trim(txtFirstNumber.Text) = "" Or Trim(txtSecondNumber.Text) = "" Then
        MsgBox "Por favor, preencha os dois campos para somar!", vbExclamation, "Aviso"
        Exit Sub
    End If

    Dim number1 As Single
    Dim number2 As Single
    Dim sum As Single
    
    number1 = 0
    number2 = 0

    ' Valida antes de converter para evitar erros
    If IsNumeric(txtFirstNumber.Text) Then number1 = CSng(txtFirstNumber.Text)
    If IsNumeric(txtSecondNumber.Text) Then number2 = CSng(txtSecondNumber.Text)
    
    sum = number1 + number2

    MsgBox "O resultado da soma é: " & CStr(sum), vbInformation, "Resultado"
    
    txtFirstNumber.Text = ""
    txtSecondNumber.Text = ""
    
End Sub

' Evento KeyPress para o primeiro TextBox
Private Sub txtFirstNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidateDecimalKeyPress(txtFirstNumber, KeyAscii)
End Sub

' Evento KeyPress para o segundo TextBox
Private Sub txtSecondNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidateDecimalKeyPress(txtSecondNumber, KeyAscii)
End Sub

' Função para validar a entrada de dados
Private Function ValidateDecimalKeyPress(txt As TextBox, ByVal KeyAscii As Integer) As Integer
    ' Troca ponto por vírgula
    If KeyAscii = Asc(".") Then
        KeyAscii = Asc(",")
    End If

    ' Permite Backspace ou números
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        ValidateDecimalKeyPress = KeyAscii
        Exit Function
    End If

    ' Permite a vírgula apenas se ainda não existir uma no TextBox
    If KeyAscii = Asc(",") And InStr(txt.Text, ",") = 0 Then
        ValidateDecimalKeyPress = KeyAscii
        Exit Function
    End If

    ' Bloqueia qualquer outro caractere
    ValidateDecimalKeyPress = 0
End Function
