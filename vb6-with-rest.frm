VERSION 5.00
Begin VB.Form frmRestCall 
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   6240
      Width           =   2415
   End
   Begin VB.TextBox txtXML 
      Height          =   4095
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   6735
   End
   Begin VB.CommandButton btnPOST 
      Caption         =   "POST"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtFactura 
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "vb6-with-rest.frx":0000
      Top             =   240
      Width           =   6735
   End
   Begin VB.CommandButton btnGET 
      Caption         =   "GET"
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Status"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
End
Attribute VB_Name = "frmRestCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGET_Click()
    Set httpURL = New WinHttp.WinHttpRequest
    Cadena = "https://jsonplaceholder.typicode.com/todos/1"
    httpURL.Open "GET", Cadena
    httpURL.Send
    Texto = httpURL.ResponseText
    txtXML.Text = Texto
    txtStatus.Text = ""
End Sub
Private Sub btnPOST_Click()
    Set httpURL = New WinHttp.WinHttpRequest
    Cadena = "https://jsonplaceholder.typicode.com/posts"
    httpURL.Open "POST", Cadena, False
    httpURL.SetRequestHeader "Content-type", "application/json"
    httpURL.Send txtFactura.Text
    Texto = httpURL.ResponseText
    txtXML.Text = Texto
    txtStatus.Text = httpURL.Status
End Sub
