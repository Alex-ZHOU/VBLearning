VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private A As String
Private myDate As Date

Private Sub Form_Load()
    A = "广东省深圳市"
    'MsgBox A & vbCrLf + Right(A, 3) & vbCrLf + Left(A, 3)
    'Debug.Print A & vbCrLf + Right(A, 3) & vbCrLf + Left(A, 3)
    myDate = #6/28/2017#
    Debug.Print myDate
End Sub
