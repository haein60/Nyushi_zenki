VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�w���v"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Label lblHelp 
      BackStyle       =   0  '����
      Caption         =   "Label1"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim sSTR As String

    sSTR = "���[�ݒ�ɂ��Ă̒��ӎ���" & vbCrLf
    sSTR = sSTR & "" & vbCrLf
    sSTR = sSTR & "�P�D���v�Ȃǂ͍����Ȗڂ�ΏۂƂȂ�Ȗڂɂ��đI����Ԃ̏�A�{�^���N���b�N�ɂ��" & vbCrLf
    sSTR = sSTR & "�@�@�ǉ����Ă��������B�I�������Ȗڂ̍��v���ƂȂ�܂��B" & vbCrLf
    sSTR = sSTR & "" & vbCrLf
    sSTR = sSTR & "�Q�D���v���͓o�^��A�^���̃��X�g��I������ƑΏۂ̉Ȗڂ��E���̃��X�g�ɕ\�����܂��B" & vbCrLf
    sSTR = sSTR & "" & vbCrLf
    sSTR = sSTR & "�R�D���v�Ȃǂ͂W�Ԗڈȍ~�łȂ���ΏW�v����܂���B" & vbCrLf
    sSTR = sSTR & "" & vbCrLf
    sSTR = sSTR & "�S�D�ׂ��������ɂ��Ă͎戵���������Q�Ƃ��Ă��������B" & vbCrLf

    lblHelp.Caption = sSTR

End Sub
