Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ALIVE = &H103
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2 '��ʾ�Ѵ������óɰ�͸����ʽ
Public Const LWA_COLORKEY = &H1 '��ʾ����ʾ�����е�͸��ɫ

'-----------------------------------����͸��-----------------------------------
Public glasseffectmode As Byte

'SetTransparentWindow(hwnd As Long, iTransparency As Integer)����˵����
'hwndΪ��Ҫ���õĴ�����
'iTransparencyΪ͸���ȣ�Ϊ0-100������0��ʾ��͸����100��ʾȫ͸��

Public Sub SetTransparentWindow(hwnd As Long, iTransparency As Byte)
    Dim rtn As Long
    Dim iTransform As Byte
    'iTransparencyת����SetLayeredWindowAttributes�ĵ�3����������͸���̶�(ȡֵ��Χ0 --255)
    iTransform = Int((100 - iTransparency) * 2.55)
    
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)      'ȡ�Ĵ���ԭ�ȵ���ʽ
    rtn = rtn Or WS_EX_LAYERED 'ʹ����������µ���ʽWS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn    '���µ���ʽ��������
    SetLayeredWindowAttributes hwnd, 0, iTransform, LWA_ALPHA 'ע��:�Ѵ������óɰ�͸����ʽ , ��3������iTransform��ʾ͸���̶ȣ�ȡֵ��Χ0 --255, Ϊ0ʱ����һ��ȫ͸���Ĵ�����
    
End Sub

'ʹ�÷�����SetTransparentWindow Me.hwnd, 60 '�޸����е�60Ϊ0����͸��
