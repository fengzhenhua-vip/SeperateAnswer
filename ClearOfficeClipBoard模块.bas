Attribute VB_Name = "ClearOfficeClipBoardģ��"
'���ƣ�ClearOfficeClipBoard (���Office���а�)
'��Դ��https://stackoverflow.com/questions/14440274/cant-clear-office-clipboard-with-vba
'˵����������������ַ����ʱ���������޸�ʹ֮������ȷѡ��VBA�汾���У�ԭʼ�汾ԭ����ȷ�����ǽṹ����
Public myVBA7 As Integer
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As Office.IAccessible, _
                                                                  ByVal iChildStart As Long, ByVal cChildren As Long, _
                                                                  ByRef rgvarChildren As Any, ByRef pcObtained As Long) As Long
Public Sub ClearOfficeClipBoard()
    If VBA7 Then
        myVBA7 = 1
    Else
        myVBA7 = 0
    End If
    Dim cmnB, IsVis As Boolean, j As Long, Arr As Variant
    Arr = Array(4, 7, 2, 0)                                     '4 and 2 for 32 bit, 7 and 0 for 64 bit
    Set cmnB = Application.CommandBars("Office Clipboard")
    IsVis = cmnB.Visible
    If Not IsVis Then
        cmnB.Visible = True
        DoEvents
    End If

    For j = 1 To Arr(0 + myVBA7)
        AccessibleChildren cmnB, Choose(j, 0, 3, 0, 3, 0, 3, 1), 1, cmnB, 1
    Next
        
    cmnB.accDoDefaultAction CLng(Arr(2 + myVBA7))

    Application.CommandBars("Office Clipboard").Visible = IsVis

End Sub
