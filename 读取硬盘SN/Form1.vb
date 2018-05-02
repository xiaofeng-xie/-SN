Imports System.Management
Imports System.Text.RegularExpressions
Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

      


        ReadHardDiskSerialNumberASCII()
    End Sub

    ' 参考资料：https://www.codeproject.com/Articles/6077/How-to-Retrieve-the-REAL-Hard-Drive-Serial-Number
    '读取字符串形式的，所有硬盘序列号，并返回这些字符的ASCII码
    Public Shared Function ReadHardDiskSerialNumberASCII() As String
        Dim SNList As New List(Of String)
        Dim strReturn As String = ""

        Dim cmicWmi As New System.Management.ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")
        For Each cmicWmiObj As ManagementObject In cmicWmi.Get()
            If Regex.IsMatch(cmicWmiObj("Model").ToString(), "usb", RegexOptions.IgnoreCase) = False Then '判断非移动硬盘
                SNList.Add(Trim(cmicWmiObj("SerialNumber").ToString()))
            End If
        Next


        For Each cmicWmiObj As ManagementObject In cmicWmi.Get()
            MsgBox(cmicWmiObj("Model").ToString()) '判断非移动硬盘
            MsgBox(cmicWmiObj("SerialNumber").ToString())
        Next
        For i = 0 To SNList.Count - 1
            For j = 0 To SNList(i).Length - 1
                strReturn = strReturn & Asc(SNList(i)(j)).ToString()
            Next
        Next

        Return strReturn
    End Function
End Class
