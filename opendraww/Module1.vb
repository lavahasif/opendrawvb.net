Imports System.Runtime.InteropServices

Module Module1
    Public Sub Main(ByVal args As String())
        Const ESC As String = ChrW(27)
        Const p As String = "p"
        Const m As String = vbNullChar
        Const t1 As String = "%"
        Const t2 As String = "ɐ"
        Const openTillCommand As String = ESC & p & m & t1 & t2
'        printer("POS-X Thermal Printer")
        RawPrinterHelper.SendStringToPrinter("POS-X Thermal Printer", openTillCommand)
    End Sub
    Public Sub printer( printe As String)
'        Const ESC As String = ChrW(27)
'        Const p As String = "p"
'        Const m As String = vbNullChar
'        Const t1 As String = "%"
'        Const t2 As String = "ɐ"
'        Const openTillCommand As String = ESC & p & m & t1 & t2
'        RawPrinterHelper.SendStringToPrinter(printe, openTillCommand)
        Dim buffer As Byte() = New Byte(4) {CByte(27), CByte(112), CByte(0), CByte(25), CByte(250)}
        SendBytesToLocalPrinter(buffer,printe)
    End Sub
    
    Public  Sub SendBytesToLocalPrinter(ByVal data As Byte(), ByVal printerName As String)
        Dim size = Marshal.SizeOf(data(0)) * data.Length
        Dim pBytes = Marshal.AllocHGlobal(size)

        Try
            RawPrinterHelper.SendBytesToPrinter(printerName, pBytes, size)
        Finally
            Marshal.FreeCoTaskMem(pBytes)
        End Try
    End Sub
End Module
