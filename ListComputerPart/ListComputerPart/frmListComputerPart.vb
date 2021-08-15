#Region "About"
' / ---------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' /
' / Purpose:  Populate any parts in computer and export to excel.
' / Microsoft Visual Basic .NET (2010) + MS Excel.
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / ---------------------------------------------------------------
#End Region

Imports System.Management
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmListComputerPart

    Private Sub frmListComputerPart_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ' Initialize ListView Control
        With ListView1
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
            .HideSelection = False
            .MultiSelect = False
            .Columns.Add("Part name", ListView1.Width \ 2 - 50)
            .Columns.Add("Description", ListView1.Width \ 2 + 20)
        End With
        Dim LV As ListViewItem
        Dim i As Integer
        ' Processor
        Dim Searcher As New ManagementObjectSearcher( _
                            "root\CIMV2", _
                            "SELECT * FROM Win32_Processor")
        For Each QueryObj As ManagementObject In Searcher.Get()
            LV = ListView1.Items.Add("System Name")
            LV.SubItems.Add(QueryObj("SystemName"))
            LV = ListView1.Items.Add("CPU Name")
            LV.SubItems.Add(QueryObj("Name"))
            LV = ListView1.Items.Add("Processor ID")
            LV.SubItems.Add(QueryObj("ProcessorID"))
        Next
        ' BaseBoard
        Searcher = New ManagementObjectSearcher( _
                            "root\CIMV2", _
                            "SELECT * FROM Win32_BaseBoard")
        For Each QueryObj As ManagementObject In Searcher.Get()
            LV = ListView1.Items.Add("MainBoard Manufacturer")
            LV.SubItems.Add(QueryObj("Manufacturer"))
            LV = ListView1.Items.Add("MainBoard Serial Number")
            LV.SubItems.Add(QueryObj("SerialNumber"))
            LV = ListView1.Items.Add("MainBoard Product Name")
            LV.SubItems.Add(QueryObj("Product"))
        Next
        ' Hard Disk Drive use PhysicalMedia Class
        Searcher = New ManagementObjectSearcher( _
                            "root\CIMV2",
                            "SELECT * FROM Win32_PhysicalMedia")
        i = 1
        For Each QueryObj As ManagementObject In Searcher.Get()
            If InStr(QueryObj("Tag"), "CDROM") = 0 Then
                LV = ListView1.Items.Add("Hard Disk Serial Number (" & i & ")")
                LV.SubItems.Add(Trim(QueryObj("SerialNumber")))
                i = i + 1
            End If
        Next
        ' CD/DVD 
        Searcher = New ManagementObjectSearcher( _
                            "root\CIMV2",
                            "SELECT * FROM Win32_CDROMDrive")
        i = 1
        For Each QueryObj As ManagementObject In Searcher.Get()
            LV = ListView1.Items.Add("CD/DVD Manufacturer (" & i & ")")
            LV.SubItems.Add(Trim(QueryObj("Name")))
            LV = ListView1.Items.Add("CD/DVD Serial Number (" & i & ")")
            LV.SubItems.Add(Trim(QueryObj("SerialNumber")))
            i = i + 1
        Next
        ' Memory
        Searcher = New ManagementObjectSearcher( _
                            "root\CIMV2",
                            "SELECT * FROM Win32_PhysicalMemory")
        i = 1
        For Each QueryObj As ManagementObject In Searcher.Get()
            LV = ListView1.Items.Add("Memory (" & i & ")")
            LV.SubItems.Add(Trim(QueryObj("Manufacturer")) & "/" & Trim(QueryObj("SerialNumber")))
            i = i + 1
        Next
        ' Network adapter
        Searcher = New ManagementObjectSearcher( _
                            "root\CIMV2", _
                            "SELECT * FROM Win32_NetworkAdapter")
        For Each QueryObj As ManagementObject In Searcher.Get()
            If Not String.IsNullOrEmpty(QueryObj("MACAddress")) And Microsoft.VisualBasic.Left(QueryObj("MACAddress"), 2) <> "00" Then
                LV = ListView1.Items.Add("MAC Address")
                LV.SubItems.Add(QueryObj("MACAddress"))
                LV = ListView1.Items.Add("Network Card Manufacturer")
                LV.SubItems.Add(QueryObj("Manufacturer"))
                LV = ListView1.Items.Add("Network Product Name")
                LV.SubItems.Add(QueryObj("ProductName"))
            End If
        Next
    End Sub

    '// Export to Excel.
    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        Dim xlApp As Excel.Application = New Excel.Application
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet
        ' Read data in ListView control and sending them to Excel Sheet.
        For i As Integer = 0 To ListView1.Items.Count - 1
            xlApp.Cells(i + 1, 1) = ListView1.Items(i).Text     ' First column
            xlApp.Cells(i + 1, 2) = ListView1.Items(i).SubItems(1).Text ' Second column
        Next
        ' Adjust autofit columns
        xlWorkSheet.Columns.AutoFit()
        ' Open Excel
        xlApp.Visible = True
        '// Release memory.
        ReleaseObject(xlApp)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlWorkSheet)
    End Sub

    ' / ------------------------------------------------------------------
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            '// GC = Garbage Collection
            GC.Collect()
        End Try
    End Sub

    Private Sub frmListComputerPart_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        ' Resize ListView Control.
        If ListView1.Columns.Count > 0 Then
            With ListView1
                .Columns(0).Width = ListView1.Width \ 2 - 50
                .Columns(1).Width = ListView1.Width \ 2 + 20
            End With
        End If
    End Sub

    Private Sub frmListComputerPart_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

End Class
