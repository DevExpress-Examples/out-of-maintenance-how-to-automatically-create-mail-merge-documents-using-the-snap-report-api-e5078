#Region "#References"
Imports System
Imports System.Data.OleDb
Imports System.IO
Imports DevExpress.Snap
Imports DevExpress.XtraRichEdit
Imports MailMergeServer.nwindDataSetTableAdapters
' ...
#End Region ' #References
#Region "#Code"
Namespace MailMergeServer
    Friend Class Program
        Private Const defaultTemplateFileName As String = "..\..\template.snx"
        Private Const defaultOutputFileName As String = "..\..\..\mailmerge.rtf"

        Shared Sub Main(ByVal args() As String)
            Console.WriteLine("Mail Merge Server")
            Dim templateFileName As String
            Dim outputFileName As String

            Select Case args.Length
                Case 0
                    templateFileName = defaultTemplateFileName
                    outputFileName = defaultOutputFileName
                Case 1
                    templateFileName = defaultTemplateFileName
                    outputFileName = args(0)
                Case 2
                    templateFileName = args(0)
                    outputFileName = args(1)
                Case Else
                    Throw New ArgumentException()
            End Select

            If Not File.Exists(templateFileName) Then
                Throw New FileNotFoundException("Template file not found", templateFileName)
            End If
            Console.WriteLine("Template file: {0}", (New FileInfo(templateFileName)).FullName)
            Console.WriteLine("Target file:   {0}", (New FileInfo(outputFileName)).FullName)

            Dim server As New SnapDocumentServer()
            server.LoadDocument(templateFileName)
            Dim dataSource As Object = CreateDataSource()
            server.Document.DataSource = dataSource
            Console.Write("Performing mail merge... ")
            server.SnapMailMerge(outputFileName, DocumentFormat.Rtf)
            Console.WriteLine("Ok!")
            Console.Write("Press any key...")
            Console.ReadKey()
        End Sub

        Private Shared Function CreateDataSource() As Object
            Dim dataSource = New nwindDataSet()
            Dim connection = New OleDbConnection()
            connection.ConnectionString = My.Settings.Default.nwindConnectionString

            Dim categories As New CategoriesTableAdapter()
            categories.Connection = connection
            categories.Fill(dataSource.Categories)

            Dim products As New ProductsTableAdapter()
            products.Connection = connection
            products.Fill(dataSource.Products)

            Return dataSource
        End Function
    End Class
End Namespace
#End Region ' #Code