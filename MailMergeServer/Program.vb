#Region "#Usings"
Imports DevExpress.Snap
Imports DevExpress.Snap.Core.API
Imports DevExpress.XtraRichEdit
Imports DevExpress.Snap.Core.Options
Imports MailMergeServer.nwindDataSetTableAdapters
Imports System
Imports System.Data.OleDb
Imports System.IO
' ...
#End Region ' #Usings

Namespace MailMergeServer
    Friend Class Program

        Private Const defaultTemplateFileName As String = "template.snx"
        Private Const defaultOutputFileName As String = "mailmerge.rtf"

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
'            #Region "#ServerCode"
            Dim server As New SnapDocumentServer()

            AddHandler server.SnapMailMergeRecordStarted, AddressOf server_SnapMailMergeRecordStarted
            AddHandler server.SnapMailMergeRecordFinished, AddressOf server_SnapMailMergeRecordFinished

            server.LoadDocument(templateFileName)
            Dim dataSource As Object = CreateDataSource()

            Dim options As SnapMailMergeExportOptions = server.Document.CreateSnapMailMergeExportOptions()
            options.DataSource = dataSource
            Console.Write("Performing mail merge... ")
            server.SnapMailMerge(options, outputFileName, DocumentFormat.Rtf)
'            #End Region ' #ServerCode
            Console.WriteLine("Ok!")
            Console.Write("Press any key...")
            Console.ReadKey()
            System.Diagnostics.Process.Start(outputFileName)
        End Sub

        #Region "#RecordFinished"
        Private Shared Sub server_SnapMailMergeRecordFinished(ByVal sender As Object, ByVal e As SnapMailMergeRecordFinishedEventArgs)
            If e.RecordIndex = 3 Then
            e.RecordDocument.AppendText("This is the third data record." & ControlChars.CrLf)
            End If
        End Sub
        #End Region ' #RecordFinished

        #Region "#RecordStarted"
        Private Shared Sub server_SnapMailMergeRecordStarted(ByVal sender As Object, ByVal e As SnapMailMergeRecordStartedEventArgs)
            If e.RecordIndex = 3 Then
                For i As Integer = 0 To e.RecordDocument.Fields.Count - 1
                    Dim item As DevExpress.XtraRichEdit.API.Native.Field = e.RecordDocument.Fields(i)
                    Dim snImage As SnapImage = TryCast(e.RecordDocument.ParseField(item), SnapImage)
                    If snImage IsNot Nothing Then
                        If snImage.DataFieldName = "Picture" Then
                            snImage.BeginUpdate()
                            snImage.ScaleX = snImage.ScaleX * 2
                            snImage.ScaleY = snImage.ScaleY * 2
                            snImage.EndUpdate()
                            item.Update()
                        End If
                    End If
                Next i
                e.RecordDocument.EndUpdate()

                ' Another code snippet for the same result:
                'e.RecordDocument.Fields[2].ShowCodes = true;
                'e.RecordDocument.Replace(e.RecordDocument.Fields[2].CodeRange, @"SNIMAGE Picture \sy 20000 \sx 20000");
                'e.RecordDocument.Fields[2].Update();
            End If
        End Sub
        #End Region ' #RecordStarted

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
