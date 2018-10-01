#region #Usings
using DevExpress.Snap;
using DevExpress.Snap.Core.API;
using DevExpress.XtraRichEdit;
using DevEspress.Snap.Core.Options;
using MailMergeServer.nwindDataSetTableAdapters;
using System;
using System.Data.OleDb;
using System.IO;
// ...
#endregion #Usings

namespace MailMergeServer {
    class Program {

        const string defaultTemplateFileName = "template.snx";
        const string defaultOutputFileName = "mailmerge.rtf";

        static void Main(string[] args) {
            Console.WriteLine("Mail Merge Server");
            string templateFileName;
            string outputFileName;

            switch (args.Length) {
                case 0:
                    templateFileName = defaultTemplateFileName;
                    outputFileName = defaultOutputFileName;
                    break;
                case 1:
                    templateFileName = defaultTemplateFileName;
                    outputFileName = args[0];
                    break;
                case 2:
                    templateFileName = args[0];
                    outputFileName = args[1];
                    break;
                default:
                    throw new ArgumentException();
            }

            if (!File.Exists(templateFileName))
                throw new FileNotFoundException("Template file not found", templateFileName);
            Console.WriteLine("Template file: {0}", new FileInfo(templateFileName).FullName);
            Console.WriteLine("Target file:   {0}", new FileInfo(outputFileName).FullName);
            #region #ServerCode
            SnapDocumentServer server = new SnapDocumentServer();

            server.SnapMailMergeRecordStarted += server_SnapMailMergeRecordStarted;
            server.SnapMailMergeRecordFinished += server_SnapMailMergeRecordFinished;

            server.LoadDocument(templateFileName);
            object dataSource = CreateDataSource();
            
            SnapMailMergeExportOptions options = server.Document.CreateSnapMailMergeExportOptions();
            options.DataSource = dataSource;
            Console.Write("Performing mail merge... ");
            server.SnapMailMerge(options, outputFileName, DocumentFormat.Rtf);
            #endregion #ServerCode
            Console.WriteLine("Ok!");
            Console.Write("Press any key...");
            Console.ReadKey();
            System.Diagnostics.Process.Start(outputFileName);
        }

        #region #RecordFinished
        static void server_SnapMailMergeRecordFinished(object sender, SnapMailMergeRecordFinishedEventArgs e)
        {
            if (e.RecordIndex == 3)
            e.RecordDocument.AppendText("This is the third data record.\r\n");
        }
        #endregion #RecordFinished

        #region #RecordStarted
        static void server_SnapMailMergeRecordStarted(object sender, SnapMailMergeRecordStartedEventArgs e)
        {            
            if (e.RecordIndex == 3) {
                for (int i = 0; i < e.RecordDocument.Fields.Count; i++)                {
                    DevExpress.XtraRichEdit.API.Native.Field item = e.RecordDocument.Fields[i];
                    SnapImage snImage = e.RecordDocument.ParseField(item) as SnapImage;
                    if (snImage != null)
                    {
                        if (snImage.DataFieldName == "Picture")
                        {
                            snImage.BeginUpdate();
                            snImage.ScaleX = snImage.ScaleX * 2;
                            snImage.ScaleY = snImage.ScaleY * 2;
                            snImage.EndUpdate();
                            item.Update();
                        }
                    }
                }
                e.RecordDocument.EndUpdate();

                // Another code snippet for the same result:
                //e.RecordDocument.Fields[2].ShowCodes = true;
                //e.RecordDocument.Replace(e.RecordDocument.Fields[2].CodeRange, @"SNIMAGE Picture \sy 20000 \sx 20000");
                //e.RecordDocument.Fields[2].Update();
            }
        }
        #endregion #RecordStarted

        static object CreateDataSource() {
            var dataSource = new nwindDataSet();
            var connection = new OleDbConnection();
            connection.ConnectionString = Properties.Settings.Default.nwindConnectionString;

            CategoriesTableAdapter categories = new CategoriesTableAdapter();
            categories.Connection = connection;
            categories.Fill(dataSource.Categories);

            ProductsTableAdapter products = new ProductsTableAdapter();
            products.Connection = connection;
            products.Fill(dataSource.Products);

            return dataSource;
        }
    }
}
