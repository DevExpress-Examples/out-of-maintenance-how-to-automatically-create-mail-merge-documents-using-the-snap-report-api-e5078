#region #References
using System;
using System.Data.OleDb;
using System.IO;
using DevExpress.Snap;
using DevExpress.XtraRichEdit;
using MailMergeServer.nwindDataSetTableAdapters;
// ...
#endregion #References
#region #Code
namespace MailMergeServer {
    class Program {
        const string defaultTemplateFileName = @"..\..\template.snx";
        const string defaultOutputFileName = @"..\..\..\mailmerge.rtf";

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

            SnapDocumentServer server = new SnapDocumentServer();
            server.LoadDocument(templateFileName);
            object dataSource = CreateDataSource();
            server.Document.DataSource = dataSource;
            Console.Write("Performing mail merge... ");
            server.SnapMailMerge(outputFileName, DocumentFormat.Rtf);
            Console.WriteLine("Ok!");
            Console.Write("Press any key...");
            Console.ReadKey();
        }

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
#endregion #Code