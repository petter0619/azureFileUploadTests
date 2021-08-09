using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Blobs.Specialized;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

/*
Sources:
    - https://microsoft.github.io/AzureTipsAndTricks/blog/tip75.html
    - https://microsoft.github.io/AzureTipsAndTricks/blog/tip76.html
    - https://microsoft.github.io/AzureTipsAndTricks/blog/tip95.html
    - https://microsoft.github.io/AzureTipsAndTricks/blog/tip78.html

BlobClient methods: https://docs.microsoft.com/en-us/dotnet/api/azure.storage.blobs.blobclient?view=azure-dotnet
EPPLus: https://epplussoftware.com/docs/5.7/api/OfficeOpenXml.ExcelPackage.html
*/

namespace MicrosoftGithubSolution
{
    class Program
    {
        static void Main(string[] args)
        {
            // ----------------------------------
            // ---------- BLOB STORAGE ----------
            // ----------------------------------

            AzureStorage container = new("testcontainer1");
            
            /* ----- Create container if it doesn't exist ----- */
            container.CreateContainerIfNotExists();

            /* ----- Upload File ----- */
            var sampleImage = @"C:\Users\cme765\OneDrive - AFRY\Desktop\Own Projects\AzureFileUploadTests\MicrosoftGithubSolution\uploads\field.jpg";
            //container.UploadFile(sampleImage, "field.jpg");

            /* ----- Download Upload File ----- */
            var downloadTo = @"C:\Users\cme765\OneDrive - AFRY\Desktop\Own Projects\AzureFileUploadTests\MicrosoftGithubSolution\downloads\field-backup.jpg";
            //container.DownloadFile("field.jpg", downloadTo);

            /* ----- Create copy of file to container ----- */
            //container.CopyToNewFile("field.jpg", "field-backup.jpg");

            /* ----- List files in container (by extension) ----- */
            //container.ListFiles(".jpg");


            // ----------------------------------
            // ------------- EXCEL --------------
            // ----------------------------------

            //container.CreateAndUploadExcelFile("TestReport");

            //container.DownloadChangeCopyExcelFile("TestReport.xlsx", "TestReportChanged");

            container.DownloadChangeReplaceExcelFile("NewTestReport.xlsx");
        }
    }

    class AzureStorage
    {
        private readonly string _conStr = "<connectionString>";
        private readonly BlobContainerClient _container;
        public AzureStorage(string containerName)
        { 
            _container = new BlobContainerClient(_conStr, containerName);
        }

        public void CreateContainerIfNotExists()
        {
            _container.CreateIfNotExists(PublicAccessType.Blob);
        }

        public void UploadFile(string filePath, string fileName)
        {
            var blockBlob = _container.GetBlobClient(fileName);
            using (var fileStream = System.IO.File.OpenRead(filePath))
            {
                blockBlob.Upload(fileStream);
                Console.WriteLine($"{blockBlob.Name} - {blockBlob.Uri}");
            }
        }

        public void DownloadFile(string fileToGet, string writeToPath)
        {
            var blockBlob = _container.GetBlobClient(fileToGet);
            using (var fileStream = System.IO.File.OpenWrite(writeToPath))
            {
                blockBlob.DownloadTo(fileStream);
            }
        }

        #nullable enable
        public void ListFiles(string? extension)
        {
            var list = _container.GetBlobs();

            if (extension != null)
            {
                var blobs = list.Where(b => Path.GetExtension(b.Name).Equals(extension));
                foreach (var item in blobs)
                {
                    string name = item.Name;
                    Console.WriteLine(name);
                }
            }
            else 
            {
                foreach (var item in list)
                {
                    string name = item.Name;
                    Console.WriteLine(name);
                }
            }
        }
        #nullable disable

        public void CopyToNewFile(string fileToCopy, string copyFileName)
        {
            BlockBlobClient blockBlob = _container.GetBlockBlobClient(fileToCopy);

            BlockBlobClient copyBlockBlob = _container.GetBlockBlobClient(copyFileName);

            //Console.WriteLine(blockBlob.Uri);

            copyBlockBlob.StartCopyFromUri(blockBlob.Uri);

            Console.WriteLine($"{copyBlockBlob.Name} - {copyBlockBlob.Uri}"); // Get new file Name && URI
        }

        // Excel Methods
        public void CreateAndUploadExcelFile(string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(fileName + ".xlsx"))) // Create Excel file
            using (var fileStream = new MemoryStream()) // Create stream
            {
                // Add content to Excel file
                var ws = package.Workbook.Worksheets.Add("MainReport");
                ws.Cells["A1"].Value = "Sample Report";

                // Add Excel file to stream
                package.SaveAs(fileStream);

                // Upload stream to Blob Storage
                fileStream.Position = 0;
                var blockBlob = _container.GetBlobClient(fileName + ".xlsx");
                blockBlob.Upload(fileStream);

                Console.WriteLine($"{blockBlob.Name} - {blockBlob.Uri}"); // Get new file Name && URI
            }
        }

        public void DownloadChangeCopyExcelFile(string fileToGet, string copyFileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var blockBlob = _container.GetBlobClient(fileToGet);

            using (var memorystream = new MemoryStream())
            {
                blockBlob.DownloadTo(memorystream);
                memorystream.Position = 0;

                using (var package = new ExcelPackage(memorystream))
                using (var newFileStream = new MemoryStream())
                {
                    var ws = package.Workbook.Worksheets[0];
                    Console.WriteLine(ws.Cells["A1"].Value); // Read current value
                    ws.Cells["A1"].Value = "New Sample Report"; // Change value

                    // Upload to Blob Storage
                    package.SaveAs(newFileStream);
                    newFileStream.Position = 0;
                    var newblockBlob = _container.GetBlobClient(copyFileName + ".xlsx");
                    newblockBlob.Upload(newFileStream);

                    Console.WriteLine($"{newblockBlob.Name} - {newblockBlob.Uri}");
                }
            }
        }

        public void DownloadChangeReplaceExcelFile(string fileToGet)
        {
            Console.WriteLine("----- Executing... -----");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var blockBlob = _container.GetBlobClient(fileToGet);

            using (var memorystream = new MemoryStream())
            {
                blockBlob.DownloadTo(memorystream);
                memorystream.Position = 0;

                // Delete current version of file...
                blockBlob.Delete();

                using (var package = new ExcelPackage(memorystream))
                using (var newFileStream = new MemoryStream())
                {
                    var ws = package.Workbook.Worksheets[0];
                    Console.WriteLine(ws.Cells["A1"].Value); // Read current value
                    ws.Cells["A1"].Value = "New Sample Report v2.0"; // Change value

                    // Upload to Blob Storage
                    package.SaveAs(newFileStream);
                    newFileStream.Position = 0;
                    var newblockBlob = _container.GetBlobClient(fileToGet);
                    newblockBlob.Upload(newFileStream);

                    Console.WriteLine($"{newblockBlob.Name} - {newblockBlob.Uri}");
                }
            }

            Console.WriteLine("----- Finished... -----");
        }

        // .excelMethods
    }
}
