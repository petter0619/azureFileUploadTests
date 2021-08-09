import React from 'react';
import { BlobServiceClient, BlobClient } from "@azure/storage-blob";

// @azure/storage-blob NPM page: https://www.npmjs.com/package/@azure/storage-blob
// BlobServiceClient docs: https://docs.microsoft.com/en-us/javascript/api/@azure/storage-blob/blobserviceclient?view=azure-node-latest
// ContainerClient docs: https://docs.microsoft.com/en-us/javascript/api/@azure/storage-blob/containerclient?view=azure-node-latest
// BlobClient docs: https://docs.microsoft.com/en-us/javascript/api/@azure/storage-blob/blobclient?view=azure-node-latest

export default function AzureButton({ downloadUrl }) {
    // const account = "teststorageaccount0619";
    // const sas = "sp=racwdl&st=2021-08-02T06:40:24Z&se=2021-08-06T14:40:24Z&spr=https&sv=2020-08-04&sr=c&sig=TCNcw6BwTH12VeEvKfoLcBdiu6gz9%2B32qdO%2FsHSq1w4%3D";
    // const containerName = "report-templates";
    // const blobName = "ProjectRapport90c4137f-4a8f-4810-928d-6fe56b146174.xlsx";

    // const blobServiceClient = new BlobServiceClient(`https://${account}.blob.core.windows.net?${sas}`);
 
    var saveBlob = (function () {
        var a = document.createElement("a");
        document.body.appendChild(a);
        a.style = "display: none";
        return function (blob, fileName) {
            var url = window.URL.createObjectURL(blob);
            a.href = url;
            a.download = fileName;
            a.click();
            window.URL.revokeObjectURL(url);
        };
    }());

    async function downloadThenDelete() {
        //const containerClient = blobServiceClient.getContainerClient(containerName);
        //const blobClient = containerClient.getBlobClient(blobName);
        
        const blobClient = new BlobClient(downloadUrl);

        // Get blob content from position 0 to the end
        // In browsers, get downloaded data by accessing downloadBlockBlobResponse.blobBody
        const downloadBlockBlobResponse = await blobClient.download();
        const blob = await downloadBlockBlobResponse.blobBody;

        // Download file to computer
        saveBlob(blob, "<ReportName>.xlsx");
        console.log("Downloaded blob content");

        // Delete file from Azure Blob Storage
        const deleteResponse = await blobClient.delete();
        console.log("Deleted blob", deleteResponse);
    }

    return (
        <button onClick={downloadThenDelete}>
            Download
        </button>
    )
}
