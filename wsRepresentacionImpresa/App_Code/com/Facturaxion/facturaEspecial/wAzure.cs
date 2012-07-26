#region "using"

using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web;
using HyperSoft.ElectronicDocumentLibrary.Document;
using Data = HyperSoft.ElectronicDocumentLibrary.Complemento.TimbreFiscalDigital.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using r3TakeCore.Data;

using Microsoft.WindowsAzure;
using Microsoft.WindowsAzure.StorageClient;

#endregion

namespace wsRepresentacionImpresa.App_Code.com.Facturaxion.facturaEspecial
{
    public class wAzure
    {

        private static readonly CultureInfo _ci = new CultureInfo("es-mx");
        private static readonly string _rutaDocs = ConfigurationManager.AppSettings["rutaDocs"];
        private static readonly string _rutaDocsExt = ConfigurationManager.AppSettings["rutaDocsExterna"];
        private static readonly CloudStorageAccount _storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureConnectionString"]);
        private static CloudBlobClient _blobClient;
        private static CloudBlobContainer _blobContainer;
        private static CloudBlob _cloudBlob;

        #region "azure"
        public static string azureUpDownLoad(int up1down2load, string pathPdf)
        {
            //C:\\Inetpub\\repositorioFacturaxion\\cfdi\\2012\\4\\23\\MTE750101S39-K66286-fa8ac37a-5dfa-496d-8c71-74b7f560ac5c.pdf
            string[] pathPdfCarpetas = pathPdf.ToString().Split(new Char[] { '\\' });
            FileStream fs = File.Open(pathPdf, FileMode.Open, FileAccess.Read, FileShare.Read);
            string pathBlob = pathPdf.Replace("C:\\Inetpub\\repositorioFacturaxion\\", "").Replace(pathPdfCarpetas[pathPdfCarpetas.Length - 1], "").Replace("\\", "/");
            Hashtable htAzureParams = new Hashtable();
            htAzureParams.Add("containerName", "boveda");

            htAzureParams.Add("path", pathBlob);
            htAzureParams.Add("filename", pathPdfCarpetas[pathPdfCarpetas.Length - 1]);
            htAzureParams.Add("fileStream", fs);
            htAzureParams.Add("contentType", "application/pdf");
            htAzureParams.Add("metaData", "FileName҈" + pathPdfCarpetas[pathPdfCarpetas.Length - 1]
                + "ᴥPublica҈FacturaxionᴥidFacturaxion҈ᴥFecha҈"+DateTime.Now.ToString("dd/MM/yyyy"));
            /*
             * htAzureParams.Add("path", "cfdi/2011/11/30/");
            htAzureParams.Add("filename", "humbertoss.jpg");
            htAzureParams.Add("fileStream", fs);
            htAzureParams.Add("contentType", "image/jpeg");
            htAzureParams.Add("metaData", "FileName҈humberto.jpgᴥPublica҈FacturaxionᴥidFacturaxion҈22ᴥFecha҈01/01/2020");
            */
            if (up1down2load == 1)
            {
                uploadFileAzure(htAzureParams);
            }
            else
            {
                byte[] buff = downloadFileAzureBuffer(htAzureParams);
                flushFile(buff, pathPdfCarpetas[pathPdfCarpetas.Length - 1]);
            }

            return string.Empty;
        }
        #endregion

        #region "uploadFileAzure"

        public static string uploadFileAzure(Hashtable htAzureParams)
        {
            try
            {
                // Validamos que el Hashtable contenga los parametros indicados y los asignamos a las variables
                string containerName = htAzureParams["containerName"].ToString();
                string path = htAzureParams["path"].ToString();
                string filename = htAzureParams["filename"].ToString();
                FileStream fs = (FileStream)htAzureParams["fileStream"];
                string contentType = htAzureParams["contentType"].ToString();
                string metaData = htAzureParams["metaData"].ToString();

                // Configuramos la conexión al Storage de Windows Azure
                _blobClient = _storageAccount.CreateCloudBlobClient();
                _blobClient.Timeout = new TimeSpan(1, 0, 0);
                _blobClient.ParallelOperationThreadCount = 4;

                // Obtiene o Crea el Contenedor
                _blobContainer = _blobClient.GetContainerReference(containerName);
                _blobContainer.CreateIfNotExist();

                // Configuramos los permisos del Contenedor para que sean Publicos
                BlobContainerPermissions permissions = new BlobContainerPermissions();
                permissions.PublicAccess = BlobContainerPublicAccessType.Container;
                _blobContainer.SetPermissions(permissions);

                // Create the Blob and upload the file
                _cloudBlob = _blobContainer.GetBlobReference(path + filename);
                _cloudBlob.UploadFromStream(fs);

                // Añadimos metadatos al objeto blob
                String[] arrMetaData = metaData.Split(new Char[] { 'ᴥ' });

                foreach (string metaDataPair in arrMetaData)
                {
                    String[] arrMetaDataPair = metaDataPair.Split(new Char[] { '҈' });
                    _cloudBlob.Metadata[arrMetaDataPair[0]] = arrMetaDataPair[1];
                }

                _cloudBlob.SetMetadata();

                // Añadimos las propiedades
                _cloudBlob.Properties.ContentType = contentType;
                _cloudBlob.SetProperties();

                return "1#Archivo Cargado Exitosamente";
            }
            catch (StorageClientException ece)
            {
                return "0#" + "Error en Storage Client: " + ece.Message;
            }
            catch (Exception ex)
            {
                return "0#" + "Error en Ejecución: " + ex.Message;
            }
        }

        #endregion

        #region "downloadFileAzureBuffer"

        public static byte[] downloadFileAzureBuffer(Hashtable htAzureParams)
        {
            try
            {
                // Validamos que el Hashtable contenga los parametros indicados y los asignamos a las variables
                string containerName = htAzureParams["containerName"].ToString();
                string path = htAzureParams["path"].ToString();
                string filename = htAzureParams["filename"].ToString();

                // Asignamos la conexión al Storage de Windows Azure
                _blobClient = _storageAccount.CreateCloudBlobClient();
                _blobContainer = _blobClient.GetContainerReference(containerName);
                _cloudBlob = _blobContainer.GetBlobReference(path + filename);

                return _cloudBlob.DownloadByteArray();
            }
            catch (StorageClientException ex)
            {
                return new byte[0];
            }
            catch (Exception)
            {
                return new byte[0];
            }
        }

        #endregion

        #region "flushFile"

        public static void flushFile(byte[] buff,string fileName)
        {
            HttpContext.Current.Response.BufferOutput = true;
            HttpContext.Current.Response.ClearContent();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.AppendHeader("Content-Length", buff.Length.ToString());
            HttpContext.Current.Response.AppendHeader("Cache-Control", "force-download");
            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + fileName);
            HttpContext.Current.Response.ContentType = "application/pdf";

            if (HttpContext.Current.Response.IsClientConnected)
            {
                HttpContext.Current.Response.BinaryWrite(buff);
                HttpContext.Current.Response.Flush();
                HttpContext.Current.Response.Close();
                HttpContext.Current.ApplicationInstance.CompleteRequest();
            }
        }

        #endregion
    }
}