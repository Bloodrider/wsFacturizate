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
    public class CED110324NN4
    {
        #region "Enum Monedas"

        enum Monedas
        {
            Peso = 1,
            Dolar = 2,
            Euro = 3,
            Yen = 4,
            LibraEsterlina = 5
        };

        #endregion

        #region "variables"

        public static CultureInfo _ci = new CultureInfo("es-mx");
        private static readonly string _rutaDocs = ConfigurationManager.AppSettings["rutaDocs"];
        private static readonly string _rutaDocsExt = ConfigurationManager.AppSettings["rutaDocsExterna"];

        private static bool timbrar;
        private static HttpContext HTC;
        private static String pathIMGLOGO;
        //private static Document document;
        private static PdfPCell cell;
        private static Cell cel;
        private static Paragraph par;
        private static Chunk dSaltoLinea;
        
        private static Color azul;
        private static Color blanco;
        private static Color Link;
        private static Color gris;
        private static Color grisOX;
        private static Color rojo;
        private static Color lbAzul;

        private static BaseFont EM;
        private static Font f5;
        private static Font f5B;
        private static Font f5BBI;
        private static Font f6;
        private static Font f6B;
        private static Font f6L;
        private static Font f7B;
        private static Font f8B;
        private static Font f8L;
        private static Font titulo;
        private static Font folio;
        private static Font f5L;
        private static Font f8LA;

        #endregion

        #region "generarPdf"

        public static string generarPdf(Hashtable htFacturaxion, HttpContext hc)
        {
            string pathPdf = htFacturaxion["rutaDocumentoPdf"].ToString();
            FileStream fs = new FileStream(pathPdf, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);

            try
            {
                DAL dal = new DAL();
                StringBuilder sbConfigFactParms = new StringBuilder();
                _ci.NumberFormat.CurrencyDecimalDigits = 2;

                ElectronicDocument electronicDocument = (ElectronicDocument)htFacturaxion["electronicDocument"];
                Data objTimbre = (Data)htFacturaxion["objTimbre"];
                timbrar = Convert.ToBoolean(htFacturaxion["timbrar"]);
                Int64 idCfdi = Convert.ToInt64(htFacturaxion["idCfdi"]);

                #region "Extraemos Datos Adicionales"

                DataTable dtOpcionalDetalle = new DataTable();
                StringBuilder sbOpcionalDetalle = new StringBuilder();

                sbOpcionalDetalle.
                    Append("SELECT ").
                    Append("COALESCE(campo1, '0') AS noIdent ").
                    Append("FROM opcionalDetalle ").
                    Append("WHERE idCFDI = @0 ");

                dtOpcionalDetalle = dal.QueryDT("DS_FE", sbOpcionalDetalle.ToString(), "F:I:" + idCfdi, hc);

                if (dtOpcionalDetalle.Rows.Count == 0)
                {
                    for (int i = 1; i <= electronicDocument.Data.Conceptos.Count; i++)
                    {
                        dtOpcionalDetalle.Rows.Add("0");
                    }
                }

                #endregion

                #region "Extraemos los datos del CFDI"

                Hashtable htDatosCfdi = new Hashtable();
                htDatosCfdi.Add("nombreEmisor", electronicDocument.Data.Emisor.Nombre.Value);
                htDatosCfdi.Add("rfcEmisor", electronicDocument.Data.Emisor.Rfc.Value);
                htDatosCfdi.Add("nombreReceptor", electronicDocument.Data.Receptor.Nombre.Value);
                htDatosCfdi.Add("rfcReceptor", electronicDocument.Data.Receptor.Rfc.Value);
                htDatosCfdi.Add("serie", electronicDocument.Data.Serie.Value);
                htDatosCfdi.Add("folio", electronicDocument.Data.Folio.Value);
                htDatosCfdi.Add("fechaCfdi", electronicDocument.Data.Fecha.Value);
                htDatosCfdi.Add("UUID", objTimbre.Uuid.Value);

                #region "Dirección Emisor"

                StringBuilder sbDirEmisor1 = new StringBuilder();
                StringBuilder sbDirEmisor2 = new StringBuilder();
                StringBuilder sbDirEmisor3 = new StringBuilder();

                if (electronicDocument.Data.Emisor.Domicilio.Calle.Value.Length > 0)
                {
                    sbDirEmisor1.Append(electronicDocument.Data.Emisor.Domicilio.Calle.Value).Append(" ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.NumeroExterior.Value.Length > 0)
                {
                    sbDirEmisor1.Append(electronicDocument.Data.Emisor.Domicilio.NumeroExterior.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.NumeroInterior.Value.Length > 0)
                {
                    sbDirEmisor1.Append(" ").Append(electronicDocument.Data.Emisor.Domicilio.NumeroInterior.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.Colonia.Value.Length > 0)
                {
                    sbDirEmisor2.Append(electronicDocument.Data.Emisor.Domicilio.Colonia.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.Localidad.Value.Length > 0)
                {
                    sbDirEmisor2.Append(electronicDocument.Data.Emisor.Domicilio.Localidad.Value);
                }
                if (electronicDocument.Data.Emisor.Domicilio.Municipio.Value.Length > 0)
                {
                    sbDirEmisor3.Append(electronicDocument.Data.Emisor.Domicilio.Municipio.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.Estado.Value.Length > 0)
                {
                    sbDirEmisor3.Append(electronicDocument.Data.Emisor.Domicilio.Estado.Value).Append(" ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.CodigoPostal.Value.Length > 0)
                {
                    sbDirEmisor3.Append("C.P. ").Append(electronicDocument.Data.Emisor.Domicilio.CodigoPostal.Value).Append(", ");
                }
                sbDirEmisor3.Append(electronicDocument.Data.Emisor.Domicilio.Pais.Value);

                #endregion

                #region "Dirección Sucursal Expedido En"

                StringBuilder sbDirExpedido1 = new StringBuilder();
                StringBuilder sbDirExpedido2 = new StringBuilder();
                StringBuilder sbDirExpedido3 = new StringBuilder();

                if (electronicDocument.Data.Emisor.ExpedidoEn.Calle.Value.Length > 0)
                {
                    sbDirExpedido1.Append(electronicDocument.Data.Emisor.ExpedidoEn.Calle.Value).Append(" ");
                }
                if (electronicDocument.Data.Emisor.ExpedidoEn.NumeroExterior.Value.Length > 0)
                {
                    sbDirExpedido1.Append(" ").Append(electronicDocument.Data.Emisor.ExpedidoEn.NumeroExterior.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.ExpedidoEn.NumeroInterior.Value.Length > 0)
                {
                    sbDirExpedido1.Append(" ").Append(electronicDocument.Data.Emisor.ExpedidoEn.NumeroInterior.Value);
                }
                if (electronicDocument.Data.Emisor.ExpedidoEn.Colonia.Value.Length > 0)
                {
                    sbDirExpedido2.Append(electronicDocument.Data.Emisor.ExpedidoEn.Colonia.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.ExpedidoEn.Localidad.Value.Length > 0)
                {
                    sbDirExpedido2.Append(electronicDocument.Data.Emisor.ExpedidoEn.Localidad.Value);
                }
                if (electronicDocument.Data.Emisor.ExpedidoEn.Municipio.Value.Length > 0)
                {
                    sbDirExpedido3.Append(electronicDocument.Data.Emisor.ExpedidoEn.Municipio.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.ExpedidoEn.Estado.Value.Length > 0)
                {
                    sbDirExpedido3.Append(electronicDocument.Data.Emisor.ExpedidoEn.Estado.Value).Append(" ");
                }
                if (electronicDocument.Data.Emisor.ExpedidoEn.CodigoPostal.Value.Length > 0)
                {
                    sbDirExpedido3.Append("C.P. ").Append(electronicDocument.Data.Emisor.ExpedidoEn.CodigoPostal.Value).Append(", ");
                }
                sbDirExpedido3.Append(electronicDocument.Data.Emisor.ExpedidoEn.Pais.Value);

                #endregion

                #region "Dirección Receptor"

                StringBuilder sbDirReceptor1 = new StringBuilder();
                StringBuilder sbDirReceptor2 = new StringBuilder();
                StringBuilder sbDirReceptor3 = new StringBuilder();

                if (electronicDocument.Data.Receptor.Domicilio.Calle.Value.Length > 0)
                {
                    sbDirReceptor1.Append(electronicDocument.Data.Receptor.Domicilio.Calle.Value).Append(" ");
                }
                if (electronicDocument.Data.Receptor.Domicilio.NumeroExterior.Value.Length > 0)
                {
                    sbDirReceptor1.Append(" ").Append(electronicDocument.Data.Receptor.Domicilio.NumeroExterior.Value).Append(" ");
                }
                if (electronicDocument.Data.Receptor.Domicilio.NumeroInterior.Value.Length > 0)
                {
                    sbDirReceptor1.Append(" ").Append(electronicDocument.Data.Receptor.Domicilio.NumeroInterior.Value);
                }
                if (electronicDocument.Data.Receptor.Domicilio.Colonia.Value.Length > 0)
                {
                    sbDirReceptor2.Append(electronicDocument.Data.Receptor.Domicilio.Colonia.Value).Append(", ");
                }
                if (electronicDocument.Data.Receptor.Domicilio.Localidad.Value.Length > 0)
                {
                    sbDirReceptor2.Append(electronicDocument.Data.Receptor.Domicilio.Localidad.Value);
                }
                if (electronicDocument.Data.Receptor.Domicilio.Municipio.Value.Length > 0)
                {
                    sbDirReceptor3.Append(electronicDocument.Data.Receptor.Domicilio.Municipio.Value).Append(", ");
                }
                if (electronicDocument.Data.Receptor.Domicilio.Estado.Value.Length > 0)
                {
                    sbDirReceptor3.Append(electronicDocument.Data.Receptor.Domicilio.Estado.Value).Append(" ");
                }
                if (electronicDocument.Data.Receptor.Domicilio.CodigoPostal.Value.Length > 0)
                {
                    sbDirReceptor3.Append("C.P. ").Append(electronicDocument.Data.Receptor.Domicilio.CodigoPostal.Value).Append(", ");
                }
                sbDirReceptor3.Append(electronicDocument.Data.Receptor.Domicilio.Pais.Value);

                #endregion

                htDatosCfdi.Add("direccionEmisor1", sbDirEmisor1.ToString());
                htDatosCfdi.Add("direccionEmisor2", sbDirEmisor2.ToString());
                htDatosCfdi.Add("direccionEmisor3", sbDirEmisor3.ToString());

                htDatosCfdi.Add("direccionExpedido1", sbDirExpedido1.ToString());
                htDatosCfdi.Add("direccionExpedido2", sbDirExpedido2.ToString());
                htDatosCfdi.Add("direccionExpedido3", sbDirExpedido3.ToString());

                htDatosCfdi.Add("direccionReceptor1", sbDirReceptor1.ToString());
                htDatosCfdi.Add("direccionReceptor2", sbDirReceptor2.ToString());
                htDatosCfdi.Add("direccionReceptor3", sbDirReceptor3.ToString());

                #endregion

                #region "Creamos el Objeto Documento y Tipos de Letra"

                Document document = new Document(PageSize.LETTER, 20, 20, 20, 40);
                document.AddAuthor("Facturaxion");
                document.AddCreator("r3Take");
                document.AddCreationDate();

                //FileStream fs = new FileStream(pathPdf, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                pdfPageEventHandlerCED110324NN4 pageEventHandler = new pdfPageEventHandlerCED110324NN4();
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                writer.PageEvent = pageEventHandler;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                //document.Open();

                HTC = hc;
                azul = new Color(22, 111, 168);
                blanco = new Color(255, 255, 255);
                Link = new Color(7, 73, 208);
                gris = new Color(236, 236, 236);
                grisOX = new Color(220, 215, 220);
                rojo = new Color(230, 7, 7);
                lbAzul = new Color(43, 145, 175);

                EM = BaseFont.CreateFont(@"C:\Windows\Fonts\VERDANA.TTF", BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
                f5 = new Font(EM, 5);
                f5B = new Font(EM, 5, Font.BOLD);
                f5BBI = new Font(EM, 5, Font.BOLDITALIC);
                f6 = new Font(EM, 6);
                f6B = new Font(EM, 6, Font.BOLD);
                f7B = new Font(EM, 7, Font.BOLD);
                f6L = new Font(EM, 6, Font.BOLD, Link);
                f5L = new Font(EM, 5, Font.BOLD, lbAzul);
                f8B = new Font(EM, 8, Font.BOLD);
                f8L = new Font(EM, 8);
                f8LA = new Font(EM, 8, Font.BOLD, lbAzul);
                titulo = new Font(EM, 6, Font.BOLD, blanco);
                folio = new Font(EM, 6, Font.BOLD, rojo);
                dSaltoLinea = new Chunk("\n\n ");

                #endregion

                #region "Generamos el Docuemto"

                formatoCED110324NN4(document, electronicDocument, objTimbre, pageEventHandler, idCfdi, dtOpcionalDetalle, htDatosCfdi, HTC);

                #endregion

                document.Close();
                writer.Close();
                fs.Close();

                string filePdfExt = pathPdf.Replace(_rutaDocs, _rutaDocsExt);
                string urlPathFilePdf = filePdfExt.Replace(@"\", "/");

                //Subimos Archivo al Azure
                string res = App_Code.com.Facturaxion.facturaEspecial.wAzure.azureUpDownLoad(1, pathPdf);

                return "1#" + urlPathFilePdf;
            }
            catch (Exception ex)
            {
                fs.Flush();
                fs.Close();
                File.Delete(pathPdf);

                return "0#" + ex.Message;
            }
        }

        #endregion

        #region "formatoCED110324NN4"

        public static void formatoCED110324NN4(Document document, ElectronicDocument electronicDocument, Data objTimbre, pdfPageEventHandlerCED110324NN4 pageEventHandler, Int64 idCfdi, DataTable dtOpcDet, Hashtable htCFDI, HttpContext hc)
        {
            try
            {
                DAL dal = new DAL();
                #region "Construimos el Documento"

                #region "Construimos el Encabezado"

                Table encabezado = new Table(3);
                float[] headerwidthsEncabezado = { 60, 20, 20 };
                encabezado.Widths = headerwidthsEncabezado;
                encabezado.WidthPercentage = 100;
                encabezado.Padding = 1;
                encabezado.Spacing = 1;
                encabezado.BorderWidth = 0;
                encabezado.DefaultCellBorder = 0;
                encabezado.BorderColor = gris;

                cel = new Cell(new Phrase("COMPROBANTE FISCAL DIGITAL POR INTERNET", f8LA));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(8f, 9f);
                par.Add(new Chunk(htCFDI["nombreEmisor"].ToString().ToUpper(), f8B));
                par.Add(new Chunk("\n\nRFC: " + htCFDI["rfcEmisor"].ToString().ToUpper(), f8L));
                par.Add(new Chunk("\n\n" + htCFDI["direccionEmisor1"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor2"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor3"].ToString().ToUpper(), f6));
                cel = new Cell(par);
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = gris;
                cel.Rowspan = 4;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Serie/Folio", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Número de Certificado", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["serie"].ToString().ToUpper() + " " + htCFDI["folio"].ToString(), folio));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.NumeroCertificado.Value.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Fecha/Hora: ", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["fechaCfdi"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Tipo:", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.TipoComprobante.Value.ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Expedido en: \n", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("FOLIO FISCAL", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                encabezado.AddCell(cel);

                StringBuilder expedido = new StringBuilder();
                expedido.
                    Append(htCFDI["direccionExpedido1"].ToString().ToUpper()).Append("\n").
                    Append(htCFDI["direccionExpedido2"].ToString().ToUpper()).Append("\n").
                    Append(htCFDI["direccionExpedido3"].ToString().ToUpper()).Append("\n");

                cel = new Cell(new Phrase(expedido.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["UUID"].ToString(), f7B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                encabezado.AddCell(cel);
                

                //Tabla Receptor y Encabezados Detalle

                Table tReceptor = new Table(5);
                float[] headerwidthsReceptor = { 15, 15, 40, 15, 15 };
                tReceptor.Widths = headerwidthsReceptor;
                tReceptor.WidthPercentage = 100;
                tReceptor.Padding = 1;
                tReceptor.Spacing = 1;
                tReceptor.BorderWidth = 0;
                tReceptor.DefaultCellBorder = 0;
                tReceptor.BorderColor = gris;

                cel = new Cell(new Phrase("Receptor:\n", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 5;
                tReceptor.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["nombreReceptor"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 5;
                tReceptor.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("R.F.C. " + electronicDocument.Data.Receptor.Rfc.Value + "\n\n", f6));
                par.Add(new Chunk(htCFDI["direccionReceptor1"].ToString() + "\n", f6));
                par.Add(new Chunk(htCFDI["direccionReceptor2"].ToString() + "\n", f6));
                par.Add(new Chunk(htCFDI["direccionReceptor3"].ToString() + "\n", f6));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthBottom = 1;
                cel.Colspan = 5;
                tReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Cantidad\n", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BackgroundColor = grisOX;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = grisOX;
                tReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Unidad", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BackgroundColor = grisOX;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = grisOX;
                tReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Descripción", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BackgroundColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = grisOX;
                tReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Precio Unitario", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BackgroundColor = grisOX;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = grisOX;
                tReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Total", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BackgroundColor = grisOX;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = grisOX;
                tReceptor.AddCell(cel);

                #endregion
                
                #region "Construimos Tablas de Partidas"

                // Creamos la tabla para insertar los conceptos de detalle de la factura
                PdfPTable tableConceptos = new PdfPTable(5);

                int[] colWithsConceptos = new int[5];
                //String[] arrColWidthConceptos = dtConfigFact.Rows[0]["conceptosColWidth"].ToString().Split(new Char[] { ',' });
                String[] arrColWidthConceptos = { "15", "15", "40", "15", "15" };

                for (int i = 0; i < arrColWidthConceptos.Length; i++)
                {
                    colWithsConceptos.SetValue(Convert.ToInt32(arrColWidthConceptos[i]), i);
                }

                tableConceptos.SetWidths(colWithsConceptos);
                tableConceptos.WidthPercentage = 100F;

                int numConceptos = electronicDocument.Data.Conceptos.Count;
                PdfPCell cellConceptos = new PdfPCell();
                PdfPCell cellMontos = new PdfPCell();

                for (int i = 0; i < numConceptos; i++)
                {
                    cellConceptos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Cantidad.Value.ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellConceptos.Border = 0;
                    cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                    tableConceptos.AddCell(cellConceptos);

                    cellConceptos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Unidad.Value, new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellConceptos.Border = 0;
                    cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                    tableConceptos.AddCell(cellConceptos);

                    cellConceptos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value + "\nNo Identificación: " + dtOpcDet.Rows[i]["noIdent"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellConceptos.Border = 0;
                    tableConceptos.AddCell(cellConceptos);

                    cellMontos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].ValorUnitario.Value.ToString("C", _ci), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellMontos.Border = 0;
                    cellMontos.HorizontalAlignment = PdfCell.ALIGN_JUSTIFIED;
                    tableConceptos.AddCell(cellMontos);

                    cellMontos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("C", _ci), new Font(Font.HELVETICA, 8, Font.NORMAL)));
                    cellMontos.Border = 0;
                    cellMontos.HorizontalAlignment = PdfCell.ALIGN_JUSTIFIED;
                    tableConceptos.AddCell(cellMontos);
                }
                
                #endregion

                #region "Construimos el Comentarios"

                Table comentarios = new Table(4);
                float[] headerwidthsComentarios = { 25, 25, 35, 15 };
                comentarios.Widths = headerwidthsComentarios;
                comentarios.WidthPercentage = 100;
                comentarios.Padding = 1;
                comentarios.Spacing = 1;
                comentarios.BorderWidth = 0;
                comentarios.DefaultCellBorder = 0;
                comentarios.BorderColor = gris;

                int idMoneda = 1;
                DataTable dtImporteLetra = dal.QueryDT("DS_FE", "SELECT dbo.convertNumToTextFunction(@0, @1) AS cantidadLetra", "F:S:" + electronicDocument.Data.Total.Value.ToString() + ";F:I:" + idMoneda, hc);

                cel = new Cell(new Phrase("Importe con Letra:\n" + dtImporteLetra.Rows[0]["cantidadLetra"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                cel.Rowspan = 3;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Sub Total:", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Total Trasladados:", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("C", _ci), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Importe Total:", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _ci), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BackgroundColor = grisOX;
                cel.BorderColor = grisOX;
                comentarios.AddCell(cel);

                #endregion

                #region "Construimos el Desglose de Impuestos"

                Table desgloseImpuestos = new Table(3);
                float[] headerwidthsDesgloce = { 15, 15, 70 };
                desgloseImpuestos.Widths = headerwidthsDesgloce;
                desgloseImpuestos.WidthPercentage = 100;
                desgloseImpuestos.Padding = 1;
                desgloseImpuestos.Spacing = 1;
                desgloseImpuestos.BorderWidth = 0;
                desgloseImpuestos.DefaultCellBorder = 0;
                desgloseImpuestos.BorderColor = gris;

                cel = new Cell(new Phrase("Desgloce de Impuestos", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                desgloseImpuestos.AddCell(cel);

                cel = new Cell(new Phrase("", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                desgloseImpuestos.AddCell(cel);

                for (int i = 0; i < electronicDocument.Data.Impuestos.Traslados.Count; i++)
                {
                    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Traslados[i].Tipo.Value.ToString() + " " + electronicDocument.Data.Impuestos.Traslados[i].Tasa.Value.ToString() + "%", f6B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 1;
                    cel.BorderWidthRight = 1;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    desgloseImpuestos.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Traslados[i].Importe.Value.ToString("C", _ci), f6B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 1;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    desgloseImpuestos.AddCell(cel);

                    cel = new Cell(new Phrase("", f6B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    desgloseImpuestos.AddCell(cel);
                }

                for (int i = 0; i < electronicDocument.Data.Impuestos.Retenciones.Count; i++)
                {
                    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Retenciones[i].Tipo.Value.ToString(), f6B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 1;
                    cel.BorderWidthRight = 1;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    desgloseImpuestos.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Retenciones[i].Importe.Value.ToString("C", _ci), f6B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 1;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    desgloseImpuestos.AddCell(cel);

                    cel = new Cell(new Phrase("", f6B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    desgloseImpuestos.AddCell(cel);
                }

                #endregion

                #region "Construimos Tabla de Datos CFDI"

                DefaultSplitCharacter split = new DefaultSplitCharacter();
                Table adicional = new Table(3);
                float[] headerwidthsAdicional = { 20, 25, 55 };
                adicional.Widths = headerwidthsAdicional;
                adicional.WidthPercentage = 100;
                adicional.Padding = 1;
                adicional.Spacing = 1;
                adicional.BorderWidth = (float).5;
                adicional.DefaultCellBorder = 1;
                adicional.BorderColor = gris;

                if (timbrar)
                {
                    #region "Generamos Quick Response Code"

                    byte[] bytesQRCode = new byte[0];

                    if (timbrar)
                    {
                        // Generamos el Quick Response Code (QRCode)
                        string re = electronicDocument.Data.Emisor.Rfc.Value;
                        string rr = electronicDocument.Data.Receptor.Rfc.Value;
                        string tt = String.Format("{0:F6}", electronicDocument.Data.Total.Value);
                        string id = objTimbre.Uuid.Value;

                        StringBuilder sbCadenaQRCode = new StringBuilder();

                        sbCadenaQRCode.
                            Append("?").
                            Append("re=").Append(re).
                            Append("&").
                            Append("rr=").Append(rr).
                            Append("&").
                            Append("tt=").Append(tt).
                            Append("&").
                            Append("id=").Append(id);

                        BarcodeLib.Barcode.QRCode.QRCode barcode = new BarcodeLib.Barcode.QRCode.QRCode();

                        barcode.Data = sbCadenaQRCode.ToString();
                        barcode.ModuleSize = 3;
                        barcode.LeftMargin = 0;
                        barcode.RightMargin = 10;
                        barcode.TopMargin = 0;
                        barcode.BottomMargin = 0;
                        barcode.Encoding = BarcodeLib.Barcode.QRCode.QRCodeEncoding.Auto;
                        barcode.Version = BarcodeLib.Barcode.QRCode.QRCodeVersion.Auto;
                        barcode.ECL = BarcodeLib.Barcode.QRCode.ErrorCorrectionLevel.L;
                        bytesQRCode = barcode.drawBarcodeAsBytes();
                    }

                    #endregion

                    Image imageQRCode = Image.GetInstance(bytesQRCode);
                    imageQRCode.Alignment = (Image.TEXTWRAP | Image.ALIGN_LEFT);
                    imageQRCode.ScaleToFit(90f, 90f);
                    imageQRCode.IndentationLeft = 9f;
                    imageQRCode.SpacingAfter = 9f;
                    imageQRCode.BorderColorTop = Color.WHITE;

                    cel = new Cell(imageQRCode);
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 6;
                    adicional.AddCell(cel);

                    par = new Paragraph();
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("SELLO DIGITAL DEL EMISOR\n", f5B));
                    par.Add(new Chunk(electronicDocument.Data.Sello.Value, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
                    cel.Colspan = 2;
                    adicional.AddCell(cel);


                    cel = new Cell(new Phrase("FOLIO FISCAL:", f5B));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(objTimbre.Uuid.Value, f5));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase("FECHA Y HORA DE CERTIFICACION:", f5B));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    string[] fechaTimbrado = Convert.ToDateTime(objTimbre.FechaTimbrado.Value).GetDateTimeFormats('s');

                    cel = new Cell(new Phrase(fechaTimbrado[0], f5));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL SAT:", f5B));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(objTimbre.NumeroCertificadoSat.Value, f5));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL EMISOR:", f5B));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.Data.NumeroCertificado.Value, f5));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    adicional.AddCell(cel);

                    par = new Paragraph();
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("TIPO DE COMPROBANTE: ", f5B));
                    par.Add(new Chunk(electronicDocument.Data.TipoComprobante.Value + "   |   ", f5));
                    par.Add(new Chunk("Moneda: ", f5B));
                    par.Add(new Chunk(electronicDocument.Data.Moneda.Value + "   |   ", f5));
                    par.Add(new Chunk("TASA DE CAMBIO: ", f5B));
                    
                    string tasaCambio = electronicDocument.Data.TipoCambio.Value;
                    string regimenes = string.Empty;

                    if (tasaCambio.Length > 0)
                    {
                        par.Add(new Chunk(Convert.ToDouble(tasaCambio).ToString("C", _ci) + "   |   ", f5));
                    }
                    else
                    {
                        par.Add(new Chunk("   |   ", f5));
                    }
                    par.Add(new Chunk("FORMA DE PAGO: ", f5B));
                    par.Add(new Chunk(electronicDocument.Data.FormaPago.Value + "   |   ", f5));

                    par.Add(new Chunk("MÉTODO DE PAGO: ", f5B));
                    par.Add(new Chunk(electronicDocument.Data.MetodoPago.Value, f5));

                    if (electronicDocument.Data.NumeroCuentaPago.Value.ToString().Length > 0)
                    {
                        par.Add(new Chunk("   |   " + "No. CUENTA: ", f5B));
                        par.Add(new Chunk(electronicDocument.Data.NumeroCuentaPago.Value, f5));
                    }

                    if (electronicDocument.Data.Emisor.Regimenes.Count > 0)
                    {
                        for (int u = 0; u < electronicDocument.Data.Emisor.Regimenes.Count; u++)
                            regimenes += electronicDocument.Data.Emisor.Regimenes[u].Regimen.Value.ToString() + ",";

                        par.Add(new Chunk("   |   " + "RÉGIMEN FISCAL: ", f5B));
                        par.Add(new Chunk(regimenes.Substring(0, regimenes.Length - 1).ToString() + "   |   ", f5));
                    }

                    if (electronicDocument.Data.FolioFiscalOriginal.Value.ToString().Length > 0)
                    {
                        par.Add(new Chunk("\nDATOS CFDI ORIGINAL - SERIE: ", f5B));
                        par.Add(new Chunk(electronicDocument.Data.SerieFolioFiscalOriginal.Value, f5));
                        par.Add(new Chunk("   FOLIO: ", f5B));
                        par.Add(new Chunk(electronicDocument.Data.FolioFiscalOriginal.Value, f5));
                        par.Add(new Chunk("   FECHA: ", f5B));
                        par.Add(new Chunk(electronicDocument.Data.FechaFolioFiscalOriginal.Value.ToString(), f5));
                        par.Add(new Chunk("   MONTO: ", f5B));
                        par.Add(new Chunk(electronicDocument.Data.MontoFolioFiscalOriginal.Value.ToString(), f5));
                    }

                    cel.BorderColor = gris;
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 2;
                    cel.BorderColor = gris;
                    adicional.AddCell(cel);

                    if (electronicDocument.Data.LugarExpedicion.Value.Length > 0)
                    {
                        par = new Paragraph();
                        par.SetLeading(7f, 0f);
                        par.Add(new Chunk("LUGAR EXPEDICIÓN: ", f5B));
                        par.Add(new Chunk(electronicDocument.Data.LugarExpedicion.Value, f5));
                        cel = new Cell(par);
                        cel.BorderColor = gris;
                        cel.BorderWidthTop = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        cel.Colspan = 3;
                        adicional.AddCell(cel);
                    }
                    
                    par = new Paragraph();
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("CADENA ORIGINAL DEL COMPLEMENTO DE CERTIFICACIÓN DIGITAL DEL SAT\n", f5B));
                    par.Add(new Chunk(electronicDocument.FingerPrintPac, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("SELLO DIGITAL DEL SAT\n", f5B));
                    par.Add(new Chunk(objTimbre.SelloSat.Value, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);
                }
                #endregion

                #region "Construimos Tabla del Footer"

                PdfPTable footer = new PdfPTable(1);
                footer.WidthPercentage = 100;
                footer.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;

                cell = new PdfPCell(new Phrase("", f5));
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                footer.AddCell(cell);

                cell = new PdfPCell(new Phrase("ESTE DOCUMENTO ES UNA REPRESENTACIÓN IMPRESA DE UN CFDI", f6B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = grisOX;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                footer.AddCell(cell);

                #endregion

                pageEventHandler.encabezado = encabezado;
                pageEventHandler.encReceptor = tReceptor;
                pageEventHandler.footer = footer;

                document.Open();
                
                //document.Add(tReceptor);
                document.Add(tableConceptos);
                document.Add(comentarios);
                document.Add(desgloseImpuestos);
                document.Add(adicional);
                

                #endregion
            }
            catch (Exception ex)
            {
                string exe = ex.Message;
            }
        }

        #endregion

    }

    public class pdfPageEventHandlerCED110324NN4 : PdfPageEventHelper
    {
        #region "variables"

        //Contentbyte del objeto writer
        PdfContentByte cb;
        // Pone el numero de pagina al final en el template
        PdfTemplate template;
        BaseFont bf = null;
        DateTime PrintTime = DateTime.Now;

        #endregion

        #region Propiedades

        public string piePaginaIdioma { get; set; }
        public string rutaImgFooter { get; set; }
        public Font FooterFont { get; set; }
        public Table encabezado { get; set; }
        public Table encReceptor { get; set; }
        public PdfPTable detalle { get; set; }
        public PdfPTable footer { get; set; }
        public Table adicional { get; set; }

        #endregion

        #region "Pdf Eventos de Página"

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            try
            {
                PrintTime = DateTime.Now;
                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb = writer.DirectContent;
                template = cb.CreateTemplate(50, 50);
            }
            catch (DocumentException)
            {
            }
            catch (IOException)
            {
            }
        }

        public override void OnStartPage(PdfWriter writer, Document document)
        {
            base.OnStartPage(writer, document);
            document.Add(encabezado);
            document.Add(encReceptor);

            cb.EndText();
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);
            footer.WriteSelectedRows(0, -1, 15, (document.BottomMargin - 0), writer.DirectContent);

            string lblPagina;
            string lblDe;
            string lblFechaImpresion;

            switch (piePaginaIdioma)
            {
                case "S":
                    lblPagina = "Página ";
                    lblDe = " de ";
                    lblFechaImpresion = "Fecha de Impresión ";
                    break;

                case "E":
                    lblPagina = "Page ";
                    lblDe = " of ";
                    lblFechaImpresion = "Printed Date ";
                    break;

                default:
                    lblPagina = "Página ";
                    lblDe = " de ";
                    lblFechaImpresion = "Fecha de Impresión ";
                    break;
            }

            int pageN = writer.PageNumber;
            String text = lblPagina + pageN + lblDe;
            float len = bf.GetWidthPoint(text, 8);

            Rectangle pageSize = document.PageSize;

            cb.SetRGBColorFill(100, 100, 100);

            cb.BeginText();
            cb.SetFontAndSize(bf, 8);
            cb.SetTextMatrix(pageSize.GetLeft(30), pageSize.GetBottom(15));
            cb.ShowText(text);
            cb.EndText();

            cb.AddTemplate(template, pageSize.GetLeft(30) + len, pageSize.GetBottom(15));

            cb.BeginText();
            cb.SetFontAndSize(bf, 8);
            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, lblFechaImpresion + PrintTime, pageSize.GetRight(30), pageSize.GetBottom(15), 0);
            cb.EndText();
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);

            template.BeginText();
            template.SetFontAndSize(bf, 8);
            template.SetTextMatrix(0, 0);
            template.ShowText("" + (writer.PageNumber - 1));
            template.EndText();
        }

        #endregion
    }

    public class DefaultSplitCharacterCED110324NN4 : ISplitCharacter
    {
        #region "ISplitCharacter"

        /**
         * An instance of the default SplitCharacter.
         */
        public static readonly ISplitCharacter DEFAULT = new DefaultSplitCharacterPfizer();

        /**
         * Checks if a character can be used to split a <CODE>PdfString</CODE>.
         * <P>
         * for the moment every character less than or equal to SPACE, the character '-'
         * and some specific unicode ranges are 'splitCharacters'.
         * 
         * @param start start position in the array
         * @param current current position in the array
         * @param end end position in the array
         * @param	cc		the character array that has to be checked
         * @param ck chunk array
         * @return	<CODE>true</CODE> if the character can be used to split a string, <CODE>false</CODE> otherwise
         */
        public bool IsSplitCharacter(int start, int current, int end, char[] cc, PdfChunk[] ck)
        {
            char c = GetCurrentCharacter(current, cc, ck);
            if (c <= ' ' || c == '\u2010')
            {
                return true;
            }
            if (c < 0x2002)
                return false;
            return ((c >= 0x2002 && c <= 0x200b)
                || (c >= 0x2e80 && c < 0xd7a0)
                || (c >= 0xf900 && c < 0xfb00)
                || (c >= 0xfe30 && c < 0xfe50)
                || (c >= 0xff61 && c < 0xffa0));
        }

        /**
         * Returns the current character
         * @param current current position in the array
         * @param	cc		the character array that has to be checked
         * @param ck chunk array
         * @return	the current character
         */
        protected char GetCurrentCharacter(int current, char[] cc, PdfChunk[] ck)
        {
            if (ck == null)
            {
                return (char)cc[current];
            }
            return (char)ck[Math.Min(current, ck.Length - 1)].GetUnicodeEquivalent(cc[current]);
        }

        #endregion
    }

}