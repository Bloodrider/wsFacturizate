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
    public class OfficeMax
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
        private static string pathPdf;

        private static HttpContext HTC;
        private static String pathIMGLOGO;
        private static String pathCedula;
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
        private static Font f4;
        private static Font f5;
        private static Font f5B;
        private static Font f5BBI;
        private static Font f6;
        private static Font f6B;
        private static Font f6L;
        private static Font titulo;
        private static Font f5Bblanco;
        private static Font folio;
        private static Font f5L;
        public static int numeroPagina;
        public static int numeroPaginasDescontar;
        public static int numeroPaginasDescontarAlm;

        #endregion

        #region "Formato Factura CFD"

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
                long idCfdi = Convert.ToInt64(htFacturaxion["idCfdi"]);

                #region "Obtenemos los datos del CFDI y Campos Opcionales"

                StringBuilder sbOpcionalEncabezado = new StringBuilder();
                DataTable dtOpcEnc = new DataTable();
                StringBuilder sbOpcionalDetalle = new StringBuilder();
                DataTable dtOpcDet = new DataTable();
                StringBuilder sbOpcionalDetalleRem = new StringBuilder();
                DataTable dtOpcDetRem = new DataTable();
                StringBuilder sbOpcionalDetalleImp = new StringBuilder();
                DataTable dtOpcDetImp = new DataTable();

                sbOpcionalEncabezado.
                    Append("SELECT ").
                    Append("campo16 AS nombreEnviarA, ").
                    Append("campo17 + ' ' + campo18 AS calleEnviarA, ").
                    Append("campo20 + ', ' + campo22 AS delMunEnviarA, ").
                    Append("campo23 + ', ' + campo24  AS edoEnviarA, ").
                    Append("'C.P.' + campo25 AS CPEnviarA, ").
                    Append("campo1 AS tipoDoc,  ").
                    Append("campo15 AS noCliente, ").
                    Append("campo14 AS gln, ").
                    Append("campo28 AS ordenCompra, ").
                    Append("campo29 AS noOrden, ").
                    Append("campo30 AS fechaOrden, ").
                    Append("campo33 AS plazo, ").
                    Append("campo32 AS folioRef, ").
                    Append("campo28 AS recibo, ").
                    Append("campo30 AS fechaRecibo, ").
                    Append("campo31 AS fechaExp, ").
                    Append("campo33 AS proyecto, ").
                    Append("campo34 AS comentarioFinDetalle, ").
                    Append("campo26 AS cantidadLetra ").
                    Append("FROM opcionalEncabezado ").
                    Append("WHERE idCFDI = @0  AND ST = 1 ");

                sbOpcionalDetalle.
                    Append("SELECT ROW_NUMBER() OVER (ORDER BY idCfdi ASC) AS numero, ").
                    Append("campo8 AS codigo, campo9 AS UPC, ").
                    Append("campo10 AS embarque, campo11 AS subtotal, campo13 AS C1, campo12 AS descuento, ").
                    Append("campo14 AS agenteAduanal ").
                    Append("FROM opcionalDetalle ").
                    Append("WHERE idCfdi = @0 ");

                sbOpcionalDetalleRem.
                    Append("SELECT campo1 AS remision, campo2 AS orden, campo3 AS codigo, campo4 AS UPC, campo5 AS descripcion, campo6 AS cantidad ").
                    Append("FROM opcionalDetalle2 ").
                    Append("WHERE idCfdi = @0 AND concepto = 1 AND ST = 1");

                sbOpcionalDetalleImp.
                    Append("SELECT campo1 AS impuesto, campo2 AS tasa, campo3 AS importe, campo4 AS baseImp, campo5 AS cl ").
                    Append("FROM opcionalDetalle2 ").
                    Append("WHERE idCfdi = @0 AND concepto = 2 AND ST = 1");

                dtOpcEnc = dal.QueryDT("DS_FE", sbOpcionalEncabezado.ToString(), "F:I:" + idCfdi, hc);
                dtOpcDet = dal.QueryDT("DS_FE", sbOpcionalDetalle.ToString(), "F:I:" + idCfdi, hc);
                dtOpcDetRem = dal.QueryDT("DS_FE", sbOpcionalDetalleRem.ToString(), "F:I:" + idCfdi, hc);
                dtOpcDetImp = dal.QueryDT("DS_FE", sbOpcionalDetalleImp.ToString(), "F:I:" + idCfdi, hc);

                if (dtOpcDet.Rows.Count == 0)
                {
                    for (int i = 1; i <= electronicDocument.Data.Conceptos.Count; i++)
                    {
                        dtOpcDet.Rows.Add("", "0.00");
                    }
                }

                #endregion

                #region "Extraemos los datos del CFDI"

                //Datos CFD
                Hashtable htDatosCfdi = new Hashtable();
                htDatosCfdi.Add("nombreEmisor", electronicDocument.Data.Emisor.Nombre.Value);
                htDatosCfdi.Add("rfcEmisor", electronicDocument.Data.Emisor.Rfc.Value);
                htDatosCfdi.Add("nombreReceptor", electronicDocument.Data.Receptor.Nombre.Value);
                htDatosCfdi.Add("rfcReceptor", electronicDocument.Data.Receptor.Rfc.Value);
                htDatosCfdi.Add("serie", electronicDocument.Data.Serie.Value);
                htDatosCfdi.Add("folio", electronicDocument.Data.Folio.Value);
                htDatosCfdi.Add("fechaCfdi", electronicDocument.Data.Fecha.Value);
                //htDatosCfdi.Add("UUID", objTimbre.Uuid.Value);

                //Datos CFD
                htDatosCfdi.Add("anioAprobacion", electronicDocument.Data.AnioAprobacion.Value);
                htDatosCfdi.Add("numeroAprobacion", electronicDocument.Data.NumeroAprobacion.Value);

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

                Document document = new Document(PageSize.LETTER, 15, 15, 15, 40);
                document.AddAuthor("Facturaxion");
                document.AddCreator("r3Take");
                document.AddCreationDate();

                //FileStream fs = new FileStream(pathPdf, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                pdfPageEventHandlerOfficeMax pageEventHandler = new pdfPageEventHandlerOfficeMax();
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                writer.PageEvent = pageEventHandler;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                //document.Open();

                HTC = hc;
                pathIMGLOGO = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\OFFICEMAX\Officemax-04.png";
                pathCedula = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\OFFICEMAX\Officemax-01.png";

                azul = new Color(22, 111, 168);
                blanco = new Color(255, 255, 255);
                Link = new Color(7, 73, 208);
                gris = new Color(236, 236, 236);
                grisOX = new Color(220, 215, 220);
                rojo = new Color(230, 7, 7);
                lbAzul = new Color(43, 145, 175);

                EM = BaseFont.CreateFont(@"C:\Windows\Fonts\VERDANA.TTF", BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
                f4 = new Font(EM, 4);
                f5 = new Font(EM, 5);
                f5B = new Font(EM, 5, Font.BOLD);
                f5BBI = new Font(EM, 5, Font.BOLDITALIC);
                f6 = new Font(EM, 6);
                f6B = new Font(EM, 6, Font.BOLD);
                f6L = new Font(EM, 6, Font.BOLD, Link);
                f5L = new Font(EM, 5, Font.BOLD, lbAzul);
                titulo = new Font(EM, 6, Font.BOLD, blanco);
                f5Bblanco = new Font(EM, 5, Font.BOLD, blanco);
                folio = new Font(EM, 6, Font.BOLD, rojo);
                dSaltoLinea = new Chunk("\n\n ");

                #endregion

                #region "Generamos el Documento"

                formatoFactura(document, electronicDocument, pageEventHandler, idCfdi, dtOpcEnc, dtOpcDet, dtOpcDetRem, dtOpcDetImp, htDatosCfdi, HTC);

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

        #region "formatoFactura"

        public static void formatoFactura(Document document, ElectronicDocument electronicDocument, pdfPageEventHandlerOfficeMax pageEventHandler, long idCfdi, DataTable dtEncabezado, DataTable dtDetalle, DataTable dtOpcDetRem, DataTable dtOpcDetImp, Hashtable htCFDI, HttpContext hc)
        {
            try
            {
                DAL dal = new DAL();
                #region "Construimos el Documento"

                for (int t = 1; t < 5; t++)
                {

                    #region "Construimos el Encabezado"

                    Table encabezado = new Table(4);
                    float[] headerwidthsEncabezado = { 28, 32, 18, 22 };
                    encabezado.Widths = headerwidthsEncabezado;
                    encabezado.WidthPercentage = 100;
                    encabezado.Padding = 1;
                    encabezado.Spacing = 1;
                    encabezado.BorderWidth = 0;
                    encabezado.DefaultCellBorder = 0;
                    encabezado.BorderColor = gris;

                    //Agregando Imagen de Logotipo
                    Image imgLogo = Image.GetInstance(pathIMGLOGO);
                    imgLogo.ScalePercent(62f);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(1f, 1f);
                    par.Add(new Chunk(imgLogo, 0, 0));
                    par.Add(new Chunk("", f6));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 3;
                    encabezado.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("VENDIDO A:", f6B));
                    par.Add(new Chunk("\n" + htCFDI["nombreReceptor"].ToString().ToUpper(), f6B));
                    par.Add(new Chunk("\n" + htCFDI["direccionReceptor1"].ToString().ToUpper(), f6));
                    par.Add(new Chunk("\n" + htCFDI["direccionReceptor2"].ToString().ToUpper(), f6));
                    par.Add(new Chunk("\n" + htCFDI["direccionReceptor3"].ToString().ToUpper(), f6));
                    par.Add(new Chunk("\nR.F.C.: " + htCFDI["rfcReceptor"].ToString().ToUpper(), f6B));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 3;
                    encabezado.AddCell(cel);

                    //Agregando Imagen de CedulaFiscal
                    Image imgCedula = Image.GetInstance(pathCedula);
                    imgCedula.ScalePercent(47f);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(1f, 1f);
                    par.Add(new Chunk(imgCedula, 0, 0));
                    par.Add(new Chunk("", f6));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 5;
                    encabezado.AddCell(cel);

                    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["tipoDoc"].ToString().ToUpper(), f6B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    encabezado.AddCell(cel);

                    cel = new Cell(new Phrase(htCFDI["serie"].ToString().ToUpper() + " " + electronicDocument.Data.Folio.Value, folio));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    encabezado.AddCell(cel);

                    cel = new Cell(new Phrase("", f6B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    encabezado.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 1f);
                    par.Add(new Chunk(htCFDI["nombreEmisor"].ToString().ToUpper() + "\n\n", f6B));
                    par.Add(new Chunk(htCFDI["direccionEmisor1"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(htCFDI["direccionEmisor2"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(htCFDI["direccionEmisor3"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk("R.F.C.: " + htCFDI["rfcEmisor"].ToString().ToUpper() + "\n\n", f6B));
                    par.Add(new Chunk("Teléfono: 91772800", f6));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 2;
                    encabezado.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 1f);
                    par.Add(new Chunk("ENVIAR A: \n", f6B));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["nombreEnviarA"].ToString().ToUpper() + "\n", f6B));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["calleEnviarA"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["delMunEnviarA"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["edoEnviarA"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["CPEnviarA"].ToString().ToUpper(), f6));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 2;
                    encabezado.AddCell(cel);

                    //Pagina n de n

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 1f);
                    par.Add(new Chunk(dtEncabezado.Rows[0]["gln"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(electronicDocument.Data.Emisor.ExpedidoEn.Calle.Value.ToString().ToUpper() + " " + electronicDocument.Data.Emisor.ExpedidoEn.NumeroExterior.ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(electronicDocument.Data.Emisor.ExpedidoEn.Colonia.Value.ToString().ToUpper() + " " + electronicDocument.Data.Emisor.ExpedidoEn.Municipio.Value.ToString().ToUpper() + " " + electronicDocument.Data.Emisor.ExpedidoEn.Estado.Value.ToString().ToUpper() + " " + electronicDocument.Data.Emisor.ExpedidoEn.Pais.Value.ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk("C.P. " + electronicDocument.Data.Emisor.ExpedidoEn.CodigoPostal.Value.ToString().ToUpper(), f6));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 2;
                    encabezado.AddCell(cel);

                    #endregion

                    #region "tabla Datos Adicionales"

                    Table datosAdic = new Table(7);
                    float[] headerWidthsDatosAdic = { 13, 13, 13, 13, 22, 13, 13 };
                    datosAdic.Widths = headerWidthsDatosAdic;
                    datosAdic.WidthPercentage = 100;
                    datosAdic.Padding = 1;
                    datosAdic.Spacing = 1;
                    datosAdic.BorderWidth = (float).5;
                    datosAdic.DefaultCellBorder = 0;
                    datosAdic.BorderColor = gris;

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("No Cliente:", f5));
                    par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["noCliente"].ToString(), f5B));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);

                    if (electronicDocument.Data.Serie.Value.ToString() == "VTASIMPORT")
                    {
                        par.Add(new Chunk("Recibo:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["recibo"].ToString(), f5B));
                    }

                    else
                    {
                        par.Add(new Chunk("O.Compra Cte.:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["ordenCompra"].ToString(), f5B));
                    }

                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("No.orden / REA:", f5));
                    par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["noOrden"].ToString(), f5B));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);

                    if (electronicDocument.Data.Serie.Value.ToString() == "VTASIMPORT")
                    {
                        par.Add(new Chunk("fecha Recibo:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["fechaRecibo"].ToString(), f5B));
                    }

                    else
                    {
                        par.Add(new Chunk("fecha Orden:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["fechaOrden"].ToString(), f5B));
                    }

                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("Lugar y fecha de Expedición:", f5));
                    par.Add(new Chunk("\n" + electronicDocument.Data.Emisor.ExpedidoEn.Calle.Value.ToString().ToUpper() + " " + electronicDocument.Data.Emisor.ExpedidoEn.NumeroExterior.Value.ToString().ToUpper() + " " + dtEncabezado.Rows[0]["fechaExp"].ToString(), f5B));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);

                    if (electronicDocument.Data.Serie.Value.ToString() == "VTASIMPORT")
                    {
                        par.Add(new Chunk("Proyecto:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["proyecto"].ToString(), f5B));
                    }

                    else
                    {
                        par.Add(new Chunk("Plazo:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["plazo"].ToString(), f5B));
                    }

                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("Folio Referencia:", f5));
                    par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["folioRef"].ToString(), f5B));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    #endregion

                    #region"Tabla Detalle"

                    Table encabezadoDetalle = new Table(12);
                    float[] headerEncabezadoDetalle = { 5, 7, 10, 20, 7, 5, 7, 8, 8, 8, 10, 5 };
                    encabezadoDetalle.Widths = headerEncabezadoDetalle;
                    encabezadoDetalle.WidthPercentage = 100F;
                    encabezadoDetalle.Padding = 1;
                    encabezadoDetalle.Spacing = 1;
                    encabezadoDetalle.BorderWidth = (float).5;
                    encabezadoDetalle.DefaultCellBorder = 0;
                    encabezadoDetalle.BorderColor = gris;

                    // NUMERO
                    cel = new Cell(new Phrase("NO.", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // CODIGO
                    cel = new Cell(new Phrase("CÓDIGO", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // UPC
                    cel = new Cell(new Phrase("UPC", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // DESCRIPCIÓN
                    cel = new Cell(new Phrase("DESCRIPCIÓN", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // CANTIDAD SOL
                    cel = new Cell(new Phrase("CANT. SOL.", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // UNIDAD DE MEDIDA
                    cel = new Cell(new Phrase("U/M", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // EMBARQUE
                    cel = new Cell(new Phrase("EMBAR.", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // PRECIO UNITARIO
                    cel = new Cell(new Phrase("PRECIO UNIT.", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // SUBTOTAL
                    cel = new Cell(new Phrase("SUBTOTAL", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // DESCUENTO
                    cel = new Cell(new Phrase("DESCUENTO", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // TOTAL
                    cel = new Cell(new Phrase("TOTAL", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // C1
                    cel = new Cell(new Phrase("C1.", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    PdfPTable tableConceptos = new PdfPTable(12);
                    tableConceptos.SetWidths(new int[12] { 5, 7, 10, 20, 7, 5, 7, 8, 8, 8, 10, 5 });
                    tableConceptos.WidthPercentage = 100F;

                    int numConceptos = electronicDocument.Data.Conceptos.Count;
                    PdfPCell cellConceptos = new PdfPCell();
                    PdfPCell cellMontos = new PdfPCell();

                    for (int i = 0; i < numConceptos; i++)
                    {
                        Concepto concepto = electronicDocument.Data.Conceptos[i];

                        cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["numero"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["codigo"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["UPC"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(concepto.Descripcion.Value.ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(concepto.Cantidad.Value.ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(concepto.Unidad.Value.ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["embarque"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tableConceptos.AddCell(cellConceptos);

                        cellMontos = new PdfPCell(new Phrase(concepto.ValorUnitario.Value.ToString("C", _ci), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellMontos.Border = 0;
                        cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                        tableConceptos.AddCell(cellMontos);

                        cellMontos = new PdfPCell(new Phrase("$" + dtDetalle.Rows[i]["subtotal"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellMontos.Border = 0;
                        cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                        tableConceptos.AddCell(cellMontos);

                        cellMontos = new PdfPCell(new Phrase("$" + dtDetalle.Rows[i]["descuento"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellMontos.Border = 0;
                        cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                        tableConceptos.AddCell(cellMontos);

                        cellMontos = new PdfPCell(new Phrase(concepto.Importe.Value.ToString("C", _ci), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellMontos.Border = 0;
                        cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                        tableConceptos.AddCell(cellMontos);

                        cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["C1"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                        tableConceptos.AddCell(cellConceptos);

                        if (concepto.InformacionAduanera.IsAssigned)
                        {
                            for (int j = 0; j < concepto.InformacionAduanera.Count; j++)
                            {
                                cellConceptos = new PdfPCell(new Phrase("Aduana", new Font(Font.HELVETICA, 5, Font.BOLD)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 2;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase("Fecha Pedimento", new Font(Font.HELVETICA, 5, Font.BOLD)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 1;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase("Pedimento", new Font(Font.HELVETICA, 5, Font.BOLD)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 1;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase("Agente Aduanal", new Font(Font.HELVETICA, 5, Font.BOLD)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 8;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase(concepto.InformacionAduanera[j].Aduana.Value.ToString(), new Font(Font.HELVETICA, 5, Font.NORMAL)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 2;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase(concepto.InformacionAduanera[j].Fecha.Value.ToString().Substring(0,9), new Font(Font.HELVETICA, 5, Font.NORMAL)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 1;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase(concepto.InformacionAduanera[j].Numero.Value.ToString(), new Font(Font.HELVETICA, 5, Font.NORMAL)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 1;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["agenteAduanal"].ToString(), new Font(Font.HELVETICA, 5, Font.NORMAL)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 8;
                                tableConceptos.AddCell(cellConceptos);

                            }
                        }
                    }

                    #endregion

                    #region "Construimos Tabla Especial para dar espacio al detalle del cfdi y el desglose de impuestos"
                    Table especial1 = new Table(7);
                    float[] headerwidthsEspecial1 = { 10, 10, 10, 12, 18, 10, 30 };
                    especial1.Widths = headerwidthsEspecial1;
                    especial1.WidthPercentage = 100;
                    especial1.Padding = 1;
                    especial1.Spacing = 1;
                    especial1.BorderWidth = 0;
                    especial1.DefaultCellBorder = 0;
                    especial1.BorderColor = blanco;

                    if (dtOpcDetRem.Rows.Count > 0)
                    {
                        cel = new Cell(new Phrase("\n\n.", titulo));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.Colspan = 7;
                        especial1.AddCell(cel);

                        //Hacemos encabezado de tabla de remision
                        cel = new Cell(new Phrase("Remisión", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("Orden", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("Código", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("UPC", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("Descripción", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("Cantidad", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("", f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        for (int i = 0; i < dtOpcDetRem.Rows.Count; i++)
                        {
                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["remision"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["orden"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["codigo"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["UPC"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["descripcion"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["cantidad"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase("", f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);
                        }
                    }

                    cel = new Cell(new Phrase("\n\n\n.", titulo));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = blanco;
                    cel.Colspan = 7;
                    especial1.AddCell(cel);

                    #endregion

                    #region "Construimos el Comentarios"

                    Table comentarios = new Table(7);
                    float[] headerwidthsComentarios = { 11, 52, 6, 8, 6, 7, 10 };
                    comentarios.Widths = headerwidthsComentarios;
                    comentarios.WidthPercentage = 100;
                    comentarios.Padding = 1;
                    comentarios.Spacing = 1;
                    comentarios.BorderWidth = (float).5;
                    comentarios.DefaultCellBorder = 0;
                    comentarios.BorderColor = gris;

                    if (dtEncabezado.Rows[0]["comentarioFinDetalle"].ToString().Length > 0)
                    {
                        cel = new Cell(new Phrase(dtEncabezado.Rows[0]["comentarioFinDetalle"].ToString().ToUpper(), f5));
                        cel.VerticalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        cel.Colspan = 7;
                        comentarios.AddCell(cel);
                    }

                    cel = new Cell(new Phrase("CADENA ORIGINAL:", f5B));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.Colspan = 7;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.FingerPrint.ToString() + "\n\n\n\n\n\n\n", f4));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    cel.Colspan = 7;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("Cantidad con letra:", f6B));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["cantidadLetra"].ToString(), f6));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("Clase", f6));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("Base Cálculo", f6));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("Tasa", f6));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("Sub Total", f6B));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f6));
                    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    #region "Codigo Alterno Impuestos"

                    //for (int i = 0; i < electronicDocument.Data.Impuestos.Traslados.Count; i++)
                    //{
                    //    cel = new Cell(new Phrase("", f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    cel.Colspan = 2;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["clase"].ToString(), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Traslados[i].Tasa.Value.ToString(), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Traslados[i].Tipo.Value.ToString(), f6B));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = (float).5;
                    //    cel.BorderWidthRight = (float).5;
                    //    cel.BorderWidthBottom = (float).5;
                    //    cel.BorderColor = gris;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Traslados[i].Importe.Value.ToString("C", _ci), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = (float).5;
                    //    cel.BorderWidthBottom = (float).5;
                    //    cel.BorderColor = gris;
                    //    comentarios.AddCell(cel);

                    //}

                    //for (int i = 0; i < electronicDocument.Data.Impuestos.Retenciones.Count; i++)
                    //{
                    //    cel = new Cell(new Phrase("", f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    cel.Colspan = 2;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["clase"].ToString(), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase("", f6B));
                    //    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Retenciones[i].Tipo.Value.ToString(), f6B));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = (float).5;
                    //    cel.BorderWidthRight = (float).5;
                    //    cel.BorderWidthBottom = (float).5;
                    //    cel.BorderColor = gris;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Retenciones[i].Importe.Value.ToString("C", _ci), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = (float).5;
                    //    cel.BorderWidthBottom = (float).5;
                    //    cel.BorderColor = gris;
                    //    comentarios.AddCell(cel);
                    //}
                    #endregion

                    for (int y = 0; y < dtOpcDetImp.Rows.Count; y++)
                    {
                        cel = new Cell(new Phrase("", f6));
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = blanco;
                        cel.Colspan = 2;
                        comentarios.AddCell(cel);

                        cel = new Cell(new Phrase(dtOpcDetImp.Rows[0]["cl"].ToString(), f6));
                        cel.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = blanco;
                        comentarios.AddCell(cel);

                        cel = new Cell(new Phrase(dtOpcDetImp.Rows[0]["baseImp"].ToString(), f6));
                        cel.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = blanco;
                        comentarios.AddCell(cel);

                        cel = new Cell(new Phrase(dtOpcDetImp.Rows[0]["tasa"].ToString(), f6));
                        cel.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = blanco;
                        comentarios.AddCell(cel);

                        cel = new Cell(new Phrase(dtOpcDetImp.Rows[0]["impuesto"].ToString(), f6B));
                        cel.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = (float).5;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        comentarios.AddCell(cel);

                        cel = new Cell(new Phrase(dtOpcDetImp.Rows[0]["importe"].ToString(), f6));
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = (float).5;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        comentarios.AddCell(cel);
                    }

                    cel = new Cell(new Phrase("", f6));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = blanco;
                    cel.Colspan = 5;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("TOTAL", f6B));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _ci), f6));
                    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    #endregion

                    #region "Construimos Tabla Especial"

                    Table especial = new Table(2);
                    float[] headerwidthsEspecial = { 65, 35 };
                    especial.Widths = headerwidthsEspecial;
                    especial.WidthPercentage = 100;
                    especial.Padding = 1;
                    especial.Spacing = 1;
                    especial.BorderWidth = 0;
                    especial.DefaultCellBorder = 0;
                    especial.BorderColor = gris;

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("\nSELLO DIGITAL:", f5B));
                    par.Add(new Chunk("\n" + electronicDocument.Data.Sello.Value.ToString(), f5));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    especial.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("\nNo. de Aprobación:  ", f5B));
                    par.Add(new Chunk(electronicDocument.Data.NumeroAprobacion.Value.ToString(), f5));
                    par.Add(new Chunk("\nNo. de Certificado:  ", f5B));
                    par.Add(new Chunk(electronicDocument.Data.NumeroCertificado.Value.ToString(), f5));
                    par.Add(new Chunk("\nAño de Certificado:  ", f5B));
                    par.Add(new Chunk(electronicDocument.Data.AnioAprobacion.Value.ToString(), f5));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    especial.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("\n\n\n\nMéxico Distrito Federal a Debo(emos) y pagaré(mos) incondicionalmente a la orden de OPERADORA OMX S.A. DE C.V. en México D.F. ", f5B));
                    par.Add(new Chunk("p en cualquier otra que se me requiera el pago, la cantidad consignada en el presente titulo misma que ha sido recibida a mi entera ", f5B));
                    par.Add(new Chunk("satisfacción en caso de incumplimiento del presente pagaré, me obligo a pagar un interés monetario mensual por el equivalente a aplicar la ", f5B));
                    par.Add(new Chunk("tasa lider publicada por el Banco de México, en el mes inmediato anterior atendiendose por tasa lider, la mayor entre las siguientes: Tasa de interés ", f5B));
                    par.Add(new Chunk("interbancario de equilibrio, tasa de interés interbancario promedio, el costo porcentual promedio de captación.", f5B));

                    par.Add(new Chunk("EL PAGO DE ESTA FACTURA SE HARÁ EN UNA SOLA EXHIBICIÓN RECIBI MERCANCIA A MI ENTERA SATISFACCIÓN PUESTO\n", f5B));
                    par.Add(new Chunk("FIRMA___________________________________________________________________________________________________\n", f5B));
                    par.Add(new Chunk("PARA EFECTOS FISCALES AL PAGO \"LA REPRODUCCIÓN NO AUTORIZADA DE ESTE COMPROBANTE CONSTITUYE UN  DELITO EN LOS TERMINOS DE LAS DISPOSICIONES FISCALES\" \n", f5B));
                    par.Add(new Chunk("ACEPTO", f5B));

                    par.Add(new Chunk("\n\nPolíticas de Cambios Físicos y Devoluciones de Mercancía \n", f4));
                    par.Add(new Chunk("Comunicandose al área de Servicios al Cliente al teléfono 5321-9921 y lada sin costo (01800) 713-7070. Es importante que al solicitar algún cambio físico, devolución parcial o total presente factura original o ticket, además de: \n", f4));
                    par.Add(new Chunk(". El artículo deberá tener su empaque original con todos sus accesorios;así como la Factura Original y/o Ticket. \n", f4));
                    par.Add(new Chunk(". Artículos eléctricos deberá realizarse dentro de los 7 días naturales siguientes a la fecha de entrega, transcurrido el plazo se deberá acudir a los centros de servicio autorizados por el fabricante. El cambio físico solo será por defecto de fabricación. \n", f4));
                    par.Add(new Chunk(". En cartuchos, toners, película tpermica, software y plumas finas, productos de exhibición, demostración, liquidación, repartos, tarjetas telefónicas, tarjetas internet, cd no hay devoluciones ni cambios físicos. Cualquier reclamación de estos productos será directamente con el fabricante.", f4));
                    par.Add(new Chunk(". En muebles, el plazo para reaizar cualwuier reclamación será de 30 días naturales contados a partir de la fecha de entrega. Solicite el armado por el personal certificado de OFFICEMAX llamando al 5321-9921. No se aceptará cambio físico o devolución si el mueble fue armado por personal no autorizado por la empresa. \n", f4));
                    par.Add(new Chunk(". En artículos de oficina el plazo para realizar cualquier reclamción será de 15 días naturales contados a partir de la fecha de entrega. \n", f4));
                    par.Add(new Chunk(". Para ventas corporativas las políticas están establecidas en contrato. (confirmar con su vendedor). \n", f4));
                    par.Add(new Chunk(". Tratándose de mercancía adquirida con tarjeta de crédito bajo el esquema de meses sin intereses, aplicará el cambio físico de mercancía (según garantías antes mencionadas). La devolución del artículo podrá efectuarse previa autorización del cliente para realizar el cargo de los gastos administrativos incurridos a la tarjeta de crédito con la que se realizó la compra. \n", f4));
                    par.Add(new Chunk(". Las devoluciones se realizarán a la forma de pago original en la que se realizó la compra. En el caso de que la compra se efectivo en cheque o efectivo la devolución procederá hasta por un monto máximo de $2,000.00 (Dos mil pesos 00/100 m.n.) en efectivo. Si el importe es mayor a esta cantidad se rembolsará a través de un cheque nominativo. (En casos especiales se aplicará vale de mercancía). \n", f4));
                    par.Add(new Chunk(". Officemax se deslinda de la responsabilidad total o parcial por pérdida o daño de mercancía una vez que se entregó al cliente y se firmó de acuse de recibo bajo su entera conformidad y satisfaciión por parte del cliente.", f4));

                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 2;
                    cel.BorderColor = gris;
                    especial.AddCell(cel);

                    if (t == 1)
                        cel = new Cell(new Phrase("ORIGINAL", f6B));

                    else if (t == 2)
                        cel = new Cell(new Phrase("CLIENTE", f6B));

                    else if (t == 3)
                        cel = new Cell(new Phrase("CRÉDITO Y COBRANZA", f6B));

                    else
                        cel = new Cell(new Phrase("CONSECUTIVO FISCAL", f6B));

                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.Colspan = 2;
                    especial.AddCell(cel);

                    #endregion

                    #region "Cuando migre a CFDI DESHABILITAR"
                    //#region "Construimos Tabla de Datos CFDI"

                    //DefaultSplitCharacter split = new DefaultSplitCharacter();
                    //Table adicional = new Table(3);
                    //float[] headerwidthsAdicional = { 20, 25, 55 };
                    //adicional.Widths = headerwidthsAdicional;
                    //adicional.WidthPercentage = 100;
                    //adicional.Padding = 1;
                    //adicional.Spacing = 1;
                    //adicional.BorderWidth = (float).5;
                    //adicional.DefaultCellBorder = 1;
                    //adicional.BorderColor = gris;

                    //if (timbrar)
                    //{
                    //    #region "Generamos Quick Response Code"

                    //    byte[] bytesQRCode = new byte[0];

                    //    if (timbrar)
                    //    {
                    //        // Generamos el Quick Response Code (QRCode)
                    //        string re = electronicDocument.Data.Emisor.Rfc.Value;
                    //        string rr = electronicDocument.Data.Receptor.Rfc.Value;
                    //        string tt = String.Format("{0:F6}", electronicDocument.Data.Total.Value);
                    //        string id = objTimbre.Uuid.Value;

                    //        StringBuilder sbCadenaQRCode = new StringBuilder();

                    //        sbCadenaQRCode.
                    //            Append("?").
                    //            Append("re=").Append(re).
                    //            Append("&").
                    //            Append("rr=").Append(rr).
                    //            Append("&").
                    //            Append("tt=").Append(tt).
                    //            Append("&").
                    //            Append("id=").Append(id);

                    //        BarcodeLib.Barcode.QRCode.QRCode barcode = new BarcodeLib.Barcode.QRCode.QRCode();

                    //        barcode.Data = sbCadenaQRCode.ToString();
                    //        barcode.ModuleSize = 3;
                    //        barcode.LeftMargin = 0;
                    //        barcode.RightMargin = 10;
                    //        barcode.TopMargin = 0;
                    //        barcode.BottomMargin = 0;
                    //        barcode.Encoding = BarcodeLib.Barcode.QRCode.QRCodeEncoding.Auto;
                    //        barcode.Version = BarcodeLib.Barcode.QRCode.QRCodeVersion.Auto;
                    //        barcode.ECL = BarcodeLib.Barcode.QRCode.ErrorCorrectionLevel.L;
                    //        bytesQRCode = barcode.drawBarcodeAsBytes();
                    //    }

                    //    #endregion

                    //    Image imageQRCode = Image.GetInstance(bytesQRCode);
                    //    imageQRCode.Alignment = (Image.TEXTWRAP | Image.ALIGN_LEFT);
                    //    imageQRCode.ScaleToFit(90f, 90f);
                    //    imageQRCode.IndentationLeft = 9f;
                    //    imageQRCode.SpacingAfter = 9f;
                    //    imageQRCode.BorderColorTop = Color.WHITE;

                    //    cel = new Cell(imageQRCode);
                    //    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    //    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthRight = (float).5;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.Rowspan = 6;
                    //    adicional.AddCell(cel);

                    //    par = new Paragraph();
                    //    par.SetLeading(7f, 0f);
                    //    par.Add(new Chunk("SELLO DIGITAL DEL EMISOR\n", f5L));
                    //    par.Add(new Chunk(electronicDocument.Data.Sello.Value, f5).SetSplitCharacter(split));
                    //    cel = new Cell(par);
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthBottom = (float).5;
                    //    cel.Colspan = 2;
                    //    adicional.AddCell(cel);


                    //    cel = new Cell(new Phrase("FOLIO FISCAL:", f5L));
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    adicional.AddCell(cel);

                    //    cel = new Cell(new Phrase(objTimbre.Uuid.Value, f5));
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = (float).5;
                    //    cel.BorderWidthBottom = 0;
                    //    adicional.AddCell(cel);

                    //    cel = new Cell(new Phrase("FECHA Y HORA DE CERTIFICACION:", f5L));
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = (float).5;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    adicional.AddCell(cel);

                    //    string[] fechaTimbrado = Convert.ToDateTime(objTimbre.FechaTimbrado.Value).GetDateTimeFormats('s');

                    //    cel = new Cell(new Phrase(fechaTimbrado[0], f5));
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = (float).5;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = (float).5;
                    //    cel.BorderWidthBottom = 0;
                    //    adicional.AddCell(cel);

                    //    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL SAT:", f5L));
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = (float).5;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    adicional.AddCell(cel);

                    //    cel = new Cell(new Phrase(objTimbre.NumeroCertificadoSat.Value, f5));
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = (float).5;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = (float).5;
                    //    cel.BorderWidthBottom = 0;
                    //    adicional.AddCell(cel);

                    //    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL EMISOR:", f5L));
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = (float).5;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthBottom = (float).5;
                    //    adicional.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.NumeroCertificado.Value, f5));
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = (float).5;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = (float).5;
                    //    cel.BorderWidthBottom = (float).5;
                    //    adicional.AddCell(cel);

                    //    par = new Paragraph();
                    //    par.SetLeading(7f, 0f);
                    //    par.Add(new Chunk("TIPO DE COMPROBANTE: ", f5L));
                    //    par.Add(new Chunk(electronicDocument.Data.TipoComprobante.Value + "   |   ", f5));
                    //    par.Add(new Chunk("Moneda: ", f5L));
                    //    par.Add(new Chunk(electronicDocument.Data.Moneda.Value + "   |   ", f5));
                    //    par.Add(new Chunk("TASA DE CAMBIO: ", f5L));
                    //    string tasaCambio = electronicDocument.Data.TipoCambio.Value;
                    //    if (tasaCambio.Length > 0)
                    //    {
                    //        par.Add(new Chunk(Convert.ToDouble(tasaCambio).ToString("C", _ci) + "   |   ", f5));
                    //    }
                    //    else
                    //    {
                    //        par.Add(new Chunk("   |   ", f5));
                    //    }
                    //    par.Add(new Chunk("FORMA DE PAGO: ", f5L));
                    //    par.Add(new Chunk(electronicDocument.Data.FormaPago.Value + "   |   ", f5));
                    //    par.Add(new Chunk("MÉTODO DE PAGO: ", f5L));
                    //    par.Add(new Chunk(electronicDocument.Data.MetodoPago.Value, f5));
                    //    cel.BorderColor = gris;
                    //    cel = new Cell(par);
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.Colspan = 2;
                    //    cel.BorderColor = gris;
                    //    adicional.AddCell(cel);

                    //    par = new Paragraph();
                    //    par.SetLeading(7f, 0f);
                    //    par.Add(new Chunk("CADENA ORIGINAL DEL COMPLEMENTO DE CERTIFICACIÓN DIGITAL DEL SAT\n", f5L));
                    //    par.Add(new Chunk(electronicDocument.FingerPrintPac, f5).SetSplitCharacter(split));
                    //    cel = new Cell(par);
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = (float).5;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthBottom = (float).5;
                    //    cel.Colspan = 3;
                    //    adicional.AddCell(cel);

                    //    par = new Paragraph();
                    //    par.KeepTogether = true;
                    //    par.SetLeading(7f, 0f);
                    //    par.Add(new Chunk("SELLO DIGITAL DEL SAT\n", f5L));
                    //    par.Add(new Chunk(objTimbre.SelloSat.Value, f5).SetSplitCharacter(split));
                    //    cel = new Cell(par);
                    //    cel.BorderColor = gris;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.Colspan = 3;
                    //    adicional.AddCell(cel);
                    //}
                    //#endregion  
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

                    cell = new PdfPCell(new Phrase("ESTE DOCUMENTO ES UNA REPRESENTACIÓN IMPRESA DE UN CFD", titulo));
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.BackgroundColor = azul;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    cell.BorderWidthBottom = 0;
                    footer.AddCell(cell);

                    #endregion

                    pageEventHandler.encabezado = encabezado;
                    pageEventHandler.encPartidas = encabezadoDetalle;
                    pageEventHandler.datosAdic = datosAdic;
                    pageEventHandler.footer = footer;

                    if (t == 1)
                    {
                        document.Open();
                        numeroPaginasDescontar = 0;
                    }

                    else
                        numeroPaginasDescontar = numeroPaginasDescontarAlm;

                    document.Add(tableConceptos);
                    document.Add(especial1);
                    document.Add(comentarios);
                    document.Add(especial);
                    //document.Add(adicional);

                    if (t < 4)
                    {
                        document.NewPage();
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                string exe = ex.Message;
            }
        }

        #endregion

        #endregion

        #region "Formato Factura CFDI"

        #region "generarPdfCFDI"

        public static string generarPdfCFDI(Hashtable htFacturaxion, HttpContext hc)
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
                long idCfdi = Convert.ToInt64(htFacturaxion["idCfdi"]);

                #region "Obtenemos los datos del CFDI y Campos Opcionales"

                StringBuilder sbOpcionalEncabezado = new StringBuilder();
                DataTable dtOpcEnc = new DataTable();
                StringBuilder sbOpcionalDetalle = new StringBuilder();
                DataTable dtOpcDet = new DataTable();
                StringBuilder sbOpcionalDetalleRem = new StringBuilder();
                DataTable dtOpcDetRem = new DataTable();
                StringBuilder sbOpcionalDetalleImp = new StringBuilder();
                DataTable dtOpcDetImp = new DataTable();

                sbOpcionalEncabezado.
                    Append("SELECT ").
                    Append("campo16 AS nombreEnviarA, ").
                    Append("campo17 + ' ' + campo18 AS calleEnviarA, ").
                    Append("campo20 + ', ' + campo22 AS delMunEnviarA, ").
                    Append("campo23 + ', ' + campo24  AS edoEnviarA, ").
                    Append("'C.P.' + campo25 AS CPEnviarA, ").
                    Append("campo1 AS tipoDoc,  ").
                    Append("campo15 AS noCliente, ").
                    Append("campo14 AS gln, ").
                    Append("campo28 AS ordenCompra, ").
                    Append("campo29 AS noOrden, ").
                    Append("campo30 AS fechaOrden, ").
                    Append("campo33 AS plazo, ").
                    Append("campo32 AS folioRef, ").
                    Append("campo28 AS recibo, ").
                    Append("campo30 AS fechaRecibo, ").
                    Append("campo31 AS fechaExp, ").
                    Append("campo33 AS proyecto, ").
                    Append("campo34 AS comentarioFinDetalle, ").
                    Append("campo26 AS cantidadLetra ").
                    Append("FROM opcionalEncabezado ").
                    Append("WHERE idCFDI = @0  AND ST = 1 ");

                sbOpcionalDetalle.
                    Append("SELECT ROW_NUMBER() OVER (ORDER BY idCfdi ASC) AS numero, ").
                    Append("campo8 AS codigo, campo9 AS UPC, ").
                    Append("campo10 AS embarque, campo11 AS subtotal, campo13 AS C1, campo12 AS descuento, ").
                    Append("campo14 AS agenteAduanal ").
                    Append("FROM opcionalDetalle ").
                    Append("WHERE idCfdi = @0 ");

                sbOpcionalDetalleRem.
                    Append("SELECT campo1 AS remision, campo2 AS orden, campo3 AS codigo, campo4 AS UPC, campo5 AS descripcion, campo6 AS cantidad ").
                    Append("FROM opcionalDetalle2 ").
                    Append("WHERE idCfdi = @0 AND concepto = 1 AND ST = 1");

                sbOpcionalDetalleImp.
                    Append("SELECT campo1 AS impuesto, campo2 AS tasa, campo3 AS importe, campo4 AS baseImp, campo5 AS cl ").
                    Append("FROM opcionalDetalle2 ").
                    Append("WHERE idCfdi = @0 AND concepto = 2 AND ST = 1");

                dtOpcEnc = dal.QueryDT("DS_FE", sbOpcionalEncabezado.ToString(), "F:I:" + idCfdi, hc);
                dtOpcDet = dal.QueryDT("DS_FE", sbOpcionalDetalle.ToString(), "F:I:" + idCfdi, hc);
                dtOpcDetRem = dal.QueryDT("DS_FE", sbOpcionalDetalleRem.ToString(), "F:I:" + idCfdi, hc);
                dtOpcDetImp = dal.QueryDT("DS_FE", sbOpcionalDetalleImp.ToString(), "F:I:" + idCfdi, hc);

                if (dtOpcDet.Rows.Count == 0)
                {
                    for (int i = 1; i <= electronicDocument.Data.Conceptos.Count; i++)
                    {
                        dtOpcDet.Rows.Add("", "0.00");
                    }
                }

                #endregion

                #region "Extraemos los datos del CFDI"

                //Datos CFDI
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

                Document document = new Document(PageSize.LETTER, 15, 15, 15, 40);
                document.AddAuthor("Facturaxion");
                document.AddCreator("r3Take");
                document.AddCreationDate();

                pdfPageEventHandlerOfficeMax pageEventHandler = new pdfPageEventHandlerOfficeMax();
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                writer.PageEvent = pageEventHandler;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                HTC = hc;
                pathIMGLOGO = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\OFFICEMAX\Officemax-04.png";
                pathCedula = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\OFFICEMAX\Officemax-01.png";

                azul = new Color(22, 111, 168);
                blanco = new Color(255, 255, 255);
                Link = new Color(7, 73, 208);
                gris = new Color(236, 236, 236);
                grisOX = new Color(220, 215, 220);
                rojo = new Color(230, 7, 7);
                lbAzul = new Color(43, 145, 175);

                EM = BaseFont.CreateFont(@"C:\Windows\Fonts\VERDANA.TTF", BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
                f4 = new Font(EM, 4);
                f5 = new Font(EM, 5);
                f5B = new Font(EM, 5, Font.BOLD);
                f5BBI = new Font(EM, 5, Font.BOLDITALIC);
                f6 = new Font(EM, 6);
                f6B = new Font(EM, 6, Font.BOLD);
                f6L = new Font(EM, 6, Font.BOLD, Link);
                f5L = new Font(EM, 5, Font.BOLD, lbAzul);
                titulo = new Font(EM, 6, Font.BOLD, blanco);
                f5Bblanco = new Font(EM, 5, Font.BOLD, blanco);
                folio = new Font(EM, 6, Font.BOLD, rojo);
                dSaltoLinea = new Chunk("\n\n ");

                #endregion

                #region "Generamos el Documento"

                formatoFacturaCFDI(document, electronicDocument, pageEventHandler, idCfdi, dtOpcEnc, dtOpcDet, dtOpcDetRem, dtOpcDetImp, htDatosCfdi, objTimbre, HTC);

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

        #region "formatoFacturaCFDI"

        public static void formatoFacturaCFDI(Document document, ElectronicDocument electronicDocument, pdfPageEventHandlerOfficeMax pageEventHandler, long idCfdi, DataTable dtEncabezado, DataTable dtDetalle, DataTable dtOpcDetRem, DataTable dtOpcDetImp, Hashtable htCFDI, Data objTimbre, HttpContext hc)
        {
            try
            {
                DAL dal = new DAL();
                #region "Construimos el Documento"

                for (int t = 1; t < 5; t++)
                {

                    #region "Construimos el Encabezado"

                    Table encabezado = new Table(4);
                    float[] headerwidthsEncabezado = { 28, 32, 18, 22 };
                    encabezado.Widths = headerwidthsEncabezado;
                    encabezado.WidthPercentage = 100;
                    encabezado.Padding = 1;
                    encabezado.Spacing = 1;
                    encabezado.BorderWidth = 0;
                    encabezado.DefaultCellBorder = 0;
                    encabezado.BorderColor = gris;

                    //Agregando Imagen de Logotipo
                    Image imgLogo = Image.GetInstance(pathIMGLOGO);
                    imgLogo.ScalePercent(62f);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(1f, 1f);
                    par.Add(new Chunk(imgLogo, 0, 0));
                    par.Add(new Chunk("", f6));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 3;
                    encabezado.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("VENDIDO A:", f6B));
                    par.Add(new Chunk("\n" + htCFDI["nombreReceptor"].ToString().ToUpper(), f6B));
                    par.Add(new Chunk("\n" + htCFDI["direccionReceptor1"].ToString().ToUpper(), f6));
                    par.Add(new Chunk("\n" + htCFDI["direccionReceptor2"].ToString().ToUpper(), f6));
                    par.Add(new Chunk("\n" + htCFDI["direccionReceptor3"].ToString().ToUpper(), f6));
                    par.Add(new Chunk("\nR.F.C.: " + htCFDI["rfcReceptor"].ToString().ToUpper(), f6B));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 3;
                    encabezado.AddCell(cel);

                    //Agregando Imagen de CedulaFiscal
                    Image imgCedula = Image.GetInstance(pathCedula);
                    imgCedula.ScalePercent(47f);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(1f, 1f);
                    par.Add(new Chunk(imgCedula, 0, 0));
                    par.Add(new Chunk("", f6));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 5;
                    encabezado.AddCell(cel);

                    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["tipoDoc"].ToString().ToUpper(), f6B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    encabezado.AddCell(cel);

                    cel = new Cell(new Phrase(htCFDI["serie"].ToString().ToUpper() + " " + electronicDocument.Data.Folio.Value, folio));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    encabezado.AddCell(cel);

                    cel = new Cell(new Phrase("UUID: \n" + htCFDI["UUID"].ToString(), f5B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    encabezado.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 1f);
                    par.Add(new Chunk(htCFDI["nombreEmisor"].ToString().ToUpper() + "\n\n", f6B));
                    par.Add(new Chunk(htCFDI["direccionEmisor1"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(htCFDI["direccionEmisor2"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(htCFDI["direccionEmisor3"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk("R.F.C.: " + htCFDI["rfcEmisor"].ToString().ToUpper() + "\n\n", f6B));
                    par.Add(new Chunk("Teléfono: 91772800", f6));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 2;
                    encabezado.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 1f);
                    par.Add(new Chunk("ENVIAR A: \n", f6B));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["nombreEnviarA"].ToString().ToUpper() + "\n", f6B));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["calleEnviarA"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["delMunEnviarA"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["edoEnviarA"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["CPEnviarA"].ToString().ToUpper(), f6));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 2;
                    encabezado.AddCell(cel);

                    //Pagina n de n

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 1f);
                    par.Add(new Chunk(dtEncabezado.Rows[0]["gln"].ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(electronicDocument.Data.Emisor.ExpedidoEn.Calle.Value.ToString().ToUpper() + " " + electronicDocument.Data.Emisor.ExpedidoEn.NumeroExterior.ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk(electronicDocument.Data.Emisor.ExpedidoEn.Colonia.Value.ToString().ToUpper() + " " + electronicDocument.Data.Emisor.ExpedidoEn.Municipio.Value.ToString().ToUpper() + " " + electronicDocument.Data.Emisor.ExpedidoEn.Estado.Value.ToString().ToUpper() + " " + electronicDocument.Data.Emisor.ExpedidoEn.Pais.Value.ToString().ToUpper() + "\n", f6));
                    par.Add(new Chunk("C.P. " + electronicDocument.Data.Emisor.ExpedidoEn.CodigoPostal.Value.ToString().ToUpper(), f6));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 2;
                    encabezado.AddCell(cel);

                    #endregion

                    #region "tabla Datos Adicionales"

                    Table datosAdic = new Table(7);
                    float[] headerWidthsDatosAdic = { 13, 13, 13, 13, 22, 13, 13 };
                    datosAdic.Widths = headerWidthsDatosAdic;
                    datosAdic.WidthPercentage = 100;
                    datosAdic.Padding = 1;
                    datosAdic.Spacing = 1;
                    datosAdic.BorderWidth = (float).5;
                    datosAdic.DefaultCellBorder = 0;
                    datosAdic.BorderColor = gris;

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("No Cliente:", f5));
                    par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["noCliente"].ToString(), f5B));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);

                    if (electronicDocument.Data.Serie.Value.ToString() == "VTASIMPORT")
                    {
                        par.Add(new Chunk("Recibo:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["recibo"].ToString(), f5B));
                    }

                    else
                    {
                        par.Add(new Chunk("O.Compra Cte.:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["ordenCompra"].ToString(), f5B));
                    }

                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("No.orden / REA:", f5));
                    par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["noOrden"].ToString(), f5B));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);

                    if (electronicDocument.Data.Serie.Value.ToString() == "VTASIMPORT")
                    {
                        par.Add(new Chunk("fecha Recibo:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["fechaRecibo"].ToString(), f5B));
                    }

                    else
                    {
                        par.Add(new Chunk("fecha Orden:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["fechaOrden"].ToString(), f5B));
                    }

                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("Lugar y fecha de Expedición:", f5));
                    par.Add(new Chunk("\n" + electronicDocument.Data.Emisor.ExpedidoEn.Calle.Value.ToString().ToUpper() + " " + electronicDocument.Data.Emisor.ExpedidoEn.NumeroExterior.Value.ToString().ToUpper() + " " + dtEncabezado.Rows[0]["fechaExp"].ToString(), f5B));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);

                    if (electronicDocument.Data.Serie.Value.ToString() == "VTASIMPORT")
                    {
                        par.Add(new Chunk("Proyecto:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["proyecto"].ToString(), f5B));
                    }

                    else
                    {
                        par.Add(new Chunk("Plazo:", f5));
                        par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["plazo"].ToString(), f5B));
                    }

                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("Folio Referencia:", f5));
                    par.Add(new Chunk("\n" + dtEncabezado.Rows[0]["folioRef"].ToString(), f5B));
                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    datosAdic.AddCell(cel);

                    #endregion

                    #region"Tabla Detalle"

                    Table encabezadoDetalle = new Table(12);
                    float[] headerEncabezadoDetalle = { 5, 7, 10, 20, 7, 5, 7, 8, 8, 8, 10, 5 };
                    encabezadoDetalle.Widths = headerEncabezadoDetalle;
                    encabezadoDetalle.WidthPercentage = 100F;
                    encabezadoDetalle.Padding = 1;
                    encabezadoDetalle.Spacing = 1;
                    encabezadoDetalle.BorderWidth = (float).5;
                    encabezadoDetalle.DefaultCellBorder = 0;
                    encabezadoDetalle.BorderColor = gris;

                    // NUMERO
                    cel = new Cell(new Phrase("NO.", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // CODIGO
                    cel = new Cell(new Phrase("CÓDIGO", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // UPC
                    cel = new Cell(new Phrase("UPC", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // DESCRIPCIÓN
                    cel = new Cell(new Phrase("DESCRIPCIÓN", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // CANTIDAD SOL
                    cel = new Cell(new Phrase("CANT. SOL.", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // UNIDAD DE MEDIDA
                    cel = new Cell(new Phrase("U/M", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // EMBARQUE
                    cel = new Cell(new Phrase("EMBAR.", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // PRECIO UNITARIO
                    cel = new Cell(new Phrase("PRECIO UNIT.", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // SUBTOTAL
                    cel = new Cell(new Phrase("SUBTOTAL", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // DESCUENTO
                    cel = new Cell(new Phrase("DESCUENTO", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // TOTAL
                    cel = new Cell(new Phrase("TOTAL", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    // C1
                    cel = new Cell(new Phrase("C1.", f5Bblanco));
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    encabezadoDetalle.AddCell(cel);

                    PdfPTable tableConceptos = new PdfPTable(12);
                    tableConceptos.SetWidths(new int[12] { 5, 7, 10, 20, 7, 5, 7, 8, 8, 8, 10, 5 });
                    tableConceptos.WidthPercentage = 100F;

                    int numConceptos = electronicDocument.Data.Conceptos.Count;
                    PdfPCell cellConceptos = new PdfPCell();
                    PdfPCell cellMontos = new PdfPCell();

                    for (int i = 0; i < numConceptos; i++)
                    {
                        Concepto concepto = electronicDocument.Data.Conceptos[i];

                        cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["numero"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["codigo"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["UPC"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(concepto.Descripcion.Value.ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(concepto.Cantidad.Value.ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(concepto.Unidad.Value.ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                        tableConceptos.AddCell(cellConceptos);

                        cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["embarque"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tableConceptos.AddCell(cellConceptos);

                        cellMontos = new PdfPCell(new Phrase(concepto.ValorUnitario.Value.ToString("C", _ci), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellMontos.Border = 0;
                        cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                        tableConceptos.AddCell(cellMontos);

                        cellMontos = new PdfPCell(new Phrase("$" + dtDetalle.Rows[i]["subtotal"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellMontos.Border = 0;
                        cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                        tableConceptos.AddCell(cellMontos);

                        cellMontos = new PdfPCell(new Phrase("$" + dtDetalle.Rows[i]["descuento"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellMontos.Border = 0;
                        cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                        tableConceptos.AddCell(cellMontos);

                        cellMontos = new PdfPCell(new Phrase(concepto.Importe.Value.ToString("C", _ci), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellMontos.Border = 0;
                        cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                        tableConceptos.AddCell(cellMontos);

                        cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["C1"].ToString(), new Font(Font.HELVETICA, 6, Font.NORMAL)));
                        cellConceptos.Border = 0;
                        cellConceptos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                        tableConceptos.AddCell(cellConceptos);

                        if (concepto.InformacionAduanera.IsAssigned)
                        {
                            for (int j = 0; j < concepto.InformacionAduanera.Count; j++)
                            {
                                cellConceptos = new PdfPCell(new Phrase("Aduana", new Font(Font.HELVETICA, 5, Font.BOLD)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 2;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase("Fecha Pedimento", new Font(Font.HELVETICA, 5, Font.BOLD)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 1;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase("Pedimento", new Font(Font.HELVETICA, 5, Font.BOLD)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 1;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase("Agente Aduanal", new Font(Font.HELVETICA, 5, Font.BOLD)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 8;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase(concepto.InformacionAduanera[j].Aduana.Value.ToString(), new Font(Font.HELVETICA, 5, Font.NORMAL)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 2;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase(concepto.InformacionAduanera[j].Fecha.Value.ToString().Substring(0, 9), new Font(Font.HELVETICA, 5, Font.NORMAL)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 1;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase(concepto.InformacionAduanera[j].Numero.Value.ToString(), new Font(Font.HELVETICA, 5, Font.NORMAL)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 1;
                                tableConceptos.AddCell(cellConceptos);

                                cellConceptos = new PdfPCell(new Phrase(dtDetalle.Rows[i]["agenteAduanal"].ToString(), new Font(Font.HELVETICA, 5, Font.NORMAL)));
                                cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                                cellConceptos.Border = 0;
                                cellConceptos.Colspan = 8;
                                tableConceptos.AddCell(cellConceptos);
                            }
                        }
                    }

                    #endregion

                    #region "Tabla Remisiones"

                    Table especial1 = new Table(7);
                    float[] headerwidthsEspecial1 = { 10, 10, 10, 12, 18, 10, 30 };
                    especial1.Widths = headerwidthsEspecial1;
                    especial1.WidthPercentage = 100;
                    especial1.Padding = 1;
                    especial1.Spacing = 1;
                    especial1.BorderWidth = 0;
                    especial1.DefaultCellBorder = 0;
                    especial1.BorderColor = blanco;

                    if (dtOpcDetRem.Rows.Count > 0)
                    {
                        cel = new Cell(new Phrase("\n\n.", titulo));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.Colspan = 7;
                        especial1.AddCell(cel);

                        //Hacemos encabezado de tabla de remision
                        cel = new Cell(new Phrase("Remisión", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("Orden", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("Código", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("UPC", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("Descripción", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("Cantidad", f5B));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        cel = new Cell(new Phrase("", f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        especial1.AddCell(cel);

                        for (int i = 0; i < dtOpcDetRem.Rows.Count; i++)
                        {
                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["remision"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["orden"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["codigo"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["UPC"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["descripcion"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase(dtOpcDetRem.Rows[i]["cantidad"].ToString(), f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);

                            cel = new Cell(new Phrase("", f5));
                            cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cel.BorderWidthTop = 0;
                            cel.BorderWidthLeft = 0;
                            cel.BorderWidthRight = 0;
                            cel.BorderWidthBottom = 0;
                            especial1.AddCell(cel);
                        }
                    }

                    cel = new Cell(new Phrase("\n\n\n.", titulo));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = blanco;
                    cel.Colspan = 7;
                    especial1.AddCell(cel);

                    #endregion

                    #region "Construimos el Comentarios"

                    Table comentarios = new Table(7);
                    float[] headerwidthsComentarios = { 11, 52, 6, 8, 6, 7, 10 };
                    comentarios.Widths = headerwidthsComentarios;
                    comentarios.WidthPercentage = 100;
                    comentarios.Padding = 1;
                    comentarios.Spacing = 1;
                    comentarios.BorderWidth = 0;
                    comentarios.DefaultCellBorder = 0;
                    comentarios.BorderColor = gris;

                    if (dtEncabezado.Rows[0]["comentarioFinDetalle"].ToString().Length > 0)
                    {
                        cel = new Cell(new Phrase(dtEncabezado.Rows[0]["comentarioFinDetalle"].ToString().ToUpper(), f5));
                        cel.VerticalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        cel.Colspan = 7;
                        comentarios.AddCell(cel);
                    }

                    cel = new Cell(new Phrase("\n\n.", titulo));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.Colspan = 7;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("Cantidad con letra:", f6B));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["cantidadLetra"].ToString(), f6));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("Clase", f6));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("Base Cálculo", f6));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("Tasa", f6));
                    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("Sub Total", f6B));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f6));
                    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    #region "Codigo Alterno Impuestos"
                    //for (int i = 0; i < electronicDocument.Data.Impuestos.Traslados.Count; i++)
                    //{
                    //    cel = new Cell(new Phrase("", f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    cel.Colspan = 2;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["clase"].ToString(), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Traslados[i].Tasa.Value.ToString(), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Traslados[i].Tipo.Value.ToString(), f6B));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = (float).5;
                    //    cel.BorderWidthRight = (float).5;
                    //    cel.BorderWidthBottom = (float).5;
                    //    cel.BorderColor = gris;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Traslados[i].Importe.Value.ToString("C", _ci), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = (float).5;
                    //    cel.BorderWidthBottom = (float).5;
                    //    cel.BorderColor = gris;
                    //    comentarios.AddCell(cel);

                    //}

                    //for (int i = 0; i < electronicDocument.Data.Impuestos.Retenciones.Count; i++)
                    //{
                    //    cel = new Cell(new Phrase("", f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    cel.Colspan = 2;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["clase"].ToString(), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase("", f6B));
                    //    cel.VerticalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = 0;
                    //    cel.BorderWidthBottom = 0;
                    //    cel.BorderColor = blanco;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Retenciones[i].Tipo.Value.ToString(), f6B));
                    //    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = (float).5;
                    //    cel.BorderWidthRight = (float).5;
                    //    cel.BorderWidthBottom = (float).5;
                    //    cel.BorderColor = gris;
                    //    comentarios.AddCell(cel);

                    //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Retenciones[i].Importe.Value.ToString("C", _ci), f6));
                    //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    //    cel.BorderWidthTop = 0;
                    //    cel.BorderWidthLeft = 0;
                    //    cel.BorderWidthRight = (float).5;
                    //    cel.BorderWidthBottom = (float).5;
                    //    cel.BorderColor = gris;
                    //    comentarios.AddCell(cel);
                    //}
                    #endregion

                    for (int y = 0; y < dtOpcDetImp.Rows.Count; y++)
                    {
                        cel = new Cell(new Phrase("", f6));
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = blanco;
                        cel.Colspan = 2;
                        comentarios.AddCell(cel);

                        cel = new Cell(new Phrase(dtOpcDetImp.Rows[0]["cl"].ToString(), f6));
                        cel.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = blanco;
                        comentarios.AddCell(cel);

                        cel = new Cell(new Phrase(dtOpcDetImp.Rows[0]["baseImp"].ToString(), f6));
                        cel.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = blanco;
                        comentarios.AddCell(cel);

                        cel = new Cell(new Phrase(dtOpcDetImp.Rows[0]["tasa"].ToString(), f6));
                        cel.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = blanco;
                        comentarios.AddCell(cel);

                        cel = new Cell(new Phrase(dtOpcDetImp.Rows[0]["impuesto"].ToString(), f6B));
                        cel.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        comentarios.AddCell(cel);

                        cel = new Cell(new Phrase(dtOpcDetImp.Rows[0]["importe"].ToString(), f6));
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        comentarios.AddCell(cel);
                    }

                    cel = new Cell(new Phrase("", f6));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = blanco;
                    cel.Colspan = 5;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase("TOTAL", f6B));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _ci), f6));
                    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    comentarios.AddCell(cel);

                    #endregion

                    #region "Construimos Tabla de Datos CFDI"

                    DefaultSplitCharacter split = new DefaultSplitCharacter();
                    Table adicional = new Table(3);
                    float[] headerwidthsAdicional = { 20, 25, 55 };
                    adicional.Widths = headerwidthsAdicional;
                    adicional.WidthPercentage = 100;
                    adicional.Padding = 1;
                    adicional.Spacing = 1;
                    adicional.BorderWidth = 0;
                    adicional.DefaultCellBorder = 0;
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
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        cel.Rowspan = 6;
                        adicional.AddCell(cel);

                        par = new Paragraph();
                        par.SetLeading(7f, 0f);
                        par.Add(new Chunk("SELLO DIGITAL DEL EMISOR\n", f5L));
                        par.Add(new Chunk(electronicDocument.Data.Sello.Value, f5).SetSplitCharacter(split));
                        cel = new Cell(par);
                        cel.BorderColor = gris;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        cel.Colspan = 2;
                        adicional.AddCell(cel);


                        cel = new Cell(new Phrase("FOLIO FISCAL:", f5L));
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
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        adicional.AddCell(cel);

                        cel = new Cell(new Phrase("FECHA Y HORA DE CERTIFICACION:", f5L));
                        cel.BorderColor = gris;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        adicional.AddCell(cel);

                        string[] fechaTimbrado = Convert.ToDateTime(objTimbre.FechaTimbrado.Value).GetDateTimeFormats('s');

                        cel = new Cell(new Phrase(fechaTimbrado[0], f5));
                        cel.BorderColor = gris;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        adicional.AddCell(cel);

                        cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL SAT:", f5L));
                        cel.BorderColor = gris;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        adicional.AddCell(cel);

                        cel = new Cell(new Phrase(objTimbre.NumeroCertificadoSat.Value, f5));
                        cel.BorderColor = gris;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        adicional.AddCell(cel);

                        cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL EMISOR:", f5L));
                        cel.BorderColor = gris;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        adicional.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.NumeroCertificado.Value, f5));
                        cel.BorderColor = gris;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        adicional.AddCell(cel);

                        par = new Paragraph();
                        par.SetLeading(7f, 0f);
                        par.Add(new Chunk("TIPO DE COMPROBANTE: ", f5L));
                        par.Add(new Chunk(electronicDocument.Data.TipoComprobante.Value + "   |   ", f5));
                        par.Add(new Chunk("Moneda: ", f5L));
                        par.Add(new Chunk(electronicDocument.Data.Moneda.Value + "   |   ", f5));
                        par.Add(new Chunk("TASA DE CAMBIO: ", f5L));
                        
                        string tasaCambio = electronicDocument.Data.TipoCambio.Value;
                        string regimenes = "";

                        if (tasaCambio.Length > 0)
                        {
                            par.Add(new Chunk(Convert.ToDouble(tasaCambio).ToString("C", _ci) + "   |   ", f5));
                        }
                        else
                        {
                            par.Add(new Chunk("   |   ", f5));
                        }
                        par.Add(new Chunk("FORMA DE PAGO: ", f5L));
                        par.Add(new Chunk(electronicDocument.Data.FormaPago.Value + "   |   ", f5));
                        par.Add(new Chunk("MÉTODO DE PAGO: ", f5L));
                        par.Add(new Chunk(electronicDocument.Data.MetodoPago.Value, f5));

                        if (electronicDocument.Data.NumeroCuentaPago.Value.ToString().Length > 0)
                        {
                            par.Add(new Chunk("   |   " + "No. CUENTA: ", f5L));
                            par.Add(new Chunk(electronicDocument.Data.NumeroCuentaPago.Value, f5));
                        }

                        if (electronicDocument.Data.Emisor.Regimenes.Count > 0)
                        {
                            for (int u = 0; u < electronicDocument.Data.Emisor.Regimenes.Count; u++)
                                regimenes += electronicDocument.Data.Emisor.Regimenes[u].Regimen.Value.ToString() + ",";

                            par.Add(new Chunk("   |   " + "RÉGIMEN FISCAL: ", f5L));
                            par.Add(new Chunk(regimenes.Substring(0,regimenes.Length-1).ToString() + "   |   ", f5));
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

                        par = new Paragraph();
                        par.SetLeading(7f, 0f);
                        par.Add(new Chunk("CADENA ORIGINAL DEL COMPLEMENTO DE CERTIFICACIÓN DIGITAL DEL SAT\n", f5L));
                        par.Add(new Chunk(electronicDocument.FingerPrintPac, f5).SetSplitCharacter(split));
                        cel = new Cell(par);
                        cel.BorderColor = gris;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthBottom = 0;
                        cel.Colspan = 3;
                        adicional.AddCell(cel);

                        par = new Paragraph();
                        par.KeepTogether = true;
                        par.SetLeading(7f, 0f);
                        par.Add(new Chunk("SELLO DIGITAL DEL SAT\n", f5L));
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

                    #region "Construimos Tabla Especial"

                    Table especial = new Table(2);
                    float[] headerwidthsEspecial = { 65, 35 };
                    especial.Widths = headerwidthsEspecial;
                    especial.WidthPercentage = 100;
                    especial.Padding = 1;
                    especial.Spacing = 1;
                    especial.BorderWidth = 0;
                    especial.DefaultCellBorder = 0;
                    especial.BorderColor = gris;

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(8f, 9f);
                    par.Add(new Chunk("\n\n\n\nMéxico Distrito Federal a Debo(emos) y pagaré(mos) incondicionalmente a la orden de OPERADORA OMX S.A. DE C.V. en México D.F. ", f5B));
                    par.Add(new Chunk("p en cualquier otra que se me requiera el pago, la cantidad consignada en el presente titulo misma que ha sido recibida a mi entera ", f5B));
                    par.Add(new Chunk("satisfacción en caso de incumplimiento del presente pagaré, me obligo a pagar un interés monetario mensual por el equivalente a aplicar la ", f5B));
                    par.Add(new Chunk("tasa lider publicada por el Banco de México, en el mes inmediato anterior atendiendose por tasa lider, la mayor entre las siguientes: Tasa de interés ", f5B));
                    par.Add(new Chunk("interbancario de equilibrio, tasa de interés interbancario promedio, el costo porcentual promedio de captación.", f5B));

                    par.Add(new Chunk("EL PAGO DE ESTA FACTURA SE HARÁ EN UNA SOLA EXHIBICIÓN RECIBI MERCANCIA A MI ENTERA SATISFACCIÓN PUESTO\n", f5B));
                    par.Add(new Chunk("FIRMA___________________________________________________________________________________________________\n", f5B));
                    par.Add(new Chunk("PARA EFECTOS FISCALES AL PAGO \"LA REPRODUCCIÓN NO AUTORIZADA DE ESTE COMPROBANTE CONSTITUYE UN  DELITO EN LOS TERMINOS DE LAS DISPOSICIONES FISCALES\" \n", f5B));
                    par.Add(new Chunk("ACEPTO", f5B));

                    par.Add(new Chunk("\n\nPolíticas de Cambios Físicos y Devoluciones de Mercancía \n", f4));
                    par.Add(new Chunk("Comunicandose al área de Servicios al Cliente al teléfono 5321-9921 y lada sin costo (01800) 713-7070. Es importante que al solicitar algún cambio físico, devolución parcial o total presente factura original o ticket, además de: \n", f4));
                    par.Add(new Chunk(". El artículo deberá tener su empaque original con todos sus accesorios;así como la Factura Original y/o Ticket. \n", f4));
                    par.Add(new Chunk(". Artículos eléctricos deberá realizarse dentro de los 7 días naturales siguientes a la fecha de entrega, transcurrido el plazo se deberá acudir a los centros de servicio autorizados por el fabricante. El cambio físico solo será por defecto de fabricación. \n", f4));
                    par.Add(new Chunk(". En cartuchos, toners, película tpermica, software y plumas finas, productos de exhibición, demostración, liquidación, repartos, tarjetas telefónicas, tarjetas internet, cd no hay devoluciones ni cambios físicos. Cualquier reclamación de estos productos será directamente con el fabricante.", f4));
                    par.Add(new Chunk(". En muebles, el plazo para reaizar cualwuier reclamación será de 30 días naturales contados a partir de la fecha de entrega. Solicite el armado por el personal certificado de OFFICEMAX llamando al 5321-9921. No se aceptará cambio físico o devolución si el mueble fue armado por personal no autorizado por la empresa. \n", f4));
                    par.Add(new Chunk(". En artículos de oficina el plazo para realizar cualquier reclamción será de 15 días naturales contados a partir de la fecha de entrega. \n", f4));
                    par.Add(new Chunk(". Para ventas corporativas las políticas están establecidas en contrato. (confirmar con su vendedor). \n", f4));
                    par.Add(new Chunk(". Tratándose de mercancía adquirida con tarjeta de crédito bajo el esquema de meses sin intereses, aplicará el cambio físico de mercancía (según garantías antes mencionadas). La devolución del artículo podrá efectuarse previa autorización del cliente para realizar el cargo de los gastos administrativos incurridos a la tarjeta de crédito con la que se realizó la compra. \n", f4));
                    par.Add(new Chunk(". Las devoluciones se realizarán a la forma de pago original en la que se realizó la compra. En el caso de que la compra se efectivo en cheque o efectivo la devolución procederá hasta por un monto máximo de $2,000.00 (Dos mil pesos 00/100 m.n.) en efectivo. Si el importe es mayor a esta cantidad se rembolsará a través de un cheque nominativo. (En casos especiales se aplicará vale de mercancía). \n", f4));
                    par.Add(new Chunk(". Officemax se deslinda de la responsabilidad total o parcial por pérdida o daño de mercancía una vez que se entregó al cliente y se firmó de acuse de recibo bajo su entera conformidad y satisfaciión por parte del cliente.", f4));

                    cel = new Cell(par);
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 2;
                    cel.BorderColor = gris;
                    especial.AddCell(cel);

                    if (t == 1)
                        cel = new Cell(new Phrase("ORIGINAL", f6B));

                    else if (t == 2)
                        cel = new Cell(new Phrase("CLIENTE", f6B));

                    else if (t == 3)
                        cel = new Cell(new Phrase("CRÉDITO Y COBRANZA", f6B));

                    else
                        cel = new Cell(new Phrase("CONSECUTIVO FISCAL", f6B));

                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.Colspan = 2;
                    especial.AddCell(cel);

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

                    cell = new PdfPCell(new Phrase("ESTE DOCUMENTO ES UNA REPRESENTACIÓN IMPRESA DE UN CFDI", titulo));
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.BackgroundColor = azul;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    cell.BorderWidthBottom = 0;
                    footer.AddCell(cell);

                    #endregion

                    pageEventHandler.encabezado = encabezado;
                    pageEventHandler.encPartidas = encabezadoDetalle;
                    pageEventHandler.datosAdic = datosAdic;
                    pageEventHandler.footer = footer;

                    if (t == 1)
                    {
                        document.Open();
                        numeroPaginasDescontar = 0;
                    }

                    else
                        numeroPaginasDescontar = numeroPaginasDescontarAlm;

                    document.Add(tableConceptos);
                    document.Add(especial1);
                    document.Add(comentarios);
                    document.Add(adicional);
                    document.Add(especial);

                    if (t < 4)
                    {
                        document.NewPage();
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                string exe = ex.Message;
            }
        }

        #endregion

        #endregion

    }

    public class pdfPageEventHandlerOfficeMax : PdfPageEventHelper
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
        public Table encPartidas { get; set; }
        public PdfPTable detalle { get; set; }
        public PdfPTable footer { get; set; }
        public Table adicional { get; set; }
        public Table datosAdic { get; set; }
        //public int numeroPagina { get; set; }
        //public int numeroPaginasDescontar { get; set; }
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
            //encabezado.WriteSelectedRows(0, -1, 15, (document.PageSize.Height - 10), writer.DirectContent);
            document.Add(encabezado);
            document.Add(datosAdic);
            document.Add(encPartidas);

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

            OfficeMax.numeroPagina = writer.PageNumber;
            OfficeMax.numeroPaginasDescontarAlm = OfficeMax.numeroPagina;
            
            int pageN = writer.PageNumber - OfficeMax.numeroPaginasDescontar;
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
            //int paginastotal = (writer.PageNumber - 1);
            int paginastotal = (writer.PageNumber - 1) / 4;
            template.ShowText(paginastotal.ToString());
            template.EndText();
        }

        #endregion
    }

    public class DefaultSplitCharacterOfficeMax : ISplitCharacter
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