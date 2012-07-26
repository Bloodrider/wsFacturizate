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
    public class PFI730206632
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
        private static Font titulo;
        private static Font folio;
        private static Font f5L;

        #endregion

        #region "generarPdf"

        public static string generarPdf(Hashtable htFacturaxion, HttpContext hc)
        {
            string pathPdf = htFacturaxion["rutaDocumentoPdf"].ToString();
            FileStream fs = new FileStream(pathPdf, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);

            try
            {
                DAL dal = new DAL();
                _ci.NumberFormat.CurrencyDecimalDigits = 2;

                ElectronicDocument electronicDocument = (ElectronicDocument)htFacturaxion["electronicDocument"];
                Data objTimbre = (Data)htFacturaxion["objTimbre"];
                timbrar = Convert.ToBoolean(htFacturaxion["timbrar"]);
                pathPdf = htFacturaxion["rutaDocumentoPdf"].ToString();
                Int64 idCfdi = Convert.ToInt64(htFacturaxion["idCfdi"]);

                #region "Obtenemos los datos del CFDI y Campos Opcionales"

                StringBuilder sbOpcionalEncabezado = new StringBuilder();
                DataTable dtOpcEnc = new DataTable();
                StringBuilder sbOpcionalDetalle = new StringBuilder();
                DataTable dtOpcDet = new DataTable();
                StringBuilder sbDataEmisor = new StringBuilder();
                DataTable dtDataEmisor = new DataTable();

                sbOpcionalEncabezado.
                    Append("SELECT ").
                    Append("campo3 AS ordenCompra, ").
                    Append("campo7 AS numCliente, ").
                    Append("campo8 AS NombreSuc, ").
                    Append("campo19 AS pedido, ").
                    Append("campo20 AS division, ").
                    Append("campo21 AS efectuarPago,  ").
                    Append("campo22 AS direccionPie,  ").
                    Append("campo23 AS paraConsultas, ").
                    Append("campo26 AS cantidadLetra, ").
                    Append("campo27 AS zona, ").
                    Append("campo28 AS vencimiento, ").
                    Append("campo29 AS corporateCode, ").
                    Append("campo31 AS hyperion, ").
                    Append("campo32 AS INTERICOM, ").
                    Append("campo30 AS oldCorporateCode, ").
                    Append("campo33 AS referencia, ").
                    Append("campo34 AS fechaRefDoc, ").
                    Append("campo36 AS codigoCorporativo, ").
                    Append("campo37 AS codigoCorporativoAnt, ").
                    Append("campo38 AS clienteNo, ").
                    Append("campo39 AS contacto, ").
                    Append("campo40 AS observaciones, ").
                    Append("campo41 AS NoAp, ").
                    Append("campo42 AS AnoAp, ").

                    Append("' ' AS [ENTREGADO-A-NOMBRE], ").
                    Append("campo10 + ' ' + campo11 + ' ' +  ").
                    Append("campo12 AS [ENTREGADO-A-CALLE], ").
                    Append("campo13 AS [ENTREGADO-A-COLONIA], ").
                    Append("campo18 AS [ENTREGADO-A-CP], ").
                    Append("campo15 AS [ENTREGADO-A-MUNIC], ").
                    Append("campo14 AS [ENTREGADO-A-LOCAL], ").
                    Append("campo16 AS [ENTREGADO-A-ESTADO], ").
                    Append("campo17 AS [ENTREGADO-A-PAIS] ").
                    Append("FROM opcionalEncabezado ").
                    Append("WHERE idCFDI = @0  AND ST = 1 ");

                sbOpcionalDetalle.
                    Append("SELECT ").
                    Append("COALESCE(campo1, '') AS codeLocal, ").
                    Append("COALESCE(campo3, '') AS lote, ").
                    Append("COALESCE(campo4, '') AS cantidad, ").
                    Append("COALESCE(campo5, '') AS expiracion, ").
                    Append("COALESCE(campo6, '') AS taxRate, ").
                    Append("COALESCE(campo7, '') AS taxPaid, ").
                    Append("COALESCE(campo10, '') AS codeOracle, ").
                    Append("COALESCE(campo11, '') AS codeISPC, ").
                    Append("COALESCE(campo12, '') AS codeImpuesto, ").
                    Append("COALESCE(campo13, '') AS centroCostos, ").
                    Append("COALESCE(campo14, '') AS clinico, ").
                    Append("COALESCE(campo15, '') AS proyecto, ").
                    Append("COALESCE(campo16, '') AS cantidadReal, ").
                    Append("COALESCE(campo17, '') AS descuento, ").
                    Append("COALESCE(campo18, '') AS codBarras ").
                    Append("FROM opcionalDetalle ").
                    Append("WHERE idCFDI = @0 ");

                sbDataEmisor.Append("SELECT nombreSucursal FROM sucursales WHERE idSucursal = @0 ");

                dtOpcEnc = dal.QueryDT("DS_FE", sbOpcionalEncabezado.ToString(), "F:I:" + idCfdi, hc);
                dtOpcDet = dal.QueryDT("DS_FE", sbOpcionalDetalle.ToString(), "F:I:" + idCfdi, hc);
                dtDataEmisor = dal.QueryDT("DS_FE", sbDataEmisor.ToString(), "F:I:" + htFacturaxion["idSucursalEmisor"].ToString(), hc);

                if (dtOpcDet.Rows.Count == 0)
                {
                    for (int i = 1; i <= electronicDocument.Data.Conceptos.Count; i++)
                    {
                        dtOpcDet.Rows.Add("", "0.00");
                    }
                }

                #endregion

                #region "Extraemos los datos del CFDI"

                Hashtable htDatosCfdi = new Hashtable();
                htDatosCfdi.Add("nombreEmisor", electronicDocument.Data.Emisor.Nombre.Value);
                htDatosCfdi.Add("rfcEmisor", electronicDocument.Data.Emisor.Rfc.Value);
                htDatosCfdi.Add("nombreReceptor", electronicDocument.Data.Receptor.Nombre.Value);
                htDatosCfdi.Add("rfcReceptor", electronicDocument.Data.Receptor.Rfc.Value);
                htDatosCfdi.Add("sucursal", dtDataEmisor.Rows[0]["nombreSucursal"]);
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
                    sbDirEmisor3.Append("CP ").Append(electronicDocument.Data.Emisor.Domicilio.CodigoPostal.Value).Append(", ");
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
                    sbDirExpedido3.Append("CP ").Append(electronicDocument.Data.Emisor.ExpedidoEn.CodigoPostal.Value).Append(", ");
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
                    sbDirReceptor3.Append("CP ").Append(electronicDocument.Data.Receptor.Domicilio.CodigoPostal.Value).Append(", ");
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
                pdfPageEventHandlerPfizer pageEventHandler = new pdfPageEventHandlerPfizer();
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                writer.PageEvent = pageEventHandler;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                //document.Open();

                HTC = hc;
                pathIMGLOGO = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\PFI730206632\logo-pfizer.jpg";

                azul = new Color(13, 142, 244);
                blanco = new Color(255, 255, 255);
                Link = new Color(7, 73, 208);
                gris = new Color(13, 142, 244);
                grisOX = new Color(74, 74, 74);
                rojo = new Color(230, 7, 7);
                lbAzul = new Color(5, 146, 230);

                EM = BaseFont.CreateFont(@"C:\Windows\Fonts\VERDANA.TTF", BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
                f5 = new Font(EM, 5, Font.NORMAL, grisOX);
                f5B = new Font(EM, 5, Font.BOLD, grisOX);
                f5BBI = new Font(EM, 5, Font.BOLDITALIC, grisOX);
                f6 = new Font(EM, 6, Font.NORMAL, grisOX);
                f6B = new Font(EM, 6, Font.BOLD, grisOX);
                f6L = new Font(EM, 6, Font.BOLD, Link);
                f5L = new Font(EM, 5, Font.BOLD, lbAzul);
                titulo = new Font(EM, 6, Font.BOLD, blanco);
                folio = new Font(EM, 6, Font.BOLD, rojo);
                dSaltoLinea = new Chunk("\n\n ");

                #endregion

                #region "Generamos el Docuemto"

                string RFC = string.Empty;
                RFC = htDatosCfdi["rfcEmisor"].ToString();
                switch (RFC)
                {
                    case "PFI730206632":
                        #region "Docuemto PFI730206632"
                        switch (htDatosCfdi["serie"].ToString())
                        {
                            case "E":
                            case "G":
                                htDatosCfdi.Add("tipoDoc", "FACTURA");
                                formatoEGFH(document, electronicDocument, objTimbre, pageEventHandler, dtOpcEnc, dtOpcDet, htDatosCfdi, HTC);
                                break;
                            case "F":
                            case "H":
                                htDatosCfdi.Add("tipoDoc", "NOTA DE CRÉDITO");
                                formatoEGFH(document, electronicDocument, objTimbre, pageEventHandler, dtOpcEnc, dtOpcDet, htDatosCfdi, HTC);
                                break;
                            case "A":
                            case "C":
                            case "I":
                            case "K":
                            case "M":
                                htDatosCfdi.Add("tipoDoc", "FACTURA");
                                formatoABCDIJKLM(document, electronicDocument, objTimbre, pageEventHandler, idCfdi, dtOpcEnc, dtOpcDet, htDatosCfdi, HTC);
                                break;

                            case "B":
                            case "D":
                            case "J":
                            case "L":
                            case "N":
                                htDatosCfdi.Add("tipoDoc", "NOTA DE CRÉDITO");
                                formatoABCDIJKLM(document, electronicDocument, objTimbre, pageEventHandler, idCfdi, dtOpcEnc, dtOpcDet, htDatosCfdi, HTC);
                                break;

                            //case "N":
                            //    htDatosCfdi.Add("tipoDoc", "NOTA DE CRÉDITO");
                            //    formatoN(document, pageEventHandler, dtOpcEnc, dtOpcDet, htDatosCfdi, HTC);
                            //    break;

                            default:
                                break;
                        }
                        break;
                        #endregion
                    case "P&U960326AG7":
                        htDatosCfdi.Add("tipoDoc", "FACTURA");
                        PYU960326AG7(document, electronicDocument, objTimbre, pageEventHandler, dtOpcEnc, dtOpcDet, htDatosCfdi, HTC);
                        break;
                    case "PME970204V63":
                        #region "Documento PME970204V63"
                        switch (htDatosCfdi["serie"].ToString())
                        {
                            case "A":
                                htDatosCfdi.Add("tipoDoc", "FACTURA");
                                formatoABCDIJKLM(document, electronicDocument, objTimbre, pageEventHandler, idCfdi, dtOpcEnc, dtOpcDet, htDatosCfdi, HTC);
                                break;
                            case "B":
                                htDatosCfdi.Add("tipoDoc", "NOTA DE CRÉDITO");
                                formatoABCDIJKLM(document, electronicDocument, objTimbre, pageEventHandler, idCfdi, dtOpcEnc, dtOpcDet, htDatosCfdi, HTC);
                                break;
                        }
                        break;
                        #endregion
                    case "SP&040526HD3":
                        break;
                }

                #endregion

                document.Close();
                writer.Close();
                fs.Close();

                string filePdfExt = pathPdf.Replace(_rutaDocs, _rutaDocsExt);
                string urlPathFilePdf = filePdfExt.Replace(@"\", "/");

                //Subimos Archivo al Azure
                wAzure.azureUpDownLoad(1, pathPdf);

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

        #region "RFC P&U960326AG7"

        public static void PYU960326AG7(Document document, ElectronicDocument electronicDocument, Data objTimbre, pdfPageEventHandlerPfizer pageEventHandler, DataTable dtEncabezado, DataTable dtDetalle, Hashtable htCFDI, HttpContext hc)
        {
            try
            {
                //DAL dal = new DAL();

                #region "Construimos el Documento"

                #region "Construimos el Encabezado"

                Table encabezado = new Table(7);
                float[] headerwidthsEncabezado = { 9, 18, 28, 28, 5, 7, 5 };
                encabezado.Widths = headerwidthsEncabezado;
                encabezado.WidthPercentage = 100;
                encabezado.Padding = 1;
                encabezado.Spacing = 1;
                encabezado.BorderWidth = 0;
                encabezado.DefaultCellBorder = 0;
                encabezado.BorderColor = gris;

                Image imgLogo = Image.GetInstance(pathIMGLOGO);
                imgLogo.ScalePercent(47f);

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
                cel.Rowspan = 5;
                cel.Colspan = 2;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(8f, 9f);
                par.Add(new Chunk(htCFDI["nombreEmisor"].ToString().ToUpper(), f6B));
                par.Add(new Chunk("\nRFC " + htCFDI["rfcEmisor"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor1"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor2"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor3"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\nTel. (52) 55 5081-8500", f6));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                StringBuilder expedido = new StringBuilder();
                expedido.
                    Append("Lugar de Expedición México DF\n").
                    Append(htCFDI["sucursal"]).Append("\n").
                    Append(htCFDI["direccionExpedido1"].ToString().ToUpper()).Append("\n").
                    Append(htCFDI["direccionExpedido2"].ToString().ToUpper()).Append("\n").
                    Append(htCFDI["direccionExpedido3"].ToString().ToUpper()).Append("\n");

                cel = new Cell(new Phrase(expedido.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 4;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["tipoDoc"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["serie"].ToString().ToUpper() + electronicDocument.Data.Folio.Value, f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Día", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Mes", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Año", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                string[] fechaCFDI = Convert.ToDateTime(htCFDI["fechaCfdi"].ToString()).GetDateTimeFormats();
                string HORAS = fechaCFDI[103];
                string DIA = Convert.ToDateTime(htCFDI["fechaCfdi"]).Day.ToString();
                string MES = Convert.ToDateTime(htCFDI["fechaCfdi"]).ToString("MMMM").ToUpper();
                string ANIO = Convert.ToDateTime(htCFDI["fechaCfdi"]).Year.ToString();

                cel = new Cell(new Phrase(DIA, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(MES, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(ANIO, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                //cel = new Cell(new Phrase("No. y Año de Aprobación: " + dtEncabezado.Rows[0]["NoAp"].ToString() + " " + dtEncabezado.Rows[0]["AnoAp"].ToString(), f6));
                cel = new Cell(new Phrase("", f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                //cel = new Cell(new Phrase(HORAS, f6));
                cel = new Cell(new Phrase("", f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("CLIENTE", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Receptor.Nombre.Value, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Termino Pago: " + electronicDocument.Data.CondicionesPago.Value, f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("R.F.C. " + electronicDocument.Data.Receptor.Rfc.Value + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor1"] + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor2"] + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor3"] + "\n", f5));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 5;
                cel.Colspan = 2;
                encabezado.AddCell(cel);

                #region "Consignado a"

                string nombreEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-NOMBRE"].ToString();//CE
                string calleEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-CALLE"].ToString();//CE 10, 11, 12
                string coloniaEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-COLONIA"].ToString();//CE13
                string cpEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-CP"].ToString();//CE18
                string municEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-MUNIC"].ToString();//CE15
                string estadoEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-ESTADO"].ToString();//CE16
                string paisEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-PAIS"].ToString();//CE17
                string localidadEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-LOCAL"].ToString();//CE14
                string saltoEntregado = "\n";
                string separador = ", ";
                string espacio = " ";

                if (nombreEntregado.Length > 0)
                    nombreEntregado = nombreEntregado + saltoEntregado;
                else
                    nombreEntregado = "" + saltoEntregado;

                if (calleEntregado.Length > 0)
                    calleEntregado = calleEntregado + saltoEntregado;
                else
                    calleEntregado = "" + saltoEntregado;

                if (coloniaEntregado.Length > 0)
                    coloniaEntregado = coloniaEntregado + espacio;
                else
                    coloniaEntregado = "" + espacio;

                if (cpEntregado.Length > 0)
                    cpEntregado = ", CP " + cpEntregado + saltoEntregado;
                else
                    cpEntregado = "" + separador + saltoEntregado;

                if (municEntregado.Length > 0)
                    municEntregado = municEntregado + separador;
                else
                    municEntregado = "";

                if (localidadEntregado.Length > 0)
                    localidadEntregado = localidadEntregado + saltoEntregado;
                else
                    localidadEntregado = "" + saltoEntregado;

                if (estadoEntregado.Length > 0)
                    estadoEntregado = estadoEntregado + espacio;
                else
                    estadoEntregado = "" + espacio;

                if (paisEntregado.Length > 0)
                    paisEntregado = paisEntregado + "";
                else
                    paisEntregado = "";

                #endregion

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Consignado a:\n", f5B));
                par.Add(new Chunk(calleEntregado, f5));
                par.Add(new Chunk(coloniaEntregado, f5));
                par.Add(new Chunk(cpEntregado, f5));
                par.Add(new Chunk(municEntregado, f5));
                par.Add(new Chunk(localidadEntregado, f5));
                par.Add(new Chunk(estadoEntregado, f5));
                par.Add(new Chunk(paisEntregado, f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Cliente No:\n", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["clienteNo"].ToString(), f5));
                par.Add(new Chunk("\nContacto:\n", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["contacto"].ToString(), f5));
                par.Add(new Chunk("\nTeléfono:\n", f5B));
                par.Add(new Chunk("" + "\n", f5));

                //if (htCFDI["serie"].ToString() == "C" || htCFDI["serie"].ToString() == "D")
                //{
                //    par.Add(new Chunk("Zona:\n", f5B));
                //    par.Add(new Chunk(dtEncabezado.Rows[0]["zona"].ToString(), f5));
                //}

                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                //if (htCFDI["serie"].ToString() == "C")
                //{
                //    par = new Paragraph();
                //    par.SetLeading(7f, 1f);
                //    par.Add(new Chunk("Vencimiento: ", f5B));
                //    par.Add(new Chunk(dtEncabezado.Rows[0]["vencimiento"].ToString(), f5B)); //CE27
                //    cel = new Cell(par);
                //    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                //    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthRight = (float).5;
                //    cel.BorderWidthBottom = (float).5;
                //    cel.BorderColor = gris;
                //    cel.Colspan = 3;
                //    encabezado.AddCell(cel);
                //}

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Vencimiento: ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["vencimiento"].ToString(), f5B)); //CE27
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Moneda: ", f5B));
                par.Add(new Chunk(electronicDocument.Data.Moneda.Value, f5));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);

                //if (htCFDI["serie"].ToString() == "D" || htCFDI["serie"].ToString() == "L" || htCFDI["serie"].ToString() == "N")
                //    par.Add(new Chunk("Referencia: ", f5B));
                //else
                //    par.Add(new Chunk("Orden Compra: ", f5B));
                par.Add(new Chunk("Orden Compra: ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["referencia"].ToString(), f5));//CE33
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                //if (htCFDI["serie"].ToString() == "N")
                //{
                //    par = new Paragraph();
                //    par.SetLeading(7f, 1f);
                //    par.Add(new Chunk("Fecha Referencia: ", f5B));
                //    par.Add(new Chunk(dtEncabezado.Rows[0]["fechaRefDoc"].ToString(), f5));//CE34
                //    cel = new Cell(par);
                //    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                //    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthRight = (float).5;
                //    cel.BorderWidthBottom = 0;
                //    cel.BorderColor = gris;
                //    cel.Colspan = 3;
                //    encabezado.AddCell(cel);
                //}

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Pedido: ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["pedido"].ToString(), f5));//CE19
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("División: ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["division"].ToString(), f5));//CE20
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                #endregion

                #region "Construimos Tablas de Partidas"

                #region "Construimos Encabezados de Partidas"

                Table encabezadoPartidas = new Table(7);
                float[] headerwidthsEncabesadoPartidas = { 5, 10, 10, 40, 11, 12, 12 };
                encabezadoPartidas.Widths = headerwidthsEncabesadoPartidas;
                encabezadoPartidas.WidthPercentage = 100;
                encabezadoPartidas.Padding = 1;
                encabezadoPartidas.Spacing = 1;
                encabezadoPartidas.BorderWidth = (float).5;
                encabezadoPartidas.DefaultCellBorder = 1;
                encabezadoPartidas.BorderColor = gris;

                cel = new Cell(new Phrase("No.", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Código", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Descripción", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Cantidad", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Precio Unitario", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Importe", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                #endregion

                #region "Construimos Contenido de las Partidas"

                Table partidas = new Table(7);
                float[] headerwidthsPartidas = { 5, 10, 10, 40, 11, 12, 12 };
                partidas.Widths = headerwidthsPartidas;
                partidas.WidthPercentage = 100;
                partidas.Padding = 1;
                partidas.Spacing = 1;
                partidas.BorderWidth = 0;
                partidas.DefaultCellBorder = 0;
                partidas.BorderColor = gris;

                if (dtEncabezado.Rows.Count > 0)
                {
                    for (int i = 0; i < electronicDocument.Data.Conceptos.Count; i++)
                    {
                        cel = new Cell(new Phrase((i + 1).ToString(), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 2;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["codeLocal"].ToString(), f5));//CD1
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 2;
                        partidas.AddCell(cel);

                        #region "Descripción"

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        cel.Colspan = 2;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Cantidad.Value + " " + electronicDocument.Data.Conceptos[i].Unidad.Value, f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 2;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].ValorUnitario.Value.ToString("C", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 2;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("C", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = (float).5;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 2;
                        partidas.AddCell(cel);


                        par = new Paragraph();
                        par.SetLeading(7f, 1f);
                        par.Add(new Chunk("Lote\n", f5L));
                        par.Add(new Chunk(dtDetalle.Rows[i]["lote"].ToString().Replace("*", "\n"), f5));//CD3
                        cel = new Cell(par);
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        par = new Paragraph();
                        par.SetLeading(7f, 1f);
                        par.Add(new Chunk("Cantidad\n", f5L));
                        par.Add(new Chunk(dtDetalle.Rows[i]["cantidad"].ToString().Replace("*", "\n"), f5));//CD4
                        cel = new Cell(par);
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        #endregion
                    }
                }

                #endregion

                #endregion

                #region "Construimos el Comentarios"

                Table comentarios = new Table(7);
                float[] headerwidthsComentarios = { 9, 18, 28, 28, 7, 5, 5 };
                comentarios.Widths = headerwidthsComentarios;
                comentarios.WidthPercentage = 100;
                comentarios.Padding = 1;
                comentarios.Spacing = 1;
                comentarios.BorderWidth = 0;
                comentarios.DefaultCellBorder = 0;
                comentarios.BorderColor = gris;

                cel = new Cell(new Phrase("Cantidad:", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Rowspan = 3;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["cantidadLetra"].ToString(), f5));//CE26  
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                cel.Rowspan = 3;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Sub Total", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                double importe = 0;
                double tasa = 0;

                for (int i = 0; i < electronicDocument.Data.Impuestos.Traslados.Count; i++)
                {
                    if (electronicDocument.Data.Impuestos.Traslados[i].Tipo.Value == "IVA")
                    {
                        importe = electronicDocument.Data.Impuestos.Traslados[i].Importe.Value;
                        tasa = electronicDocument.Data.Impuestos.Traslados[i].Tasa.Value;
                        break;
                    }
                }

                cel = new Cell(new Phrase("IVA " + tasa + " %", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("C", _ci), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Total", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _ci), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Observaciones:", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Rowspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["observaciones"].ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                cel.Rowspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Descuento Pronto Pago", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Descuento.Value.ToString("C", _ci), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Importe Pago Neto", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Descuento.Value.ToString("C", _ci), f5));//Duda
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.FormaPago.Value.ToUpper(), f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 7;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 7;
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
                    par.Add(new Chunk("SELLO DIGITAL DEL EMISOR\n", f5L));
                    par.Add(new Chunk(electronicDocument.Data.Sello.Value, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
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
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase("FECHA Y HORA DE CERTIFICACION:", f5L));
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

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL SAT:", f5L));
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

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL EMISOR:", f5L));
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

                    string lbRegimen = "REGIMEN FISCAL APLICABLE: ";
                    StringBuilder regimenes = new StringBuilder();
                    if (electronicDocument.Data.Emisor.Regimenes.IsAssigned)
                    {
                        for (int i = 0; i < electronicDocument.Data.Emisor.Regimenes.Count; i++)
                        {
                            regimenes.Append(electronicDocument.Data.Emisor.Regimenes[i].Regimen.Value).Append("\n");
                        }
                    }
                    else
                    {
                        regimenes.Append("");
                        lbRegimen = "";
                    }

                    string cuenta = electronicDocument.Data.NumeroCuentaPago.IsAssigned
                                    ? electronicDocument.Data.NumeroCuentaPago.Value
                                    : "";

                    par = new Paragraph();
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("TIPO DE COMPROBANTE: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.TipoComprobante.Value + "   |   ", f5));
                    par.Add(new Chunk("Moneda: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.Moneda.Value + "   |   ", f5));
                    par.Add(new Chunk("TASA DE CAMBIO: ", f5L));
                    string tasaCambio = electronicDocument.Data.TipoCambio.Value;

                    if (tasaCambio.Length > 0)
                    {
                        par.Add(new Chunk(Convert.ToDouble(tasaCambio).ToString("C", _ci) + "   |   ", f5));
                    }
                    else
                    {
                        par.Add(new Chunk("   |   ", f5));
                    }
                    par.Add(new Chunk("FORMA DE PAGO: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.FormaPago.Value + "\n", f5));
                    par.Add(new Chunk("MÉTODO DE PAGO: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.MetodoPago.Value + "   |   ", f5));
                    par.Add(new Chunk("NÚMERO DE CUENTA: ", f5L));
                    par.Add(new Chunk(cuenta + "   |   ", f5));
                    par.Add(new Chunk(lbRegimen, f5L));
                    par.Add(new Chunk(regimenes.ToString(), f5));

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
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
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

                #region "Construimos Tabla Adicional Datos de Pago"

                Table pago = new Table(5);
                float[] headerwidthsPago = { 40, 10, 20, 10, 20 };
                pago.Widths = headerwidthsPago;
                pago.WidthPercentage = 100;
                pago.Padding = 1;
                pago.Spacing = 1;
                pago.BorderWidth = (float).5;
                pago.DefaultCellBorder = 1;
                pago.BorderColor = gris;


                cel = new Cell(new Phrase("Para ejecutar su pago para consultas", f5L));
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Para efectuar su pago", f5L));
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Información del Cliente", f5L));
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                pago.AddCell(cel);

                string[] ejecutarPago = dtEncabezado.Rows[0]["paraConsultas"].ToString().Split(new Char[] { '*' });//CE23
                string[] pagoDatos = dtEncabezado.Rows[0]["efectuarPago"].ToString().Split(new Char[] { '*' });//CE21

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (ejecutarPago.Length > 0)
                    par.Add(new Chunk(ejecutarPago[0] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                if (ejecutarPago.Length > 1)
                    par.Add(new Chunk(ejecutarPago[1] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                if (ejecutarPago.Length > 2)
                    par.Add(new Chunk(ejecutarPago[2] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                if (ejecutarPago.Length > 3)
                    par.Add(new Chunk(ejecutarPago[3] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                cel.Rowspan = 6;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Banco:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 0)
                    par.Add(new Chunk(pagoDatos[0], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Cliente:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Receptor.Nombre.Value, f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Cuenta:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 1)
                    par.Add(new Chunk(pagoDatos[1], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("No. Cliente:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["clienteNo"].ToString(), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                //if (htCFDI["serie"].ToString() == "B" || htCFDI["serie"].ToString() == "D" || htCFDI["serie"].ToString() == "J" || htCFDI["serie"].ToString() == "L" || htCFDI["serie"].ToString() == "N")
                //{
                //    cel = new Cell(new Phrase("", f5));
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthRight = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthBottom = 0;
                //    cel.Colspan = 2;
                //    pago.AddCell(cel);
                //}

                //else
                //{
                //    cel = new Cell(new Phrase("SWIFT:", f5L));
                //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthRight = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthBottom = 0;
                //    pago.AddCell(cel);

                //    cel = new Cell(new Phrase(pagoDatos[3], f5));
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthRight = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthBottom = 0;
                //    pago.AddCell(cel);
                //}

                cel = new Cell(new Phrase("", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Factura No:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["serie"].ToString().ToUpper() + "-" + electronicDocument.Data.Folio.Value.ToString(), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Moneda:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 4)
                    par.Add(new Chunk(pagoDatos[4], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Fecha Factura", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                string[] fechaFactura = Convert.ToDateTime(objTimbre.FechaTimbrado.Value).GetDateTimeFormats('s');

                cel = new Cell(new Phrase(DIA + "/" + MES + "/" + ANIO, f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Beneficiario:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 5)
                    par.Add(new Chunk(pagoDatos[5], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Valor Total:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _ci), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Dirección:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 4)
                    par.Add(new Chunk(dtEncabezado.Rows[0]["direccionPie"].ToString(), f5));//CE22
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                pago.AddCell(cel);

                //if (htCFDI["serie"].ToString() == "C")
                //{
                //    cel = new Cell(new Phrase("Fecha Vencimiento:", f5L));
                //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                //    cel.BorderColor = gris;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthRight = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthBottom = (float).5;
                //    pago.AddCell(cel);

                //    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["vencimiento"].ToString(), f5B));
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthRight = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthBottom = (float).5;
                //    cel.BorderColor = gris;
                //    pago.AddCell(cel);
                //}
                //else
                //{
                //    cel = new Cell(new Phrase("Moneda:", f5L));
                //    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                //    cel.BorderColor = gris;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthRight = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthBottom = (float).5;
                //    pago.AddCell(cel);

                //    cel = new Cell(new Phrase(electronicDocument.Data.Moneda.Value, f5));
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthRight = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthBottom = (float).5;
                //    cel.BorderColor = gris;
                //    pago.AddCell(cel);
                //}

                cel = new Cell(new Phrase("Moneda:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Moneda.Value, f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                pago.AddCell(cel);

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
                pageEventHandler.encPartidas = encabezadoPartidas;
                pageEventHandler.footer = footer;

                document.Open();

                document.Add(partidas);
                document.Add(comentarios);
                document.Add(adicional);
                document.Add(pago);

                #endregion
            }
            catch (Exception ex)
            {
                string exe = ex.Message;
            }
        }

        #endregion

        #region "formatoEGFH"

        public static void formatoEGFH(Document document, ElectronicDocument electronicDocument, Data objTimbre, pdfPageEventHandlerPfizer pageEventHandler, DataTable dtEncabezado, DataTable dtDetalle, Hashtable htCFDI, HttpContext hc)
        {
            try
            {
                DAL dal = new DAL();

                #region "Construimos el Documento"

                #region "Construimos el Encabezado"

                Table encabezado = new Table(7);
                float[] headerwidthsEncabezado = { 9, 18, 28, 28, 5, 7, 5 };
                encabezado.Widths = headerwidthsEncabezado;
                encabezado.WidthPercentage = 100;
                encabezado.Padding = 1;
                encabezado.Spacing = 1;
                encabezado.BorderWidth = 0;
                encabezado.DefaultCellBorder = 0;
                encabezado.BorderColor = gris;

                //Agregando Imagen de Logotipo
                Image imgLogo = Image.GetInstance(pathIMGLOGO);
                imgLogo.ScalePercent(47f);

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
                cel.Rowspan = 5;
                cel.Colspan = 2;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(8f, 9f);
                par.Add(new Chunk(htCFDI["nombreEmisor"].ToString().ToUpper(), f6B));
                par.Add(new Chunk("\nRFC " + htCFDI["rfcEmisor"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor1"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor2"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor3"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\nTel. (52) 55 5081-8500", f6));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                StringBuilder expedido = new StringBuilder();
                expedido.
                    Append("Lugar de Expedición México DF\n").
                    Append(htCFDI["sucursal"]).Append("\n").
                    Append(htCFDI["direccionExpedido1"].ToString().ToUpper()).Append("\n").
                    Append(htCFDI["direccionExpedido2"].ToString().ToUpper()).Append("\n").
                    Append(htCFDI["direccionExpedido3"].ToString().ToUpper()).Append("\n").
                    Append("Corporate Code: " + dtEncabezado.Rows[0]["corporateCode"]).Append("\n").//CE29
                    Append("Old Corporate Code: " + dtEncabezado.Rows[0]["oldCorporateCode"].ToString()).Append("\n");//CE30;

                cel = new Cell(new Phrase(expedido.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 4;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["tipoDoc"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["serie"].ToString().ToUpper() + electronicDocument.Data.Folio.Value, f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Día", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Mes", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Año", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                string[] fechaCFDI = Convert.ToDateTime(htCFDI["fechaCfdi"].ToString()).GetDateTimeFormats();
                string HORAS = fechaCFDI[103];
                string DIA = Convert.ToDateTime(htCFDI["fechaCfdi"]).Day.ToString();
                string MES = Convert.ToDateTime(htCFDI["fechaCfdi"]).ToString("MMMM").ToUpper();
                string ANIO = Convert.ToDateTime(htCFDI["fechaCfdi"]).Year.ToString();

                cel = new Cell(new Phrase(DIA, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(MES, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(ANIO, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("No. y Año de Aprobación: " + dtEncabezado.Rows[0]["NoAp"] + " " + dtEncabezado.Rows[0]["AnoAp"], f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(HORAS, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("CLIENTE", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Receptor.Nombre.Value, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Termino Pago: " + electronicDocument.Data.CondicionesPago.Value, f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("R.F.C. " + electronicDocument.Data.Receptor.Rfc.Value + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor1"] + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor2"] + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor3"] + "\n", f5));
                par.Add(new Chunk("Código Corporativo: " + dtEncabezado.Rows[0]["codigoCorporativo"] + "\n" + "\n", f5));//36
                par.Add(new Chunk("Código Corporativo Anterior: " + dtEncabezado.Rows[0]["codigoCorporativoAnt"].ToString(), f5));//37
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 5;
                cel.Colspan = 2;
                encabezado.AddCell(cel);

                #region "Consignado a"

                string nombreEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-NOMBRE"].ToString();//CE
                string calleEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-CALLE"].ToString();//CE 10, 11, 12
                string coloniaEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-COLONIA"].ToString();//CE13
                string cpEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-CP"].ToString();//CE18
                string municEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-MUNIC"].ToString();//CE15
                string estadoEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-ESTADO"].ToString();//CE16
                string paisEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-PAIS"].ToString();//CE17
                string localidadEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-LOCAL"].ToString();//CE14
                string intericom = dtEncabezado.Rows[0]["INTERICOM"].ToString().ToUpper();//CE32
                string saltoEntregado = "\n";
                string separador = ", ";
                string espacio = " ";

                if (nombreEntregado.Length > 0)
                    nombreEntregado = nombreEntregado + saltoEntregado;
                else
                    nombreEntregado = "" + saltoEntregado;

                if (calleEntregado.Length > 0)
                    calleEntregado = calleEntregado + saltoEntregado;
                else
                    calleEntregado = "" + saltoEntregado;

                if (coloniaEntregado.Length > 0)
                    coloniaEntregado = coloniaEntregado + espacio;
                else
                    coloniaEntregado = "" + espacio;

                if (cpEntregado.Length > 0)
                    cpEntregado = ", CP " + cpEntregado + saltoEntregado;
                else
                    cpEntregado = "" + separador + saltoEntregado;

                if (municEntregado.Length > 0)
                    municEntregado = municEntregado + separador;
                else
                    municEntregado = "";

                if (localidadEntregado.Length > 0)
                    localidadEntregado = localidadEntregado + saltoEntregado;
                else
                    localidadEntregado = "" + saltoEntregado;

                if (estadoEntregado.Length > 0)
                    estadoEntregado = estadoEntregado + espacio;
                else
                    estadoEntregado = "" + espacio;

                if (paisEntregado.Length > 0)
                    paisEntregado = paisEntregado + "";
                else
                    paisEntregado = "";

                if (intericom.Length > 0)
                    intericom = intericom + espacio;
                else
                    intericom = "";

                #endregion

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Consignado a:\n", f5B));
                par.Add(new Chunk(calleEntregado, f5));
                par.Add(new Chunk(coloniaEntregado, f5));
                par.Add(new Chunk(cpEntregado, f5));
                par.Add(new Chunk(municEntregado, f5));
                par.Add(new Chunk(localidadEntregado, f5));
                par.Add(new Chunk(estadoEntregado, f5));
                par.Add(new Chunk(paisEntregado, f5));
                par.Add(new Chunk("\nINCOTERM: " + intericom, f5));//CE32
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Contacto:\n", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["contacto"].ToString(), f5));
                par.Add(new Chunk("\nTeléfono:\n", f5B));
                par.Add(new Chunk("" + "\n\n", f5));
                par.Add(new Chunk("Método de Pago:\n", f5B));
                par.Add(new Chunk(electronicDocument.Data.MetodoPago.Value, f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Hyperion Acc: ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["hyperion"].ToString(), f5));//CE31 
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Moneda: ", f5B));
                par.Add(new Chunk(electronicDocument.Data.Moneda.Value, f5));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Referencia: ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["referencia"].ToString(), f5));//CE33
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Pedido: ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["pedido"].ToString(), f5));//CE19
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("División: ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["division"].ToString(), f5));//CE20
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                #endregion

                #region "Construimos Tablas de Partidas"

                #region "Construimos Encabezados de Partidas"

                Table encabezadoPartidas = new Table(14);
                float[] headerwidthsEncabesadoPartidas = { 4, 6, 6, 6, 7, 6, 7, 8, 8, 8, 9, 6, 9, 10 };
                encabezadoPartidas.Widths = headerwidthsEncabesadoPartidas;
                encabezadoPartidas.WidthPercentage = 100;
                encabezadoPartidas.Padding = 1;
                encabezadoPartidas.Spacing = 1;
                encabezadoPartidas.BorderWidth = (float).5;
                encabezadoPartidas.DefaultCellBorder = 1;
                encabezadoPartidas.BorderColor = gris;

                cel = new Cell(new Phrase("No.", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Código Local", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Código Oracle", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Código ISPC", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Código Impuesto", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Centro Costos", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("# Clinicos", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("# Proyecto", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Descripción", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 3;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Cantidad", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Precio Unitario", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Importe", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                #endregion

                #region "Construimos Contenido de las Partidas"

                Table partidas = new Table(14);
                float[] headerwidthsPartidas = { 4, 6, 6, 6, 7, 6, 7, 8, 8, 8, 9, 6, 9, 10 };
                partidas.Widths = headerwidthsPartidas;
                partidas.WidthPercentage = 100;
                partidas.Padding = 1;
                partidas.Spacing = 1;
                partidas.BorderWidth = (float).5;
                partidas.DefaultCellBorder = 1;
                partidas.BorderColor = gris;

                if (dtEncabezado.Rows.Count > 0)
                {
                    for (int i = 0; i < electronicDocument.Data.Conceptos.Count; i++)
                    {
                        cel = new Cell(new Phrase((i + 1).ToString(), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["codeLocal"].ToString(), f5));//CD1
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["codeOracle"].ToString(), f5));//CD10
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["codeISPC"].ToString(), f5));//CD11
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["codeImpuesto"].ToString(), f5));//CD12
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["centroCostos"].ToString(), f5));//CD13
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["clinico"].ToString(), f5));//CD14
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["proyecto"].ToString(), f5));//CD15
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        #region "Descripción"

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        cel.Colspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Cantidad.Value + " " + electronicDocument.Data.Conceptos[i].Unidad.Value, f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].ValorUnitario.Value.ToString("C", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("C", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        cel.Rowspan = 3;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase("Lote", f5L));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase("Cantidad", f5L));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase("Expiración", f5L));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["lote"].ToString().Replace("*", "\n"), f5));//CD3
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["cantidad"].ToString().Replace("*", "\n"), f5));//CD4
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["expiracion"].ToString().Replace("*", "\n"), f5));//CD5
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        #endregion
                    }
                }

                #endregion

                #endregion

                #region "Construimos el Comentarios"

                Table comentarios = new Table(7);
                float[] headerwidthsComentarios = { 9, 18, 28, 28, 7, 5, 5 };
                comentarios.Widths = headerwidthsComentarios;
                comentarios.WidthPercentage = 100;
                comentarios.Padding = 1;
                comentarios.Spacing = 1;
                comentarios.BorderWidth = 0;
                comentarios.DefaultCellBorder = 0;
                comentarios.BorderColor = gris;

                cel = new Cell(new Phrase("Cantidad:", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Rowspan = 3;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["cantidadLetra"].ToString(), f5));//CE26  
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                cel.Rowspan = 3;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Sub Total", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                double importe = 0;
                double tasa = 0;

                for (int i = 0; i < electronicDocument.Data.Impuestos.Traslados.Count; i++)
                {
                    if (electronicDocument.Data.Impuestos.Traslados[i].Tipo.Value == "IVA")
                    {
                        importe = electronicDocument.Data.Impuestos.Traslados[i].Importe.Value;
                        tasa = electronicDocument.Data.Impuestos.Traslados[i].Tasa.Value;
                        break;
                    }
                }

                cel = new Cell(new Phrase("IVA " + tasa + " %", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("C", _ci), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Total", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _ci), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Observaciones:", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["observaciones"].ToString(), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 6;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 7;
                comentarios.AddCell(cel);

                #endregion

                #region "Construimos Tabla Especial"
                Table especial = new Table(5);
                float[] headerwidthsEspecial = { 20, 20, 20, 20, 20 };
                especial.Widths = headerwidthsEspecial;
                especial.WidthPercentage = 100;
                especial.Padding = 1;
                especial.Spacing = 1;
                especial.BorderWidth = 0;
                especial.DefaultCellBorder = 0;
                especial.BorderColor = gris;

                cel = new Cell(new Phrase("TAX SUMMARY", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("TAX RATE", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("TAXABLE AMOUNT", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("TAX PAID", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("TOTAL AMOUNT", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                Table detalleImpuestos = new Table(5);
                float[] headerwidthsDetalleImpuestos = { 20, 20, 20, 20, 20 };
                detalleImpuestos.Widths = headerwidthsDetalleImpuestos;
                detalleImpuestos.WidthPercentage = 100;
                detalleImpuestos.Padding = 1;
                detalleImpuestos.Spacing = 1;
                detalleImpuestos.BorderWidth = 0;
                detalleImpuestos.DefaultCellBorder = 0;
                detalleImpuestos.BorderColor = gris;

                if (dtDetalle.Rows.Count > 0)
                {
                    for (int j = 0; j < electronicDocument.Data.Conceptos.Count; j++)
                    {
                        cel = new Cell(new Phrase("", f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = (float).5;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        detalleImpuestos.AddCell(cel);

                        cel = new Cell(new Phrase(Convert.ToDouble(dtDetalle.Rows[j]["taxRate"]).ToString("N", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = (float).5;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = (float).5;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        detalleImpuestos.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[j].Importe.Value.ToString("C", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = (float).5;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = (float).5;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        detalleImpuestos.AddCell(cel);

                        cel = new Cell(new Phrase(Convert.ToDouble(dtDetalle.Rows[j]["taxPaid"]).ToString("N", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = (float).5;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = (float).5;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        detalleImpuestos.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[j].Importe.Value.ToString("C", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = (float).5;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = (float).5;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        detalleImpuestos.AddCell(cel);
                    }
                }

                Table especialTotal = new Table(5);
                float[] headerwidthsEspeciaTootal = { 20, 20, 20, 20, 20 };
                especialTotal.Widths = headerwidthsEspeciaTootal;
                especialTotal.WidthPercentage = 100;
                especialTotal.Padding = 1;
                especialTotal.Spacing = 1;
                especialTotal.BorderWidth = 0;
                especialTotal.DefaultCellBorder = 0;
                especialTotal.BorderColor = gris;

                cel = new Cell(new Phrase("", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                especialTotal.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                especialTotal.AddCell(cel);

                cel = new Cell(new Phrase(importe.ToString("C", _ci), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especialTotal.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _ci), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especialTotal.AddCell(cel);

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
                    par.Add(new Chunk("SELLO DIGITAL DEL EMISOR\n", f5L));
                    par.Add(new Chunk(electronicDocument.Data.Sello.Value, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
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
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase("FECHA Y HORA DE CERTIFICACION:", f5L));
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

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL SAT:", f5L));
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

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL EMISOR:", f5L));
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

                    string lbRegimen = "REGIMEN FISCAL APLICABLE: ";
                    StringBuilder regimenes = new StringBuilder();
                    if (electronicDocument.Data.Emisor.Regimenes.IsAssigned)
                    {
                        for (int i = 0; i < electronicDocument.Data.Emisor.Regimenes.Count; i++)
                        {
                            regimenes.Append(electronicDocument.Data.Emisor.Regimenes[i].Regimen.Value).Append("\n");
                        }
                    }
                    else
                    {
                        regimenes.Append("");
                        lbRegimen = "";
                    }

                    string cuenta = electronicDocument.Data.NumeroCuentaPago.IsAssigned
                                    ? electronicDocument.Data.NumeroCuentaPago.Value
                                    : "";

                    par = new Paragraph();
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("TIPO DE COMPROBANTE: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.TipoComprobante.Value + "   |   ", f5));
                    par.Add(new Chunk("Moneda: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.Moneda.Value + "   |   ", f5));
                    par.Add(new Chunk("TASA DE CAMBIO: ", f5L));
                    string tasaCambio = electronicDocument.Data.TipoCambio.Value;
                    if (tasaCambio.Length > 0)
                    {
                        par.Add(new Chunk(Convert.ToDouble(tasaCambio).ToString("C", _ci) + "   |   ", f5));
                    }
                    else
                    {
                        par.Add(new Chunk("   |   ", f5));
                    }
                    par.Add(new Chunk("FORMA DE PAGO: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.FormaPago.Value + "\n", f5));
                    par.Add(new Chunk("MÉTODO DE PAGO: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.MetodoPago.Value + "   |   ", f5));
                    par.Add(new Chunk("NÚMERO DE CUENTA: ", f5L));
                    par.Add(new Chunk(cuenta + "   |   ", f5));
                    par.Add(new Chunk(lbRegimen, f5L));
                    par.Add(new Chunk(regimenes.ToString(), f5));

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
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
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

                #region "Construimos Tabla Adicional Datos de Pago"

                Table pago = new Table(5);
                float[] headerwidthsPago = { 40, 10, 20, 10, 20 };
                pago.Widths = headerwidthsPago;
                pago.WidthPercentage = 100;
                pago.Padding = 1;
                pago.Spacing = 1;
                pago.BorderWidth = (float).5;
                pago.DefaultCellBorder = 1;
                pago.BorderColor = gris;


                cel = new Cell(new Phrase("Para ejecutar su pago para consultas", f5L));
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Para efectuar su pago", f5L));
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Información del Cliente", f5L));
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                pago.AddCell(cel);

                string[] ejecutarPago = dtEncabezado.Rows[0]["paraConsultas"].ToString().Split(new Char[] { '*' });//CE23
                string[] pagoDatos = dtEncabezado.Rows[0]["efectuarPago"].ToString().Split(new Char[] { '*' });//CE21

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (ejecutarPago.Length > 0)
                    par.Add(new Chunk(ejecutarPago[0] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                if (ejecutarPago.Length > 1)
                    par.Add(new Chunk(ejecutarPago[1] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                if (ejecutarPago.Length > 2)
                    par.Add(new Chunk(ejecutarPago[2] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                if (ejecutarPago.Length > 3)
                    par.Add(new Chunk(ejecutarPago[3] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                cel.Rowspan = 6;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Banco:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 0)
                    par.Add(new Chunk(pagoDatos[0], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Cliente:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Receptor.Nombre.Value, f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Cuenta:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 1)
                    par.Add(new Chunk(pagoDatos[1], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("No. Cliente:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["clienteNo"].ToString(), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Factura No:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["serie"].ToString().ToUpper() + htCFDI["folio"], f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Moneda:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 4)
                    par.Add(new Chunk(pagoDatos[4], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Fecha Factura:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                string[] fechaFactura = Convert.ToDateTime(objTimbre.FechaTimbrado.Value).GetDateTimeFormats('s');

                cel = new Cell(new Phrase(DIA + "/" + MES + "/" + ANIO, f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Beneficiario:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 5)
                    par.Add(new Chunk(pagoDatos[5], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Valor Total:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _ci), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Dirección:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 4)
                    par.Add(new Chunk(dtEncabezado.Rows[0]["direccionPie"].ToString(), f5));//CE22
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Moneda:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Moneda.Value, f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                pago.AddCell(cel);
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
                pageEventHandler.encPartidas = encabezadoPartidas;
                pageEventHandler.footer = footer;

                document.Open();

                document.Add(partidas);
                document.Add(comentarios);
                document.Add(especial);
                document.Add(detalleImpuestos);
                document.Add(especialTotal);
                document.Add(adicional);
                document.Add(pago);

                #endregion
            }
            catch (Exception ex)
            {
                string exe = ex.Message;
            }
        }

        #endregion

        #region "formatoABCDIJKLM"

        public static void formatoABCDIJKLM(Document document, ElectronicDocument electronicDocument, Data objTimbre, pdfPageEventHandlerPfizer pageEventHandler, Int64 idCfdi, DataTable dtEncabezado, DataTable dtDetalle, Hashtable htCFDI, HttpContext hc)
        {
            try
            {
                DAL dal = new DAL();
                #region "Construimos el Documento"

                #region "Construimos el Encabezado"

                Table encabezado = new Table(7);
                float[] headerwidthsEncabezado = { 9, 18, 28, 28, 5, 7, 5 };
                encabezado.Widths = headerwidthsEncabezado;
                encabezado.WidthPercentage = 100;
                encabezado.Padding = 1;
                encabezado.Spacing = 1;
                encabezado.BorderWidth = 0;
                encabezado.DefaultCellBorder = 0;
                encabezado.BorderColor = gris;

                //Agregando Imagen de Logotipo
                Image imgLogo = Image.GetInstance(pathIMGLOGO);
                imgLogo.ScalePercent(47f);

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
                cel.Rowspan = 5;
                cel.Colspan = 2;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(8f, 9f);
                par.Add(new Chunk(htCFDI["nombreEmisor"].ToString().ToUpper(), f6B));
                par.Add(new Chunk("\nRFC " + htCFDI["rfcEmisor"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor1"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor2"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor3"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\nTel. (52) 55 5081-8500", f6));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                StringBuilder expedido = new StringBuilder();
                expedido.
                    Append("Lugar de Expedición México DF\n").
                    Append(htCFDI["sucursal"]).Append("\n").
                    Append(htCFDI["direccionExpedido1"].ToString().ToUpper()).Append("\n").
                    Append(htCFDI["direccionExpedido2"].ToString().ToUpper()).Append("\n").
                    Append(htCFDI["direccionExpedido3"].ToString().ToUpper()).Append("\n");

                cel = new Cell(new Phrase(expedido.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 4;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["tipoDoc"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["serie"].ToString().ToUpper() + electronicDocument.Data.Folio.Value, f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Día", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Mes", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Año", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                string[] fechaCFDI = Convert.ToDateTime(htCFDI["fechaCfdi"].ToString()).GetDateTimeFormats();
                string HORAS = fechaCFDI[103];
                string DIA = Convert.ToDateTime(htCFDI["fechaCfdi"]).Day.ToString();
                string MES = Convert.ToDateTime(htCFDI["fechaCfdi"]).ToString("MMMM").ToUpper();
                string ANIO = Convert.ToDateTime(htCFDI["fechaCfdi"]).Year.ToString();

                cel = new Cell(new Phrase(DIA, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(MES, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(ANIO, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("No. y Año de Aprobación: " + dtEncabezado.Rows[0]["NoAp"].ToString() + " " + dtEncabezado.Rows[0]["AnoAp"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(HORAS, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("CLIENTE", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Receptor.Nombre.Value, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Termino Pago: " + electronicDocument.Data.CondicionesPago.Value, f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("R.F.C. " + electronicDocument.Data.Receptor.Rfc.Value + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor1"] + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor2"] + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor3"] + "\n", f5));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 5;
                cel.Colspan = 2;
                encabezado.AddCell(cel);

                #region "Consignado a"

                string nombreEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-NOMBRE"].ToString();//CE
                string calleEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-CALLE"].ToString();//CE 10, 11, 12
                string coloniaEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-COLONIA"].ToString();//CE13
                string cpEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-CP"].ToString();//CE18
                string municEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-MUNIC"].ToString();//CE15
                string estadoEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-ESTADO"].ToString();//CE16
                string paisEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-PAIS"].ToString();//CE17
                string localidadEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-LOCAL"].ToString();//CE14
                string saltoEntregado = "\n";
                string separador = ", ";
                string espacio = " ";

                if (nombreEntregado.Length > 0)
                    nombreEntregado = nombreEntregado + saltoEntregado;
                else
                    nombreEntregado = "" + saltoEntregado;

                if (calleEntregado.Length > 0)
                    calleEntregado = calleEntregado + saltoEntregado;
                else
                    calleEntregado = "" + saltoEntregado;

                if (coloniaEntregado.Length > 0)
                    coloniaEntregado = coloniaEntregado + espacio;
                else
                    coloniaEntregado = "" + espacio;

                if (cpEntregado.Length > 0)
                    cpEntregado = ", CP " + cpEntregado + saltoEntregado;
                else
                    cpEntregado = "" + separador + saltoEntregado;

                if (municEntregado.Length > 0)
                    municEntregado = municEntregado + separador;
                else
                    municEntregado = "";

                if (localidadEntregado.Length > 0)
                    localidadEntregado = localidadEntregado + saltoEntregado;
                else
                    localidadEntregado = "" + saltoEntregado;

                if (estadoEntregado.Length > 0)
                    estadoEntregado = estadoEntregado + espacio;
                else
                    estadoEntregado = "" + espacio;

                if (paisEntregado.Length > 0)
                    paisEntregado = paisEntregado + "";
                else
                    paisEntregado = "";

                #endregion

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Consignado a:\n", f5B));
                par.Add(new Chunk(calleEntregado, f5));
                par.Add(new Chunk(coloniaEntregado, f5));
                par.Add(new Chunk(cpEntregado, f5));
                par.Add(new Chunk(municEntregado, f5));
                par.Add(new Chunk(localidadEntregado, f5));
                par.Add(new Chunk(estadoEntregado, f5));
                par.Add(new Chunk(paisEntregado, f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Cliente No:\n", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["clienteNo"].ToString(), f5));
                par.Add(new Chunk("\nContacto:\n", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["contacto"].ToString(), f5));
                par.Add(new Chunk("\nTeléfono:\n", f5B));
                par.Add(new Chunk("" + "\n", f5));

                if (htCFDI["serie"].ToString() == "C" || htCFDI["serie"].ToString() == "D")
                {
                    par.Add(new Chunk("Zona:\n", f5B));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["zona"].ToString(), f5));
                }

                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                if (htCFDI["serie"].ToString() == "C")
                {
                    par = new Paragraph();
                    par.SetLeading(7f, 1f);
                    par.Add(new Chunk("Vencimiento: ", f5B));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["vencimiento"].ToString(), f5B)); //CE27
                    cel = new Cell(par);
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    cel.Colspan = 3;
                    encabezado.AddCell(cel);
                }

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Moneda: ", f5B));
                par.Add(new Chunk(electronicDocument.Data.Moneda.Value, f5));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);

                if (htCFDI["serie"].ToString() == "D" || htCFDI["serie"].ToString() == "L" || htCFDI["serie"].ToString() == "N")
                    par.Add(new Chunk("Referencia: ", f5B));
                else
                    par.Add(new Chunk("Orden Compra: ", f5B));

                par.Add(new Chunk(dtEncabezado.Rows[0]["referencia"].ToString(), f5));//CE33
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                if (htCFDI["serie"].ToString() == "N")
                {
                    par = new Paragraph();
                    par.SetLeading(7f, 1f);
                    par.Add(new Chunk("Fecha Referencia: ", f5B));
                    par.Add(new Chunk(dtEncabezado.Rows[0]["fechaRefDoc"].ToString(), f5));//CE34
                    cel = new Cell(par);
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    cel.Colspan = 3;
                    encabezado.AddCell(cel);
                }

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Pedido: ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["pedido"].ToString(), f5));//CE19
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("División: ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["division"].ToString(), f5));//CE20
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                #endregion

                #region "Construimos Tablas de Partidas"

                #region "Añadimos Detalle de Conceptos"

                // Creamos la tabla para insertar los conceptos de detalle de la factura

                StringBuilder sbOpcionalDetalle = new StringBuilder();
                sbOpcionalDetalle.
                    Append("SELECT ROW_NUMBER() OVER (ORDER BY idCfdi ASC) AS numero, ").
                    Append("campo1 AS codigoProducto, campo18 AS codigoBarras, ").
                    Append("campo3 AS lote, campo5 AS fechaLote, campo4 AS cantidadLote ").
                    Append("FROM opcionalDetalle ").
                    Append("WHERE idCfdi = @0 ");

                DataTable dtOpcionalDetalle = dal.QueryDT("DS_FE", sbOpcionalDetalle.ToString(), "F:I:" + idCfdi, hc);

                // Obtenemos los conceptos guardados en el Xml 
                int numConceptosXml = electronicDocument.Data.Conceptos.Count;

                // Debido a que no siempre hacen match el número de conceptos de xml con los de opcionalDetalle armamos un Hashtable
                // Armamos la tabla que contendra los conceptops finales a imprimir en el Pdf

                DataTable dtConceptosFinal = new DataTable();
                dtConceptosFinal.Columns.Add("numero", typeof(int));
                dtConceptosFinal.Columns.Add("codigoBarras", typeof(string));
                dtConceptosFinal.Columns.Add("codigoProducto", typeof(string));
                dtConceptosFinal.Columns.Add("lote", typeof(string));
                dtConceptosFinal.Columns.Add("fechaLote", typeof(string));
                dtConceptosFinal.Columns.Add("cantidadLote", typeof(string));
                dtConceptosFinal.Columns.Add("descripcion", typeof(string));
                dtConceptosFinal.Columns.Add("cantidad", typeof(double));
                dtConceptosFinal.Columns.Add("unidadMedida", typeof(string));
                dtConceptosFinal.Columns.Add("precioUnitario", typeof(double));
                dtConceptosFinal.Columns.Add("importe", typeof(double));

                int contConceptos = 1;

                foreach (DataRow rowConceptos in dtOpcionalDetalle.Rows)
                {
                    for (int i = contConceptos; i <= numConceptosXml; i++)
                    {
                        object[] arrayConceptos = new object[11];

                        if (Convert.ToInt32(rowConceptos["numero"]) == i)
                        {
                            arrayConceptos[0] = Convert.ToInt32(rowConceptos["numero"]);
                            arrayConceptos[1] = rowConceptos["codigoBarras"].ToString();
                            arrayConceptos[2] = rowConceptos["codigoProducto"].ToString();
                            arrayConceptos[3] = rowConceptos["lote"].ToString();
                            arrayConceptos[4] = rowConceptos["fechaLote"].ToString();
                            arrayConceptos[5] = rowConceptos["cantidadLote"].ToString();
                            arrayConceptos[6] = electronicDocument.Data.Conceptos[i - 1].Descripcion.Value;
                            arrayConceptos[7] = electronicDocument.Data.Conceptos[i - 1].Cantidad.Value;
                            arrayConceptos[8] = electronicDocument.Data.Conceptos[i - 1].Unidad.Value;
                            arrayConceptos[9] = electronicDocument.Data.Conceptos[i - 1].ValorUnitario.Value;
                            arrayConceptos[10] = electronicDocument.Data.Conceptos[i - 1].Importe.Value;
                        }
                        else
                        {
                            arrayConceptos[0] = 0;
                            arrayConceptos[1] = string.Empty;
                            arrayConceptos[2] = string.Empty;
                            arrayConceptos[3] = string.Empty;
                            arrayConceptos[4] = string.Empty;
                            arrayConceptos[5] = string.Empty;
                            arrayConceptos[6] = electronicDocument.Data.Conceptos[i - 1].Descripcion.Value;
                            arrayConceptos[7] = electronicDocument.Data.Conceptos[i - 1].Cantidad.Value;
                            arrayConceptos[8] = electronicDocument.Data.Conceptos[i - 1].Unidad.Value;
                            arrayConceptos[9] = electronicDocument.Data.Conceptos[i - 1].ValorUnitario.Value;
                            arrayConceptos[10] = electronicDocument.Data.Conceptos[i - 1].Importe.Value;
                        }

                        dtConceptosFinal.Rows.Add(arrayConceptos);
                        break;
                    }

                    contConceptos++;
                }

                // Una vez llena la tabla final de conceptos procedemos a generar la tabla de detalle.

                #region"Tabla Detalle"

                Table encabezadoDetalle = new Table(7);
                float[] headerEncabezadoDetalle = { 5, 12, 11, 40, 9, 10, 12 };
                encabezadoDetalle.Widths = headerEncabezadoDetalle;
                encabezadoDetalle.WidthPercentage = 100F;
                encabezadoDetalle.Padding = 1;
                encabezadoDetalle.Spacing = 1;
                encabezadoDetalle.BorderWidth = 0;
                encabezadoDetalle.DefaultCellBorder = 0;
                encabezadoDetalle.BorderColor = gris;

                // Número
                cel = new Cell(new Phrase("No", titulo));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Còdigo de Barras
                cel = new Cell(new Phrase("Cod. Barras", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Código
                cel = new Cell(new Phrase("Código", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Descripción
                cel = new Cell(new Phrase("Descripción", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Cantidad
                cel = new Cell(new Phrase("Cantidad", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Precio Unitario
                cel = new Cell(new Phrase("Precio Unitario", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Importe
                cel = new Cell(new Phrase("Importe", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                PdfPCell cell;
                PdfPTable tableConceptos = new PdfPTable(7);
                tableConceptos.SetWidths(new int[7] { 5, 12, 11, 40, 9, 10, 12 });
                //                tableConceptos.WidthPercentage = 91.5F;
                tableConceptos.WidthPercentage = 100F;

                Font fontLbl = new Font(Font.HELVETICA, 7, Font.BOLD, new Color(43, 145, 175));
                Font fontVal = new Font(Font.HELVETICA, 7, Font.NORMAL);

                foreach (DataRow rowConceptos in dtConceptosFinal.Rows)
                {
                    // Número
                    cell = new PdfPCell(new Phrase(rowConceptos["numero"].ToString(), f5));
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = (float).5;
                    cell.BorderWidthRight = 0;
                    cell.BorderWidthBottom = (float).5;
                    cell.BorderColor = gris;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tableConceptos.AddCell(cell);

                    // Código de Barras
                    cell = new PdfPCell(new Phrase(rowConceptos["codigoBarras"].ToString(), f5));
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = (float).5;
                    cell.BorderWidthRight = 0;
                    cell.BorderWidthBottom = (float).5;
                    cell.BorderColor = gris;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tableConceptos.AddCell(cell);

                    // Código de Producto
                    cell = new PdfPCell(new Phrase(rowConceptos["codigoProducto"].ToString(), f5));
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = (float).5;
                    cell.BorderWidthRight = 0;
                    cell.BorderWidthBottom = (float).5;
                    cell.BorderColor = gris;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tableConceptos.AddCell(cell);

                    // Tabla de Descripción
                    PdfPTable tableDesc = new PdfPTable(3);
                    tableDesc.WidthPercentage = 100;

                    // Descripción
                    cell = new PdfPCell(new Phrase(rowConceptos["descripcion"].ToString(), f5));
                    cell.Border = 0;
                    cell.NoWrap = true;
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cell.Colspan = 3;
                    tableDesc.AddCell(cell);

                    // Lote
                    //Phrase pLote = new Phrase();
                    //Chunk cLoteLbl = new Chunk("Lote : \n", fontLbl);
                    //Chunk cLoteVal = new Chunk(rowConceptos["lote"].ToString().Replace("*", "\n"), fontVal);

                    //pLote.Add(cLoteLbl);
                    //pLote.Add(cLoteVal);

                    //par.KeepTogether = true;
                    //par.SetLeading(7f, 1f);

                    par = new Paragraph();
                    par.Add(new Chunk("Lote\n\n", f5L));
                    par.Add(new Chunk(rowConceptos["lote"].ToString().Replace("*", "\n"), f5));
                    cell = new PdfPCell(par);
                    cell.Border = 0;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    tableDesc.AddCell(cell);

                    //// Cantidad 
                    //Phrase pCantidad = new Phrase();
                    //Chunk cCantidadLbl = new Chunk("Cantidad\n", fontLbl);
                    //Chunk cCantidadVal = new Chunk(rowConceptos["cantidadLote"].ToString().Replace("*", "\n"), fontVal);

                    //pCantidad.Add(cCantidadLbl);
                    //pCantidad.Add(cCantidadVal);
                    //cell = new PdfPCell(pCantidad);
                    //cell.Border = 1;
                    //cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    //tableDesc.AddCell(cell);

                    par = new Paragraph();
                    par.Add(new Chunk("Cantidad\n\n", f5L));
                    par.Add(new Chunk(rowConceptos["cantidadLote"].ToString().Replace("*", "\n"), f5));
                    cell = new PdfPCell(par);
                    cell.Border = 0;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    tableDesc.AddCell(cell);

                    //// Fecha Lote
                    //Phrase pFecLote = new Phrase();
                    //Chunk cFecLoteLbl;
                    //Chunk cFecLoteVal;

                    par = new Paragraph();
                    //if (electronicDocument.Data.Serie.Value == "C" || electronicDocument.Data.Serie.Value == "D" || electronicDocument.Data.Serie.Value == "K" || electronicDocument.Data.Serie.Value == "L" || electronicDocument.Data.Serie.Value == "M" || electronicDocument.Data.Serie.Value == "N")
                    //{
                    //    par.Add(new Chunk("", fontLbl));
                    //    par.Add(new Chunk("", fontVal));
                    //    //cFecLoteLbl = new Chunk("11", fontLbl);
                    //    //cFecLoteVal = new Chunk("22", fontVal);
                    //}
                    //else 
                    if (electronicDocument.Data.Serie.Value == "A")
                    {
                        par.Add(new Chunk("Expiración\n\n", f5L));
                        par.Add(new Chunk(rowConceptos["fechaLote"].ToString().Replace("*", "\n"), f5));
                        //cFecLoteLbl = new Chunk("Fecha Lote : ", fontLbl);
                        //cFecLoteVal = new Chunk(rowConceptos["fechaLote"].ToString().Replace("*", "\n"), fontVal);
                    }
                    else
                    {
                        par.Add(new Chunk(""));
                        par.Add(new Chunk("", f5));
                        //cFecLoteLbl = new Chunk("Fecha Lote : ", fontLbl);
                        //cFecLoteVal = new Chunk(rowConceptos["fechaLote"].ToString().Replace("*", "\n"), fontVal);
                    }
                    //pFecLote.Add(cFecLoteLbl);
                    //pFecLote.Add(cFecLoteVal);

                    cell = new PdfPCell(par);
                    cell.Border = 0;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    tableDesc.AddCell(cell);

                    // Número de Cantidad
                    cell = new PdfPCell(new Phrase(""));
                    cell.Border = 0;
                    cell.Colspan = 3;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tableDesc.AddCell(cell);

                    cell = new PdfPCell(tableDesc);
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = (float).5;
                    cell.BorderWidthRight = 0;
                    cell.BorderWidthBottom = (float).5;
                    cell.BorderColor = gris;
                    tableConceptos.AddCell(cell);

                    // Cantidad
                    cell = new PdfPCell(new Phrase(rowConceptos["cantidad"] + " " + rowConceptos["unidadMedida"], f5));
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = (float).5;
                    cell.BorderWidthRight = 0;
                    cell.BorderWidthBottom = (float).5;
                    cell.BorderColor = gris;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tableConceptos.AddCell(cell);

                    // Precio Unitario 
                    cell = new PdfPCell(new Phrase(Convert.ToDouble(rowConceptos["precioUnitario"]).ToString("C", _ci), f5));
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = (float).5;
                    cell.BorderWidthRight = 0;
                    cell.BorderWidthBottom = (float).5;
                    cell.BorderColor = gris;
                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tableConceptos.AddCell(cell);

                    // Importe
                    cell = new PdfPCell(new Phrase(Convert.ToDouble(rowConceptos["importe"]).ToString("C", _ci), f5));
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = (float).5;
                    cell.BorderWidthRight = (float).5;
                    cell.BorderWidthBottom = (float).5;
                    cell.BorderColor = gris;
                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tableConceptos.AddCell(cell);
                }

                #endregion

                #endregion

                #endregion

                #region "Construimos el Comentarios"

                Table comentarios = new Table(7);
                float[] headerwidthsComentarios = { 9, 18, 28, 28, 7, 5, 5 };
                comentarios.Widths = headerwidthsComentarios;
                comentarios.WidthPercentage = 100;
                comentarios.Padding = 1;
                comentarios.Spacing = 1;
                comentarios.BorderWidth = 0;
                comentarios.DefaultCellBorder = 0;
                comentarios.BorderColor = gris;

                cel = new Cell(new Phrase("Cantidad:", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Rowspan = 3;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["cantidadLetra"].ToString(), f5));//CE26  
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                cel.Rowspan = 3;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Sub Total", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                double importe = 0;
                double tasa = 0;

                for (int i = 0; i < electronicDocument.Data.Impuestos.Traslados.Count; i++)
                {
                    if (electronicDocument.Data.Impuestos.Traslados[i].Tipo.Value == "IVA")
                    {
                        importe = electronicDocument.Data.Impuestos.Traslados[i].Importe.Value;
                        tasa = electronicDocument.Data.Impuestos.Traslados[i].Tasa.Value;
                        break;
                    }
                }

                cel = new Cell(new Phrase("IVA " + tasa + " %", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("C", _ci), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Total", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _ci), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Observaciones:", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["observaciones"].ToString(), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 6;
                comentarios.AddCell(cel);

                #endregion

                #region "Construimos Tabla Especial"
                Table especial = new Table(5);
                float[] headerwidthsEspecial = { 20, 20, 20, 20, 20 };
                especial.Widths = headerwidthsEspecial;
                especial.WidthPercentage = 100;
                especial.Padding = 1;
                especial.Spacing = 1;
                especial.BorderWidth = 0;
                especial.DefaultCellBorder = 0;
                especial.BorderColor = blanco;

                cel = new Cell(new Phrase("", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = blanco;
                cel.Colspan = 5;
                especial.AddCell(cel);

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
                    par.Add(new Chunk("SELLO DIGITAL DEL EMISOR\n", f5L));
                    par.Add(new Chunk(electronicDocument.Data.Sello.Value, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
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
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase("FECHA Y HORA DE CERTIFICACION:", f5L));
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

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL SAT:", f5L));
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

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL EMISOR:", f5L));
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

                    string lbRegimen = "REGIMEN FISCAL APLICABLE: ";
                    StringBuilder regimenes = new StringBuilder();
                    if (electronicDocument.Data.Emisor.Regimenes.IsAssigned)
                    {
                        for (int i = 0; i < electronicDocument.Data.Emisor.Regimenes.Count; i++)
                        {
                            regimenes.Append(electronicDocument.Data.Emisor.Regimenes[i].Regimen.Value).Append("\n");
                        }
                    }
                    else
                    {
                        regimenes.Append("");
                        lbRegimen = "";
                    }

                    string cuenta = electronicDocument.Data.NumeroCuentaPago.IsAssigned
                                    ? electronicDocument.Data.NumeroCuentaPago.Value
                                    : "";

                    par = new Paragraph();
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("TIPO DE COMPROBANTE: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.TipoComprobante.Value + "   |   ", f5));
                    par.Add(new Chunk("Moneda: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.Moneda.Value + "   |   ", f5));
                    par.Add(new Chunk("TASA DE CAMBIO: ", f5L));
                    string tasaCambio = electronicDocument.Data.TipoCambio.Value;
                    if (tasaCambio.Length > 0)
                    {
                        par.Add(new Chunk(Convert.ToDouble(tasaCambio).ToString("C", _ci) + "   |   ", f5));
                    }
                    else
                    {
                        par.Add(new Chunk("   |   ", f5));
                    }
                    par.Add(new Chunk("FORMA DE PAGO: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.FormaPago.Value + "\n", f5));
                    par.Add(new Chunk("MÉTODO DE PAGO: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.MetodoPago.Value + "   |   ", f5));
                    par.Add(new Chunk("NÚMERO DE CUENTA: ", f5L));
                    par.Add(new Chunk(cuenta + "   |   ", f5));
                    par.Add(new Chunk(lbRegimen, f5L));
                    par.Add(new Chunk(regimenes.ToString(), f5));
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
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
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

                #region "Construimos Tabla Adicional Datos de Pago"

                Table pago = new Table(5);
                float[] headerwidthsPago = { 40, 10, 20, 10, 20 };
                pago.Widths = headerwidthsPago;
                pago.WidthPercentage = 100;
                pago.Padding = 1;
                pago.Spacing = 1;
                pago.BorderWidth = (float).5;
                pago.DefaultCellBorder = 1;
                pago.BorderColor = gris;


                cel = new Cell(new Phrase("Para ejecutar su pago para consultas", f5L));
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Para efectuar su pago", f5L));
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Información del Cliente", f5L));
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                pago.AddCell(cel);

                string[] ejecutarPago = dtEncabezado.Rows[0]["paraConsultas"].ToString().Split(new Char[] { '*' });//CE23
                string[] pagoDatos = dtEncabezado.Rows[0]["efectuarPago"].ToString().Split(new Char[] { '*' });//CE21

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (ejecutarPago.Length > 0)
                    par.Add(new Chunk(ejecutarPago[0] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                if (ejecutarPago.Length > 1)
                    par.Add(new Chunk(ejecutarPago[1] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                if (ejecutarPago.Length > 2)
                    par.Add(new Chunk(ejecutarPago[2] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                if (ejecutarPago.Length > 3)
                    par.Add(new Chunk(ejecutarPago[3] + "\n", f5));
                else
                    par.Add(new Chunk("\n", f5));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                cel.Rowspan = 6;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Banco:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 0)
                    par.Add(new Chunk(pagoDatos[0], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Cliente:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Receptor.Nombre.Value, f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Cuenta:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 1)
                    par.Add(new Chunk(pagoDatos[1], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("No. Cliente:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["clienteNo"].ToString(), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                if (htCFDI["serie"].ToString() == "B" || htCFDI["serie"].ToString() == "D" || htCFDI["serie"].ToString() == "J" || htCFDI["serie"].ToString() == "L" || htCFDI["serie"].ToString() == "N")
                {
                    cel = new Cell(new Phrase("", f5));
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 2;
                    pago.AddCell(cel);
                }
                else
                {
                    cel = new Cell(new Phrase("SWIFT:", f5L));
                    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    pago.AddCell(cel);

                    if (pagoDatos.Length > 3)
                        cel = new Cell(new Phrase(pagoDatos[3], f5));
                    else
                        cel = new Cell(new Phrase("", f5));
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    pago.AddCell(cel);
                }

                cel = new Cell(new Phrase("Factura No:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["serie"].ToString().ToUpper() + "-" + electronicDocument.Data.Folio.Value.ToString(), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Moneda:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 4)
                    par.Add(new Chunk(pagoDatos[4], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Fecha Factura", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                string[] fechaFactura = Convert.ToDateTime(objTimbre.FechaTimbrado.Value).GetDateTimeFormats('s');

                cel = new Cell(new Phrase(DIA + "/" + MES + "/" + ANIO, f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Beneficiario:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 5)
                    par.Add(new Chunk(pagoDatos[5], f5));
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Valor Total:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _ci), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                pago.AddCell(cel);

                cel = new Cell(new Phrase("Dirección:", f5L));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                pago.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                if (pagoDatos.Length > 4)
                    par.Add(new Chunk(dtEncabezado.Rows[0]["direccionPie"].ToString(), f5));//CE22
                else
                    par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                pago.AddCell(cel);

                if (htCFDI["serie"].ToString() == "C")
                {
                    cel = new Cell(new Phrase("Fecha Vencimiento:", f5L));
                    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
                    pago.AddCell(cel);

                    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["vencimiento"].ToString(), f5B));
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    pago.AddCell(cel);
                }
                else
                {
                    cel = new Cell(new Phrase("Moneda:", f5L));
                    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
                    pago.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.Data.Moneda.Value, f5));
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    pago.AddCell(cel);
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
                pageEventHandler.footer = footer;

                document.Open();

                document.Add(tableConceptos);
                document.Add(comentarios);
                document.Add(especial);
                document.Add(adicional);
                document.Add(pago);

                #endregion
            }
            catch (Exception ex)
            {
                string exe = ex.Message;
            }
        }

        #endregion

        #region "formatoN"

        public static void formatoN(Document document, ElectronicDocument electronicDocument, Data objTimbre, pdfPageEventHandlerPfizer pageEventHandler, DataTable dtEncabezado, DataTable dtDetalle, Hashtable htCFDI, HttpContext hc)
        {
            try
            {
                DAL dal = new DAL();

                #region "Construimos el Documento"

                #region "Construimos el Encabezado"

                Table encabezado = new Table(7);
                float[] headerwidthsEncabezado = { 9, 18, 28, 28, 5, 7, 5 };
                encabezado.Widths = headerwidthsEncabezado;
                encabezado.WidthPercentage = 100;
                encabezado.Padding = 1;
                encabezado.Spacing = 1;
                encabezado.BorderWidth = 0;
                encabezado.DefaultCellBorder = 0;
                encabezado.BorderColor = gris;

                //Agregando Imagen de Logotipo
                Image imgLogo = Image.GetInstance(pathIMGLOGO);
                imgLogo.ScalePercent(47f);

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
                cel.Rowspan = 5;
                cel.Colspan = 2;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(8f, 9f);
                par.Add(new Chunk(htCFDI["nombreEmisor"].ToString().ToUpper(), f6B));
                par.Add(new Chunk("\nRFC " + htCFDI["rfcEmisor"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor1"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor2"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htCFDI["direccionEmisor3"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\nTel. (52) 55 5081-8500", f6));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                StringBuilder expedido = new StringBuilder();
                expedido.
                    Append("Lugar de Expedición México DF\n").
                    Append(htCFDI["sucursal"]).Append("\n").
                    Append(htCFDI["direccionExpedido1"].ToString().ToUpper()).Append("\n").
                    Append(htCFDI["direccionExpedido2"].ToString().ToUpper()).Append("\n").
                    Append(htCFDI["direccionExpedido3"].ToString().ToUpper()).Append("\n").
                    Append(dtEncabezado.Rows[0]["corporateCode"]).Append("\n").//CE29
                    Append(dtEncabezado.Rows[0]["oldCorporateCode"].ToString()).Append("\n");//CE30;

                cel = new Cell(new Phrase(expedido.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["tipoDoc"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["serie"].ToString().ToUpper() + electronicDocument.Data.Folio.Value, f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Día", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Mes", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Año", f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                string[] fechaCFDI = Convert.ToDateTime(htCFDI["fechaCfdi"].ToString()).GetDateTimeFormats();
                string HORAS = fechaCFDI[103];
                string DIA = Convert.ToDateTime(htCFDI["fechaCfdi"]).Day.ToString();
                string MES = Convert.ToDateTime(htCFDI["fechaCfdi"]).ToString("MMMM").ToUpper();
                string ANIO = Convert.ToDateTime(htCFDI["fechaCfdi"]).Year.ToString();

                cel = new Cell(new Phrase(DIA, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(MES, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(ANIO, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(HORAS, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("CLIENTE", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Receptor.Nombre.Value, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("Referencia Documento: " + dtEncabezado.Rows[0]["referencia"].ToString(), f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("R.F.C. " + electronicDocument.Data.Receptor.Rfc.Value + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor1"].ToString() + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor2"].ToString() + "\n", f5));
                par.Add(new Chunk(htCFDI["direccionReceptor3"].ToString() + "\n", f5));
                par.Add(new Chunk(dtEncabezado.Rows[0]["codigoCorporativo"] + "\n" + "\n", f5));//36
                par.Add(new Chunk(dtEncabezado.Rows[0]["codigoCorporativoAnt"].ToString(), f5));//37
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 5;
                cel.Colspan = 2;
                encabezado.AddCell(cel);

                #region "Consignado a"

                string nombreEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-NOMBRE"].ToString();//CE
                string calleEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-CALLE"].ToString();//CE 10, 11, 12
                string coloniaEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-COLONIA"].ToString();//CE13
                string cpEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-CP"].ToString();//CE18
                string municEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-MUNIC"].ToString();//CE15
                string estadoEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-ESTADO"].ToString();//CE16
                string paisEntregado = dtEncabezado.Rows[0]["ENTREGADO-A-PAIS"].ToString();//CE14
                string intericom = dtEncabezado.Rows[0]["INTERICOM"].ToString().ToUpper();//CE32
                string saltoEntregado = "\n";
                string separador = ", ";
                string espacio = " ";

                if (nombreEntregado.Length > 0)
                    nombreEntregado = nombreEntregado + saltoEntregado;
                else
                    nombreEntregado = "" + saltoEntregado;

                if (calleEntregado.Length > 0)
                    calleEntregado = calleEntregado + saltoEntregado;
                else
                    calleEntregado = "" + saltoEntregado;

                if (coloniaEntregado.Length > 0)
                    coloniaEntregado = coloniaEntregado + espacio;
                else
                    coloniaEntregado = "" + espacio;

                if (cpEntregado.Length > 0)
                    cpEntregado = cpEntregado + ", CP " + saltoEntregado;
                else
                    cpEntregado = "" + separador + saltoEntregado;

                if (municEntregado.Length > 0)
                    municEntregado = municEntregado + saltoEntregado;
                else
                    municEntregado = "";

                if (estadoEntregado.Length > 0)
                    estadoEntregado = estadoEntregado + espacio;
                else
                    estadoEntregado = "" + espacio;

                if (paisEntregado.Length > 0)
                    paisEntregado = paisEntregado + "";
                else
                    paisEntregado = "";

                if (intericom.Length > 0)
                    intericom = intericom + espacio;
                else
                    intericom = "";

                #endregion

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Consignado a:\n", f5B));
                par.Add(new Chunk(calleEntregado, f5));
                par.Add(new Chunk(coloniaEntregado, f5));
                par.Add(new Chunk(cpEntregado, f5));
                par.Add(new Chunk(municEntregado, f5));
                par.Add(new Chunk(estadoEntregado, f5));
                par.Add(new Chunk(paisEntregado, f5));
                par.Add(new Chunk("INTERICOM:\n", f5));
                par.Add(new Chunk(intericom, f5));//CE32
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Contacto:\n", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["contacto"].ToString(), f5));
                par.Add(new Chunk("\nTeléfono:\n", f5B));
                par.Add(new Chunk("" + "\n", f5));
                par.Add(new Chunk("Método de Pago:\n", f5B));
                par.Add(new Chunk(electronicDocument.Data.FormaPago.Value, f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Rowspan = 5;
                encabezado.AddCell(cel);


                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Fecha de Ref. Doc:  ", f5B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["fechaRefDoc"].ToString(), f5));//CE34
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Moneda: ", f5B));
                par.Add(new Chunk(electronicDocument.Data.Moneda.Value, f5));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Forma de Pago: ", f5B));
                par.Add(new Chunk(electronicDocument.Data.FormaPago.Value.ToString(), f5));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Condiciones de Pago: ", f5B));
                par.Add(new Chunk(electronicDocument.Data.CondicionesPago.Value.ToString(), f5));//CE19
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk("Proveedor: ", f5B));
                par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                #endregion

                #region "Construimos Tablas de Partidas"

                #region "Construimos Encabezados de Partidas"

                Table encabezadoPartidas = new Table(9);
                float[] headerwidthsEncabesadoPartidas = { 10, 34, 7, 7, 7, 9, 8, 9, 9 };
                encabezadoPartidas.Widths = headerwidthsEncabesadoPartidas;
                encabezadoPartidas.WidthPercentage = 100;
                encabezadoPartidas.Padding = 1;
                encabezadoPartidas.Spacing = 1;
                encabezadoPartidas.BorderWidth = (float).5;
                encabezadoPartidas.DefaultCellBorder = 1;
                encabezadoPartidas.BorderColor = gris;

                cel = new Cell(new Phrase("Codigo Producto.", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Descripción", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Cantidad", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Cantidad Real (Kg,Lt,Mt)", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Unidad Medida", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Precio Unitario", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Descuento", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Importe", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Importe Neto", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                #endregion

                #region "Construimos Contenido de las Partidas"

                Table partidas = new Table(9);
                float[] headerwidthsPartidas = { 10, 34, 7, 7, 7, 9, 8, 9, 9 };
                partidas.Widths = headerwidthsPartidas;
                partidas.WidthPercentage = 100;
                partidas.Padding = 1;
                partidas.Spacing = 1;
                partidas.BorderWidth = (float).5;
                partidas.DefaultCellBorder = 1;
                partidas.BorderColor = gris;

                double sumDescuento = 0;

                if (dtEncabezado.Rows.Count > 0)
                {
                    for (int i = 0; i < electronicDocument.Data.Conceptos.Count; i++)
                    {
                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["codeLocal"].ToString(), f5));//CD1
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Cantidad.Value.ToString("N", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["cantidadReal"].ToString(), f5));//CD16
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Unidad.Value, f5));//CD10
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].ValorUnitario.Value.ToString("N", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(dtDetalle.Rows[i]["descuento"].ToString(), f5));//CD12
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        if (double.TryParse(dtDetalle.Rows[i]["descuento"].ToString(), out sumDescuento))
                            sumDescuento += double.Parse(dtDetalle.Rows[i]["descuento"].ToString());
                        else
                            sumDescuento += 0;

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("N", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("N", _ci), f5));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = (float).5;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                    }
                }

                #endregion

                #endregion

                #region "Construimos el Comentarios"

                Table comentarios = new Table(7);
                float[] headerwidthsComentarios = { 9, 18, 28, 28, 7, 5, 5 };
                comentarios.Widths = headerwidthsComentarios;
                comentarios.WidthPercentage = 100;
                comentarios.Padding = 1;
                comentarios.Spacing = 1;
                comentarios.BorderWidth = 0;
                comentarios.DefaultCellBorder = 0;
                comentarios.BorderColor = gris;

                cel = new Cell(new Phrase("Observaciones:", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Cantidad:", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Sub Total", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Descuento", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(sumDescuento.ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("I.E.P.S.", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("0", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["observaciones"].ToString(), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                cel.Rowspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["cantidadLetra"].ToString(), f5));//CE26           
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                cel.Rowspan = 2;
                comentarios.AddCell(cel);

                double importe = 0;
                double tasa = 0;

                for (int i = 0; i < electronicDocument.Data.Impuestos.Traslados.Count; i++)
                {
                    if (electronicDocument.Data.Impuestos.Traslados[i].Tipo.Value == "IVA")
                    {
                        importe = electronicDocument.Data.Impuestos.Traslados[i].Importe.Value;
                        tasa = electronicDocument.Data.Impuestos.Traslados[i].Tasa.Value;
                        break;
                    }
                }

                cel = new Cell(new Phrase("IVA " + tasa.ToString() + " %", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Total", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Total de Articulos", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Conceptos.Count.ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                #endregion

                #region "Construimos Tabla Especial"
                Table especial = new Table(5);
                float[] headerwidthsEspecial = { 20, 20, 20, 20, 20 };
                especial.Widths = headerwidthsEspecial;
                especial.WidthPercentage = 100;
                especial.Padding = 1;
                especial.Spacing = 1;
                especial.BorderWidth = 0;
                especial.DefaultCellBorder = 0;
                especial.BorderColor = gris;

                cel = new Cell(new Phrase("", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Rowspan = 2;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

                cel = new Cell(new Phrase("", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                especial.AddCell(cel);

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
                    par.Add(new Chunk("SELLO DIGITAL DEL EMISOR\n", f5L));
                    par.Add(new Chunk(electronicDocument.Data.Sello.Value, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
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
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase("FECHA Y HORA DE CERTIFICACION:", f5L));
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

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL SAT:", f5L));
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

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL EMISOR:", f5L));
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

                    string lbRegimen = "REGIMEN FISCAL APLICABLE: ";
                    StringBuilder regimenes = new StringBuilder();
                    if (electronicDocument.Data.Emisor.Regimenes.IsAssigned)
                    {
                        for (int i = 0; i < electronicDocument.Data.Emisor.Regimenes.Count; i++)
                        {
                            regimenes.Append(electronicDocument.Data.Emisor.Regimenes[i].Regimen.Value).Append("\n");
                        }
                    }
                    else
                    {
                        regimenes.Append("");
                        lbRegimen = "";
                    }

                    string cuenta = electronicDocument.Data.NumeroCuentaPago.IsAssigned
                                    ? electronicDocument.Data.NumeroCuentaPago.Value
                                    : "";
                    par = new Paragraph();
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("TIPO DE COMPROBANTE: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.TipoComprobante.Value + "   |   ", f5));
                    par.Add(new Chunk("Moneda: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.Moneda.Value + "   |   ", f5));
                    par.Add(new Chunk("TASA DE CAMBIO: ", f5L));
                    string tasaCambio = electronicDocument.Data.TipoCambio.Value;
                    if (tasaCambio.Length > 0)
                    {
                        par.Add(new Chunk(Convert.ToDouble(tasaCambio).ToString("C", _ci) + "   |   ", f5));
                    }
                    else
                    {
                        par.Add(new Chunk("   |   ", f5));
                    }
                    par.Add(new Chunk("FORMA DE PAGO: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.FormaPago.Value + "\n", f5));
                    par.Add(new Chunk("MÉTODO DE PAGO: ", f5L));
                    par.Add(new Chunk(electronicDocument.Data.MetodoPago.Value + "   |   ", f5));
                    par.Add(new Chunk("NÚMERO DE CUENTA: ", f5L));
                    par.Add(new Chunk(cuenta + "   |   ", f5));
                    par.Add(new Chunk(lbRegimen, f5L));
                    par.Add(new Chunk(regimenes.ToString(), f5));
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
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = (float).5;
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
                pageEventHandler.encPartidas = encabezadoPartidas;
                pageEventHandler.footer = footer;

                document.Open();

                document.Add(partidas);
                document.Add(comentarios);
                document.Add(especial);
                document.Add(adicional);

                document.Close();

                #endregion
            }
            catch (Exception ex)
            {
                string exe = ex.Message;
            }
        }

        #endregion
    }

    public class pdfPageEventHandlerPfizer : PdfPageEventHelper
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
            document.Add(encPartidas);
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

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);

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

    public class DefaultSplitCharacterPfizer : ISplitCharacter
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