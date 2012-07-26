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

#endregion

namespace wsRepresentacionImpresa.App_Code.com.Facturaxion.facturaEspecial
{
    public class JIGAFRA
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

        private static String pathIMGLOGO;
        private static String pathIMGNOMBRE;
        private static Cell cel;
        private static Paragraph par;
        private static Color gris;
        private static Color blanco;
        private static BaseFont EM;
        private static Font f5;
        private static Font f6U;
        private static Font f4U;
        private static Font f6;
        private static Font f5B;
        private static Font f6B;
        private static Font f7;
        private static Font f7B;
        private static Font titulo;
        private static Font folio;

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
                StringBuilder sbDataEmisor = new StringBuilder();
                DataTable dtDataEmisor = new DataTable();

                sbOpcionalEncabezado.
                    Append("SELECT ").
                    //Append("campo1 AS [IV], ").
                    //Append("campo2 AS [ATZ], ").
                    //Append("campo3 AS [ON], ").
                    Append("campo4 AS [ACE], ").
                    Append("campo5 AS [REFERENCIA], ").
                    Append("campo6 AS [VIA_DE_EMBARQUE], ").
                    Append("campo8 AS [OBSERVACIONES], ").
                    Append("campo9 AS [CODIGO_DE_CLIENTE], ").
                    Append("campo10 AS [TELEFONO], ").
                    Append("campo11 AS [FECHA_DE_VENCIMIENTO], ").
                    Append("campo12 AS [SUCURSAL], ").
                    Append("campo13 AS [CANTIDAD-LETRA], ").
                    Append("campo12 AS [CONDICIONES_PAGO], ").
                    Append("campo14 AS [SUBTOTAL], ").
                    Append("campo15 AS [I] ").
                    Append("FROM opcionalEncabezado ").
                    Append("WHERE idCFDI = @0  AND ST = 1 ");

                sbDataEmisor.Append("SELECT nombreSucursal FROM sucursales WHERE idSucursal = @0 ");

                dtOpcEnc = dal.QueryDT("DS_FE", sbOpcionalEncabezado.ToString(), "F:I:" + idCfdi, hc);
                dtDataEmisor = dal.QueryDT("DS_FE", sbDataEmisor.ToString(), "F:I:" + htFacturaxion["idSucursalEmisor"], hc);

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
                    sbDirEmisor1.Append("CALLE: ").Append(electronicDocument.Data.Emisor.Domicilio.Calle.Value).Append(" ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.NumeroExterior.Value.Length > 0)
                {
                    sbDirEmisor1.Append(electronicDocument.Data.Emisor.Domicilio.NumeroExterior.Value).Append(" ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.NumeroInterior.Value.Length > 0)
                {
                    sbDirEmisor1.Append(electronicDocument.Data.Emisor.Domicilio.NumeroInterior.Value);
                }
                if (electronicDocument.Data.Emisor.Domicilio.Colonia.Value.Length > 0)
                {
                    sbDirEmisor2.Append(" COL. ").Append(electronicDocument.Data.Emisor.Domicilio.Colonia.Value);
                }
                if (electronicDocument.Data.Emisor.Domicilio.CodigoPostal.Value.Length > 0)
                {
                    sbDirEmisor2.Append(", C.P. ").Append(electronicDocument.Data.Emisor.Domicilio.CodigoPostal.Value);
                }
                if (electronicDocument.Data.Emisor.Domicilio.Municipio.Value.Length > 0)
                {
                    sbDirEmisor3.Append(" DEL. ").Append(electronicDocument.Data.Emisor.Domicilio.Municipio.Value).Append(" ");
                }
                sbDirEmisor3.Append(electronicDocument.Data.Emisor.Domicilio.Pais.Value).Append(", ");
                sbDirEmisor3.Append(electronicDocument.Data.Emisor.Domicilio.Estado.Value);

                #endregion

                #region "Dirección Receptor"

                StringBuilder sbDirReceptor1 = new StringBuilder();
                StringBuilder sbDirReceptor2 = new StringBuilder();
                StringBuilder sbDirReceptor3 = new StringBuilder();
                StringBuilder sbDirReceptor4 = new StringBuilder();
                StringBuilder sbDirReceptor5 = new StringBuilder();
                StringBuilder sbDirReceptor6 = new StringBuilder();

                if (electronicDocument.Data.Receptor.Domicilio.Calle.Value.Length > 0)
                {
                    sbDirReceptor1.Append("DIRECCIÓN. ").Append(electronicDocument.Data.Receptor.Domicilio.Calle.Value);
                }
                if (electronicDocument.Data.Receptor.Domicilio.NumeroExterior.Value.Length > 0)
                {
                    sbDirReceptor1.Append(" NO. ").Append(electronicDocument.Data.Receptor.Domicilio.NumeroExterior.Value);
                }
                if (electronicDocument.Data.Receptor.Domicilio.NumeroInterior.Value.Length > 0)
                {
                    sbDirReceptor1.Append(" INT. ").Append(electronicDocument.Data.Receptor.Domicilio.NumeroInterior.Value);
                }
                if (electronicDocument.Data.Receptor.Domicilio.Colonia.Value.Length > 0)
                {
                    sbDirReceptor2.Append("COLONIA: ").Append(electronicDocument.Data.Receptor.Domicilio.Colonia.Value);
                }
                if (electronicDocument.Data.Receptor.Domicilio.Municipio.Value.Length > 0)
                {
                    sbDirReceptor3.Append("DELEGACIÓN/MUNICIPIO: ").Append(electronicDocument.Data.Receptor.Domicilio.Municipio.Value);
                }
                if (electronicDocument.Data.Receptor.Domicilio.Localidad.Value.Length > 0)
                {
                    sbDirReceptor4.Append("CIUDAD./LOCALIDAD.: ").Append(electronicDocument.Data.Receptor.Domicilio.Localidad.Value);
                }
                //if (electronicDocument.Data.Receptor.Domicilio.Estado.Value.Length > 0)
                //{
                //    sbDirReceptor2.Append("\nCIUDAD: ").Append(electronicDocument.Data.Receptor.Domicilio.Estado.Value);
                //}
                if (electronicDocument.Data.Receptor.Domicilio.Estado.Value.Length > 0)
                {
                    sbDirReceptor5.Append("ESTADO: ").Append(electronicDocument.Data.Receptor.Domicilio.Estado.Value);
                }
                if (electronicDocument.Data.Receptor.Domicilio.CodigoPostal.Value.Length > 0)
                {
                    sbDirReceptor6.Append("C.P. ").Append(electronicDocument.Data.Receptor.Domicilio.CodigoPostal.Value);
                }
                //sbDirReceptor3.Append(electronicDocument.Data.Receptor.Domicilio.Pais.Value);

                #endregion

                htDatosCfdi.Add("direccionEmisor1", sbDirEmisor1.ToString());
                htDatosCfdi.Add("direccionEmisor2", sbDirEmisor2.ToString());
                htDatosCfdi.Add("direccionEmisor3", sbDirEmisor3.ToString());

                htDatosCfdi.Add("direccionReceptor1", sbDirReceptor1.ToString());
                htDatosCfdi.Add("direccionReceptor2", sbDirReceptor2.ToString());
                htDatosCfdi.Add("direccionReceptor3", sbDirReceptor3.ToString());
                htDatosCfdi.Add("direccionReceptor4", sbDirReceptor4.ToString());
                htDatosCfdi.Add("direccionReceptor5", sbDirReceptor5.ToString());
                htDatosCfdi.Add("direccionReceptor6", sbDirReceptor6.ToString());

                #endregion

                #region "Creamos el Objeto Documento y Tipos de Letra"

                Document document = new Document(PageSize.LETTER, 15, 15, 15, 15);
                document.AddAuthor("Facturaxion");
                document.AddCreator("r3Take");
                document.AddCreationDate();

                pdfPageEventHandlerJIGAFRA pageEventHandler = new pdfPageEventHandlerJIGAFRA();
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                writer.PageEvent = pageEventHandler;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                pathIMGLOGO = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\JIGO311112L67\logo.jpg";
                pathIMGNOMBRE = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\JIGO311112L67\NOMBRE.png";

                gris = new Color(13, 142, 244);
                blanco = new Color(255, 255, 255);


                EM = BaseFont.CreateFont(@"C:\Windows\Fonts\VERDANA.TTF", BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
                f5 = new Font(EM, 5, Font.NORMAL);
                f6U = new Font(EM, 6, Font.UNDERLINE);
                f4U = new Font(EM, 5, Font.UNDERLINE);
                f5B = new Font(EM, 5, Font.BOLD);
                f6 = new Font(EM, 6, Font.NORMAL);
                f6B = new Font(EM, 6, Font.BOLD);
                f7 = new Font(EM, 7, Font.NORMAL);
                f7B = new Font(EM, 7, Font.BOLD);
                titulo = new Font(EM, 7, Font.BOLD);
                folio = new Font(EM, 7, Font.BOLD);

                #endregion

                #region "Generamos el Docuemto"

                string SERIE = string.Empty;
                SERIE = electronicDocument.Data.Serie.Value;

                #region "Documento JIGAFRA"
                switch (SERIE)
                {
                    case "FE":
                        htDatosCfdi.Add("tipoDoc", "FACTURA");
                        factura(document, electronicDocument, objTimbre, pageEventHandler, dtOpcEnc, null, htDatosCfdi, hc);
                        break;
                    case "D":
                    case "C":
                        htDatosCfdi.Add("tipoDoc", "NOTA DE CRÉDITO");
                        factura(document, electronicDocument, objTimbre, pageEventHandler, dtOpcEnc, null, htDatosCfdi, hc);
                        break;
                    case "NC":
                        htDatosCfdi.Add("tipoDoc", "NOTA DE CARGO");
                        factura(document, electronicDocument, objTimbre, pageEventHandler, dtOpcEnc, null, htDatosCfdi, hc);
                        break;
                    default:
                        break;
                }

                #endregion

                document.Close();
                writer.Close();
                fs.Close();

                string filePdfExt = pathPdf.Replace(_rutaDocs, _rutaDocsExt);
                string urlPathFilePdf = filePdfExt.Replace(@"\", "/");

                //Subimos Archivo al Azure
                //wAzure.azureUpDownLoad(1, pathPdf);

                return "1#" + urlPathFilePdf;

                #endregion
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

        #region "factura"

        public static void factura(Document document, ElectronicDocument electronicDocument, Data objTimbre, pdfPageEventHandlerJIGAFRA pageEventHandler, DataTable dtEncabezado, DataTable dtDetalle, Hashtable htCFDI, HttpContext hc)
        {
            try
            {
                #region "Construimos el Documento"

                #region "Construimos el Encabezado"

                Table datosFiscales = new Table(3);
                float[] headerwidthsdatosFiscales = { 30, 40, 30 };
                datosFiscales.Widths = headerwidthsdatosFiscales;
                datosFiscales.WidthPercentage = 100;
                datosFiscales.Padding = 0;
                datosFiscales.Spacing = 0;
                datosFiscales.BorderWidth = 0;
                datosFiscales.DefaultCellBorder = 0;
                datosFiscales.BorderColor = gris;

                Image imgLogo = Image.GetInstance(pathIMGLOGO);
                imgLogo.ScalePercent(47f);

                Image imgNombre = Image.GetInstance(pathIMGNOMBRE);
                imgNombre.ScalePercent(47f);

                cel = new Cell(imgLogo);
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales.AddCell(cel);

                StringBuilder regimenes = new StringBuilder();

                for (int i = 0; i < electronicDocument.Data.Emisor.Regimenes.Count; i++)
                    regimenes.Append(electronicDocument.Data.Emisor.Regimenes[i].Regimen.Value).Append("\n");

                string regimen = regimenes.ToString();

                if (regimen == "No Aplica\n")
                    regimen = "RÉGIMEN GENERAL DE LEY PERSONAS MORALES";

                cel = new Cell(imgNombre);
                cel.Add(new Phrase(htCFDI["direccionEmisor1"].ToString().ToUpper(), f7));
                cel.Add(new Phrase(htCFDI["direccionEmisor2"].ToString().ToUpper() + "\n", f7));
                cel.Add(new Phrase(htCFDI["direccionEmisor3"].ToString().ToUpper(), f7));
                cel.Add(new Phrase("\nTELS: 5338-8969 al 5338-8973 Y 01-800-7019570", f7));
                cel.Add(new Phrase("\nRFC: " + htCFDI["rfcEmisor"].ToString().ToUpper(), f7));
                cel.Add(new Phrase("\n" + regimen.ToString(), f7));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales.AddCell(cel);

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
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales.AddCell(cel);

                //Generar nueva tabla para acomodar datos
                #endregion

                #region Datos fiscales2
                Table datosFiscales2 = new Table(8);
                float[] headerwidthsdatosFiscales2 = { 15, 8, 15, 10, 12, 10, 15, 15 };
                datosFiscales2.Widths = headerwidthsdatosFiscales2;
                datosFiscales2.WidthPercentage = 100;
                datosFiscales2.Padding = 0;
                datosFiscales2.Spacing = 0;
                datosFiscales2.BorderWidth = 0;
                datosFiscales2.DefaultCellBorder = 0;
                datosFiscales2.BorderColor = gris;

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["CODIGO_DE_CLIENTE"].ToString(), f6B));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 6;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase("                                     " + htCFDI["tipoDoc"].ToString() + "  " + htCFDI["serie"].ToString().ToUpper() + " " + electronicDocument.Data.Folio.Value, f6B));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                datosFiscales2.AddCell(cel);

                //1er nivel
                cel = new Cell(new Phrase("Nombre o Razón Social", f6));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase("R.F.C. ", f6));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase(htCFDI["rfcReceptor"].ToString().ToUpper(), f6));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase("SUCURSAL: ", f6));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase(dtEncabezado.Rows[0]["SUCURSAL"].ToString(), f6));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase("Folio SAT", f6B));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                datosFiscales2.AddCell(cel);

                //2o.nivel
                cel = new Cell(new Phrase("CLIENTE: " + electronicDocument.Data.Receptor.Nombre.Value, f6));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 6;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase(objTimbre.Uuid.Value, f6));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                datosFiscales2.AddCell(cel);

                //3ER NIVEL
                cel = new Cell(new Phrase(htCFDI["direccionReceptor1"].ToString(), f6));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 6;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase("Cerificado SAT", f6B));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase("Cerificado Emisor", f6B));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                //4o nivel
                cel = new Cell(new Phrase(htCFDI["direccionReceptor2"].ToString(), f6));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 6;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase(objTimbre.NumeroCertificadoSat.Value, f6));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.NumeroCertificado.Value, f6));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                //5o nivel
                cel = new Cell(new Phrase(htCFDI["direccionReceptor3"].ToString(), f6));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 6;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase("Fecha emisión", f6B));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase("Fecha Timbrado", f6B));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                //6o Nivel

                cel = new Cell(new Phrase(htCFDI["direccionReceptor4"].ToString(), f6));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 6;
                datosFiscales2.AddCell(cel);

                string[] fechaCFDI = Convert.ToDateTime(htCFDI["fechaCfdi"].ToString()).GetDateTimeFormats('s');
                cel = new Cell(new Phrase(fechaCFDI[0], f6));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                string[] fechaTimbrado = Convert.ToDateTime(objTimbre.FechaTimbrado.Value).GetDateTimeFormats('s');
                cel = new Cell(new Phrase(fechaTimbrado[0], f6));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                datosFiscales2.AddCell(cel);

                //7o nivel
                cel = new Cell(new Phrase(htCFDI["direccionReceptor5"].ToString(), f6));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 6;
                datosFiscales2.AddCell(cel);

                cel = new Cell(new Phrase("Lugar y Fecha de Expedición", f6B));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                datosFiscales2.AddCell(cel);

                //8o nivel
                cel = new Cell(new Phrase(htCFDI["direccionReceptor6"].ToString(), f6));
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 6;
                datosFiscales2.AddCell(cel);

                DateTime fecha = Convert.ToDateTime(htCFDI["fechaCfdi"].ToString());
                cel = new Cell(new Phrase("MÉX. D.F. A " + fecha.ToString("dd/MM/yyyy"), f6));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                datosFiscales2.AddCell(cel);

                //9,10,11vo. nivel


                if (htCFDI["tipoDoc"].ToString() == "FACTURA")
                {
                    cel = new Cell(new Phrase("TELEFONO: " + dtEncabezado.Rows[0]["TELEFONO"].ToString() + "     ", f6));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 8;
                    datosFiscales2.AddCell(cel);

                    cel = new Cell(new Phrase("OBSERVACIONES: " + dtEncabezado.Rows[0]["OBSERVACIONES"].ToString(), f6));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 8;
                    datosFiscales2.AddCell(cel);
                }
                else
                {

                    cel = new Cell(new Phrase("TELEFONO: " + dtEncabezado.Rows[0]["TELEFONO"].ToString() + "     ", f6));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 6;
                    datosFiscales2.AddCell(cel);

                    cel = new Cell(new Phrase("Referencia: ", f6B));
                    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    datosFiscales2.AddCell(cel);

                    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["ACE"].ToString(), f6));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    datosFiscales2.AddCell(cel);

                    cel = new Cell(new Phrase("CONCEPTO: " + dtEncabezado.Rows[0]["OBSERVACIONES"].ToString(), f6));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 8;
                    datosFiscales2.AddCell(cel);
                }

                if (htCFDI["tipoDoc"].ToString() == "FACTURA")
                {
                    cel = new Cell(new Phrase("ENVIAR A:", f6));
                    cel.Add(new Phrase("\n A", new Font(Font.HELVETICA, 2, Font.NORMAL, blanco)));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 8;
                    datosFiscales2.AddCell(cel);
                }

                if (htCFDI["tipoDoc"].ToString() == "NOTA DE CRÉDITO")
                {
                    cel = new Cell(new Phrase("REF: " + dtEncabezado.Rows[0]["REFERENCIA"].ToString(), f6));
                    cel.Add(new Phrase("\n A", new Font(Font.HELVETICA, 2, Font.NORMAL, blanco)));
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 8;
                    datosFiscales2.AddCell(cel);
                }

                #endregion

                #region "Datos opcionales encabezado"

                Table datosOpcionales = new Table(7);
                float[] headerwidthsdatosOpcionales = { 8, 8, 10, 22, 25, 17, 10 };
                datosOpcionales.Widths = headerwidthsdatosOpcionales;
                datosOpcionales.WidthPercentage = 100;
                datosOpcionales.Padding = 1;
                datosOpcionales.Spacing = 1;
                datosOpcionales.BorderWidth = 0;
                datosOpcionales.DefaultCellBorder = 0;
                datosOpcionales.BorderColor = gris;

                if (htCFDI["tipoDoc"].ToString() == "FACTURA")
                {
                    cel = new Cell(new Phrase("REF", f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase("RUTA", f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase("ZONA", f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase("VIA DE EMBARQUE", f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase("CONDICIONES DE PAGO", f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase("VENCIMIENTO", f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase("MÉXICO D.F. A", f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["REFERENCIA"].ToString(), f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase("", f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase("", f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["VIA_DE_EMBARQUE"].ToString(), f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.Data.CondicionesPago.Value, f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase(dtEncabezado.Rows[0]["FECHA_DE_VENCIMIENTO"].ToString(), f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    datosOpcionales.AddCell(cel);

                    cel = new Cell(new Phrase(fecha.ToString("dd/MM/yyyy"), f6));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    datosOpcionales.AddCell(cel);
                }

                #endregion

                #region "Construimos Tablas de Partidas"

                #region "Construimos Encabezados de Partidas"

                Table datosPartidas = new Table(7);
                float[] headerwidthsdatosPartidas = { 8, 6, 11, 24, 30, 13, 9 };
                datosPartidas.Widths = headerwidthsdatosPartidas;
                datosPartidas.WidthPercentage = 100;
                datosPartidas.Padding = 1;
                datosPartidas.Spacing = 1;
                datosPartidas.BorderWidth = 0;
                datosPartidas.DefaultCellBorder = 0;
                datosPartidas.BorderColor = gris;

                cel = new Cell(new Phrase("CANTIDAD", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                datosPartidas.AddCell(cel);

                cel = new Cell(new Phrase("UNIDAD", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                datosPartidas.AddCell(cel);

                cel = new Cell(new Phrase("CÓDIGO", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                datosPartidas.AddCell(cel);

                cel = new Cell(new Phrase("DESCRIPCIÓN", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.Colspan = 2;
                datosPartidas.AddCell(cel);

                cel = new Cell(new Phrase("PRECIO UNITARIO", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                datosPartidas.AddCell(cel);

                cel = new Cell(new Phrase("IMPORTE", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                datosPartidas.AddCell(cel);

                #endregion

                #region "Construimos Contenido de las Partidas"

                Table partidas = new Table(7);
                float[] headerwidthsPartidas = { 8, 6, 11, 26, 27, 12, 9 };
                partidas.Widths = headerwidthsPartidas;
                partidas.WidthPercentage = 100;
                partidas.Padding = 0;
                partidas.Spacing = 0;
                partidas.BorderWidth = 0;
                partidas.DefaultCellBorder = 0;
                partidas.BorderColor = gris;

                if (dtEncabezado.Rows.Count > 0)
                {
                    for (int i = 0; i < electronicDocument.Data.Conceptos.Count; i++)
                    {
                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Cantidad.Value.ToString(), f7));
                        cel.VerticalAlignment = Element.ALIGN_TOP;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Unidad.Value, f7));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].NumeroIdentificacion.Value, f7));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, f7));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.Colspan = 2;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].ValorUnitario.Value.ToString("C", _ci), f7));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("C", _ci), f7));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        partidas.AddCell(cel);

                        if (electronicDocument.Data.Conceptos[i].InformacionAduanera.IsAssigned)
                        {
                            for (int j = 0; j < electronicDocument.Data.Conceptos[i].InformacionAduanera.Count; j++)
                            {
                                StringBuilder sbInfoAduanera = new StringBuilder();

                                sbInfoAduanera.Append("ADUANA: " + electronicDocument.Data.Conceptos[i].InformacionAduanera[j].Aduana.Value.ToString() + "       PEDIMENTO: " + electronicDocument.Data.Conceptos[i].InformacionAduanera[j].Numero.Value.ToString() + "       FECHA: " + electronicDocument.Data.Conceptos[i].InformacionAduanera[j].Fecha.Value.ToString());

                                cel = new Cell(new Phrase(sbInfoAduanera.ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                                cel.Border = 0;
                                cel.Colspan = 7;
                                partidas.AddCell(cel);
                            }
                        }
                    }
                }
                cel = new Cell(new Phrase("A", new Font(Font.HELVETICA, 2, Font.NORMAL, blanco)));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                partidas.AddCell(cel);

                #endregion

                #endregion

                #region "Construimos Tabla de Datos CFDI"

                DefaultSplitCharacter split = new DefaultSplitCharacter();
                Table adicional = new Table(4);
                float[] headerwidthsAdicional = { 30, 40, 20, 10 };
                adicional.Widths = headerwidthsAdicional;
                adicional.WidthPercentage = 100;
                adicional.Padding = 0;
                adicional.Spacing = 0;
                adicional.BorderWidth = 0;
                adicional.DefaultCellBorder = 0;
                adicional.BorderColor = gris;

                if (timbrar)
                {
                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("CADENA ORIGINAL DEL COMPLEMENTO DE CERTIFICACIÓN DIGITAL DEL SAT\n", f5B));
                    par.Add(new Chunk(electronicDocument.FingerPrintPac, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 4;
                    adicional.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("SELLO DIGITAL DEL EMISOR\n", f5B));
                    par.Add(new Chunk(electronicDocument.Data.Sello.Value, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 4;
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
                    cel.Colspan = 4;
                    adicional.AddCell(cel);
                }

                if (htCFDI["tipoDoc"].ToString() == "NOTA DE CRÉDITO")
                {
                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("\n-TODA ACLARACIÓN O RECLAMACIÓN DEBERA HACERSE DENTRO DE LAS 48 HORAS SIGUIENTES A LA FECHA DE RECEPCIÓN DE LA MERCANCÍA\n", f5));
                    par.Add(new Chunk("-TODO CHEQUE DEVUELTO CAUSARÁ UN CARGO DEL 20% ADICIONAL SOBRE EL VALOR DEL MISMO DE ACUERDO AL ARTICULO 193 DE LA LEY GENERAL\n ", f5));
                    par.Add(new Chunk(" DE TITULOS Y OPERACIONES DE CRÉDITO.", f5));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 4;
                    adicional.AddCell(cel);
                }

                if (htCFDI["tipoDoc"].ToString() == "FACTURA")
                {
                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("\nDEPOSITE A : ", f6B));
                    par.Add(new Chunk("BANCOMER: 0142346478", f6B));
                    par.Add(new Chunk("         CUENTA CLABE: 012180001423464787", f6B));
                    par.Add(new Chunk("\nBANAMEX: 128411 SUC. 4087", f6B));
                    par.Add(new Chunk("  CUENTA CLABE: 002180408701284111", f6B));
                    par.Add(new Chunk("\nA NOMBRE DE JIGAFRA S.A. DE C.V.", f6B));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 4;
                    adicional.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk("-TODA ACLARACIÓN O RECLAMACIÓN DEBERA HACERSE DENTRO DE LAS 48 HORAS SIGUIENTES A LA FECHA DE RECEPCIÓN DE LA MERCANCÍA\n", f5));
                    par.Add(new Chunk("-TODO CHEQUE DEVUELTO CAUSARÁ UN CARGO DEL 20% ADICIONAL SOBRE EL VALOR DEL MISMO DE ACUERDO AL ARTICULO 193 DE LA LEY GENERAL\n ", f5));
                    par.Add(new Chunk(" DE TITULOS Y OPERACIONES DE CRÉDITO.", f5));
                    cel = new Cell(par);
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 4;
                    adicional.AddCell(cel);
                }
                string cuenta = electronicDocument.Data.NumeroCuentaPago.IsAssigned
                                ? electronicDocument.Data.NumeroCuentaPago.Value
                                : "";

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                par.Add(new Chunk("\nTIPO DE PAGO: ", f6));
                par.Add(new Chunk(electronicDocument.Data.MetodoPago.Value, f6));

                if (cuenta.Length > 0)
                {
                    par.Add(new Chunk("\nNÚMERO DE CUENTA: ", f6));
                    par.Add(new Chunk(cuenta, f6));
                }

                par.Add(new Chunk("\nCANTIDAD CON LETRA:\n", f6B));
                par.Add(new Chunk(dtEncabezado.Rows[0]["CANTIDAD-LETRA"].ToString(), f6));

                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                adicional.AddCell(cel);

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

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 8f);
                par.Add(new Chunk("Subtotal:\n ", f7B));
                par.Add(new Chunk("Descuento:\n ", f7B));
                par.Add(new Chunk("Subtotal:\n ", f7B));
                par.Add(new Chunk("I.V.A: ", f7B));
                par.Add(new Chunk(tasa + " %\n ", f7B));
                par.Add(new Chunk("Total:", f7B));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                double valorSubtotal = 0;
                valorSubtotal = Convert.ToDouble(electronicDocument.Data.SubTotal.Value) - Convert.ToDouble(electronicDocument.Data.Descuento.Value);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 8f);
                par.Add(new Chunk(electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f7));
                par.Add(new Chunk("\n" + electronicDocument.Data.Descuento.Value.ToString("C", _ci), f7));
                par.Add(new Chunk("\n" + valorSubtotal.ToString("C", _ci), f7));
                par.Add(new Chunk("\n" + electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("C", _ci), f7));
                par.Add(new Chunk("\n" + electronicDocument.Data.Total.Value.ToString("C", _ci), f7));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                adicional.AddCell(cel);


                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 8f);
                if (htCFDI["tipoDoc"].ToString() == "FACTURA")
                {
                    par.Add(new Chunk("LUGAR DE EXPEDICIÓN:  Mexico DF. a ", f6));
                    par.Add(new Chunk("     " + fecha.ToString("dd/MM/yyyy") + "     ", f6U));
                    par.Add(new Chunk("\nPor este pagaré señalo que debo y pagaré incondicionalmente a la orden de JIGAFRA S.A. de C.V. en México\nD.F. el día", f6));
                    par.Add(new Chunk("     " + dtEncabezado.Rows[0]["FECHA_DE_VENCIMIENTO"] + "     ", f6U));
                    par.Add(new Chunk("la cantidad de", f6));
                    par.Add(new Chunk("   " + electronicDocument.Data.Total.Value.ToString("C", _ci) + " (" + dtEncabezado.Rows[0]["CANTIDAD-LETRA"].ToString() + ")  " + "" + " ", f4U));
                    par.Add(new Chunk("valor recibido a mi entera satisfacción, desde la fecha de vencimiento de este documento hasta el día de la liquidación correrá", f6));
                    par.Add(new Chunk(" un interés moratorio al tipo del", f6));
                    par.Add(new Chunk("   10%   ", f6U));
                    par.Add(new Chunk("mensual, pagadero en esta ciudad.\n", f6));
                    //par.Add(new Chunk(dtEncabezado.Rows[0]["CANTIDAD-LETRA"].ToString() + "\n", f6));
                    par.Add(new Chunk("\nNO ACEPTAMOS DEVOLUCIÓN", folio));
                }
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                adicional.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 8f);
                if (htCFDI["tipoDoc"].ToString() == "NOTA DE CRÉDITO")
                {
                    par.Add(new Chunk("\n\n***" + electronicDocument.Data.FormaPago.Value + "***\n", f6));
                    par.Add(new Chunk("EXIJA SU RECIBO DE PAGO\n", f6));
                    par.Add(new Chunk("Este documento es una representación impresa de un CFDI", folio));
                }
                if (htCFDI["tipoDoc"].ToString() == "FACTURA")
                {
                    par.Add(new Chunk("***" + electronicDocument.Data.FormaPago.Value + "***\n", f6));
                    par.Add(new Chunk("EXIJA SU RECIBO DE PAGO\n", f6));
                    par.Add(new Chunk("Este documento es una representación impresa de un CFDI", folio));
                }
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                adicional.AddCell(cel);


                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("                                                                      \n", f6U));
                par.Add(new Chunk(" ACEPTO DE CONFORMIDAD NOMBRE Y FIRMA ", f6));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 4;
                adicional.AddCell(cel);

                #endregion

                #endregion

                pageEventHandler.datosFiscales = datosFiscales;
                pageEventHandler.datosFiscales2 = datosFiscales2;
                pageEventHandler.datosOpcionales = datosOpcionales;
                pageEventHandler.datosPartidas = datosPartidas;

                document.Open();
                document.Add(partidas);
                document.Add(adicional);

            }
            catch (Exception ex)
            {
                string exe = ex.Message;
            }
        }

        #endregion


        public class pdfPageEventHandlerJIGAFRA : PdfPageEventHelper
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

            public Phrase p2 { get; set; }
            public Table datosFiscales { get; set; }
            public Table datosFiscales2 { get; set; }
            public Table datosOpcionales { get; set; }
            public Table datosPartidas { get; set; }

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

                document.Add(datosFiscales);
                document.Add(datosFiscales2);
                document.Add(datosOpcionales);
                document.Add(datosPartidas);

                base.OnEndPage(writer, document);

                string lblPagina = "Página ";
                string lblDe = " de ";
                //string lblFechaImpresion = "Fecha de Impresión ";

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
                //cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, lblFechaImpresion + PrintTime, pageSize.GetRight(30), pageSize.GetBottom(15), 0);
                cb.EndText();
            }

            public override void OnEndPage(PdfWriter writer, Document document)
            {

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

        public class DefaultSplitCharacter : ISplitCharacter
        {
            #region "ISplitCharacter"

            public static readonly ISplitCharacter DEFAULT = new DefaultSplitCharacter();

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
}