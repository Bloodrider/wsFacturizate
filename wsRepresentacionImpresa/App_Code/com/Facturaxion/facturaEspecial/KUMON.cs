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
    public class KUMON
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

        private static HttpContext HTC;
        private static String pathIMGLOGO;
        //private static Document document;
        private static PdfPCell cell;
        private static Cell cel;
        private static Paragraph par;
        private static Chunk dSaltoLinea;

        private static Color azul;
        private static Color azulClaro;
        private static Color blanco;
        private static Color Link;
        private static Color gris;
        private static Color grisOX;
        private static Color rojo;
        private static Color lbAzul;

        private static BaseFont EM;
        private static Font f5;
        private static Font f5B;
        private static Font f5R;
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
                StringBuilder sbConfigFact = new StringBuilder();
                StringBuilder sbConfigFactParms = new StringBuilder();
                _ci.NumberFormat.CurrencyDecimalDigits = 2;

                DAL dal = new DAL();
                ElectronicDocument electronicDocument = (ElectronicDocument)htFacturaxion["electronicDocument"];
                Data objTimbre = (Data)htFacturaxion["objTimbre"];
                bool timbrar = Convert.ToBoolean(htFacturaxion["timbrar"]);

                // Obtenemos el logo y plantilla

                StringBuilder sbLogo = new StringBuilder();

                sbLogo.
                    Append("SELECT S.rutaLogo, P.rutaEncabezado, P.rutaFooter ").
                    Append("FROM sucursales S ").
                    Append("LEFT OUTER JOIN tipoPlantillas P ON S.idTipoPlantilla = P.idTipoPlantilla AND P.ST = 1 ").
                    Append("WHERE idSucursal = @0 AND S.ST = 1");

                DataTable dtLogo = dal.QueryDT("DS_FE", sbLogo.ToString(), "F:I:" + htFacturaxion["idSucursalEmisor"], hc);

                string rutaLogo = dtLogo.Rows[0]["rutaLogo"].ToString();
                string rutaHeader = dtLogo.Rows[0]["rutaEncabezado"].ToString();
                string rutaFooter = dtLogo.Rows[0]["rutaFooter"].ToString();

                if (electronicDocument.Data.Moneda.Value == "USD")
                    htFacturaxion["idMoneda"] = 2;

                sbConfigFactParms.
                    Append("F:I:").Append(Convert.ToInt64(htFacturaxion["idSucursalEmisor"])).
                    Append(";").
                    Append("F:I:").Append(Convert.ToInt32(htFacturaxion["tipoComprobante"])).
                    Append(";").
                    Append("F:S:").Append(electronicDocument.Data.Total.Value).
                    Append(";").
                    Append("F:I:").Append(Convert.ToInt32(htFacturaxion["idMoneda"])).
                    Append(";").
                    Append("F:S:").Append(rutaLogo).
                    Append(";").
                    Append("F:S:").Append(rutaHeader).
                    Append(";").
                    Append("F:S:").Append(rutaFooter).
                    Append(";").
                    Append("F:I:").Append(Convert.ToInt64(htFacturaxion["idEmisor"]));

                sbConfigFact.
                    Append("DECLARE @idEmpresa AS INT;").
                    Append("IF EXISTS (SELECT * FROM configuracionFacturas WHERE idEmpresa = @7) ").
                    Append("SET @idEmpresa = @7; ").
                    Append("ELSE ").
                    Append("SET @idEmpresa = 0; ").
                    Append("SELECT rutaTemplateHeader, rutaTemplateFooter, S.rutaLogo, objDesc, posX, posY, fontSize, dbo.convertNumToTextFunction( @2, @3) AS cantidadLetra, ").
                    Append("logoPosX, logoPosY, headerPosX, headerPosY, footerPosX, footerPosY, conceptosColWidth, desgloseColWidth, S.nombreSucursal ").
                    Append("FROM configuracionFacturas CF ").
                    Append("LEFT OUTER JOIN sucursales S ON S.idSucursal = @0 ").
                    Append("LEFT OUTER JOIN configuracionFactDet CFD ON CF.idConFact = CFD.idConFact ").
                    Append("WHERE CF.ST = 1 AND CF.idEmpresa = @idEmpresa AND CF.idTipoComp = @1 AND idCFDProcedencia = 1 AND objDesc NOT LIKE 'nuevoLbl%' ");

                DataTable dtConfigFact = dal.QueryDT("DS_FE", sbConfigFact.ToString(), sbConfigFactParms.ToString(), hc);

                // Creamos el Objeto Documento
                Document document = new Document(PageSize.LETTER, 25, 25, 25, 40);
                document.AddAuthor("Facturaxion");
                document.AddCreator("r3Take");
                document.AddCreationDate();
                pdfPageEventHandler pageEventHandler = new pdfPageEventHandler();
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                writer.PageEvent = pageEventHandler;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                HTC = hc;
                azul = new Color(7, 157, 198);
                azulClaro = new Color(27, 180, 240);
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
                f5R = new Font(EM, 5, Font.NORMAL, rojo);
                f6 = new Font(EM, 6);
                f6B = new Font(EM, 6, Font.BOLD);
                f6L = new Font(EM, 6, Font.BOLD, Link);
                f5L = new Font(EM, 5, Font.BOLD, lbAzul);
                titulo = new Font(EM, 6, Font.BOLD, blanco);
                folio = new Font(EM, 6, Font.BOLD, rojo);
                dSaltoLinea = new Chunk("\n\n ");

                pageEventHandler.rutaImgFooter = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\KUMON\pieMod219.png";

                Chunk cSaltoLinea = new Chunk("\n");
                Chunk cLineaSpace = new Chunk(cSaltoLinea + "________________________________________________________________________________________________________________________________________________________________________", new Font(Font.HELVETICA, 6, Font.BOLD));
                Chunk cLineaDiv = new Chunk(cSaltoLinea + "________________________________________________________________________________________________________________________________________________________________________" + cSaltoLinea, new Font(Font.HELVETICA, 6, Font.BOLD));
                Chunk cDataSpacer = new Chunk("      |      ", new Font(Font.HELVETICA, 6, Font.BOLD));

                BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                BaseFont bfB = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                document.Open();
                PdfContentByte cb = writer.DirectContent;
                cb.BeginText();

                // Almacenamos en orden los objetos del Electronic Document para posteriormente añadidos al documento

                #region "Armamos las etiquetas que tienen posiciones absolutas para ser insertadas en el documento"

                Hashtable htDatosCfdi = new Hashtable();

                // Armamos las direcciones

                #region "Dirección Emisor"

                StringBuilder sbDirEmisor1 = new StringBuilder();
                StringBuilder sbDirEmisor2 = new StringBuilder();
                StringBuilder sbDirEmisor3 = new StringBuilder();

                if (electronicDocument.Data.Emisor.Domicilio.Calle.Value.Length > 0)
                {
                    sbDirEmisor1.Append("Calle ").Append(electronicDocument.Data.Emisor.Domicilio.Calle.Value).Append(" ");
                }

                if (electronicDocument.Data.Emisor.Domicilio.NumeroExterior.Value.Length > 0)
                {
                    sbDirEmisor1.Append(", No. Ext ").Append(electronicDocument.Data.Emisor.Domicilio.NumeroExterior.Value).Append(" ");
                }

                if (electronicDocument.Data.Emisor.Domicilio.NumeroInterior.Value.Length > 0)
                {
                    sbDirEmisor1.Append(", No. Int ").Append(electronicDocument.Data.Emisor.Domicilio.NumeroInterior.Value);
                }

                if (electronicDocument.Data.Emisor.Domicilio.Colonia.Value.Length > 0)
                {
                    sbDirEmisor2.Append("Col. ").Append(electronicDocument.Data.Emisor.Domicilio.Colonia.Value).Append(" ");
                }

                if (electronicDocument.Data.Emisor.Domicilio.CodigoPostal.Value.Length > 0)
                {
                    sbDirEmisor2.Append(", C.P. ").Append(electronicDocument.Data.Emisor.Domicilio.CodigoPostal.Value).Append(" ");
                }

                if (electronicDocument.Data.Emisor.Domicilio.Localidad.Value.Length > 0)
                {
                    sbDirEmisor2.Append(", ").Append(electronicDocument.Data.Emisor.Domicilio.Localidad.Value);
                }

                if (electronicDocument.Data.Emisor.Domicilio.Municipio.Value.Length > 0)
                {
                    sbDirEmisor3.Append("Mpio. / Del. ").Append(electronicDocument.Data.Emisor.Domicilio.Municipio.Value).Append(" ");
                }

                if (electronicDocument.Data.Emisor.Domicilio.Estado.Value.Length > 0)
                {
                    sbDirEmisor3.Append(", Estado ").Append(electronicDocument.Data.Emisor.Domicilio.Estado.Value).Append(" ");
                }

                sbDirEmisor3.Append(", ").Append(electronicDocument.Data.Emisor.Domicilio.Pais.Value);


                #endregion

                #region "Dirección Sucursal Expedido En"

                StringBuilder sbDirExpedido1 = new StringBuilder();
                StringBuilder sbDirExpedido2 = new StringBuilder();
                StringBuilder sbDirExpedido3 = new StringBuilder();

                if (electronicDocument.Data.Emisor.ExpedidoEn.Calle.Value.Length > 0)
                {
                    sbDirExpedido1.Append("Calle ").Append(electronicDocument.Data.Emisor.ExpedidoEn.Calle.Value).Append(" ");
                }

                if (electronicDocument.Data.Emisor.ExpedidoEn.NumeroExterior.Value.Length > 0)
                {
                    sbDirExpedido1.Append(", No. Ext ").Append(electronicDocument.Data.Emisor.ExpedidoEn.NumeroExterior.Value).Append(" ");
                }

                if (electronicDocument.Data.Emisor.ExpedidoEn.NumeroInterior.Value.Length > 0)
                {
                    sbDirExpedido1.Append(", No. Int ").Append(electronicDocument.Data.Emisor.ExpedidoEn.NumeroInterior.Value);
                }

                if (electronicDocument.Data.Emisor.ExpedidoEn.Colonia.Value.Length > 0)
                {
                    sbDirExpedido2.Append("Col. ").Append(electronicDocument.Data.Emisor.ExpedidoEn.Colonia.Value).Append(" ");
                }

                if (electronicDocument.Data.Emisor.ExpedidoEn.CodigoPostal.Value.Length > 0)
                {
                    sbDirExpedido2.Append(", C.P. ").Append(electronicDocument.Data.Emisor.ExpedidoEn.CodigoPostal.Value).Append(" ");
                }

                if (electronicDocument.Data.Emisor.ExpedidoEn.Localidad.Value.Length > 0)
                {
                    sbDirExpedido2.Append(", ").Append(electronicDocument.Data.Emisor.ExpedidoEn.Localidad.Value);
                }

                if (electronicDocument.Data.Emisor.ExpedidoEn.Municipio.Value.Length > 0)
                {
                    sbDirExpedido3.Append("Mpio. / Del. ").Append(electronicDocument.Data.Emisor.ExpedidoEn.Municipio.Value).Append(" ");
                }

                if (electronicDocument.Data.Emisor.ExpedidoEn.Estado.Value.Length > 0)
                {
                    sbDirExpedido3.Append(", Estado ").Append(electronicDocument.Data.Emisor.ExpedidoEn.Estado.Value).Append(" ");
                }

                sbDirExpedido3.Append(", ").Append(electronicDocument.Data.Emisor.ExpedidoEn.Pais.Value);

                #endregion

                #region "Dirección Receptor"

                StringBuilder sbDirReceptor1 = new StringBuilder();
                StringBuilder sbDirReceptor2 = new StringBuilder();
                StringBuilder sbDirReceptor3 = new StringBuilder();

                if (electronicDocument.Data.Receptor.Domicilio.Calle.Value.Length > 0)
                {
                    sbDirReceptor1.Append("Calle ").Append(electronicDocument.Data.Receptor.Domicilio.Calle.Value).Append(" ");
                }

                if (electronicDocument.Data.Receptor.Domicilio.NumeroExterior.Value.Length > 0)
                {
                    sbDirReceptor1.Append(", No. Ext ").Append(electronicDocument.Data.Receptor.Domicilio.NumeroExterior.Value).Append(" ");
                }

                if (electronicDocument.Data.Receptor.Domicilio.NumeroInterior.Value.Length > 0)
                {
                    sbDirReceptor1.Append(", No. Int ").Append(electronicDocument.Data.Receptor.Domicilio.NumeroInterior.Value);
                }

                if (electronicDocument.Data.Receptor.Domicilio.Colonia.Value.Length > 0)
                {
                    sbDirReceptor2.Append("Col. ").Append(electronicDocument.Data.Receptor.Domicilio.Colonia.Value).Append(" ");
                }

                if (electronicDocument.Data.Receptor.Domicilio.CodigoPostal.Value.Length > 0)
                {
                    sbDirReceptor2.Append(", C.P. ").Append(electronicDocument.Data.Receptor.Domicilio.CodigoPostal.Value).Append(" ");
                }

                if (electronicDocument.Data.Receptor.Domicilio.Localidad.Value.Length > 0)
                {
                    sbDirReceptor2.Append(", ").Append(electronicDocument.Data.Receptor.Domicilio.Localidad.Value);
                }

                if (electronicDocument.Data.Receptor.Domicilio.Municipio.Value.Length > 0)
                {
                    sbDirReceptor3.Append("Mpio. / Del. ").Append(electronicDocument.Data.Receptor.Domicilio.Municipio.Value).Append(" ");
                }

                if (electronicDocument.Data.Receptor.Domicilio.Estado.Value.Length > 0)
                {
                    sbDirReceptor3.Append(", Estado ").Append(electronicDocument.Data.Receptor.Domicilio.Estado.Value).Append(" ");
                }

                sbDirReceptor3.Append(", ").Append(electronicDocument.Data.Receptor.Domicilio.Pais.Value);

                #endregion

                htDatosCfdi.Add("rfcEmisor", electronicDocument.Data.Emisor.Rfc.Value);
                htDatosCfdi.Add("rfcEmpresa", electronicDocument.Data.Emisor.Rfc.Value);

                htDatosCfdi.Add("nombreEmisor", electronicDocument.Data.Emisor.Nombre.Value);
                htDatosCfdi.Add("empresa", electronicDocument.Data.Emisor.Nombre.Value);

                htDatosCfdi.Add("rfcReceptor", electronicDocument.Data.Receptor.Rfc.Value);
                htDatosCfdi.Add("rfcCliente", electronicDocument.Data.Receptor.Rfc.Value);

                htDatosCfdi.Add("nombreReceptor", electronicDocument.Data.Receptor.Nombre.Value);
                htDatosCfdi.Add("cliente", electronicDocument.Data.Receptor.Nombre.Value);

                htDatosCfdi.Add("sucursal", "Sucursal " + dtConfigFact.Rows[0]["nombreSucursal"]);

                htDatosCfdi.Add("serie", electronicDocument.Data.Serie.Value);
                htDatosCfdi.Add("folio", electronicDocument.Data.Folio.Value);

                htDatosCfdi.Add("fechaCfdi", electronicDocument.Data.Fecha.Value);
                htDatosCfdi.Add("fechaFactura", electronicDocument.Data.Fecha.Value);

                htDatosCfdi.Add("UUID", objTimbre.Uuid.Value);
                htDatosCfdi.Add("folioFiscal", objTimbre.Uuid.Value);

                htDatosCfdi.Add("direccionEmisor1", sbDirEmisor1.ToString());
                htDatosCfdi.Add("direccionEmpresa1", sbDirEmisor1.ToString());

                htDatosCfdi.Add("direccionEmisor2", sbDirEmisor2.ToString());
                htDatosCfdi.Add("direccionEmpresa2", sbDirEmisor2.ToString());

                htDatosCfdi.Add("direccionEmisor3", sbDirEmisor3.ToString());
                htDatosCfdi.Add("direccionEmpresa3", sbDirEmisor3.ToString());

                htDatosCfdi.Add("direccionExpedido1", sbDirExpedido1.ToString());
                htDatosCfdi.Add("direccionSucursal1", sbDirExpedido1.ToString());

                htDatosCfdi.Add("direccionExpedido2", sbDirExpedido2.ToString());
                htDatosCfdi.Add("direccionSucursal2", sbDirExpedido2.ToString());

                htDatosCfdi.Add("direccionExpedido3", sbDirExpedido3.ToString());
                htDatosCfdi.Add("direccionSucursal3", sbDirExpedido3.ToString());

                htDatosCfdi.Add("direccionReceptor1", sbDirReceptor1.ToString());
                htDatosCfdi.Add("direccionCliente1", sbDirReceptor1.ToString());

                htDatosCfdi.Add("direccionReceptor2", sbDirReceptor2.ToString());
                htDatosCfdi.Add("direccionCliente2", sbDirReceptor2.ToString());

                htDatosCfdi.Add("direccionReceptor3", sbDirReceptor3.ToString());
                htDatosCfdi.Add("direccionCliente3", sbDirReceptor3.ToString());

                // Leemos los objetos que se situaran en posiciones absolutas en el documento
                //foreach (DataRow row in dtConfigFact.Rows)
                //{
                //    cb.SetFontAndSize(bf, Convert.ToInt32(row["fontSize"]));
                //    cb.SetTextMatrix(Convert.ToSingle(row["posX"]), Convert.ToSingle(row["posY"]));
                //    cb.ShowText(htDatosCfdi[row["objDesc"].ToString()].ToString());
                //}

                string tipoComp = electronicDocument.Data.TipoComprobante.Value.ToString();

                if (tipoComp.Length > 0)
                    tipoComp = tipoComp.Substring(0, 1).ToUpper() + tipoComp.Substring(1);

                cb.SetFontAndSize(bfB, 9);
                cb.SetTextMatrix(518, 763);
                cb.ShowText(tipoComp);

                cb.SetFontAndSize(bfB, 8);
                cb.SetTextMatrix(400, 749);
                cb.ShowText(htDatosCfdi["UUID"].ToString());

                cb.SetFontAndSize(bf, 8);
                cb.SetTextMatrix(450, 716);
                cb.ShowText("México, D.F.");

                cb.SetFontAndSize(bfB, 8);
                cb.SetColorFill(iTextSharp.text.Color.RED);
                cb.SetTextMatrix(527, 717);
                cb.ShowText(htDatosCfdi["folio"].ToString());

                cb.SetFontAndSize(bfB, 8);
                cb.SetTextMatrix(455, 649);
                cb.ShowText(electronicDocument.Data.NumeroCertificado.Value);

                cb.SetFontAndSize(bf, 8);
                cb.SetColorFill(iTextSharp.text.Color.BLACK);
                cb.SetTextMatrix(461, 686);
                cb.ShowText(htDatosCfdi["fechaCfdi"].ToString());

                cb.SetFontAndSize(bfB, 8);
                cb.SetTextMatrix(40, 725);
                cb.ShowText("Kumon Instituto de Educación, S.A. de C.V.");

                cb.SetFontAndSize(bfB, 6);
                cb.SetTextMatrix(40, 715);
                cb.ShowText("MATRIZ");

                cb.SetTextMatrix(40, 705);
                cb.ShowText("Arquímedes No.130,Piso 7 y 8-801,");

                cb.SetTextMatrix(40, 695);
                cb.ShowText("Col.Polanco, Del. Miguel Hidalgo, México, D.F. 11560");

                cb.SetTextMatrix(40, 682);
                cb.ShowText(htDatosCfdi["rfcEmisor"].ToString());

                cb.SetTextMatrix(40, 649);
                cb.ShowText(htDatosCfdi["nombreReceptor"].ToString());

                cb.SetFontAndSize(bf, 6);
                cb.SetTextMatrix(40, 635);
                cb.ShowText(htDatosCfdi["direccionReceptor1"].ToString());

                cb.SetTextMatrix(40, 625);
                cb.ShowText(htDatosCfdi["direccionReceptor2"].ToString());

                cb.SetTextMatrix(40, 615);
                cb.ShowText(htDatosCfdi["direccionReceptor3"].ToString());

                cb.SetTextMatrix(40, 605);
                cb.ShowText(htDatosCfdi["rfcReceptor"].ToString());

                cb.SetFontAndSize(bfB, 6);
                cb.SetTextMatrix(200, 715);
                cb.ShowText("LOCAL");

                cb.SetTextMatrix(200, 705);
                cb.ShowText("Centro Kumon Newton");

                cb.SetTextMatrix(200, 695);
                cb.ShowText("Homero No.301 Mezzanine, Col Chapultepec Morales");

                cb.SetTextMatrix(200, 685);
                cb.ShowText("Del.Miguel Hidalgo, México, D.F. 11570");

                #endregion

                cb.EndText();

                #region "Header"

                //Agregando Imagen de Encabezado de Página
                Image imgHeader = Image.GetInstance(@"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\KUMON\encabezadoMod226.png");
                imgHeader.ScalePercent(47f);

                double posXH = Convert.ToDouble(34);
                double posYH = Convert.ToDouble(577);

                double PXH = posXH;
                double PYH = posYH;
                imgHeader.SetAbsolutePosition(Convert.ToSingle(PXH), Convert.ToSingle(PYH));
                document.Add(imgHeader);

                #endregion

                #region "Logotipo"

                //Agregando Imagen de Logotipo
                Image imgLogo = Image.GetInstance(@"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\KUMON\logo-kumon.png");
                float imgLogoWidth = 90;
                float imgLogoHeight = 28;

                imgLogo.ScaleAbsolute(imgLogoWidth, imgLogoHeight);
                imgLogo.SetAbsolutePosition(50, 735);
                document.Add(imgLogo);

                #endregion

                #region "Espaciador entre Header y Conceptos"

                Paragraph pRelleno = new Paragraph();
                Chunk cRelleno = new Chunk();

                for (int lineas = 0; lineas < 12; lineas++)
                {
                    cRelleno.Append("\n");
                }

                pRelleno.Add(cRelleno);
                document.Add(pRelleno);

                #endregion

                #region "Añadimos Detalle de Conceptos"

                // Creamos la tabla para insertar los conceptos de detalle de la factura
                PdfPTable tableConceptos = new PdfPTable(5);

                int[] colWithsConceptos = new int[5];
                String[] arrColWidthConceptos = { "5", "9", "46", "20", "20" };

                for (int i = 0; i < arrColWidthConceptos.Length; i++)
                {
                    colWithsConceptos.SetValue(Convert.ToInt32(arrColWidthConceptos[i]), i);
                }

                tableConceptos.SetWidths(colWithsConceptos);
                tableConceptos.WidthPercentage = 93F;

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
                    //cellConceptos = new PdfPCell(new Phrase("Colegiatura", new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellConceptos.Border = 0;
                    cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                    tableConceptos.AddCell(cellConceptos);

                    cellConceptos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellConceptos.Border = 0;
                    tableConceptos.AddCell(cellConceptos);

                    cellMontos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].ValorUnitario.Value.ToString("C", _ci), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellMontos.Border = 0;
                    cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                    tableConceptos.AddCell(cellMontos);

                    cellMontos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("C", _ci), new Font(Font.HELVETICA, 8, Font.NORMAL)));
                    cellMontos.Border = 0;
                    cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                    tableConceptos.AddCell(cellMontos);

                    if (electronicDocument.Data.Conceptos[i].InformacionAduanera.IsAssigned)
                    {
                        for (int j = 0; j < electronicDocument.Data.Conceptos[i].InformacionAduanera.Count; j++)
                        {
                            StringBuilder sbInfoAduanera = new StringBuilder();

                            sbInfoAduanera.Append("PEDIMENTO: " + electronicDocument.Data.Conceptos[i].InformacionAduanera[j].Numero.Value.ToString() + "           FECHA: " + electronicDocument.Data.Conceptos[i].InformacionAduanera[j].Fecha.Value.ToString() + "\nADUANA: " + electronicDocument.Data.Conceptos[i].InformacionAduanera[j].Aduana.Value.ToString());

                            cellConceptos = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL)));
                            cellConceptos.Border = 0;
                            cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                            tableConceptos.AddCell(cellConceptos);

                            cellConceptos = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL)));
                            cellConceptos.Border = 0;
                            cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                            tableConceptos.AddCell(cellConceptos);

                            cellConceptos = new PdfPCell(new Phrase(sbInfoAduanera.ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                            cellConceptos.Border = 0;
                            tableConceptos.AddCell(cellConceptos);

                            cellMontos = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL)));
                            cellMontos.Border = 0;
                            cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                            tableConceptos.AddCell(cellMontos);

                            cellMontos = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL)));
                            cellMontos.Border = 0;
                            cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                            tableConceptos.AddCell(cellMontos);
                        }
                    }
                }

                document.Add(tableConceptos);

                #endregion

                #region "Espaciador entre Conceptos y Desglose"

                Paragraph pRelleno2 = new Paragraph();
                Chunk cRelleno2 = new Chunk("\n");
                pRelleno2.Add(cRelleno2);
                document.Add(pRelleno2);

                #endregion

                #region "Desglose"

                PdfPTable tableDesglose = new PdfPTable(4);
                int[] colWithsDesglose = new int[4];
                String[] arrColWidthDesglose = { "55", "15", "20", "15" };

                for (int i = 0; i < arrColWidthDesglose.Length; i++)
                {
                    colWithsDesglose.SetValue(Convert.ToInt32(arrColWidthDesglose[i]), i);
                }

                tableDesglose.SetWidths(colWithsDesglose);
                tableDesglose.WidthPercentage = 93F;

                PdfPCell cellDesgloseRelleno = new PdfPCell();
                PdfPCell cellDesgloseRelleno1 = new PdfPCell();
                PdfPCell cellDesgloseDescripcion = new PdfPCell();
                PdfPCell cellDesgloseMonto = new PdfPCell();
                PdfPCell cellCantidadLetra = new PdfPCell();
                PdfPCell cellCantidadLetra1 = new PdfPCell();
                PdfPCell cellCantidadLetra2 = new PdfPCell();
                PdfPCell cellCantidadLetra3 = new PdfPCell();

                //Armamnos el Hashtable que conntiene el desglose del Cfdi

                Hashtable htDesglose = new Hashtable();
                ArrayList alDesgloseOrden = new ArrayList();

                alDesgloseOrden.Add("Subtotal");

                if (electronicDocument.Data.Descuento.Value != 0)
                {
                    alDesgloseOrden.Add("Descuento");
                }

                if (electronicDocument.Data.Impuestos.TotalTraslados.Value != 0)
                {
                    alDesgloseOrden.Add("IVA 16%");
                }

                if (electronicDocument.Data.Impuestos.TotalRetenciones.Value != 0)
                {
                    alDesgloseOrden.Add("Impuestos Retenidos");
                }

                alDesgloseOrden.Add("Total");

                htDesglose.Add("Subtotal", electronicDocument.Data.SubTotal.Value.ToString("C", _ci));
                htDesglose.Add("Descuento", electronicDocument.Data.Descuento.Value.ToString("C", _ci));
                htDesglose.Add("IVA 16%", electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("C", _ci));
                htDesglose.Add("Impuestos Retenidos", electronicDocument.Data.Impuestos.TotalRetenciones.Value.ToString("C", _ci));
                htDesglose.Add("Total", electronicDocument.Data.Total.Value.ToString("C", _ci));

                foreach (string desglose in alDesgloseOrden)
                {
                    cellDesgloseRelleno = new PdfPCell(new Phrase(string.Empty, new Font(Font.HELVETICA, 7, Font.BOLD)));
                    cellDesgloseRelleno.Border = 0;

                    cellDesgloseDescripcion = new PdfPCell(new Phrase(desglose, new Font(Font.HELVETICA, 7, Font.BOLD)));
                    cellDesgloseDescripcion.Border = 0;
                    cellDesgloseDescripcion.BackgroundColor = azulClaro;

                    cellDesgloseRelleno1 = new PdfPCell(new Phrase(string.Empty, new Font(Font.HELVETICA, 7, Font.BOLD)));
                    cellDesgloseRelleno1.Border = 0;

                    cellDesgloseMonto = new PdfPCell(new Phrase(htDesglose[desglose].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellDesgloseMonto.Border = 0;
                    cellDesgloseMonto.HorizontalAlignment = PdfCell.ALIGN_RIGHT;

                    if (desglose == "Total")
                    {
                        cellDesgloseRelleno = new PdfPCell(new Phrase("Cantidad con Letra", new Font(Font.HELVETICA, 7, Font.BOLD)));
                        cellDesgloseRelleno.BackgroundColor = azulClaro;
                        cellDesgloseRelleno.Border = 0;

                        cellDesgloseMonto.BorderWidthTop = 1;
                    }

                    tableDesglose.AddCell(cellDesgloseRelleno);
                    tableDesglose.AddCell(cellDesgloseRelleno1);
                    tableDesglose.AddCell(cellDesgloseDescripcion);
                    tableDesglose.AddCell(cellDesgloseMonto);
                }

                string cantidadLetra = dtConfigFact.Rows[0]["cantidadLetra"].ToString();

                cellCantidadLetra = new PdfPCell(new Phrase(cantidadLetra, new Font(Font.HELVETICA, 7, Font.BOLD)));
                cellCantidadLetra.Border = 0;

                cellCantidadLetra1 = new PdfPCell(new Phrase(string.Empty, new Font(Font.HELVETICA, 7, Font.BOLD)));
                cellCantidadLetra1.Border = 0;

                cellCantidadLetra2 = new PdfPCell(new Phrase(string.Empty, new Font(Font.HELVETICA, 7, Font.BOLD)));
                cellCantidadLetra2.Border = 0;

                cellCantidadLetra3 = new PdfPCell(new Phrase(string.Empty, new Font(Font.HELVETICA, 7, Font.BOLD)));
                cellCantidadLetra3.Border = 0;

                tableDesglose.AddCell(cellCantidadLetra);
                tableDesglose.AddCell(cellCantidadLetra1);
                tableDesglose.AddCell(cellCantidadLetra2);
                tableDesglose.AddCell(cellCantidadLetra3);

                document.Add(tableDesglose);

                #endregion

                //#region "Construimos el Desglose de Impuestos"

                //Table desgloseImpuestos = new Table(3);
                //float[] headerwidthsDesgloce = { 15, 15, 70 };
                //desgloseImpuestos.Widths = headerwidthsDesgloce;
                //desgloseImpuestos.WidthPercentage = 95;
                //desgloseImpuestos.Padding = 1;
                //desgloseImpuestos.Spacing = 1;
                //desgloseImpuestos.BorderWidth = 0;
                //desgloseImpuestos.DefaultCellBorder = 0;
                //desgloseImpuestos.BorderColor = azulClaro;

                //cel = new Cell(new Phrase("Desgloce de Impuestos", f6B));
                //cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                //cel.HorizontalAlignment = Element.ALIGN_CENTER;
                //cel.BorderWidthTop = 1;
                //cel.BorderWidthLeft = 1;
                //cel.BorderWidthRight = 1;
                //cel.BorderWidthBottom = 1;
                //cel.BorderColor = azulClaro;
                //cel.Colspan = 2;
                //desgloseImpuestos.AddCell(cel);

                //cel = new Cell(new Phrase("", f6B));
                //cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                //cel.BorderWidthTop = 0;
                //cel.BorderWidthLeft = 0;
                //cel.BorderWidthRight = 0;
                //cel.BorderWidthBottom = 0;
                //desgloseImpuestos.AddCell(cel);

                //for (int i = 0; i < electronicDocument.Data.Impuestos.Traslados.Count; i++)
                //{
                //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Traslados[i].Tipo.Value.ToString() + " " + electronicDocument.Data.Impuestos.Traslados[i].Tasa.Value.ToString() + "%", f6B));
                //    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthLeft = 1;
                //    cel.BorderWidthRight = 1;
                //    cel.BorderWidthBottom = 1;
                //    cel.BorderColor = azulClaro;
                //    desgloseImpuestos.AddCell(cel);

                //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Traslados[i].Importe.Value.ToString("C", _ci), f6B));
                //    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthRight = 1;
                //    cel.BorderWidthBottom = 1;
                //    cel.BorderColor = azulClaro;
                //    desgloseImpuestos.AddCell(cel);

                //    cel = new Cell(new Phrase("", f6B));
                //    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthRight = 0;
                //    cel.BorderWidthBottom = 0;
                //    cel.BorderColor = azulClaro;
                //    desgloseImpuestos.AddCell(cel);
                //}

                //for (int i = 0; i < electronicDocument.Data.Impuestos.Retenciones.Count; i++)
                //{
                //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Retenciones[i].Tipo.Value.ToString(), f6B));
                //    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthLeft = 1;
                //    cel.BorderWidthRight = 1;
                //    cel.BorderWidthBottom = 1;
                //    cel.BorderColor = azulClaro;
                //    desgloseImpuestos.AddCell(cel);

                //    cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.Retenciones[i].Importe.Value.ToString("C", _ci), f6B));
                //    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthRight = 1;
                //    cel.BorderWidthBottom = 1;
                //    cel.BorderColor = azulClaro;
                //    desgloseImpuestos.AddCell(cel);

                //    cel = new Cell(new Phrase("", f6B));
                //    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                //    cel.BorderWidthTop = 0;
                //    cel.BorderWidthLeft = 0;
                //    cel.BorderWidthRight = 0;
                //    cel.BorderWidthBottom = 0;
                //    cel.BorderColor = azulClaro;
                //    desgloseImpuestos.AddCell(cel);
                //}

                ////document.Add(desgloseImpuestos);

                //#endregion

                #region "Construimos Tabla de Datos CFDI"

                DefaultSplitCharacter split = new DefaultSplitCharacter();
                Table adicional = new Table(3);
                float[] headerwidthsAdicional = { 20, 25, 55 };
                adicional.Widths = headerwidthsAdicional;
                adicional.WidthPercentage = 100;
                adicional.Padding = 1;
                adicional.Spacing = 1;
                adicional.BorderWidth = 2;
                adicional.DefaultCellBorder = 2;
                adicional.BorderColor = grisOX;

                //string cantidadLetra = dtConfigFact.Rows[0]["cantidadLetra"].ToString();

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                par.Add(new Chunk("", f5B));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                par = new Paragraph();
                par.SetLeading(7f, 0f);
                par.Add(new Chunk("", f5B));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                adicional.AddCell(cel);

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
                    cel.BorderWidthBottom = (float).5;
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
                    cel.BorderWidthBottom = 0;
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
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase("FECHA Y HORA DE CERTIFICACION:", f5B));
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

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL SAT:", f5B));
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

                    cel = new Cell(new Phrase("No. DE SERIE DEL CERTIFICADO DEL EMISOR:", f5B));
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
                    par.Add(new Chunk("TIPO DE COMPROBANTE: ", f5B));
                    par.Add(new Chunk(electronicDocument.Data.TipoComprobante.Value + "   |   ", f5));
                    par.Add(new Chunk("MONEDA: ", f5B));
                    par.Add(new Chunk(electronicDocument.Data.Moneda.Value + "   |   ", f5));
                    par.Add(new Chunk("TASA DE CAMBIO: ", f5B));

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
                        cel.BorderWidthTop = 0;
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
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
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

                    cel = new Cell(new Phrase("Considere porfavor su responsabilidad ambiental antes de imprimir este documento\nwww.kumon.com.mx", f5B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
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
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                footer.AddCell(cell);

                #endregion

                //document.Add(desgloseImpuestos);
                document.Add(adicional);
                document.Add(footer);

                document.Close();
                writer.Close();

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
    }
}