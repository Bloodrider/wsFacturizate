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
    public class OVMON
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
                Int64 idCfdi = Convert.ToInt64(htFacturaxion["idCfdi"]);

                #region "Extraemos los datos del CFDI"

                StringBuilder sbConfigFact = new StringBuilder();
                StringBuilder sbConfigFactParms = new StringBuilder();

                sbConfigFactParms.
                    Append("F:I:").Append(Convert.ToInt64(htFacturaxion["idSucursalEmisor"])).
                    Append(";").
                    Append("F:I:").Append(Convert.ToInt32(htFacturaxion["tipoComprobante"])).
                    Append(";").
                    Append("F:S:").Append(electronicDocument.Data.Total.Value).
                    Append(";").
                    Append("F:I:").Append(Convert.ToInt32(htFacturaxion["idMoneda"])).
                    Append(";").
                    Append("F:I:").Append(Convert.ToInt64(htFacturaxion["idEmisor"]));

                sbConfigFact.
                    Append("DECLARE @idEmpresa AS INT; SET @idEmpresa = 0; ").
                    Append("SELECT rutaTemplateHeader, CF.rutaLogo, objDesc, posX, posY, fontSize, dbo.convertNumToTextFunction( @2, @3) AS cantidadLetra, ").
                    Append("logoPosX, logoPosY, headerPosX, headerPosY, footerPosX, footerPosY, conceptosColWidth, desgloseColWidth, S.nombreSucursal ").
                    Append("FROM configuracionFacturas CF ").
                    Append("LEFT OUTER JOIN sucursales S ON S.idSucursal = @0 ").
                    Append("LEFT OUTER JOIN configuracionFactDet CFD ON CF.idConFact = CFD.idConFact ").
                    Append("WHERE CF.ST = 1 AND CF.idEmpresa = -1 AND CF.idTipoComp = @1 AND idCFDProcedencia = 1 AND objDesc NOT LIKE 'nuevoLbl%' ");

                DataTable dtConfigFact = dal.QueryDT("DS_FE", sbConfigFact.ToString(), sbConfigFactParms.ToString(), hc);

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

                Document documento = new Document(PageSize.LETTER, 15, 15, 15, 40);
                documento.AddAuthor("Facturaxion");
                documento.AddCreator("r3Take");
                documento.AddCreationDate();

                PdfWriter writer = PdfWriter.GetInstance(documento, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                Chunk cSaltoLinea = new Chunk("\n");
                Chunk cDataSpacer = new Chunk("      |      ", new Font(Font.HELVETICA, 6, Font.BOLD));

                //Creamos Colores para el PDF
                Color azul = new Color(37, 133, 230);
                Color blanco = new Color(255, 255, 255);
                Color Link = new Color(7, 73, 208);
                Color gris = new Color(240, 240, 240);
                Color rojo = new Color(208, 7, 7);

                //BaseFont EM = BaseFont.CreateFont(hc.Request.PhysicalApplicationPath + "estilos" + "\\verdana.ttf", BaseFont.WINANSI, BaseFont.EMBEDDED);
                BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                BaseFont bfB = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                //Creamos Fuentes
                Font f6 = new Font(bf, 6);
                Font f6B = new Font(bf, 6, Font.BOLD, blanco);
                Font f7 = new Font(bf, 7);
                Font f7B = new Font(bf, 7, Font.BOLD, blanco);
                Font f7BB = new Font(bf, 7, Font.BOLD);
                Font titulo = new Font(bf, 6, Font.BOLD, blanco);
                Font f8 = new Font(bf, 8);
                Font f8B = new Font(bf, 8, Font.BOLD);
                Font f9 = new Font(bf, 9);
                Font f9B = new Font(bf, 9, Font.BOLD);

                Cell cell;
                documento.Open();

                PdfContentByte cb = writer.DirectContent;
                cb.BeginText();

                #endregion

                #region "Generamos el Docuemto"

                #region "insertamos Objetos del PDF"

                //Extraemos informacion de datos opcionales
                StringBuilder sbOpcionalEncabezado = new StringBuilder();
                DataTable dtOpcEnc = new DataTable();

                sbOpcionalEncabezado.
                    Append("SELECT campo1 AS datosEstudiante ").
                    Append("FROM opcionalEncabezado WHERE idCFDI = @0 ");

                dtOpcEnc = dal.QueryDT("DS_FE", sbOpcionalEncabezado.ToString(), "F:I:" + idCfdi, hc);

                string datosEstudiante = dtOpcEnc.Rows[0]["datosEstudiante"].ToString();
                string serieFolio = htDatosCfdi["serie"].ToString() + " - " + htDatosCfdi["folio"].ToString();

                //Insertamos Tabla que contendra los datos del PDF
                Table tEncabezadoDetail = new Table(4);
                tEncabezadoDetail.WidthPercentage = 95;
                tEncabezadoDetail.SetWidths(new int[4] { 25, 25, 25, 20 });
                tEncabezadoDetail.DefaultCell.Border = 0;
                tEncabezadoDetail.BorderColor = blanco;
                tEncabezadoDetail.Padding = 1;
                tEncabezadoDetail.Spacing = 1;

                //Agregando Imagen de Logotipo
                Image imgLogo = Image.GetInstance(@"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\OVM800704940\logo_OVMON.png");
                imgLogo.ScalePercent(65f);

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
                tEncabezadoDetail.AddCell(cel);

                cell = new Cell(new Phrase(htDatosCfdi["nombreEmisor"].ToString(), f8B));
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                cell.Colspan = 3;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("RFC: " + htDatosCfdi["rfcEmisor"].ToString() + "\n" + htDatosCfdi["direccionEmisor1"].ToString() + htDatosCfdi["direccionEmisor2"].ToString() + "\n" + htDatosCfdi["direccionEmisor3"].ToString(), f8));
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                cell.Colspan = 3;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("Tels. (55)5586-8212 y 5586-5955", f8));
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                cell.Colspan = 3;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("INCORPORACIÓN SEP", f7B));
                cell.VerticalAlignment = Element.ALIGN_CENTER;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                cell.Colspan = 4;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("Primaria Latacunga\nAcuerdo No. 00911871\nde Fecha 25 de Julio de 1991\na Nivel Primaria", f7));
                cell.VerticalAlignment = Element.ALIGN_LEFT;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("Primaria Torres Lindavista\nAcuerdo No. 00911649\nde Fecha 9 de Mayo de 1991\na Nivel Primaria", f7));
                cell.VerticalAlignment = Element.ALIGN_LEFT;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("Preescolar\nAcuerdo No. 09060102\nde Fecha 17 de Marzo de 2006\na Nivel Preescolar", f7));
                cell.VerticalAlignment = Element.ALIGN_LEFT;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("Secundaria\nAcuerdo No. SECR-09100130\nde Fecha 25 de Agosto de 2010\na Nivel Secundaria", f7));
                cell.VerticalAlignment = Element.ALIGN_LEFT;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("CLIENTE", f7B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 2;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = blanco;
                cell.Colspan = 3;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("FACTURA", f7B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 2;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = blanco;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("Nombre: " + htDatosCfdi["nombreReceptor"].ToString() + "\nRFC: " + htDatosCfdi["rfcReceptor"].ToString() + "\nDireccion: " + htDatosCfdi["direccionReceptor1"].ToString() + "\n" + htDatosCfdi["direccionReceptor2"].ToString() + "\n" + htDatosCfdi["direccionReceptor3"].ToString() + "\n\n" + datosEstudiante, f7));
                cell.VerticalAlignment = Element.ALIGN_LEFT;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.Rowspan = 5;
                cell.Colspan = 3;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase(serieFolio, f7));
                cell.VerticalAlignment = Element.ALIGN_CENTER;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("Fecha y Hora", f7B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 2;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = blanco;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase(htDatosCfdi["fechaCfdi"].ToString(), f7));
                cell.VerticalAlignment = Element.ALIGN_CENTER;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase("UUID", f7B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 2;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = blanco;
                tEncabezadoDetail.AddCell(cell);

                cell = new Cell(new Phrase(htDatosCfdi["UUID"].ToString(), f7));
                cell.VerticalAlignment = Element.ALIGN_CENTER;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                tEncabezadoDetail.AddCell(cell);

                documento.Add(tEncabezadoDetail);

                //Tabla que contiene los encabezados del detalle de las partidas
                Table tDetail = new Table(5);
                tDetail.WidthPercentage = 95;
                tDetail.SetWidths(new int[5] { 13, 25, 25, 22, 20 });
                tDetail.DefaultCell.Border = 0;
                tDetail.BorderColor = blanco;
                tDetail.Padding = 2;
                tDetail.Spacing = 1;

                cell = new Cell(new Phrase("Cantidad", f7B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 2;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = blanco;
                tDetail.AddCell(cell);

                cell = new Cell(new Phrase("Descripción", f7B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 2;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = blanco;
                cell.Colspan = 2;
                tDetail.AddCell(cell);

                cell = new Cell(new Phrase("Valor Unitario", f7B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 2;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = blanco;
                tDetail.AddCell(cell);

                cell = new Cell(new Phrase("Importe", f7B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 2;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = blanco;
                tDetail.AddCell(cell);

                documento.Add(tDetail);

                #endregion

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

                #region "Añadimos Detalle de Conceptos"

                // Creamos la tabla para insertar los conceptos de detalle de la factura
                PdfPTable tableConceptos = new PdfPTable(4);

                int[] colWithsConceptos = new int[4];
                String[] arrColWidthConceptos = { "15", "45", "20", "20" };

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
                    cellConceptos.HorizontalAlignment = PdfCell.ALIGN_MIDDLE;
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
                }

                documento.Add(tableConceptos);

                #endregion

                #region "Espaciador entre Conceptos y Desglose"

                Paragraph pRelleno2 = new Paragraph();
                Chunk cRelleno2 = new Chunk("\n");
                pRelleno2.Add(cRelleno2);
                documento.Add(pRelleno2);

                #endregion

                #region "Desglose"

                PdfPTable tableDesglose = new PdfPTable(3);
                int[] colWithsDesglose = new int[3];
                String[] arrColWidthDesglose = dtConfigFact.Rows[0]["desgloseColWidth"].ToString().Split(new Char[] { ',' });

                for (int i = 0; i < arrColWidthDesglose.Length; i++)
                {
                    colWithsDesglose.SetValue(Convert.ToInt32(arrColWidthDesglose[i]), i);
                }

                tableDesglose.SetWidths(colWithsDesglose);
                tableDesglose.WidthPercentage = 93F;

                PdfPCell cellDesgloseRelleno = new PdfPCell();
                PdfPCell cellDesgloseDescripcion = new PdfPCell();
                PdfPCell cellDesgloseMonto = new PdfPCell();

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
                    alDesgloseOrden.Add("Impuestos Trasladados");
                }

                if (electronicDocument.Data.Impuestos.TotalRetenciones.Value != 0)
                {
                    alDesgloseOrden.Add("Impuestos Retenidos");
                }

                alDesgloseOrden.Add("Total");

                htDesglose.Add("Subtotal", electronicDocument.Data.SubTotal.Value.ToString("C", _ci));
                htDesglose.Add("Descuento", electronicDocument.Data.Descuento.Value.ToString("C", _ci));
                htDesglose.Add("Impuestos Trasladados", electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("C", _ci));
                htDesglose.Add("Impuestos Retenidos", electronicDocument.Data.Impuestos.TotalRetenciones.Value.ToString("C", _ci));
                htDesglose.Add("Total", electronicDocument.Data.Total.Value.ToString("C", _ci));

                foreach (string desglose in alDesgloseOrden)
                {
                    cellDesgloseRelleno = new PdfPCell(new Phrase(string.Empty, new Font(Font.HELVETICA, 7, Font.BOLD)));
                    cellDesgloseRelleno.Border = 0;

                    cellDesgloseDescripcion = new PdfPCell(new Phrase(desglose, new Font(Font.HELVETICA, 7, Font.BOLD)));
                    cellDesgloseDescripcion.Border = 0;

                    cellDesgloseMonto = new PdfPCell(new Phrase(htDesglose[desglose].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellDesgloseMonto.Border = 0;
                    cellDesgloseMonto.HorizontalAlignment = PdfCell.ALIGN_RIGHT;

                    tableDesglose.AddCell(cellDesgloseRelleno);
                    tableDesglose.AddCell(cellDesgloseDescripcion);
                    tableDesglose.AddCell(cellDesgloseMonto);
                }

                documento.Add(tableDesglose);

                #endregion

                #region "Desglose Detalle"

                PdfPTable tableDesgloseDetalle = new PdfPTable(3);
                int[] colWithsDesgloseDetalle = new int[3];
                string colDesgloseDetalle = "8,8,84";
                String[] arrColWidthDesgloseDetalle = colDesgloseDetalle.Split(new Char[] { ',' });

                for (int i = 0; i < arrColWidthDesgloseDetalle.Length; i++)
                {
                    colWithsDesgloseDetalle.SetValue(Convert.ToInt32(arrColWidthDesgloseDetalle[i]), i);
                }

                tableDesgloseDetalle.SetWidths(colWithsDesgloseDetalle);
                tableDesgloseDetalle.WidthPercentage = 100F;

                #endregion

                #region "Creación e Inserción de Chunks en Paragraph"

                Paragraph pFooter = new Paragraph();
                pFooter.KeepTogether = true;
                pFooter.Alignment = PdfCell.ALIGN_LEFT;
                pFooter.SetLeading(1.6f, 1.6f);

                if (electronicDocument.Data.TipoCambio.Value.Length == 0)
                    electronicDocument.Data.TipoCambio.Value = "1";

                Font fontLbl = new Font(Font.HELVETICA, 5, Font.BOLD, new Color(43, 145, 175));
                Font fontVal = new Font(Font.HELVETICA, 5, Font.NORMAL);

                Font fontLbl2 = new Font(Font.HELVETICA, 6, Font.BOLD, new Color(43, 145, 175));
                Font fontVal2 = new Font(Font.HELVETICA, 6, Font.NORMAL);

                Font fontVal3 = new Font(Font.HELVETICA, 5, Font.BOLD);

                string cantidadLetra = dtConfigFact.Rows[0]["cantidadLetra"].ToString();
                Chunk cCantidadLetraLbl = new Chunk("Cantidad :  ", fontLbl2);
                Chunk cCantidadLetraVal = new Chunk(cantidadLetra, new Font(Font.HELVETICA, 6, Font.BOLD));

                Chunk cTipoComprobanteLbl = new Chunk("Tipo de Comprobante: ", fontLbl);
                Chunk cTipoComprobanteVal = new Chunk(electronicDocument.Data.TipoComprobante.Value, fontVal);

                Chunk cFormaPagoLbl = new Chunk("Forma de Pago: ", fontLbl);
                Chunk cFormaPagoVal = new Chunk(electronicDocument.Data.FormaPago.Value, fontVal);

                Chunk cMetodoPagoLbl = new Chunk("Método de Pago: ", fontLbl);
                Chunk cMetodoPagoVal = new Chunk(electronicDocument.Data.MetodoPago.Value, fontVal);

                Chunk cMonedaLbl = new Chunk("Moneda: ", fontLbl);
                Chunk cMonedaVal = new Chunk(electronicDocument.Data.Moneda.Value, fontVal);

                Chunk cTasaCambioLbl = new Chunk("Tasa de Cambio: ", fontLbl);
                Chunk cTasaCambioVal = new Chunk(Convert.ToDouble(electronicDocument.Data.TipoCambio.Value).ToString("C", _ci), fontVal);

                Chunk cCertificadoLbl = new Chunk("Certificado del Emisor: ", fontLbl);
                Chunk cCertificadoVal = new Chunk(electronicDocument.Data.NumeroCertificado.Value, fontVal);

                Chunk cCadenaOriginalLbl = new Chunk("Cadena Original", fontLbl);
                Chunk cCadenaOriginalVal = new Chunk(electronicDocument.FingerPrint, fontVal);

                Chunk cCadenaOriginalPACLbl = new Chunk("Cadena Original del Complemento de Certificación Digital del SAT", fontLbl);
                Chunk cCadenaOriginalPACVal = new Chunk(electronicDocument.FingerPrintPac, fontVal);

                Chunk cSelloDigitalLbl = new Chunk("Sello Digital del Emisor", fontLbl);
                Chunk cSelloDigitalVal = new Chunk(electronicDocument.Data.Sello.Value, fontVal);

                Chunk cFolioFiscalLbl = new Chunk("Folio Fiscal: ", fontLbl);
                Chunk cFolioFiscalVal = new Chunk(objTimbre.Uuid.Value, fontVal);

                Chunk cFechaTimbradoLbl = new Chunk("Fecha y Hora de Certificación: ", fontLbl);
                DateTime fechaTimbrado = Convert.ToDateTime(objTimbre.FechaTimbrado.Value);

                string formatoFechaTimbrado = fechaTimbrado.ToString("yyyy-MM-dd") + "T" + fechaTimbrado.ToString("HH:mm:ss");
                Chunk cFechaTimbradoVal = new Chunk(formatoFechaTimbrado, fontVal);

                Chunk cCertificadoSatLbl = new Chunk("Certificado SAT: ", fontLbl);
                Chunk cCertificadoSatVal = new Chunk(objTimbre.NumeroCertificadoSat.Value, fontVal);

                Chunk cSelloDigitalSatLbl = new Chunk("Sello Digital SAT", fontLbl);
                Chunk cSelloDigitalSatVal = new Chunk(objTimbre.SelloSat.Value, fontVal);

                string regimenes = "";

                for (int u = 0; u < electronicDocument.Data.Emisor.Regimenes.Count; u++)
                    regimenes += electronicDocument.Data.Emisor.Regimenes[u].Regimen.Value.ToString() + ",";

                Chunk cNoTarjetaLbl = new Chunk("No. Tarjeta: ", fontLbl);
                Chunk cNoTarjetaVal = new Chunk(electronicDocument.Data.NumeroCuentaPago.Value, fontVal);

                Chunk cExpedidoEnLbl = new Chunk("Expedido En: ", fontLbl);
                Chunk cExpedidoEnVal = new Chunk(electronicDocument.Data.LugarExpedicion.Value, fontVal);

                #endregion

                #region "Añadimos Código Bidimensional y Pie de Pagina"

                pFooter.Add(cCantidadLetraLbl);
                pFooter.Add(cCantidadLetraVal);

                if (htFacturaxion["observaciones"].ToString().Length > 0)
                {
                    Chunk cObsLbl = new Chunk("Observaciones :  ", fontLbl2);
                    Chunk cObsVal = new Chunk(htFacturaxion["observaciones"].ToString(), fontVal2);
                    pFooter.Add(cSaltoLinea);
                    pFooter.Add(cObsLbl);
                    pFooter.Add(cObsVal);
                }

                documento.Add(pFooter);
                pFooter.Clear();

                if (timbrar)
                {
                    Image imageQRCode = Image.GetInstance(bytesQRCode);
                    imageQRCode.Alignment = (Image.TEXTWRAP | Image.ALIGN_LEFT);
                    imageQRCode.ScaleToFit(90f, 90f);
                    imageQRCode.IndentationLeft = 9f;
                    imageQRCode.SpacingAfter = 9f;
                    imageQRCode.BorderColorTop = Color.WHITE;
                    documento.Add(imageQRCode);
                }

                pFooter.Add(cSaltoLinea);
                pFooter.Add(cFolioFiscalLbl);
                pFooter.Add(cFolioFiscalVal);
                pFooter.Add(cSaltoLinea);
                pFooter.Add(cFechaTimbradoLbl);
                pFooter.Add(cFechaTimbradoVal);
                pFooter.Add(cSaltoLinea);
                pFooter.Add(cCertificadoLbl);
                pFooter.Add(cCertificadoVal);
                pFooter.Add(cSaltoLinea);

                pFooter.Add(cCertificadoSatLbl);
                pFooter.Add(cCertificadoSatVal);
                pFooter.Add(cSaltoLinea);

                pFooter.Add(cSelloDigitalLbl);
                pFooter.Add(cSaltoLinea);
                pFooter.Add(cSelloDigitalVal);
                pFooter.Add(cSaltoLinea);

                pFooter.Add(cSelloDigitalSatLbl);
                pFooter.Add(cSaltoLinea);
                pFooter.Add(cSelloDigitalSatVal);
                pFooter.Add(cSaltoLinea);

                if (timbrar)
                {
                    pFooter.Add(cCadenaOriginalPACLbl);
                    pFooter.Add(cSaltoLinea);
                    pFooter.Add(cCadenaOriginalPACVal);
                }

                documento.Add(pFooter);
                pFooter.Clear();

                pFooter.Add(cTipoComprobanteLbl);
                pFooter.Add(cTipoComprobanteVal);
                pFooter.Add(cDataSpacer);
                pFooter.Add(cMonedaLbl);
                pFooter.Add(cMonedaVal);
                pFooter.Add(cDataSpacer);
                pFooter.Add(cTasaCambioLbl);
                pFooter.Add(cTasaCambioVal);
                pFooter.Add(cDataSpacer);
                pFooter.Add(cFormaPagoLbl);
                pFooter.Add(cFormaPagoVal);
                pFooter.Add(cDataSpacer);
                pFooter.Add(cMetodoPagoLbl);
                pFooter.Add(cMetodoPagoVal);

                if (electronicDocument.Data.NumeroCuentaPago.Value.ToString().Length > 0)
                {
                    pFooter.Add(cDataSpacer);
                    pFooter.Add(cNoTarjetaLbl);
                    pFooter.Add(cNoTarjetaVal);
                }

                if (electronicDocument.Data.Emisor.Regimenes.Count > 0)
                {
                    Chunk cRegimenLbl = new Chunk("Régimen Fiscal: ", fontLbl);
                    Chunk cRegimenVal = new Chunk(regimenes.Substring(0, regimenes.Length - 1).ToString(), fontVal);

                    pFooter.Add(cDataSpacer);
                    pFooter.Add(cRegimenLbl);
                    pFooter.Add(cRegimenVal);
                }

                if (electronicDocument.Data.FolioFiscalOriginal.Value.ToString().Length > 0)
                {
                    Chunk cFolioOriginal1Lbl = new Chunk("Datos CFDI Original - Serie: ", fontLbl);
                    Chunk cFolioOriginal1Val = new Chunk(electronicDocument.Data.SerieFolioFiscalOriginal.Value + "   ", fontVal);

                    Chunk cFolioOriginal2Lbl = new Chunk("Folio: ", fontLbl);
                    Chunk cFolioOriginal2Val = new Chunk(electronicDocument.Data.FolioFiscalOriginal.Value + "   ", fontVal);

                    Chunk cFolioOriginal3Lbl = new Chunk("Fecha: ", fontLbl);
                    Chunk cFolioOriginal3Val = new Chunk(electronicDocument.Data.FechaFolioFiscalOriginal.Value.ToString() + "   ", fontVal);

                    Chunk cFolioOriginal4Lbl = new Chunk("Monto: ", fontLbl);
                    Chunk cFolioOriginal4Val = new Chunk(electronicDocument.Data.MontoFolioFiscalOriginal.Value.ToString(), fontVal);

                    pFooter.Add(cSaltoLinea);
                    pFooter.Add(cFolioOriginal1Lbl);
                    pFooter.Add(cFolioOriginal1Val);
                    pFooter.Add(cFolioOriginal2Lbl);
                    pFooter.Add(cFolioOriginal2Val);
                    pFooter.Add(cFolioOriginal3Lbl);
                    pFooter.Add(cFolioOriginal3Val);
                    pFooter.Add(cFolioOriginal4Lbl);
                    pFooter.Add(cFolioOriginal4Val);
                }

                pFooter.Add(cSaltoLinea);

                if (electronicDocument.Data.LugarExpedicion.Value.Length > 0)
                {
                    pFooter.Add(cExpedidoEnLbl);
                    pFooter.Add(cExpedidoEnVal);
                }

                pFooter.Add(cSaltoLinea);

                documento.Add(pFooter);

                #endregion

                #region "Añadimos leyenda de CFDI"

                Paragraph pLeyendaCfdi = new Paragraph();
                string leyenda;

                if (timbrar)
                {
                    leyenda = "                                                                                                        Este documento es una representación impresa de un CFDI";
                    Chunk cLeyendaCfdi = new Chunk(leyenda, new Font(Font.HELVETICA, 6, Font.BOLD | Font.ITALIC));
                    pLeyendaCfdi.Add(cLeyendaCfdi);
                    pLeyendaCfdi.SetLeading(1.6f, 1.6f);
                    documento.Add(pLeyendaCfdi);
                }

                #endregion

                #region "Footer"

                //Agregando Imagen de Pie de Página
                Image imgFooter = Image.GetInstance(@"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\OVM800704940\footer.png");
                float imgFooterWidth = documento.PageSize.Width - 70;
                float imgFooterHeight = imgFooter.Height / (imgFooter.Width / imgFooterWidth);

                imgFooter.ScaleAbsolute(imgFooterWidth, imgFooterHeight);
                imgFooter.SetAbsolutePosition(Convert.ToSingle(dtConfigFact.Rows[0]["footerPosX"]), Convert.ToSingle(dtConfigFact.Rows[0]["footerPosY"]));
                documento.Add(imgFooter);

                #endregion


                #endregion

                documento.Close();
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
    }
}