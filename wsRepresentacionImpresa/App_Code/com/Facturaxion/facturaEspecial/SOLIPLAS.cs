
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
    public class SOLIPLAS
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
        public static CultureInfo _c2 = new CultureInfo("es-mx");
      
        private static readonly string _rutaDocs = ConfigurationManager.AppSettings["rutaDocs"];
        private static readonly string _rutaDocsExt = ConfigurationManager.AppSettings["rutaDocsExterna"];

        private static bool timbrar;
        private static string pathPdf;

        private static HttpContext HTC;
        private static String pathIMGLOGO;
        private static String pathIMGFX;        
        private static Chunk dSaltoLinea;
        private static ElectronicDocument electronicDocument;
        private static Data objTimbre;

        private static Color azul;
        private static Color azul1;
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
                StringBuilder sbConfigFactParms = new StringBuilder();
                StringBuilder sbConfigFact = new StringBuilder();
                StringBuilder sbDataEmisor = new StringBuilder();
                StringBuilder sbOpcionalEncabezado = new StringBuilder();
                DataTable dtConfigFact = new DataTable();
                DataTable dtDataEmisor = new DataTable();
                DataTable dtOpcEncabezado = new DataTable();
                _ci.NumberFormat.CurrencyDecimalDigits = 4;
                _c2.NumberFormat.CurrencyDecimalDigits = 2;
                

                electronicDocument = (ElectronicDocument)htFacturaxion["electronicDocument"];
                objTimbre = (Data)htFacturaxion["objTimbre"];
                timbrar = Convert.ToBoolean(htFacturaxion["timbrar"]);
                //pathPdf = htFacturaxion["rutaDocumentoPdf"].ToString();
                Int64 idCfdi = Convert.ToInt64(htFacturaxion["idCfdi"]);      

                #region "Extraemos los datos del CFDI"


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

                sbDataEmisor.
                    Append("SELECT nombreSucursal FROM sucursales WHERE idSucursal = @0 ");

                sbOpcionalEncabezado.
                    Append("SELECT campo1 AS observaciones ").
                    //Append("campo2 AS observaciones2").
                    //Append("campo3 AS observaciones3").
                    Append("FROM opcionalEncabezado WHERE idCFDI = @0 and ST = 1");

                dtConfigFact = dal.QueryDT("DS_FE", sbConfigFact.ToString(), sbConfigFactParms.ToString(), hc);
                dtDataEmisor = dal.QueryDT("DS_FE", sbDataEmisor.ToString(), "F:I:" + htFacturaxion["idSucursalEmisor"].ToString(), hc);
                dtOpcEncabezado = dal.QueryDT("DS_FE", sbOpcionalEncabezado.ToString(), "F:I:" + idCfdi, hc);

                Hashtable htDatosCfdi = new Hashtable();

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

                htDatosCfdi.Add("nombreEmisor", "Razón Social " + electronicDocument.Data.Emisor.Nombre.Value);
                htDatosCfdi.Add("empresa", "Razón Social " + electronicDocument.Data.Emisor.Nombre.Value);

                htDatosCfdi.Add("rfcReceptor", electronicDocument.Data.Receptor.Rfc.Value);
                htDatosCfdi.Add("rfcCliente", electronicDocument.Data.Receptor.Rfc.Value);

                htDatosCfdi.Add("nombreReceptor", "Razón Social " + electronicDocument.Data.Receptor.Nombre.Value);
                htDatosCfdi.Add("cliente", "Razón Social " + electronicDocument.Data.Receptor.Nombre.Value);

                htDatosCfdi.Add("sucursal", "Sucursal " + dtDataEmisor.Rows[0]["nombreSucursal"]);

                htDatosCfdi.Add("serie", electronicDocument.Data.Serie.Value);
                htDatosCfdi.Add("folio", electronicDocument.Data.Folio.Value);

                htDatosCfdi.Add("fechaCfdi", electronicDocument.Data.Fecha.Value);
                htDatosCfdi.Add("fechaFactura", electronicDocument.Data.Fecha.Value);

                htDatosCfdi.Add("UUID", objTimbre.Uuid.Value.ToUpper());
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

                #endregion

                #region "Creamos el Objeto Documento y Tipos de Letra"

                Document document = new Document(PageSize.LETTER, 15, 15, 15, 40);
                document.AddAuthor("Facturaxion");
                document.AddCreator("r3Take");
                document.AddCreationDate();

               //pdfPageEventHandlerProbell pageEventHandler = new pdfPageEventHandlerProbell();
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                //writer.PageEvent = pageEventHandler;
                PdfPageEventHandlerSoliplas sol = new PdfPageEventHandlerSoliplas();
                writer.PageEvent = sol;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                azul = new Color(22, 111, 168);
                azul1 = new Color(43, 145, 175);
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
                f6L = new Font(EM, 6, Font.BOLD, Link);
                f5L = new Font(EM, 5, Font.BOLD, lbAzul);
                titulo = new Font(EM, 6, Font.BOLD, blanco);
                folio = new Font(EM, 6, Font.BOLD, rojo);
                PdfPCell cell;
                Paragraph par;
                Cell cel;
                dSaltoLinea = new Chunk("\n\n ");

                #endregion

                PdfPTable encabezado = new PdfPTable(3);
                encabezado.WidthPercentage = 100;
                encabezado.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                encabezado.SetWidths(new int[3] { 30, 10, 60 });
                encabezado.DefaultCell.Border = 0;
                encabezado.LockedWidth = true;

                pathIMGLOGO = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\SOLIPLAS\logoSoliplas.png";
                pathIMGFX = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\SOLIPLAS\cfdifx.png";
                
                Image imgLogo = Image.GetInstance(pathIMGLOGO);
                imgLogo.ScalePercent(48f);
                Image imgFx = Image.GetInstance(pathIMGFX);
                imgFx.ScalePercent(25f);
                 
                      
                #region "Construimos el encabezado y Detalles del documento"
                //Encabezado Folio Fiscal
                Table encabezadoFolio = new Table(3);
                float[] headerEncabezadoFolio = { 60, 10, 30 };
                encabezadoFolio.Widths = headerEncabezadoFolio;
                encabezadoFolio.WidthPercentage = 100F;
                encabezadoFolio.Padding = 1;
                encabezadoFolio.Spacing = 1;
                encabezadoFolio.BorderWidth = 0;
                encabezadoFolio.DefaultCellBorder = 0;
                encabezadoFolio.BorderColor = gris;

                cel = new Cell(imgFx);
                cel.HorizontalAlignment = Element.ALIGN_LEFT;                
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                encabezadoFolio.AddCell(cel);

                cel = new Cell(new Phrase("Folio Fiscal", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoFolio.AddCell(cel);

                cel = new Cell(new Phrase(htDatosCfdi["folioFiscal"].ToString().ToUpper(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = azul;
                encabezadoFolio.AddCell(cel);

                //Encabezado Comprobante
                Table encabezadoComprobante = new Table(6);
                float[] headerEncabezadoComprobante = { 30,40,5,5,10,10 };
                encabezadoComprobante.Widths = headerEncabezadoComprobante;
                encabezadoComprobante.WidthPercentage = 100F;
                encabezadoComprobante.Padding = 1;
                encabezadoComprobante.Spacing = 1;
                encabezadoComprobante.BorderWidth = 0;
                encabezadoComprobante.DefaultCellBorder = 0;
                encabezadoComprobante.BorderColor = gris;

                //LOGO DELA EMPRESA
                cel = new Cell(imgLogo);
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 4;
                //cel.Colspan = 2;
                encabezadoComprobante.AddCell(cel);       
                
                //EMISOR
                StringBuilder emisor = new StringBuilder();
                emisor.                    
                    Append("RFC: ").
                    Append(htDatosCfdi["rfcEmisor"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["nombreEmisor"]).Append("\n").
                    Append(htDatosCfdi["direccionEmisor1"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["direccionEmisor2"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["direccionEmisor3"].ToString().ToUpper()).Append("\n");

                cel = new Cell(new Phrase(emisor.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;                
                cel.Rowspan = 4;
                cel.Colspan = 3;
                encabezadoComprobante.AddCell(cel);                
                

                // Serie         
                cel = new Cell(new Phrase("Serie", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                //cel.Colspan = 3;
                encabezadoComprobante.AddCell(cel);       
                
                // Folio
                cel = new Cell(new Phrase("Folio", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                encabezadoComprobante.AddCell(cel);

                cel = new Cell(new Phrase(htDatosCfdi["serie"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezadoComprobante.AddCell(cel); 

                cel = new Cell(new Phrase(htDatosCfdi["folio"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezadoComprobante.AddCell(cel);

                // Fecha de emisión del comprobante
                cel = new Cell(new Phrase("Fecha", titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                encabezadoComprobante.AddCell(cel);

                cel = new Cell(new Phrase(htDatosCfdi["fechaFactura"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                encabezadoComprobante.AddCell(cel);


                Table encabezadoComprobante1 = new Table(2);
                float[] headerEncabezadoComprobante1 = { 50, 50 };
                encabezadoComprobante1.Widths = headerEncabezadoComprobante1;
                encabezadoComprobante1.WidthPercentage = 100F;
                encabezadoComprobante1.Padding = 1;
                encabezadoComprobante1.Spacing = 1;
                encabezadoComprobante1.BorderWidth = 0;
                encabezadoComprobante1.DefaultCellBorder = 0;
                encabezadoComprobante1.BorderColor = gris;
                
                //Datos discales del cliente
                cel = new Cell(new Phrase("Datos Fiscales del Cliente", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoComprobante1.AddCell(cel);

                //Expedido en:
                cel = new Cell(new Phrase("Expedido en:", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoComprobante1.AddCell(cel);

                //Receptor
                StringBuilder cliente = new StringBuilder();
                cliente.
                    Append("\nRFC: ").
                    Append(htDatosCfdi["rfcCliente"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["cliente"]).Append("\n").
                    Append(htDatosCfdi["direccionCliente1"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["direccionCliente2"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["direccionCliente3"].ToString().ToUpper()).Append("\n").
                    Append("\n");

                cel = new Cell(new Phrase(cliente.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 5;
                encabezadoComprobante1.AddCell(cel);
                
                
                //Expedido en:
                StringBuilder expedido = new StringBuilder();
                expedido.
                    //Append(htDatosCfdi["sucursal"]).Append("\n").
                    Append(htDatosCfdi["direccionExpedido1"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["direccionExpedido2"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["direccionExpedido3"].ToString().ToUpper()).Append("\n").
                    Append("\n");

                cel = new Cell(new Phrase(expedido.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.UseBorderPadding = true;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;                
                cel.Rowspan = 5;
                encabezadoComprobante1.AddCell(cel);

                #endregion
                
                sol.encabezado = encabezado;
                sol.dSaltoLinea = dSaltoLinea;                

                #region "Tabla Detalle"

                Table encabezadoDetalle = new Table(6);
                float[] headerEncabezadoDetalle = { 7, 8, 11, 52, 13, 13 };
                encabezadoDetalle.Widths = headerEncabezadoDetalle;
                encabezadoDetalle.WidthPercentage = 100F;
                encabezadoDetalle.Padding = 1;
                encabezadoDetalle.Spacing = 1;
                encabezadoDetalle.BorderWidth = 0;
                encabezadoDetalle.DefaultCellBorder = 0;
                encabezadoDetalle.BorderColor = gris;

                // Número
                cel = new Cell(new Phrase("Cantidad", titulo));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Còdigo de Barras
                cel = new Cell(new Phrase("Unidad", titulo));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Código
                cel = new Cell(new Phrase("Código", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Descripción
                cel = new Cell(new Phrase("Descripción", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);
                
                // Precio Unitario
                cel = new Cell(new Phrase("Precio Unitario", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Importe
                cel = new Cell(new Phrase("Importe", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                encabezadoDetalle.AddCell(cel);

                // Creamos la tabla para insertar los conceptos de detalle de la factura
                PdfPTable tableConceptos = new PdfPTable(6);

                int[] colWithsConceptos = new int[6];
                //String[] arrColWidthConceptos = dtConfigFact.Rows[0]["conceptosColWidth"].ToString().Split(new Char[] { ',' });
                String[] arrColWidthConceptos = { "12", "8", "15", "51", "13", "17" };

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
                    cellConceptos = new PdfPCell(new Phrase("         "+ electronicDocument.Data.Conceptos[i].Cantidad.Value.ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellConceptos.Border = 0;
                    cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                    tableConceptos.AddCell(cellConceptos);

                    cellConceptos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Unidad.Value, new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellConceptos.Border = 0;
                    cellConceptos.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                    tableConceptos.AddCell(cellConceptos);

                    cellConceptos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].NumeroIdentificacion.Value, new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellConceptos.Border = 0;
                    cellConceptos.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    tableConceptos.AddCell(cellConceptos);

                    cellConceptos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellConceptos.Border = 0;
                    tableConceptos.AddCell(cellConceptos);

                    cellMontos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].ValorUnitario.Value.ToString("C", _ci), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellMontos.Border = 0;
                    cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                    tableConceptos.AddCell(cellMontos);

                    cellMontos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("C", _ci), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                    cellMontos.Border = 0;
                    cellMontos.HorizontalAlignment = PdfCell.ALIGN_RIGHT;
                    tableConceptos.AddCell(cellMontos);
                }
                dSaltoLinea = new Chunk("\n\n ");

                #endregion

                #region "Construimos el Comentarios"

                Table comentarios = new Table(7);
                float[] headerwidthsComentarios = { 18, 18, 28, 28, 7, 5, 5 };
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

                cel = new Cell(new Phrase(electronicDocument.Data.SubTotal.Value.ToString("C", _c2), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["observaciones"].ToString(), f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                cel.Rowspan = 3;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(dtConfigFact.Rows[0]["cantidadLetra"].ToString(), f5));//CE26           
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                cel.Rowspan = 3;
                comentarios.AddCell(cel);


                cel = new Cell(new Phrase("Descuento",f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Descuento.Value.ToString("C", _c2), f5));
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
                                
                cel = new Cell(new Phrase("IVA " + tasa.ToString() + " %", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("C", _c2), f5));
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

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("C", _c2), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                #endregion

                #region "Construimos Tabla Datos CFDi"

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

                    cel = new Cell(new Phrase(objTimbre.Uuid.Value.ToUpper(), f5));
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


                    if (electronicDocument.Data.Emisor.Regimenes.Count > 0)
                    {
                        for (int u = 0; u < electronicDocument.Data.Emisor.Regimenes.Count; u++)
                            regimenes += electronicDocument.Data.Emisor.Regimenes[u].Regimen.Value.ToString() + ",";

                        par.Add(new Chunk("   |   ", f5));
                        par.Add(new Chunk("RÉGIMEN FISCAL: ", f5L));
                        par.Add(new Chunk(regimenes.Substring(0, regimenes.Length - 1).ToString(), f5));
                        par.Add(new Chunk("   |   ", f5));
                    }

                    if (electronicDocument.Data.CondicionesPago.Value.ToString().Length > 0)
                    {
                        par.Add(new Chunk("   |   ", f5));
                        par.Add(new Chunk("CONDICIONES DE PAGO: ", f5L));
                        par.Add(new Chunk(electronicDocument.Data.CondicionesPago.Value.ToString(), f5));
                        par.Add(new Chunk("   |   ", f5));
                    }

                    if (electronicDocument.Data.NumeroCuentaPago.Value.ToString().Length > 0)
                    {
                        par.Add(new Chunk("   |   ", f5));
                        par.Add(new Chunk("No. CUENTA: ", f5L));
                        par.Add(new Chunk(electronicDocument.Data.NumeroCuentaPago.Value, f5));
                        par.Add(new Chunk("   |   ", f5));
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
                        par.Add(new Chunk("LUGAR EXPEDICIÓN: ", f5L));
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

                sol.encFolioFiscal = encabezadoFolio;
                sol.encComprobante = encabezadoComprobante;
                sol.encComprobante1 = encabezadoComprobante1;
                sol.encTitulos = encabezadoDetalle;                
                sol.adicional = adicional;
                sol.footer = footer;

                document.Open();  

                document.Add(tableConceptos);
                document.Add(comentarios);
                document.Add(adicional);
                

                string filePdfExt = pathPdf.Replace(_rutaDocs, _rutaDocsExt);
                string urlPathFilePdf = filePdfExt.Replace(@"\", "/");
                document.Close();
                writer.Close();
                fs.Close();

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

    public class PdfPageEventHandlerSoliplas: PdfPageEventHelper
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

        public string rutaImgFooter { get; set; }
        public PdfPTable footer { get; set; }
        public Table encTitulos { get; set; }
        public Table encFolioFiscal { get; set; }
        public Table encComprobante { get; set; }
        public Table encComprobante1 { get; set; }
        public PdfPTable encabezado { get; set; }
        public Chunk dSaltoLinea { get; set; }
        public Table adicional { get; set; }

        #endregion

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

            /*base.OnStartPage(writer, document);
            Chunk ch = new Chunk("This is my Stack Overflow Header on page " + writer.PageNumber);
            document.Add(ch);*/

            base.OnStartPage(writer, document);
            document.Add(encabezado);
            document.Add(dSaltoLinea);
            document.Add(encFolioFiscal);
            document.Add(encComprobante);
            document.Add(encComprobante1);
            document.Add(encTitulos);

            cb.EndText();

        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);
            footer.WriteSelectedRows(0, -1, 15, (document.BottomMargin - 0), writer.DirectContent);

            string lblPagina;
            string lblDe;
            string lblFechaImpresion;

            lblPagina = "Página ";
            lblDe = " de ";
            lblFechaImpresion = "Fecha de Impresión ";

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
    }

    public class DefaultSplitCharacterSoliplas : ISplitCharacter
    {
        #region "ISplitCharacter"

        /**
         * An instance of the default SplitCharacter.
         */
        public static readonly ISplitCharacter DEFAULT = new DefaultSplitCharacterSoliplas();

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