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
    public class opamComprobante
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

        private static String pathIMGLOGO;
        private static String pathIMGRFC;
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
        private static Color negro;

        private static BaseFont EM;
        private static Font f5;
        private static Font f5B;
        private static Font f5BBI;
        private static Font f6;
        private static Font f6B;
        private static Font f6L;
        private static Font titulo;
        private static Font titulo1;
        private static Font folio;
        private static Font f5L;

        #endregion

        #region generarPdf

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
                StringBuilder sbRegimenFiscal = new StringBuilder();
                StringBuilder sbAnexoEncabezado = new StringBuilder();
                StringBuilder sbAnexoDetalle = new StringBuilder();
                DataTable dtOpcEncabezado = new DataTable();
                DataTable dtConfigFact = new DataTable();
                DataTable dtDataEmisor = new DataTable();
                DataTable dtRegimenFiscal = new DataTable();
                DataTable dtAnexoEncabezado = new DataTable();
                DataTable dtAnexoDetalle = new DataTable();
                _ci.NumberFormat.CurrencyDecimalDigits = 3;
                _c2.NumberFormat.CurrencyDecimalDigits = 2;

                electronicDocument = (ElectronicDocument)htFacturaxion["electronicDocument"];
                objTimbre = (Data)htFacturaxion["objTimbre"];
                timbrar = Convert.ToBoolean(htFacturaxion["timbrar"]);
                pathPdf = htFacturaxion["rutaDocumentoPdf"].ToString();
                Int64 idCfdi = Convert.ToInt64(htFacturaxion["idCfdi"]);
                //Int64 idEmpresa = Convert.ToInt64(htFacturaxion["idEmpresa"]);

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

                sbRegimenFiscal.
                    Append("SELECT regimen FROM regimenesXEmpresa WHERE idEmpresa = @0 and ST=1");

                sbOpcionalEncabezado.
                    Append("SELECT campo1 AS noCliente, ").
                    Append("campo2 AS pedido, ").
                    Append("campo3 AS fechaCompra, ").
                    Append("campo4 AS fechaEntrega, ").
                    Append("campo5 AS cantidad, ").
                    Append("campo6 AS descripcion, ").
                    Append("campo7 AS dispersion, ").
                    Append("campo8 AS descuento, ").
                    Append("campo9 AS cantidadDescuento, ").
                    Append("campo10 AS comisionConLetra, ").
                    Append("campo11 AS granTotal, ").
                    Append("campo12 AS granTotalConLetra, ").
                    Append("campo13 AS telefono, ").
                    Append("campo14 AS fax, ").
                    Append("campo15 AS ejecutivo, ").
                    Append("campo16 AS contacto, ").
                    Append("campo17 AS observaciones, ").
                    Append("campo18 AS tipoDocumento, ").
                    Append("campo19 AS cuentaBanamex, ").
                    Append("campo20 AS referenciaBanamex, ").
                    Append("campo21 AS cuentaBanorte, ").
                    Append("campo22 AS referenciaBanorte, ").
                    Append("campo23 AS cuentaBancomer, ").
                    Append("campo24 AS referenciaBancomer, ").
                    Append("campo25 AS cuentaHSBC, ").
                    Append("campo26 AS referenciaHSBC, ").
                    Append("campo27 AS cuentaSantander, ").
                    Append("campo28 AS referenciaSantander, ").
                    Append("campo43 AS calle, ").
                    Append("campo44 AS noExterior, ").
                    Append("campo45 AS noInterior, ").
                    Append("campo46 AS colonia, ").
                    Append("campo47 AS municipio, ").
                    Append("campo48 AS ciudad, ").
                    Append("campo49 AS codPostal, ").
                    Append("campo50 AS estado ").
                    Append("FROM opcionalEncabezado WHERE idCFDI = @0 and ST = 1");

                sbAnexoEncabezado.
                    Append("SELECT campo2 AS fechaEmision, ").
                    Append("campo3 AS numeroCliente, ").
                    Append("campo4 AS razonSocial, ").
                    Append("campo5 AS folioFactura ").
                    Append("FROM opcionalDetalle2 WHERE idCFDI = @0 and concepto = 1 and ST = 1");

                sbAnexoDetalle.
                    Append("SELECT campo2 AS tipoTarjeta, ").
                    Append("campo3 AS numeroEmpleado, ").
                    Append("campo4 AS nombreEmpleado, ").
                    Append("campo5 AS numeroProducto, ").
                    Append("campo6 AS producto, ").
                    Append("campo7 AS fechaAlta, ").
                    Append("campo8 AS pedido ").
                    Append("FROM opcionalDetalle2 WHERE idCFDI = @0 and concepto >= 2 and ST = 1");

                dtConfigFact = dal.QueryDT("DS_FE", sbConfigFact.ToString(), sbConfigFactParms.ToString(), hc);
                dtDataEmisor = dal.QueryDT("DS_FE", sbDataEmisor.ToString(), "F:I:" + htFacturaxion["idSucursalEmisor"].ToString(), hc);
                dtOpcEncabezado = dal.QueryDT("DS_FE", sbOpcionalEncabezado.ToString(), "F:I:" + idCfdi, hc);
                dtAnexoEncabezado = dal.QueryDT("DS_FE", sbAnexoEncabezado.ToString(), "F:I:" + idCfdi, hc);
                dtAnexoDetalle = dal.QueryDT("DS_FE", sbAnexoDetalle.ToString(), "F:I:" + idCfdi, hc);
                //dtRegimenFiscal = dal.QueryDT("DS_FE", sbRegimenFiscal.ToString(), "F:I:" + idEmpresa, hc);

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
                    sbDirExpedido3.Append("Municipio. / Del. ").Append(electronicDocument.Data.Emisor.ExpedidoEn.Municipio.Value).Append(" ");
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
                    sbDirReceptor1.Append(electronicDocument.Data.Receptor.Domicilio.Calle.Value).Append(" ");
                }

                if (electronicDocument.Data.Receptor.Domicilio.NumeroExterior.Value.Length > 0)
                {
                    sbDirReceptor1.Append(", ").Append(electronicDocument.Data.Receptor.Domicilio.NumeroExterior.Value).Append(" ");
                }

                if (electronicDocument.Data.Receptor.Domicilio.NumeroInterior.Value.Length > 0)
                {
                    sbDirReceptor1.Append(", ").Append(electronicDocument.Data.Receptor.Domicilio.NumeroInterior.Value);
                }

                if (electronicDocument.Data.Receptor.Domicilio.Colonia.Value.Length > 0)
                {
                    sbDirReceptor2.Append(", ").Append(electronicDocument.Data.Receptor.Domicilio.Colonia.Value).Append(" ");
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
                    sbDirReceptor3.Append(", ").Append(electronicDocument.Data.Receptor.Domicilio.Municipio.Value).Append(" ");
                }

                if (electronicDocument.Data.Receptor.Domicilio.Estado.Value.Length > 0)
                {
                    sbDirReceptor3.Append(",  ").Append(electronicDocument.Data.Receptor.Domicilio.Estado.Value).Append(" ");
                }

                sbDirReceptor3.Append(", ").Append(electronicDocument.Data.Receptor.Domicilio.Pais.Value);

                #endregion

                htDatosCfdi.Add("rfcEmisor", electronicDocument.Data.Emisor.Rfc.Value);
                htDatosCfdi.Add("rfcEmpresa", electronicDocument.Data.Emisor.Rfc.Value);


                htDatosCfdi.Add("nombreEmisor", "Razón Social " + electronicDocument.Data.Emisor.Nombre.Value);
                htDatosCfdi.Add("empresa", "Razón Social " + electronicDocument.Data.Emisor.Nombre.Value);

                htDatosCfdi.Add("rfcReceptor", electronicDocument.Data.Receptor.Rfc.Value);
                htDatosCfdi.Add("rfcCliente", electronicDocument.Data.Receptor.Rfc.Value);

                htDatosCfdi.Add("nombreReceptor", electronicDocument.Data.Receptor.Nombre.Value);
                htDatosCfdi.Add("cliente", electronicDocument.Data.Receptor.Nombre.Value);

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

                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                pdfPageEventHandlerOPAM ev = new pdfPageEventHandlerOPAM();
                writer.PageEvent = ev;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                azul = new Color(22, 111, 168);
                azul1 = new Color(43, 145, 175);
                blanco = new Color(255, 255, 255);
                Link = new Color(7, 73, 208);
                gris = new Color(236, 236, 236);
                grisOX = new Color(220, 215, 220);
                rojo = new Color(230, 7, 7);
                lbAzul = new Color(43, 145, 175);
                negro = new Color(0, 0, 0);

                EM = BaseFont.CreateFont(@"C:\Windows\Fonts\VERDANA.TTF", BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
                f5 = new Font(EM, 5);
                f5B = new Font(EM, 5, Font.BOLD);
                f5BBI = new Font(EM, 5, Font.BOLDITALIC);
                f6 = new Font(EM, 6);
                f6B = new Font(EM, 6, Font.BOLD);
                f6L = new Font(EM, 6, Font.BOLD, Link);
                f5L = new Font(EM, 5, Font.BOLD, azul);
                titulo = new Font(EM, 6, Font.BOLD, blanco);
                titulo1 = new Font(EM, 6, Font.BOLD, negro);
                folio = new Font(EM, 6, Font.BOLD, rojo);
                PdfPCell cell;
                Paragraph par;
                Cell cel;
                dSaltoLinea = new Chunk("\n\n ");

                #endregion

                #region "Generamos el Documento"

                switch (htDatosCfdi["serie"].ToString())
                {
                    case "D":
                        htDatosCfdi.Add("tipoDoc", "FACTURA");
                        break;
                    case "NC":
                        htDatosCfdi.Add("tipoDoc", "Nota de Credito");
                        break;
                    case "ND":
                        htDatosCfdi.Add("tipoDoc", "Nota de Cargo");
                        break;
                }

                #endregion

                #region "Creamos encabezado Del Documento"

                pathIMGLOGO = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\OPAM\logoOPAM.png";
                pathIMGRFC = @"C:\inetpub\RepositorioFacturaxion\imagesFacturaEspecial\OPAM\rfcOPAM.png";

                Image imgLogo = Image.GetInstance(pathIMGLOGO);
                imgLogo.ScalePercent(25f);
                Image imgRFC = Image.GetInstance(pathIMGRFC);
                imgRFC.ScalePercent(30f);


                Table encabezadoLogo = new Table(6);
                float[] headerEncabezadoFolio = { 40, 15, 15, 10, 10, 10 };
                encabezadoLogo.Widths = headerEncabezadoFolio;
                encabezadoLogo.WidthPercentage = 100F;
                encabezadoLogo.Padding = 1;
                encabezadoLogo.Spacing = 1;
                encabezadoLogo.BorderWidth = 0;
                encabezadoLogo.DefaultCellBorder = 0;
                encabezadoLogo.BorderColor = gris;

                cel = new Cell(imgLogo);
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                encabezadoLogo.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["tipoDocumento"].ToString().ToUpper(), new Font(Font.BOLD, 14, Font.NORMAL)));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = blanco;
                encabezadoLogo.AddCell(cel);

                cel = new Cell(new Phrase("Folio Fiscal", titulo1));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = blanco;
                cel.BackgroundColor = blanco;
                encabezadoLogo.AddCell(cel);

                cel = new Cell(new Phrase(htDatosCfdi["folioFiscal"].ToString().ToUpper(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 1;
                cel.BorderWidthLeft = 1;
                cel.BorderWidthRight = 1;
                cel.BorderWidthBottom = 1;
                cel.BorderColor = blanco;
                cel.Colspan = 3;
                encabezadoLogo.AddCell(cel);

                //Datos del Emisor
                StringBuilder emisor = new StringBuilder();
                emisor.
                    Append("RFC: ").
                    Append(htDatosCfdi["rfcEmisor"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["nombreEmisor"]).Append("\n").
                    Append(htDatosCfdi["direccionEmisor1"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["direccionEmisor2"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["direccionEmisor3"].ToString().ToUpper()).Append("\n").
                    Append("Tel: 50 91 51 00 R.F.C. OPA010719SF0 www.opam.com.mx");

                //DAtos DEl Emisor
                cel = new Cell(new Phrase(emisor.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 4;
                cel.Colspan = 4;
                encabezadoLogo.AddCell(cel);


                // Serie         
                cel = new Cell(new Phrase("Serie", titulo1));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                //cel.Colspan = 3;
                encabezadoLogo.AddCell(cel);

                // Folio
                cel = new Cell(new Phrase("Folio", titulo1));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezadoLogo.AddCell(cel);

                cel = new Cell(new Phrase(htDatosCfdi["serie"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezadoLogo.AddCell(cel);

                cel = new Cell(new Phrase(htDatosCfdi["folio"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezadoLogo.AddCell(cel);

                // Fecha de emisión del comprobante
                cel = new Cell(new Phrase("Fecha", titulo1));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                encabezadoLogo.AddCell(cel);

                cel = new Cell(new Phrase(htDatosCfdi["fechaFactura"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                encabezadoLogo.AddCell(cel);

                #endregion

                #region "Datos del Receptor y cuerpo del comprobante"

                Table datosReceptor = new Table(4);
                float[] headerDatosReceptor = { 25, 25, 25, 25 };
                datosReceptor.Widths = headerDatosReceptor;
                datosReceptor.WidthPercentage = 100F;
                datosReceptor.Padding = 1;
                datosReceptor.Spacing = 1;
                datosReceptor.BorderWidth = 0;
                datosReceptor.DefaultCellBorder = 0;
                datosReceptor.BorderColor = gris;

                cel = new Cell(new Phrase("Razón Social", f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase(htDatosCfdi["cliente"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Dirección: ", f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);

                StringBuilder direccionReceptor = new StringBuilder();
                direccionReceptor.
                    Append("Calle, No. Exterior e Interior, Colonia, Delegación/Municipio, Ciudad, Código Postal, Estado, País").Append("\n").
                    Append(htDatosCfdi["direccionCliente1"].ToString().ToUpper()).
                    Append(htDatosCfdi["direccionCliente2"].ToString().ToUpper()).
                    Append(htDatosCfdi["direccionCliente3"].ToString().ToUpper());

                cel = new Cell(new Phrase(direccionReceptor.ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase("R.F.C: " + htDatosCfdi["rfcCliente"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Teléfonos: " + dtOpcEncabezado.Rows[0]["telefono"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Fáx: " + dtOpcEncabezado.Rows[0]["fax"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase("No. de Cliente: ", f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase("No. de Pedido: ", f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Fecha de Compra: ", f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Fecha de Entrega: ", f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["noCliente"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["pedido"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);


                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["fechaCompra"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["fechaEntrega"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                datosReceptor.AddCell(cel);
                ev.dSaltoLinea = dSaltoLinea;

                cel = new Cell(new Phrase("Contacto: " + dtOpcEncabezado.Rows[0]["contacto"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Nombre de Ejecutico de Ventas: " + dtOpcEncabezado.Rows[0]["ejecutivo"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                datosReceptor.AddCell(cel);

                cel = new Cell(new Phrase("Lugar de Expedición (Calle, No.Exterior e Interior, Colonia, Delegación/Municipio, Ciudad, Código Postal, Estado): " + "\n", f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                datosReceptor.AddCell(cel);

                /*StringBuilder lugarExpedicion = new StringBuilder();
                lugarExpedicion.
                    Append(dtOpcEncabezado.Rows[0]["calle"].ToString() + ", ").
                    Append(dtOpcEncabezado.Rows[0]["noExterior"].ToString() + ", ").
                    Append(dtOpcEncabezado.Rows[0]["noInterior"].ToString() + ", ").
                    Append(dtOpcEncabezado.Rows[0]["colonia"].ToString() + ", ").
                    Append(dtOpcEncabezado.Rows[0]["municipio"].ToString() + ", ").
                    Append(dtOpcEncabezado.Rows[0]["ciudad"].ToString() + ", ").
                    Append(dtOpcEncabezado.Rows[0]["codPostal"].ToString() + ", ").
                    Append(dtOpcEncabezado.Rows[0]["estado"].ToString());*/

                StringBuilder expedido = new StringBuilder();
                expedido.
                    //Append(htDatosCfdi["sucursal"]).Append("\n").                    
                    Append(electronicDocument.Data.Emisor.ExpedidoEn.Calle.Value).Append(", ").
                    Append(electronicDocument.Data.Emisor.ExpedidoEn.NumeroExterior.Value).Append(" ").
                    Append(electronicDocument.Data.Emisor.ExpedidoEn.NumeroInterior.Value).Append(", ").
                    Append(electronicDocument.Data.Emisor.ExpedidoEn.Colonia.Value).Append(", ").
                    Append(electronicDocument.Data.Emisor.ExpedidoEn.Municipio.Value).Append(", ").
                    Append(electronicDocument.Data.Emisor.ExpedidoEn.Localidad.Value).Append(", ").
                    Append(electronicDocument.Data.Emisor.ExpedidoEn.CodigoPostal.Value).Append(", ").
                    Append(electronicDocument.Data.Emisor.ExpedidoEn.Estado.Value).Append(", ").
                    Append(electronicDocument.Data.Emisor.ExpedidoEn.Pais.Value);

                cel = new Cell(new Phrase(expedido.ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                datosReceptor.AddCell(cel);

                #endregion

                ev.dSaltoLinea = dSaltoLinea;

                #region "Tabla Detalles"

                Table encabezadoDetalle = new Table(6);
                float[] headerEncabezadoDetalle = { 10, 45, 10, 10, 10, 15 };
                encabezadoDetalle.Widths = headerEncabezadoDetalle;
                encabezadoDetalle.WidthPercentage = 100F;
                encabezadoDetalle.Padding = 1;
                encabezadoDetalle.Spacing = 1;
                encabezadoDetalle.BorderWidth = 0;
                encabezadoDetalle.DefaultCellBorder = 0;
                encabezadoDetalle.BorderColor = gris;

                cel = new Cell(new Phrase("\n", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = blanco;
                cel.Colspan = 6;
                encabezadoDetalle.AddCell(cel);

                // Número
                cel = new Cell(new Phrase("Cantidad", titulo1));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BackgroundColor = gris;
                encabezadoDetalle.AddCell(cel);

                // Descripción
                cel = new Cell(new Phrase("Descripción", titulo1));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.BackgroundColor = gris;
                cel.Colspan = 4;
                encabezadoDetalle.AddCell(cel);

                // Precio Unitario
                cel = new Cell(new Phrase("Total Valor      ", titulo1));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezadoDetalle.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["cantidad"].ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezadoDetalle.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["descripcion"].ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                encabezadoDetalle.AddCell(cel);

                cel = new Cell(new Phrase(Convert.ToDouble(dtOpcEncabezado.Rows[0]["dispersion"]).ToString("C", _c2), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                encabezadoDetalle.AddCell(cel);

                #endregion

                #region "Construimos el Comentarios"

                Table comentarios = new Table(6);
                float[] headerwidthsComentarios = { 10, 45, 10, 10, 10, 15 };
                comentarios.Widths = headerwidthsComentarios;
                comentarios.WidthPercentage = 100;
                comentarios.Padding = 1;
                comentarios.Spacing = 1;
                comentarios.BorderWidth = 0;
                comentarios.DefaultCellBorder = 0;
                comentarios.BorderColor = gris;

                //Concepto del comprobate 
                cel = new Cell(new Phrase("Concepto", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios.AddCell(cel);

                //unidad del comprobante
                cel = new Cell(new Phrase("Unidad.", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Unit", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                cel = new Cell(new Phrase("Importe", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
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

                cel = new Cell(new Phrase("I.V.A.  " + tasa.ToString() + " %", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios.AddCell(cel);

                // Creamos la tabla para insertar los conceptos de detalle de la factura
                PdfPTable tableConceptos = new PdfPTable(5);

                int[] colWithsConceptos = new int[5];
                //String[] arrColWidthConceptos = dtConfigFact.Rows[0]["conceptosColWidth"].ToString().Split(new Char[] { ',' });
                String[] arrColWidthConceptos = { "55", "10", "10", "10", "15" };

                for (int i = 0; i < arrColWidthConceptos.Length; i++)
                {
                    colWithsConceptos.SetValue(Convert.ToInt32(arrColWidthConceptos[i]), i);
                }

                tableConceptos.SetWidths(colWithsConceptos);
                tableConceptos.WidthPercentage = 100F;

                int numConceptos = electronicDocument.Data.Conceptos.Count;
                PdfPCell cellConceptos = new PdfPCell();
                PdfPCell cellMontos = new PdfPCell();
                PdfPCell celladicional = new PdfPCell();

                for (int i = 0; i < numConceptos; i++)
                {

                    cellConceptos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, f5));
                    cellConceptos.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                    cellConceptos.VerticalAlignment = PdfCell.ALIGN_MIDDLE;
                    cellConceptos.BorderWidthTop = (float).5;
                    cellConceptos.BorderWidthLeft = (float).5;
                    cellConceptos.BorderWidthRight = (float).5;
                    cellConceptos.BorderWidthBottom = (float).5;
                    cellConceptos.BorderColor = gris;
                    tableConceptos.AddCell(cellConceptos);

                    cellConceptos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Unidad.Value.ToString(), f5));
                    cellConceptos.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                    cellConceptos.VerticalAlignment = PdfCell.ALIGN_MIDDLE;
                    cellConceptos.BorderWidthTop = (float).5;
                    cellConceptos.BorderWidthLeft = (float).5;
                    cellConceptos.BorderWidthRight = (float).5;
                    cellConceptos.BorderWidthBottom = (float).5;
                    cellConceptos.BorderColor = gris;
                    tableConceptos.AddCell(cellConceptos);

                    cellMontos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].ValorUnitario.Value.ToString("C", _c2), f5));
                    cellMontos.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                    cellMontos.VerticalAlignment = PdfCell.ALIGN_MIDDLE;
                    cellMontos.BorderWidthTop = (float).5;
                    cellMontos.BorderWidthLeft = (float).5;
                    cellMontos.BorderWidthRight = (float).5;
                    cellMontos.BorderWidthBottom = (float).5;
                    cellMontos.BorderColor = gris;
                    tableConceptos.AddCell(cellMontos);

                    cellMontos = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("C", _c2), f5));
                    cellMontos.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                    cellMontos.VerticalAlignment = PdfCell.ALIGN_MIDDLE;
                    cellMontos.BorderWidthTop = (float).5;
                    cellMontos.BorderWidthLeft = (float).5;
                    cellMontos.BorderWidthRight = (float).5;
                    cellMontos.BorderWidthBottom = (float).5;
                    cellMontos.BorderColor = gris;
                    tableConceptos.AddCell(cellMontos);
                }

                //dSaltoLinea = new Chunk("\n\n ");

                //Dato de IVA
                celladicional = new PdfPCell(new Phrase(electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("C", _c2), f5));
                celladicional.VerticalAlignment = Element.ALIGN_MIDDLE;
                celladicional.HorizontalAlignment = Element.ALIGN_CENTER;
                celladicional.BorderWidthTop = (float).5;
                celladicional.BorderWidthLeft = (float).5;
                celladicional.BorderWidthRight = (float).5;
                celladicional.BorderWidthBottom = (float).5;
                celladicional.BorderColor = gris;
                tableConceptos.AddCell(celladicional);


                Table comentarios2 = new Table(6);
                float[] headerwidthsComentarios2 = { 10, 45, 10, 10, 10, 15 };
                comentarios2.Widths = headerwidthsComentarios2;
                comentarios2.WidthPercentage = 100;
                comentarios2.Padding = 1;
                comentarios2.Spacing = 1;
                comentarios2.BorderWidth = 0;
                comentarios2.DefaultCellBorder = 0;
                comentarios2.BorderColor = gris;

                //Descuento
                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["descuento"].ToString(), f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                comentarios2.AddCell(cel);

                cel = new Cell(new Phrase("", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios2.AddCell(cel);

                cel = new Cell(new Phrase("", f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios2.AddCell(cel);


                //Cantidad de Descuento
                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["cantidadDescuento"].ToString(), f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios2.AddCell(cel);

                dSaltoLinea = new Chunk("\n\n");

                //imagen RFC
                #endregion

                #region "Adicionales imagen RFC"

                Table comentarios1 = new Table(6);
                float[] headerwidthsComentarios1 = { 20, 25, 15, 10, 15, 15 };
                comentarios1.Widths = headerwidthsComentarios1;
                comentarios1.WidthPercentage = 100;
                comentarios1.Padding = 1;
                comentarios1.Spacing = 1;
                comentarios1.BorderWidth = 0;
                comentarios1.DefaultCellBorder = 0;
                comentarios1.BorderColor = gris;

                cel = new Cell(new Phrase("\n", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = blanco;
                cel.Colspan = 6;
                comentarios1.AddCell(cel);

                //imgen del RFC
                cel = new Cell(imgRFC);
                cel.VerticalAlignment = Element.ALIGN_CENTER;
                cel.HorizontalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 6;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase("IMPORTE DE CONTRAPRESTACIÓN EN LETRA:   ", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["comisionConLetra"].ToString(), f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios1.AddCell(cel);


                cel = new Cell(new Phrase("GRAN TOTAL:  ", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios1.AddCell(cel);


                cel = new Cell(new Phrase(Convert.ToDouble(dtOpcEncabezado.Rows[0]["granTotal"]).ToString("C", _c2), f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase("IMPORTE GRAN TOTAL EN LETRA:  ", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["granTotalConLetra"].ToString(), f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase("OBSERVACIONES:  ", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["observaciones"].ToString(), f5B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase("RÉGIMEN FISCAL:  ", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase("Régimen de persona Moral", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase("MÉTODO DE PAGO:  ", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.MetodoPago.Value, f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase("NO. DE CUENTA:  ", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.NumeroCuentaPago.Value, f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.BorderColor = gris;
                comentarios1.AddCell(cel);

                cel = new Cell(new Phrase("\n\n", f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_LEFT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 6;
                comentarios1.AddCell(cel);

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

                    cel = new Cell(new Phrase(objTimbre.Uuid.Value.ToUpper(), f5));
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
                    par.Add(new Chunk("MONEDA: ", f5B));
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

                dSaltoLinea = new Chunk("\n\n");

                Table adicional1 = new Table(4);
                float[] headerwidthsAdicional1 = { 25, 25, 25, 25 };
                adicional1.Widths = headerwidthsAdicional1;
                adicional1.WidthPercentage = 100;
                adicional1.Padding = 1;
                adicional1.Spacing = 1;
                adicional1.BorderWidth = (float).5;
                adicional1.DefaultCellBorder = 1;
                adicional1.BorderColor = blanco;

                if (dtOpcEncabezado.Rows[0]["cuentaBanamex"].ToString().Length > 0 ||
                        dtOpcEncabezado.Rows[0]["cuentaBanorte"].ToString().Length > 0 ||
                            dtOpcEncabezado.Rows[0]["cuentaBancomer"].ToString().Length > 0 ||
                                dtOpcEncabezado.Rows[0]["cuentaHSBC"].ToString().Length > 0 ||
                                    dtOpcEncabezado.Rows[0]["cuentaSantander"].ToString().Length > 0)
                {

                    cel = new Cell(new Phrase("\n\n\n", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = blanco;
                    cel.Colspan = 4;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase("Cuenta", f5B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase("Referencia", f5B));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["cuentaBanamex"].ToString(), f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["referenciaBanamex"].ToString(), f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["cuentaBanorte"].ToString(), f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["referenciaBanorte"].ToString(), f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["cuentaBancomer"].ToString(), f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["referenciaBancomer"].ToString(), f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["cuentaHSBC"].ToString(), f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["referenciaHSBC"].ToString(), f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["cuentaSantander"].ToString(), f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_LEFT;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(dtOpcEncabezado.Rows[0]["referenciaSantander"].ToString(), f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase(" ", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);

                    cel = new Cell(new Phrase("\n\n\n\n\n", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 4;
                    cel.BorderColor = gris;
                    adicional1.AddCell(cel);
                }
                #endregion

                #region "anexoEncabezado"
                Table encabezadoAnexoE = new Table(4);
                float[] headerEncabezadoAnexoE = { 20, 20, 40, 20 };
                encabezadoAnexoE.Widths = headerEncabezadoAnexoE;
                encabezadoAnexoE.WidthPercentage = 100F;
                encabezadoAnexoE.Padding = 1;
                encabezadoAnexoE.Spacing = 1;
                encabezadoAnexoE.BorderWidth = 0;
                encabezadoAnexoE.DefaultCellBorder = 0;
                encabezadoAnexoE.BorderColor = gris;

                PdfPTable tablaAnexoEncabezado = new PdfPTable(4);

                if (dtAnexoEncabezado.Rows.Count > 0)
                {
                    cel = new Cell(new Phrase("DETALLE DE TARJETAS POR FACTURA", new Font(Font.HELVETICA, 12, Font.NORMAL)));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = blanco;
                    cel.BackgroundColor = blanco;
                    cel.Colspan = 4;
                    encabezadoAnexoE.AddCell(cel);

                    cel = new Cell(new Phrase("\n\n", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = blanco;
                    cel.BackgroundColor = blanco;
                    cel.Colspan = 4;
                    encabezadoAnexoE.AddCell(cel);

                    cel = new Cell(new Phrase("Fecha de emisión", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoE.AddCell(cel);

                    cel = new Cell(new Phrase("Cliente", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoE.AddCell(cel);

                    cel = new Cell(new Phrase("Razón Social", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoE.AddCell(cel);

                    cel = new Cell(new Phrase("Folio Factura", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoE.AddCell(cel);



                    int[] colWithsAnexoE = new int[4];
                    String[] arrColWidthAnexoE = { "20", "20", "40", "20" };

                    for (int i = 0; i < arrColWidthAnexoE.Length; i++)
                    {
                        colWithsAnexoE.SetValue(Convert.ToInt32(arrColWidthAnexoE[i]), i);
                    }

                    tablaAnexoEncabezado.SetWidths(colWithsAnexoE);
                    tablaAnexoEncabezado.WidthPercentage = 100F;

                    int numConceptosAnexoE = dtAnexoEncabezado.Rows.Count;
                    PdfPCell cellConceptosAnexoE = new PdfPCell();

                    for (int i = 0; i < numConceptosAnexoE; i++)
                    {
                        cellConceptosAnexoE = new PdfPCell(new Phrase(dtAnexoEncabezado.Rows[i]["fechaEmision"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexoE.Border = 0;
                        cellConceptosAnexoE.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tablaAnexoEncabezado.AddCell(cellConceptosAnexoE);

                        cellConceptosAnexoE = new PdfPCell(new Phrase(dtAnexoEncabezado.Rows[i]["numeroCliente"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexoE.Border = 0;
                        cellConceptosAnexoE.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tablaAnexoEncabezado.AddCell(cellConceptosAnexoE);

                        cellConceptosAnexoE = new PdfPCell(new Phrase(dtAnexoEncabezado.Rows[i]["razonsocial"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexoE.Border = 0;
                        cellConceptosAnexoE.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                        tablaAnexoEncabezado.AddCell(cellConceptosAnexoE);

                        cellConceptosAnexoE = new PdfPCell(new Phrase(dtAnexoEncabezado.Rows[i]["folioFactura"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexoE.Border = 0;
                        cellConceptosAnexoE.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tablaAnexoEncabezado.AddCell(cellConceptosAnexoE);

                    }

                    cellConceptosAnexoE = new PdfPCell(new Phrase("\n\n", f5));
                    cellConceptosAnexoE.BorderColor = blanco;
                    cellConceptosAnexoE.Colspan = 7;
                    tablaAnexoEncabezado.AddCell(cellConceptosAnexoE);
                }

                #endregion

                #region "anexoDetalle"

                Table encabezadoAnexoDetalle = new Table(7);
                float[] headerEncabezadoAnexoDetalle = { 12, 12, 20, 12, 12, 12, 10 };
                encabezadoAnexoDetalle.Widths = headerEncabezadoAnexoDetalle;
                encabezadoAnexoDetalle.WidthPercentage = 100F;
                encabezadoAnexoDetalle.Padding = 1;
                encabezadoAnexoDetalle.Spacing = 1;
                encabezadoAnexoDetalle.BorderWidth = 0;
                encabezadoAnexoDetalle.DefaultCellBorder = 0;
                encabezadoAnexoDetalle.BorderColor = gris;

                PdfPTable tablaAnexoDetalle = new PdfPTable(7);

                if (dtAnexoDetalle.Rows.Count > 0)
                {
                    cel = new Cell(new Phrase("Tipo de Tarjeta", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoDetalle.AddCell(cel);

                    cel = new Cell(new Phrase("Numero de empleado", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoDetalle.AddCell(cel);

                    cel = new Cell(new Phrase("Nombre de empleado", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoDetalle.AddCell(cel);

                    cel = new Cell(new Phrase("Numero de Producto", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoDetalle.AddCell(cel);

                    cel = new Cell(new Phrase("Producto", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoDetalle.AddCell(cel);

                    cel = new Cell(new Phrase("Fecha Alta", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoDetalle.AddCell(cel);

                    cel = new Cell(new Phrase("Pedido", f5));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = gris;
                    encabezadoAnexoDetalle.AddCell(cel);



                    int[] colWithsAnexo = new int[7];
                    String[] arrColWidthAnexo = { "12", "12", "20", "12", "12", "12", "10" };

                    for (int i = 0; i < arrColWidthAnexo.Length; i++)
                    {
                        colWithsAnexo.SetValue(Convert.ToInt32(arrColWidthAnexo[i]), i);
                    }

                    tablaAnexoDetalle.SetWidths(colWithsAnexo);
                    tablaAnexoDetalle.WidthPercentage = 100F;

                    int numConceptosAnexo = dtAnexoDetalle.Rows.Count;
                    PdfPCell cellConceptosAnexo = new PdfPCell();

                    cellConceptosAnexo = new PdfPCell(new Phrase("\n", f5));
                    cellConceptosAnexo.BorderColor = blanco;
                    cellConceptosAnexo.Colspan = 7;
                    tablaAnexoDetalle.AddCell(cellConceptosAnexo);

                    for (int i = 0; i < numConceptosAnexo; i++)
                    {
                        cellConceptosAnexo = new PdfPCell(new Phrase(dtAnexoDetalle.Rows[i]["tipoTarjeta"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexo.Border = 0;
                        cellConceptosAnexo.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tablaAnexoDetalle.AddCell(cellConceptosAnexo);

                        cellConceptosAnexo = new PdfPCell(new Phrase(dtAnexoDetalle.Rows[i]["numeroEmpleado"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexo.Border = 0;
                        cellConceptosAnexo.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tablaAnexoDetalle.AddCell(cellConceptosAnexo);

                        cellConceptosAnexo = new PdfPCell(new Phrase(dtAnexoDetalle.Rows[i]["nombreEmpleado"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexo.Border = 0;
                        cellConceptosAnexo.HorizontalAlignment = PdfCell.ALIGN_LEFT;
                        tablaAnexoDetalle.AddCell(cellConceptosAnexo);

                        cellConceptosAnexo = new PdfPCell(new Phrase(dtAnexoDetalle.Rows[i]["numeroProducto"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexo.Border = 0;
                        cellConceptosAnexo.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tablaAnexoDetalle.AddCell(cellConceptosAnexo);

                        cellConceptosAnexo = new PdfPCell(new Phrase(dtAnexoDetalle.Rows[i]["producto"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexo.Border = 0;
                        cellConceptosAnexo.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tablaAnexoDetalle.AddCell(cellConceptosAnexo);

                        cellConceptosAnexo = new PdfPCell(new Phrase(dtAnexoDetalle.Rows[i]["fechaAlta"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexo.Border = 0;
                        cellConceptosAnexo.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tablaAnexoDetalle.AddCell(cellConceptosAnexo);

                        cellConceptosAnexo = new PdfPCell(new Phrase(dtAnexoDetalle.Rows[i]["pedido"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL)));
                        cellConceptosAnexo.Border = 0;
                        cellConceptosAnexo.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                        tablaAnexoDetalle.AddCell(cellConceptosAnexo);
                    }
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

                cell = new PdfPCell(new Phrase("ESTE DOCUMENTO ES UNA REPRESENTACIÓN IMPRESA DE UN CFDI http://www.opam.com.mx", titulo1));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = gris;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                footer.AddCell(cell);

                #endregion

                ev.footer = footer;

                document.Open();

                document.Add(encabezadoLogo);
                document.Add(datosReceptor);
                document.Add(encabezadoDetalle);
                document.Add(comentarios);
                document.Add(tableConceptos);
                document.Add(comentarios2);
                document.Add(comentarios1);
                document.Add(adicional);
                document.Add(adicional1);
                document.Add(encabezadoAnexoE);
                document.Add(tablaAnexoEncabezado);
                document.Add(encabezadoAnexoDetalle);
                document.Add(tablaAnexoDetalle);


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
                //string exe = ex.Message;
            }
        }


        #endregion
    }

    public class pdfPageEventHandlerOPAM : PdfPageEventHelper
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
        public Table encLogo { get; set; }
        public Table encTitulos { get; set; }
        public Table datosCliente { get; set; }
        public Chunk dSaltoLinea { get; set; }
        public Table adicional { get; set; }
        public Table adicional1 { get; set; }

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
            base.OnStartPage(writer, document);
            //document.Add(encLogo);
            //document.Add(datosCliente);            
            document.Add(dSaltoLinea);
            //document.Add(encTitulos);

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

    public class DefaultSplitCharacterOPAM : ISplitCharacter
    {
        #region "ISplitCharacter"

        /**
         * An instance of the default SplitCharacter.
         */
        public static readonly ISplitCharacter DEFAULT = new DefaultSplitCharacterOPAM();

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