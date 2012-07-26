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
    public class DENTEGRA
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
        private static String pathIMGCEDULA;
        private static Cell cel;
        private static Paragraph par;

        private static Color azul;
        private static Color gris;

        private static BaseFont EM;
        private static Font f5;
        private static Font f8;
        private static Font f8B;
        private static Font f8L;
        private static Font f9;
        private static Font f9B;
        private static Font titulo;

        #endregion

        #region "generarPdf"

        public static string generarPdf(Hashtable htFacturaxion, HttpContext hc)
        {
            try
            {
                DAL dal = new DAL();
                _ci.NumberFormat.CurrencyDecimalDigits = 2;
                string pathPdf = htFacturaxion["rutaDocumentoPdf"].ToString();

                ElectronicDocument electronicDocument = (ElectronicDocument)htFacturaxion["electronicDocument"];
                Data objTimbre = (Data)htFacturaxion["objTimbre"];
                timbrar = Convert.ToBoolean(htFacturaxion["timbrar"]);
                pathPdf = htFacturaxion["rutaDocumentoPdf"].ToString();
                Int64 idCfdi = Convert.ToInt64(htFacturaxion["idCfdi"]);

                #region "Obtenemos los datos del CFDI y Campos Opcionales"

                StringBuilder sbOpcionalEncabezado = new StringBuilder();
                DataTable dtOpcEnc = new DataTable();
                //StringBuilder sbOpcionalDetalle = new StringBuilder();
                DataTable dtOpcDet = new DataTable();
                StringBuilder sbDataEmisor = new StringBuilder();
                DataTable dtDataEmisor = new DataTable();

                sbOpcionalEncabezado.
                    Append("SELECT ").
                    Append("campo1 AS [NUMERO-POLIZA], ").
                    Append("campo2 AS [NUMERO-ENDOSO], ").
                    Append("campo3 AS [INICIO-VIGENCIA], ").
                    Append("campo4 AS [FIN-VIGENCIA], ").
                    Append("campo5 AS [CLAVE], ").
                    Append("campo6 AS [NOMBRE-AGENTE], ").
                    Append("campo9 AS [NUMERO-RECIBO], ").
                    Append("campo10 AS [FECHA-LIMITE], ").
                    Append("campo11 AS [SERIE-RECIBO], ").
                    Append("campo12 AS [ESQUEMA-PAGO], ").
                    Append("campo14 AS [BANCO], ").
                    Append("campo15 AS [PRIMA-NETA], ").
                    Append("campo16 AS [DERECHOS], ").
                    Append("campo17 AS [RECARGOS], ").
                    Append("campo18 AS [CANTIDAD-LETRA] ").
                    Append("FROM opcionalEncabezado ").
                    Append("WHERE idCFDI = @0  AND ST = 1 ");

                //sbOpcionalDetalle.
                //    Append("SELECT ").
                //    Append("COALESCE(campo1, '') AS codeLocal, ").
                //    Append("COALESCE(campo3, '') AS lote, ").
                //    Append("COALESCE(campo4, '') AS cantidad, ").
                //    Append("COALESCE(campo5, '') AS expiracion, ").
                //    Append("COALESCE(campo6, '') AS taxRate, ").
                //    Append("COALESCE(campo7, '') AS taxPaid, ").
                //    Append("COALESCE(campo10, '') AS codeOracle, ").
                //    Append("COALESCE(campo11, '') AS codeISPC, ").
                //    Append("COALESCE(campo12, '') AS codeImpuesto, ").
                //    Append("COALESCE(campo13, '') AS centroCostos, ").
                //    Append("COALESCE(campo14, '') AS clinico, ").
                //    Append("COALESCE(campo15, '') AS proyecto, ").
                //    Append("COALESCE(campo16, '') AS cantidadReal, ").
                //    Append("COALESCE(campo17, '') AS descuento, ").
                //    Append("COALESCE(campo18, '') AS codBarras ").
                //    Append("FROM opcionalDetalle ").
                //    Append("WHERE idCFDI = @0 ");

                sbDataEmisor.Append("SELECT nombreSucursal FROM sucursales WHERE idSucursal = @0 ");

                dtOpcEnc = dal.QueryDT("DS_FE", sbOpcionalEncabezado.ToString(), "F:I:" + idCfdi, hc);
                //dtOpcDet = dal.QueryDT("DS_FE", sbOpcionalDetalle.ToString(), "F:I:" + idCfdi, hc);
                dtDataEmisor = dal.QueryDT("DS_FE", sbDataEmisor.ToString(), "F:I:" + htFacturaxion["idSucursalEmisor"],
                                           hc);

                //if (dtOpcDet.Rows.Count == 0)
                //{
                //    for (int i = 1; i <= electronicDocument.Data.Conceptos.Count; i++)
                //    {
                //        dtOpcDet.Rows.Add("", "0.00");
                //    }
                //}

                #endregion

                #region "Extraemos los datos del CFDI"

                htFacturaxion.Add("nombreEmisor", electronicDocument.Data.Emisor.Nombre.Value);
                htFacturaxion.Add("rfcEmisor", electronicDocument.Data.Emisor.Rfc.Value);
                htFacturaxion.Add("nombreReceptor", electronicDocument.Data.Receptor.Nombre.Value);
                htFacturaxion.Add("rfcReceptor", electronicDocument.Data.Receptor.Rfc.Value);
                htFacturaxion.Add("sucursal", dtDataEmisor.Rows[0]["nombreSucursal"]);
                htFacturaxion.Add("serie", electronicDocument.Data.Serie.Value);
                htFacturaxion.Add("folio", electronicDocument.Data.Folio.Value);
                htFacturaxion.Add("fechaCfdi", electronicDocument.Data.Fecha.Value);
                htFacturaxion.Add("UUID", objTimbre.Uuid.Value);

                #region "Dirección Emisor"

                StringBuilder sbDirEmisor1 = new StringBuilder();
                StringBuilder sbDirEmisor2 = new StringBuilder();
                StringBuilder sbDirEmisor3 = new StringBuilder();

                if (electronicDocument.Data.Emisor.Domicilio.Calle.Value.Length > 0)
                {
                    sbDirEmisor1.Append("Av. ").Append(electronicDocument.Data.Emisor.Domicilio.Calle.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.NumeroExterior.Value.Length > 0)
                {
                    sbDirEmisor1.Append(electronicDocument.Data.Emisor.Domicilio.NumeroExterior.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.NumeroInterior.Value.Length > 0)
                {
                    sbDirEmisor1.Append("Piso ").Append(electronicDocument.Data.Emisor.Domicilio.NumeroInterior.Value);
                }
                if (electronicDocument.Data.Emisor.Domicilio.Colonia.Value.Length > 0)
                {
                    sbDirEmisor2.Append("Col. ").Append(electronicDocument.Data.Emisor.Domicilio.Colonia.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.CodigoPostal.Value.Length > 0)
                {
                    sbDirEmisor2.Append("C.P. ").Append(electronicDocument.Data.Emisor.Domicilio.CodigoPostal.Value);
                }
                if (electronicDocument.Data.Emisor.Domicilio.Municipio.Value.Length > 0)
                {
                    sbDirEmisor3.Append(electronicDocument.Data.Emisor.Domicilio.Municipio.Value).Append(", ");
                }
                if (electronicDocument.Data.Emisor.Domicilio.Estado.Value.Length > 0)
                {
                    sbDirEmisor3.Append(electronicDocument.Data.Emisor.Domicilio.Estado.Value).Append(", ");
                }

                sbDirEmisor3.Append(electronicDocument.Data.Emisor.Domicilio.Pais.Value);

                #endregion

                #region "Dirección Receptor"

                StringBuilder sbDirReceptor1 = new StringBuilder();
                StringBuilder sbDirReceptor2 = new StringBuilder();
                StringBuilder sbDirReceptor3 = new StringBuilder();

                if (electronicDocument.Data.Receptor.Domicilio.Calle.Value.Length > 0)
                {
                    sbDirReceptor1.Append("Calle:                             ").Append(electronicDocument.Data.Receptor.Domicilio.Calle.Value);
                }
                if (electronicDocument.Data.Receptor.Domicilio.Colonia.Value.Length > 0)
                {
                    sbDirReceptor1.Append("\n         Colonia:                         ").Append(electronicDocument.Data.Receptor.Domicilio.Colonia.Value);
                }

                sbDirReceptor2.Append("\nNo. Ext:  ").Append(electronicDocument.Data.Receptor.Domicilio.NumeroExterior.Value).Append(" ");
                sbDirReceptor2.Append("                    No. Int:  ").Append(electronicDocument.Data.Receptor.Domicilio.NumeroInterior.Value).Append(" ");

                if (electronicDocument.Data.Receptor.Domicilio.CodigoPostal.Value.Length > 0)
                {
                    sbDirReceptor2.Append("\nC.P:        ").Append(
                        electronicDocument.Data.Receptor.Domicilio.CodigoPostal.Value);
                }
                if (electronicDocument.Data.Receptor.Domicilio.Municipio.Value.Length > 0)
                {
                    sbDirReceptor3.Append("Delegación / Municipio:    ").Append(
                        electronicDocument.Data.Receptor.Domicilio.Municipio.Value).Append(" ");
                }
                if (electronicDocument.Data.Receptor.Domicilio.Estado.Value.Length > 0)
                {
                    sbDirReceptor3.Append("\n         Estado:                          ").Append(electronicDocument.Data.Receptor.Domicilio.Estado.Value).
                        Append(" ");
                }

                sbDirReceptor2.Append("\n\nPaís:       ").Append(electronicDocument.Data.Receptor.Domicilio.Pais.Value);

                #endregion

                htFacturaxion.Add("direccionEmisor1", sbDirEmisor1.ToString());
                htFacturaxion.Add("direccionEmisor2", sbDirEmisor2.ToString());
                htFacturaxion.Add("direccionEmisor3", sbDirEmisor3.ToString());

                htFacturaxion.Add("direccionReceptor1", sbDirReceptor1.ToString());
                htFacturaxion.Add("direccionReceptor2", sbDirReceptor2.ToString());
                htFacturaxion.Add("direccionReceptor3", sbDirReceptor3.ToString());

                #endregion

                #region "Creamos el Objeto Documento y Tipos de Letra"

                Document document = new Document(PageSize.LETTER, 40, 40, 40, 40);
                document.AddAuthor("Facturaxion");
                document.AddCreator("r3Take");
                document.AddCreationDate();

                FileStream fs = new FileStream(pathPdf, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                pdfPageEventHandlerDENTEGRA pageEventHandler = new pdfPageEventHandlerDENTEGRA();
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                writer.PageEvent = pageEventHandler;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                pathIMGLOGO = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\DSD0611086V4\logo.jpg";
                pathIMGCEDULA = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\DSD0611086V4\RFC.jpg";

                azul = new Color(66, 138, 205);
                gris = new Color(74, 74, 74);

                EM = BaseFont.CreateFont(@"C:\Windows\Fonts\Tahoma.TTF", BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
                f5 = new Font(EM, 5, Font.NORMAL);
                f8 = new Font(EM, 6, Font.NORMAL);
                f8B = new Font(EM, 6, Font.BOLD);
                f9 = new Font(EM, 7, Font.NORMAL);
                f9B = new Font(EM, 7, Font.BOLD);
                titulo = new Font(EM, 7, Font.BOLD);

                #endregion

                #region "Construimos el Documento"

                #region "Construimos el Encabezado"

                Table encabezado = new Table(7);
                float[] headerwidthsEncabezado = { 10, 20, 30, 10, 15, 10, 5 };
                encabezado.Widths = headerwidthsEncabezado;
                encabezado.WidthPercentage = 100;
                encabezado.Padding = (float).5;
                encabezado.Spacing = 1;
                encabezado.BorderWidth = (float).5;
                encabezado.DefaultCellBorder = 0;
                encabezado.BorderColor = gris;

                Image imgLogo = Image.GetInstance(pathIMGLOGO);
                imgLogo.ScalePercent(47f);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(1f, 1f);
                par.Add(new Chunk(imgLogo, 0, 0));
                par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 2;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                string tagTipoComprobante = string.Empty;
                if (electronicDocument.Data.TipoComprobante.Value.ToUpper() == "INGRESO")
                {
                    tagTipoComprobante = "RECIBO DE PRIMAS";
                }
                else if (electronicDocument.Data.TipoComprobante.Value.ToUpper() == "EGRESO")
                {
                    tagTipoComprobante = "NOTA DE CREDITO";
                }
                else
                {
                    tagTipoComprobante = "";
                }
                cel = new Cell(new Phrase("     " + tagTipoComprobante, titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Serie: ", f9));
                par.Add(new Chunk(htFacturaxion["serie"].ToString().ToUpper(), f9));
                par.Add(new Chunk("\nFolio: ", f9));
                par.Add(new Chunk(electronicDocument.Data.Folio.Value, f9));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk(htFacturaxion["nombreEmisor"].ToString().ToUpper(), f9B));
                par.Add(new Chunk("\n" + htFacturaxion["direccionEmisor1"].ToString().ToUpper(), f9));
                par.Add(new Chunk("\n" + htFacturaxion["direccionEmisor2"].ToString().ToUpper(), f9));
                par.Add(new Chunk("\n" + htFacturaxion["direccionEmisor3"].ToString().ToUpper(), f9));
                par.Add(new Chunk("\nTel. +52 (55) 5002-3100\n", f9));
                par.Add(new Chunk("\n", f9));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

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
                    regimenes.Append("Ley General Regimen Persona Moral");
                }

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("RFC: ", f9));
                par.Add(new Chunk(htFacturaxion["rfcEmisor"].ToString().ToUpper(), f9B));
                par.Add(new Chunk("\n\nRégimen Fiscal: ", f9));
                par.Add(new Chunk(regimenes.ToString(), f9));
                par.Add(new Chunk("\n\nLugar Expedicion: ", f9));
                par.Add(new Chunk(electronicDocument.Data.LugarExpedicion.IsAssigned
                                      ? electronicDocument.Data.LugarExpedicion.Value
                                      : "Mexico, D.F.", f9));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 4;
                encabezado.AddCell(cel);

                cel = new Cell(new Phrase("DATOS DEL CONTRATANTE", titulo));
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 7;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Nombre o Razón Social: ", f9));
                par.Add(new Chunk(electronicDocument.Data.Receptor.Nombre.Value, f9));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("RFC: ", f9));
                par.Add(new Chunk(electronicDocument.Data.Receptor.Rfc.Value, f9));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 4;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Domicilio:\n", f9));
                par.Add(new Chunk("         " + htFacturaxion["direccionReceptor1"] + "\n", f9));
                par.Add(new Chunk("         " + htFacturaxion["direccionReceptor3"], f9));
                par.Add(new Chunk("\n ", f9));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 3;
                encabezado.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk(htFacturaxion["direccionReceptor2"].ToString(), f9));
                par.Add(new Chunk("\n ", f9));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 4;
                encabezado.AddCell(cel);

                #endregion

                #region "Opcionales"

                #region "Datos Poliza y Agente"

                Table datosPolizaAgente = new Table(4);
                float[] headerwidthsPolizaAgente = { 18, 42, 25, 15 };
                datosPolizaAgente.Widths = headerwidthsPolizaAgente;
                datosPolizaAgente.WidthPercentage = 100;
                datosPolizaAgente.Padding = 1;
                datosPolizaAgente.Spacing = 1;
                datosPolizaAgente.BorderWidth = (float).5;
                datosPolizaAgente.DefaultCellBorder = 0;
                datosPolizaAgente.BorderColor = gris;

                cel = new Cell(new Phrase("DATOS DE LA POLIZA", titulo));
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 4;
                datosPolizaAgente.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Número de Póliza:   ", f9));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["NUMERO-POLIZA"].ToString(), f9)); //1
                par.Add(new Chunk("\nNúmero de Endoso: ", f9));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["NUMERO-ENDOSO"].ToString(), f9)); //2
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                datosPolizaAgente.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Inicio de Vigencia:      ", f9));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["INICIO-VIGENCIA"].ToString(), f9)); //3
                par.Add(new Chunk("\nFin de Vigencia:         ", f9));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["FIN-VIGENCIA"].ToString(), f9)); //4
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 2;
                cel.BorderColor = gris;
                datosPolizaAgente.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                datosPolizaAgente.AddCell(cel);

                #endregion

                #region "Datos del Recibo"

                Table datosRecibo = new Table(4);
                float[] headerwidthsDatosRecibo = { 20, 40, 20, 20 };
                datosRecibo.Widths = headerwidthsDatosRecibo;
                datosRecibo.WidthPercentage = 100;
                datosRecibo.Padding = 1;
                datosRecibo.Spacing = 1;
                datosRecibo.BorderWidth = (float).5;
                datosRecibo.DefaultCellBorder = 0;
                datosRecibo.BorderColor = gris;

                cel = new Cell(new Phrase("DATOS DEL AGENTE", titulo));
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 4;
                datosRecibo.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Clave: ", f9));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["CLAVE"].ToString(), f9)); //5
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosRecibo.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Nombre: ", f9));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["NOMBRE-AGENTE"].ToString(), f9)); //6
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 3;
                datosRecibo.AddCell(cel);

                cel = new Cell(new Phrase("", f5)); //4
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                datosPolizaAgente.AddCell(cel);

                cel = new Cell(new Phrase("DATOS DEL RECIBO", titulo));
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 4;
                datosRecibo.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Folio Fiscal: ", f9));
                par.Add(new Chunk("\nCertificado del Emisor: ", f9));
                par.Add(new Chunk("\nNúmero de recibo: ", f9));
                par.Add(new Chunk("\nParcialidad o Serie del Recibo: ", f9));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosRecibo.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk(objTimbre.Uuid.Value, f9));//Folio Fiscal
                par.Add(new Chunk("\n", f9));
                par.Add(new Chunk(electronicDocument.Data.NumeroCertificado.Value, f9));//Certificado del Emisor
                par.Add(new Chunk("\n", f9));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["NUMERO-RECIBO"].ToString(), f9)); //9 Número de recibo
                par.Add(new Chunk("\n", f9));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["SERIE-RECIBO"].ToString(), f9)); //11 Parcialidad 
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosRecibo.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Fecha y Hora de Emision: ", f9));
                par.Add(new Chunk("\nFecha y Hora de Certificación: ", f9));
                par.Add(new Chunk("\nFecha Límite de Pago: ", f9));
                par.Add(new Chunk("\nEsquema Pago: ", f9));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosRecibo.AddCell(cel);

                string[] fechaEmisionCFDI = Convert.ToDateTime(htFacturaxion["fechaCfdi"].ToString()).GetDateTimeFormats('s');
                string[] fechaTimbrado = Convert.ToDateTime(objTimbre.FechaTimbrado.Value).GetDateTimeFormats('s');
                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk(fechaEmisionCFDI[0], f9));
                par.Add(new Chunk("\n", f9));
                par.Add(new Chunk(fechaTimbrado[0], f9));
                par.Add(new Chunk("\n", f9));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["FECHA-LIMITE"].ToString(), f9)); //10
                par.Add(new Chunk("\n", f9));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["ESQUEMA-PAGO"].ToString().ToUpper(), f9)); //12
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosRecibo.AddCell(cel);

                cel = new Cell(new Phrase("", f5)); //4
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                datosRecibo.AddCell(cel);

                #endregion

                #region "Datos Bancarios"

                Table datosBancarios = new Table(4);
                float[] headerwidthsDatosBancarios = { 12, 48, 15, 25 };
                datosBancarios.Widths = headerwidthsDatosBancarios;
                datosBancarios.WidthPercentage = 100;
                datosBancarios.Padding = 1;
                datosBancarios.Spacing = 1;
                datosBancarios.BorderWidth = (float).5;
                datosBancarios.DefaultCellBorder = 0;
                datosBancarios.BorderColor = gris;

                cel = new Cell(new Phrase("DATOS BANCARIOS", titulo));
                cel.BorderColor = gris;
                cel.BackgroundColor = azul;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 4;
                datosBancarios.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Banco: ", f9));
                par.Add(new Chunk("\nForma de Pago: ", f9));
                par.Add(new Chunk("\nTipo de Pago: ", f9));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosBancarios.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk(dtOpcEnc.Rows[0]["BANCO"].ToString(), f9)); //14
                par.Add(new Chunk("\n", f9));
                par.Add(new Chunk(electronicDocument.Data.MetodoPago.Value, f9)); //METODO DE PAGO
                par.Add(new Chunk("\n", f9));
                par.Add(new Chunk(electronicDocument.Data.CondicionesPago.Value, f9)); //TIPO DE PAGO
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosBancarios.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("Número de Cuenta: ", f9));
                par.Add(new Chunk("\n\nMoneda: ", f9));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosBancarios.AddCell(cel);

                string cuenta = electronicDocument.Data.NumeroCuentaPago.IsAssigned
                                    ? electronicDocument.Data.NumeroCuentaPago.Value
                                    : "";
                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk(cuenta, f9));
                par.Add(new Chunk("\n\n", f9));
                par.Add(new Chunk(electronicDocument.Data.Moneda.Value, f9));
                cel = new Cell(par);
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                datosBancarios.AddCell(cel);

                cel = new Cell(new Phrase("", f5));
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 4;
                datosBancarios.AddCell(cel);

                #endregion

                #endregion

                #region "Construimos Tablas de Partidas"

                #region "Construimos Encabezados de Partidas"

                Table encabezadoPartidas = new Table(5);
                float[] headerwidthsEncabesadoPartidas = { 10, 15, 35, 20, 20 };
                encabezadoPartidas.Widths = headerwidthsEncabesadoPartidas;
                encabezadoPartidas.WidthPercentage = 100;
                encabezadoPartidas.Padding = 1;
                encabezadoPartidas.Spacing = 1;
                encabezadoPartidas.BorderWidth = (float).5;
                encabezadoPartidas.DefaultCellBorder = 1;
                encabezadoPartidas.BorderColor = gris;

                cel = new Cell(new Phrase("Cantidad.", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Unidad", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Concepto", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Valor Unitario", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                cel = new Cell(new Phrase("Importe", titulo));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                encabezadoPartidas.AddCell(cel);

                #endregion

                #region "Construimos Contenido de las Partidas"

                Table partidas = new Table(5);
                float[] headerwidthsPartidas = { 10, 15, 35, 20, 20 };
                partidas.Widths = headerwidthsPartidas;
                partidas.WidthPercentage = 100;
                partidas.Padding = 1;
                partidas.Spacing = 1;
                partidas.BorderWidth = 0;
                partidas.DefaultCellBorder = 0;
                partidas.BorderColor = gris;
                int rowPartidas = electronicDocument.Data.Conceptos.Count;
                if (rowPartidas > 0)
                {
                    for (int i = 0; i < rowPartidas; i++)
                    {
                        cel = new Cell(new Phrase((i + 1).ToString(), f9));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = (float).5;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Unidad.Value, f9));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        cel.Rowspan = 2;
                        partidas.AddCell(cel);

                        cel = new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, f9));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel =
                            new Cell(
                                new Phrase(electronicDocument.Data.Conceptos[i].ValorUnitario.Value.ToString("C", _ci), f9));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = 0;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);

                        cel =
                            new Cell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("C", _ci),
                                                f9));
                        cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cel.BorderWidthTop = 0;
                        cel.BorderWidthLeft = 0;
                        cel.BorderWidthRight = (float).5;
                        cel.BorderWidthBottom = 0;
                        cel.BorderColor = gris;
                        partidas.AddCell(cel);
                    }
                }

                #endregion

                #endregion

                #region "Construimos el Cantidades"

                Table cantidades = new Table(2);
                float[] headerwidthsComentarios = { 85, 15 };
                cantidades.Widths = headerwidthsComentarios;
                cantidades.WidthPercentage = 100;
                cantidades.Padding = (float).5;
                cantidades.Spacing = 1;
                cantidades.BorderWidth = 0;
                cantidades.DefaultCellBorder = 0;
                cantidades.BorderColor = gris;

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
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("\nPrima Neta:\n", f9B));
                par.Add(new Chunk("Derechos:\n ", f9B));
                par.Add(new Chunk("Recargos:\n ", f9B));
                par.Add(new Chunk("Descuento por comisión:\n ", f9B));
                par.Add(new Chunk("Sub Total:\n ", f9B));
                par.Add(new Chunk("I.V.A: ", f9B));
                par.Add(new Chunk(tasa + " %\n ", f9));
                par.Add(new Chunk("TOTAL: ", f9B));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cantidades.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk("\n" + Convert.ToDouble(dtOpcEnc.Rows[0]["PRIMA-NETA"]).ToString("C", _ci), f9));
                par.Add(new Chunk("\n" + Convert.ToDouble(dtOpcEnc.Rows[0]["DERECHOS"]).ToString("C", _ci), f9));
                par.Add(new Chunk("\n" + Convert.ToDouble(dtOpcEnc.Rows[0]["RECARGOS"]).ToString("C", _ci), f9));
                par.Add(new Chunk("\n" + electronicDocument.Data.Descuento.Value.ToString("C", _ci), f9));
                par.Add(new Chunk("\n" + electronicDocument.Data.SubTotal.Value.ToString("C", _ci), f9));
                par.Add(new Chunk("\n" + electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("C", _ci), f9));
                par.Add(new Chunk("\n" + electronicDocument.Data.Total.Value.ToString("C", _ci), f9B));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cantidades.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(9f, 10f);
                par.Add(new Chunk(" Importe con Letra: ", f9B));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["CANTIDAD-LETRA"].ToString(), f9));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.BorderColor = gris;
                cel.Colspan = 2;
                cantidades.AddCell(cel);

                #endregion

                #region "Datos Fiscales"

                Table datosFiscales = new Table(3);
                float[] headerwidthsDatosFiscales = { 15, 70, 15 };
                datosFiscales.Widths = headerwidthsDatosFiscales;
                datosFiscales.WidthPercentage = 100;
                datosFiscales.Padding = 1;
                datosFiscales.Spacing = 1;
                datosFiscales.BorderWidth = 0;
                datosFiscales.DefaultCellBorder = 0;
                datosFiscales.BorderColor = gris;

                #region "Generamos Quick Response Code"

                byte[] bytesQRCode = new byte[0];

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

                #endregion

                Image imageQRCode = Image.GetInstance(bytesQRCode);
                imageQRCode.Alignment = (Image.TEXTWRAP | Image.ALIGN_LEFT);
                imageQRCode.ScaleToFit(90f, 90f);
                imageQRCode.IndentationLeft = 9f;
                imageQRCode.SpacingAfter = 9f;
                imageQRCode.BorderColorTop = Color.WHITE;

                cel = new Cell(new Phrase("ESTE DOCUMENTO ES UNA REPRESENTACIÓN IMPRESA DE UN CFDI", f9B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 3;
                datosFiscales.AddCell(cel);

                cel = new Cell(imageQRCode);
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                datosFiscales.AddCell(cel);

                DefaultSplitCharacter split = new DefaultSplitCharacter();
                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(8f, 9f);
                par.Add(new Chunk("Certificado SAT: ", f8B));
                par.Add(new Chunk(objTimbre.NumeroCertificadoSat.Value, f8));
                par.Add(new Chunk("\n\nCadena Original del Complemento de Certificado Digital del SAT\n", f8B));
                par.Add(new Chunk(electronicDocument.FingerPrintPac, f8).SetSplitCharacter(split));
                par.Add(new Chunk("\nSello Digital del Emisor\n", f8B));
                par.Add(new Chunk(electronicDocument.Data.Sello.Value, f8).SetSplitCharacter(split));
                par.Add(new Chunk("\nSello Digital del SAT\n", f8B));
                par.Add(new Chunk(objTimbre.SelloSat.Value, f8).SetSplitCharacter(split));
                par.Add(new Chunk("\n ", f5));
                cel = new Cell(par);
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthBottom = (float).5;
                datosFiscales.AddCell(cel);

                Image imgCedula = Image.GetInstance(pathIMGCEDULA);
                imgCedula.ScalePercent(47f);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(1f, 1f);
                par.Add(new Chunk(imgCedula, 0, 0));
                par.Add(new Chunk("", f5));
                cel = new Cell(par);
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderWidthTop = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthBottom = (float).5;
                datosFiscales.AddCell(cel);

                #endregion

                #endregion

                pageEventHandler.encabezado = encabezado;
                document.Open();
                document.Add(datosPolizaAgente);
                document.Add(datosRecibo);
                document.Add(datosBancarios);
                document.Add(encabezadoPartidas);
                document.Add(partidas);
                document.Add(cantidades);
                document.Add(datosFiscales);
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
                return "0#" + ex.Message;
            }
        }

        #endregion

        public class pdfPageEventHandlerDENTEGRA : PdfPageEventHelper
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

            public Table encabezado { get; set; }

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

                base.OnEndPage(writer, document);

                string lblPagina = "Página ";
                string lblDe = " de ";
                string lblFechaImpresion = "Fecha de Impresión ";

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