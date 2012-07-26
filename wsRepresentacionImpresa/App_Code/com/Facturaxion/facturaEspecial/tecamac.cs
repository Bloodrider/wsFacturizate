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
    public class tecamac
    {
        private static readonly CultureInfo _ci = new CultureInfo("es-mx");
        private static readonly string _rutaDocs = ConfigurationManager.AppSettings["rutaDocs"];
        private static readonly string _rutaDocsExt = ConfigurationManager.AppSettings["rutaDocsExterna"];
   
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
                //string pathPdf = htFacturaxion["rutaDocumentoPdf"].ToString();

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

                //Obtenemos Rol de la empresa
                StringBuilder sbRol = new StringBuilder();

                sbRol.
                    Append("DECLARE @iduser INT; ").
                    Append("SELECT DISTINCT @iduser = UxR.idUser FROM sucursales S ").
                    Append("LEFT OUTER JOIN sucursalesXUsuario SxU ON S.idSucursal = SxU.idSucursal LEFT OUTER JOIN r3TakeCore.dbo.SYS_UserXRol UxR ON UxR.idUser = SxU.idUser ").
                    Append("WHERE idEmpresa = @0 AND UxR.idRol IN(22,15) AND S.ST = 1 AND SxU.ST = 1; ").
                    Append("SELECT UxR.idRol FROM r3TakeCore.dbo.SYS_UserXRol UxR WHERE idUser = @idUser");  //Rol 15 > MIT ; 22 > Facturizate

                int idRol = dal.ExecuteScalar("DS_FE", sbRol.ToString(), "F:S:" + htFacturaxion["idEmisor"], hc);

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

                if (idRol == 15)
                {
                    if (rutaHeader.Length > 0)
                    {
                        sbConfigFact.
                            Append("SELECT @5 AS rutaTemplateHeader, @6 AS rutaTemplateFooter, @4 AS rutaLogo, objDesc, posX, posY, fontSize, dbo.convertNumToTextFunction( @2, @3) AS cantidadLetra, ").
                            Append("logoPosX, logoPosY, headerPosX, headerPosY, footerPosX, footerPosY, conceptosColWidth, desgloseColWidth ").
                            Append("FROM configuracionFacturas CF ").
                            Append("LEFT OUTER JOIN sucursales S ON S.idSucursal = @0 ").
                            Append("LEFT OUTER JOIN configuracionFactDet CFD ON CF.idConFact = CFD.idConFact ").
                            Append("WHERE CF.ST = 1 AND CF.idEmpresa = -1 AND CF.idTipoComp = @1 AND idCFDProcedencia = 1 AND objDesc NOT LIKE 'nuevoLbl%' ");
                    }
                    else
                    {
                        sbConfigFact.
                            Append("SELECT rutaTemplateHeader, rutaTemplateFooter, @4 AS rutaLogo, objDesc, posX, posY, fontSize, dbo.convertNumToTextFunction( @2, @3) AS cantidadLetra, ").
                            Append("logoPosX, logoPosY, headerPosX, headerPosY, footerPosX, footerPosY, conceptosColWidth, desgloseColWidth ").
                            Append("FROM configuracionFacturas CF ").
                            Append("LEFT OUTER JOIN sucursales S ON S.idSucursal = @0 ").
                            Append("LEFT OUTER JOIN configuracionFactDet CFD ON CF.idConFact = CFD.idConFact ").
                            Append("WHERE CF.ST = 1 AND CF.idEmpresa = -1 AND CF.idTipoComp = @1 AND idCFDProcedencia = 1 AND objDesc NOT LIKE 'nuevoLbl%' ");
                    }
                }
                else
                {
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
                }

                DataTable dtConfigFact = dal.QueryDT("DS_FE", sbConfigFact.ToString(), sbConfigFactParms.ToString(), hc);

                // Creamos el Objeto Documento
                Document document = new Document(PageSize.LETTER, 25, 25, 25, 25);
                document.AddAuthor("Facturaxion");
                document.AddCreator("r3Take");
                document.AddCreationDate();
                //FileStream fs = new FileStream(pathPdf, FileMode.Create);
                //FileStream fs = new FileStream(pathPdf, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                Chunk cSaltoLinea = new Chunk("\n");
                Chunk cLineaSpace = new Chunk(cSaltoLinea + "________________________________________________________________________________________________________________________________________________________________________", new Font(Font.HELVETICA, 6, Font.BOLD));
                Chunk cLineaDiv = new Chunk(cSaltoLinea + "________________________________________________________________________________________________________________________________________________________________________" + cSaltoLinea, new Font(Font.HELVETICA, 6, Font.BOLD));
                Chunk cDataSpacer = new Chunk("      |      ", new Font(Font.HELVETICA, 6, Font.BOLD));

                BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
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

                htDatosCfdi.Add("nombreEmisor", "Razón Social " + electronicDocument.Data.Emisor.Nombre.Value);
                htDatosCfdi.Add("empresa", "Razón Social " + electronicDocument.Data.Emisor.Nombre.Value);

                htDatosCfdi.Add("rfcReceptor", electronicDocument.Data.Receptor.Rfc.Value);
                htDatosCfdi.Add("rfcCliente", electronicDocument.Data.Receptor.Rfc.Value);

                htDatosCfdi.Add("nombreReceptor", "Razón Social " + electronicDocument.Data.Receptor.Nombre.Value);
                htDatosCfdi.Add("cliente", "Razón Social " + electronicDocument.Data.Receptor.Nombre.Value);

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
                foreach (DataRow row in dtConfigFact.Rows)
                {
                    cb.SetFontAndSize(bf, Convert.ToInt32(row["fontSize"]));
                    cb.SetTextMatrix(Convert.ToSingle(row["posX"]), Convert.ToSingle(row["posY"]));
                    cb.ShowText(htDatosCfdi[row["objDesc"].ToString()].ToString());
                }
                /////////////////////////////////////////////////////////////
                DataTable dtGetOptional = new DataTable();

                dtGetOptional = dal.QueryDT("DS_FE", "SELECT idCFDI,campo1,campo2,campo3,campo4,campo5 FROM dbo.opcionalEncabezado WHERE idCFDI = @0", "F:I:" + htFacturaxion["idCfdi"], hc);

                if (dtGetOptional.Rows.Count > 0)
                {
                    cb.SetFontAndSize(bf, 7);
                    cb.SetTextMatrix(40, 606);
                    cb.ShowText(dtGetOptional.Rows[0]["campo1"].ToString());

                    cb.SetFontAndSize(bf, 7);
                    cb.SetTextMatrix(40, 597);
                    cb.ShowText(dtGetOptional.Rows[0]["campo2"].ToString());


                    cb.SetFontAndSize(bf, 7);
                    cb.SetTextMatrix(345, 615);
                    cb.ShowText(dtGetOptional.Rows[0]["campo3"].ToString());


                    cb.SetFontAndSize(bf, 7);
                    cb.SetTextMatrix(345, 606);
                    cb.ShowText(dtGetOptional.Rows[0]["campo4"].ToString());

                    cb.SetFontAndSize(bf, 7);
                    cb.SetTextMatrix(40, 597);
                    cb.ShowText(dtGetOptional.Rows[0]["campo5"].ToString());
                }
                ////////////////////////////////////////////////////////////////////

                #endregion

                cb.EndText();

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

                #region "Header"

                //Agregando Imagen de Encabezado de Página
                Image imgHeader = Image.GetInstance(dtConfigFact.Rows[0]["rutaTemplateHeader"].ToString());
                imgHeader.ScalePercent(47f);

                double posXH = Convert.ToDouble(dtConfigFact.Rows[0]["headerPosX"].ToString());
                double posYH = Convert.ToDouble(dtConfigFact.Rows[0]["headerPosY"].ToString());

                double PXH = posXH;
                double PYH = posYH;
                imgHeader.SetAbsolutePosition(Convert.ToSingle(PXH), Convert.ToSingle(PYH));
                document.Add(imgHeader);

                #endregion

                #region "Logotipo"

                //Agregando Imagen de Logotipo
                Image imgLogo = Image.GetInstance(dtConfigFact.Rows[0]["rutaLogo"].ToString());
                float imgLogoWidth = 100;
                float imgLogoHeight = 50;

                imgLogo.ScaleAbsolute(imgLogoWidth, imgLogoHeight);
                imgLogo.SetAbsolutePosition(Convert.ToSingle(dtConfigFact.Rows[0]["logoPosX"]), Convert.ToSingle(dtConfigFact.Rows[0]["logoPosY"]));
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
                String[] arrColWidthConceptos = dtConfigFact.Rows[0]["conceptosColWidth"].ToString().Split(new Char[] { ',' });

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

                document.Add(tableDesglose);

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

                //AZUL Font fontLbl = new Font(Font.HELVETICA, 6, Font.BOLD, new Color(43, 145, 175));
                Font fontLbl = new Font(Font.HELVETICA, 6, Font.BOLD, new Color(25, 71, 6));
                Font fontVal = new Font(Font.HELVETICA, 6, Font.NORMAL);

                string cantidadLetra = dtConfigFact.Rows[0]["cantidadLetra"].ToString();
                Chunk cCantidadLetraVal = new Chunk(cantidadLetra, new Font(Font.HELVETICA, 7, Font.BOLD));

                Chunk cTipoComprobanteLbl = new Chunk("Tipo de Comprobante: ", fontLbl);
                Chunk cTipoComprobanteVal = new Chunk(electronicDocument.Data.TipoComprobante.Value, fontVal);

                Chunk cFormaPagoLbl = new Chunk("Forma de Pago: ", fontLbl);
                Chunk cFormaPagoVal = new Chunk(electronicDocument.Data.FormaPago.Value, fontVal);

                Chunk cMetodoPagoLbl = new Chunk("Método de Pago: ", fontLbl);
                Chunk cMetodoPagoVal = new Chunk(electronicDocument.Data.MetodoPago.Value, fontVal);

                Chunk cMonedaLbl = new Chunk("Moneda: ", fontLbl);
                Chunk cMonedaVal = new Chunk(electronicDocument.Data.Moneda.Value, fontVal);

                if (electronicDocument.Data.TipoCambio.Value.Length == 0)
                    electronicDocument.Data.TipoCambio.Value = "1";

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

                string regimenes = "";

                for (int u = 0; u < electronicDocument.Data.Emisor.Regimenes.Count; u++)
                    regimenes += electronicDocument.Data.Emisor.Regimenes[u].Regimen.Value.ToString() + ",";

                Chunk cNoTarjetaLbl = new Chunk("No. Tarjeta: ", fontLbl);
                Chunk cNoTarjetaVal = new Chunk(electronicDocument.Data.NumeroCuentaPago.Value, fontVal);

                Chunk cExpedidoEnLbl = new Chunk("Expedido En: ", fontLbl);
                Chunk cExpedidoEnVal = new Chunk(electronicDocument.Data.LugarExpedicion.Value, fontVal);

                pFooter.Add(cCantidadLetraVal);
                pFooter.Add(cSaltoLinea);

                pFooter.Add(cTipoComprobanteLbl);
                pFooter.Add(cTipoComprobanteVal);
                pFooter.Add(cDataSpacer);
                pFooter.Add(cMonedaLbl);
                pFooter.Add(cMonedaVal);
                pFooter.Add(cDataSpacer);
                pFooter.Add(cTasaCambioLbl);
                pFooter.Add(cTasaCambioVal);
                pFooter.Add(cDataSpacer);

                if (htFacturaxion["noOrdenCompra"].ToString().Length > 0)
                {
                    Chunk cOrdenCompraLbl = new Chunk("Orden de Compra: ", fontLbl);
                    Chunk cOrdenCompraVal = new Chunk(htFacturaxion["noOrdenCompra"].ToString(), fontVal);
                    pFooter.Add(cOrdenCompraLbl);
                    pFooter.Add(cOrdenCompraVal);
                    pFooter.Add(cDataSpacer);
                }

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

                pFooter.Add(cLineaDiv);

                if (htFacturaxion["observaciones"].ToString().Length > 0)
                {
                    Chunk cObsLbl = new Chunk("Observaciones ", fontLbl);
                    Chunk cObsVal = new Chunk(htFacturaxion["observaciones"].ToString(), fontVal);
                    pFooter.Add(cObsLbl);
                    pFooter.Add(cSaltoLinea);
                    pFooter.Add(cObsVal);
                    pFooter.Add(cLineaDiv);
                }

                if (electronicDocument.Data.LugarExpedicion.Value.Length > 0)
                {
                    pFooter.Add(cExpedidoEnLbl);
                    pFooter.Add(cExpedidoEnVal);
                    pFooter.Add(cLineaDiv);
                }

                pFooter.Add(cCertificadoLbl);
                pFooter.Add(cCertificadoVal);
                pFooter.Add(cLineaDiv);

                if (timbrar)
                {
                    pFooter.Add(cCadenaOriginalPACLbl);
                    pFooter.Add(cSaltoLinea);
                    pFooter.Add(cCadenaOriginalPACVal);
                }
                else
                {
                    pFooter.Add(cCadenaOriginalLbl);
                    pFooter.Add(cSaltoLinea);
                    pFooter.Add(cCadenaOriginalVal);
                }

                pFooter.Add(cLineaDiv);
                pFooter.Add(cSelloDigitalLbl);
                pFooter.Add(cSaltoLinea);
                pFooter.Add(cSelloDigitalVal);
                pFooter.Add(cLineaSpace);

                document.Add(pFooter);

                #endregion

                #region "Añadimos Código Bidimensional"

                if (timbrar)
                {
                    Image imageQRCode = Image.GetInstance(bytesQRCode);
                    imageQRCode.Alignment = (Image.TEXTWRAP | Image.ALIGN_LEFT);
                    imageQRCode.ScaleToFit(90f, 90f);
                    imageQRCode.IndentationLeft = 9f;
                    imageQRCode.SpacingAfter = 9f;
                    imageQRCode.BorderColorTop = Color.WHITE;
                    document.Add(imageQRCode);
                    pFooter.Clear();

                    #region "Creación e Inserción de Chunks de Timbrado en Paragraph"

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

                    pFooter.Add(cSaltoLinea);
                    pFooter.Add(cFolioFiscalLbl);
                    pFooter.Add(cFolioFiscalVal);
                    pFooter.Add(cDataSpacer);
                    pFooter.Add(cCertificadoSatLbl);
                    pFooter.Add(cCertificadoSatVal);
                    pFooter.Add(cSaltoLinea);
                    pFooter.Add(cFechaTimbradoLbl);
                    pFooter.Add(cFechaTimbradoVal);
                    pFooter.Add(cSaltoLinea);
                    pFooter.Add(cSelloDigitalSatLbl);
                    pFooter.Add(cSaltoLinea);
                    pFooter.Add(cSelloDigitalSatVal);

                    #endregion

                    document.Add(pFooter);
                }

                #endregion

                #region "Añadimos leyenda de CFDI"

                Paragraph pLeyendaCfdi = new Paragraph();
                string leyenda;

                if (timbrar)
                {
                    leyenda = "Este documento es una representación impresa de un CFDI";
                }
                else
                {
                    leyenda = "Este documento es una representación impresa de un Comprobante Fiscal Digital";
                }

                Chunk cLeyendaCfdi = new Chunk(leyenda, new Font(Font.HELVETICA, 8, Font.BOLD | Font.ITALIC));
                pLeyendaCfdi.Add(cLeyendaCfdi);
                pLeyendaCfdi.SetLeading(1.6f, 1.6f);
                document.Add(pLeyendaCfdi);

                #endregion

                #region "Footer"

                //Agregando Imagen de Pie de Página
                Image imgFooter = Image.GetInstance(dtConfigFact.Rows[0]["rutaTemplateFooter"].ToString());
                float imgFooterWidth = document.PageSize.Width - 70;
                float imgFooterHeight = imgFooter.Height / (imgFooter.Width / imgFooterWidth);

                imgFooter.ScaleAbsolute(imgFooterWidth, imgFooterHeight);
                imgFooter.SetAbsolutePosition(Convert.ToSingle(dtConfigFact.Rows[0]["footerPosX"]), Convert.ToSingle(dtConfigFact.Rows[0]["footerPosY"]));
                document.Add(imgFooter);

                // Si el rol del usuario es gratuito añadimos el footer las imagenes de facturaxion y r3take

                if (idRol == 16)
                {
                    Image facturaxionImgFooter = Image.GetInstance(ConfigurationManager.AppSettings["logoFacturaxion"]);
                    Image r3TakeImgFooter = Image.GetInstance(ConfigurationManager.AppSettings["logor3Take"]);

                    facturaxionImgFooter.ScaleAbsolute(70, 25);
                    r3TakeImgFooter.ScaleAbsolute(70, 25);

                    facturaxionImgFooter.SetAbsolutePosition(25, 10);
                    r3TakeImgFooter.SetAbsolutePosition(600, 10);

                    document.Add(facturaxionImgFooter);
                    document.Add(r3TakeImgFooter);
                }

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
    }
}