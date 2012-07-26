#region "using"

using System;
using System.Globalization;
using System.IO;
using System.Collections;
using System.Configuration;
using System.Web;
using System.Text;
using System.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using r3TakeCore.Data;
using HyperSoft.ElectronicDocumentLibrary.Document;
using Data = HyperSoft.ElectronicDocumentLibrary.Complemento.TimbreFiscalDigital.Data;

#endregion

namespace wsRepresentacionImpresa.App_Code.com.Facturaxion.facturaEspecial
{
    public class corporativoGalgo
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

        private static CultureInfo _ci = new CultureInfo("es-mx");
        private static CultureInfo _ce = new CultureInfo("en-US");
        private static readonly string _rutaDocs = ConfigurationManager.AppSettings["rutaDocs"];
        private static readonly string _rutaDocsExt = ConfigurationManager.AppSettings["rutaDocsExterna"];

        #endregion

        #region "generarPDF"

        public static string generarPdf(Hashtable htFacturaxion, HttpContext hc)
        {
            try
            {
                DAL dal = new DAL();
                StringBuilder sbConfigFactParms = new StringBuilder();
                _ci.NumberFormat.CurrencyDecimalDigits = 2;

                ElectronicDocument electronicDocument = (ElectronicDocument)htFacturaxion["electronicDocument"];
                Data objTimbre = (Data)htFacturaxion["objTimbre"];
                bool timbrar = Convert.ToBoolean(htFacturaxion["timbrar"]);
                string pathPdf = htFacturaxion["rutaDocumentoPdf"].ToString();
                Int64 idCfdi = Convert.ToInt64(htFacturaxion["idCfdi"]);

                #region "Obtenemos los datos del CFDI y Campos Opcionales"

                StringBuilder sbOpcionalEncabezado = new StringBuilder();
                DataTable dtOpcEnc = new DataTable();
                StringBuilder sbOpcionalDetalle = new StringBuilder();
                DataTable dtOpcDet = new DataTable();

                sbOpcionalEncabezado.
                    Append("SELECT ").
                    Append("campo27 AS tipoIDoc, ").
                    Append("campo28 AS IDIOMA, ").
                    Append("campo29 AS tipoDoc, ").
                    Append("COALESCE(campo30, '') AS PEDIDO, ").
                    Append("COALESCE(campo31, '') AS REMISION, ").
                    Append("campo32 AS [NO-DE-PEDIDO-CLIENTE], ").
                    Append("campo33 AS VENCIMIENTO, ").
                    Append("campo34 AS PAGO, ").
                    Append("COALESCE(campo35, '')  AS TRANSPORTE, ").
                    Append("COALESCE(campo36, '')  AS [CARTA-PORTE], ").
                    Append("COALESCE(campo37, '')  AS CONTENEDOR, ").
                    Append("COALESCE(campo38, '')  AS INCOTERMS, ").
                    Append("COALESCE(campo39, '') AS [PAIS-ORIGEN], ").
                    Append("COALESCE(campo40, '') AS [PUERTO-ORIGEN], ").
                    Append("COALESCE(campo41, '') AS [PUERTO-DESTINO], ").
                    Append("campo42 AS [FACTURADO-A], ").
                    Append("campo77 AS nombreSucursal, ").

                    Append("campo8  AS [ENTREGADO-A-NOMBRE], ").
                    Append("campo11 AS [ENTREGADO-A-CALLE], ").
                    Append("campo13 AS [ENTREGADO-A-COLONIA], ").
                    Append("CASE campo15 WHEN '' THEN '' ").
                    Append("ELSE campo15 END AS [ENTREGADO-A-MUNIC], ").
                    Append("CASE campo16 WHEN '' THEN '' ").
                    Append("ELSE campo16 END AS [ENTREGADO-A-ESTADO], ").
                     Append("campo17 AS [ENTREGADO-A-PAIS], ").
                    Append("campo18 AS [ENTREGADO-A-CP], ").
                    Append("CASE campo43 WHEN '' THEN '' ").
                    Append("ELSE campo43 END AS [ENTREGADO-A-TEL1], ").
                    Append("CASE campo44 WHEN  '' THEN '' ").
                    Append("ELSE campo44 END AS [ENTREGADO-A-TEL2], ").
                    Append("CASE campo45 WHEN '' THEN '' ").
                    Append("ELSE campo45 END AS [ENTREGADO-FAX], ").
                    Append("campo46 AS [FECHA-PAGO-ANT], ").

                    Append("CASE WHEN LEN(campo47) > 0 ").
                    Append("THEN COALESCE(campo47, '0.00') ").
                    Append("ELSE '0.00' END AS ADUANA, ").
                    Append("CASE WHEN LEN(campo48) > 0 ").
                    Append("THEN COALESCE(campo48, '0.00') ").
                    Append("ELSE '0.00' END AS FLETE, ").
                    Append("CASE WHEN LEN(campo49) > 0 ").
                    Append("THEN COALESCE(campo49, '0.00') ").
                    Append("ELSE '0.00' END AS SEGURO, ").

                    Append("COALESCE(campo50, '') AS [TOTAL-TARIMAS], ").
                    Append("COALESCE(campo51, '') AS [NUMERO-CAJAS], ").
                    Append("COALESCE(campo52, '') AS [NUMERO-PIEZAS], ").
                    Append("COALESCE(campo53, '') AS [PESO-BRUTO], ").
                    Append("COALESCE(campo54, '') AS [PESO-NETO], ").
                    Append("COALESCE(campo55, '') AS COMENTARIOS,  ").
                    Append("COALESCE(campo56, '') AS BANCO, ").
                    Append("COALESCE(campo57, '') AS [CUENTA-CIE], ").
                    Append("COALESCE(campo58, '') AS REFERENCIA2, ").
                    Append("campo59 AS [CANTIDAD-LETRA], ").
                    Append("campo60 AS [NUM-FOLIO], ").
                    Append("campo61 AS IVA, ").
                    Append("COALESCE(campo62, '') AS [FORMA-PAGO], ").
                    Append("campo63 AS [TEL-EMISOR], ").
                    Append("campo64 AS [FAX-EMISOR], ").
                    Append("campo65 AS [LUGAR-PAGO], ").
                    Append("campo66 + ', ' + campo67 AS [EXPEDIDO-EN], ").
                    Append("campo68 AS [TERMINOS-CREDITO], ").
                    Append("campo69 AS [ENTREGADO-CODE], ").
                    Append("campo70 AS [TEL-EXPEDIDO], ").
                    Append("campo71 AS [FAX-EXPEDIDO], ").
                    Append("campo72 AS [TEL-FACTURADO], ").
                    Append("campo73 AS [FAX-FACTURADO], ").
                    Append("COALESCE(campo74, '0.00') AS [SUMA], ").
                    Append("COALESCE(campo75, '0.00') AS [SUB-TOTAL], ").
                    Append("campo76 AS [TEL2-FACTURADO], ").
                    Append("campo78 AS [BANCO-DEPOSITO] ").//
                    Append("FROM opcionalEncabezado ").
                    Append("WHERE idCFDI = @0  AND ST = 1 ");

                sbOpcionalDetalle.
                    Append("SELECT ").
                    Append("COALESCE(campo10, '0.00') AS PS, ").
                    Append("COALESCE(campo11, '') AS DESCRIP ").
                    Append("FROM opcionalDetalle ").
                    Append("WHERE idCFDI = @0 ");

                dtOpcEnc = dal.QueryDT("DS_FE", sbOpcionalEncabezado.ToString(), "F:I:" + idCfdi, hc);
                dtOpcDet = dal.QueryDT("DS_FE", sbOpcionalDetalle.ToString(), "F:I:" + idCfdi, hc);

                if (dtOpcDet.Rows.Count == 0)
                {
                    for (int i = 1; i <= electronicDocument.Data.Conceptos.Count; i++)
                    {
                        dtOpcDet.Rows.Add("", "0.00");
                    }
                }

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

                #endregion

                #region "Cargamos Etiquetas de las Tablas"

                Hashtable htDatosEtiquetas = new Hashtable();
                string tipoIDoc = dtOpcEnc.Rows[0]["tipoIDoc"].ToString();
                string IDIOMA = dtOpcEnc.Rows[0]["IDIOMA"].ToString();
                string rfcEmisor = electronicDocument.Data.Emisor.Rfc.Value;
                string tipoDoc = string.Empty;
                string CondicionesPago = string.Empty;
                string nombreSucursal = string.Empty;

                if (tipoIDoc.StartsWith("INVOIC"))
                {
                    if (rfcEmisor == "IHG8212239Z4")
                    {
                        tipoDoc = dtOpcEnc.Rows[0]["tipoDoc"].ToString().Substring(0, 5);
                    }
                    else
                    {
                        tipoDoc = dtOpcEnc.Rows[0]["tipoDoc"].ToString().Substring(0, 4);
                    }
                    CondicionesPago = dtOpcEnc.Rows[0]["TERMINOS-CREDITO"].ToString();
                    nombreSucursal = dtOpcEnc.Rows[0]["nombreSucursal"].ToString().ToUpper();
                }
                else if (tipoIDoc.StartsWith("FIDCCP02"))
                {
                    tipoDoc = dtOpcEnc.Rows[0]["tipoDoc"].ToString().Substring(0, 2);
                    CondicionesPago = electronicDocument.Data.CondicionesPago.Value;
                    nombreSucursal = electronicDocument.Data.Emisor.Nombre.Value;
                }

                if (IDIOMA == "S")
                {
                    htDatosEtiquetas.Add("REFERENCIA", "REFERENCIA");
                    htDatosEtiquetas.Add("FECHA-HORA-EMISION", "FECHA Y HORA DE EMISION");
                    htDatosEtiquetas.Add("PEDIDO", "PEDIDO:");
                    htDatosEtiquetas.Add("REMISION", "REMISION:");
                    htDatosEtiquetas.Add("NO-DE-PEDIDO-CLIENTE", "NO. DE PEDIDO DEL CLIENTE");
                    htDatosEtiquetas.Add("VENCIMIENTO", "VENCIMIENTO");

                    switch (tipoDoc)
                    {
                        case "00003":
                        case "1811":
                        case "21":
                            htDatosEtiquetas.Add("FACTURA", "FACTURA");
                            break;
                        case "01906":
                        case "1812":
                        case "22":
                            htDatosEtiquetas.Add("FACTURA", "NOTA DE CREDITO");
                            break;
                        case "01902":
                        case "1813":
                        case "23":
                            htDatosEtiquetas.Add("FACTURA", "NOTA DE CARGO");
                            break;
                        default:
                            htDatosEtiquetas.Add("FACTURA", "FACTURA");
                            break;
                    }

                    htDatosEtiquetas.Add("EXPEDIDO-EN", "EXPEDIDO EN:");
                    htDatosEtiquetas.Add("TRANSPORTE", "TRANSPORTE");
                    htDatosEtiquetas.Add("CARTA-PORTE", "CARTA PORTE");
                    htDatosEtiquetas.Add("CONTENEDOR", "CONTENEDOR");
                    htDatosEtiquetas.Add("INCOTERMS", "INCOTERM");
                    htDatosEtiquetas.Add("PAIS-ORIGEN", "PAIS DE ORIGEN");
                    htDatosEtiquetas.Add("TERMINOS-CREDITO", "PLAZO DE CREDITO");
                    htDatosEtiquetas.Add("PUERTO-ORIGEN", "PUERTO ORIGEN");
                    htDatosEtiquetas.Add("PUERTO-DESTINO", "PUERTO DESTINO");
                    htDatosEtiquetas.Add("FACTURADO-A", "FACTURADO A:");
                    htDatosEtiquetas.Add("ENTREGADO-A", "ENTREGADO A:");
                    htDatosEtiquetas.Add("CLAVE", "CLAVE");
                    htDatosEtiquetas.Add("DESCRIPCION", "DESCRIPCION");
                    htDatosEtiquetas.Add("PS", "PS");

                    htDatosEtiquetas.Add("UNIDAD-MEDIDA-EMPAQUE", "UNIDAD DE MEDIDA EMPAQUE");
                    htDatosEtiquetas.Add("CANTIDAD", "CANTIDAD");
                    htDatosEtiquetas.Add("UNIDAD-MEDIDA", "UNIDAD DE MEDIDA");
                    htDatosEtiquetas.Add("PRECIO-UNITARIO", "PRECIO UNITARIO");
                    htDatosEtiquetas.Add("SUMA", "SUMA");
                    htDatosEtiquetas.Add("DESCUENTO", "MENOS DESCUENTO");
                    htDatosEtiquetas.Add("IMPORTE", "IMPORTE");
                    htDatosEtiquetas.Add("SUB-TOTAL", "SUB TOTAL");
                    htDatosEtiquetas.Add("ADUANA", "ADUANA");
                    htDatosEtiquetas.Add("FLETE", "FLETE");
                    htDatosEtiquetas.Add("SEGURO", "SEGURO");
                    htDatosEtiquetas.Add("IVA", "IVA: " + dtOpcEnc.Rows[0]["IVA"] + "% ");
                    htDatosEtiquetas.Add("TOTAL", "TOTAL");

                    if (electronicDocument.Data.Receptor.Domicilio.Pais.Value.TrimStart().TrimEnd() == "México" && htDatosEtiquetas["FACTURA"].ToString() == "FACTURA")
                    {
                        htDatosEtiquetas.Add("NOTA1", "A FALTA DE PAGO DE LA CANTIDAD DE ");
                        htDatosEtiquetas.Add("NOTA2", electronicDocument.Data.Total.Value.ToString("C", _ci) + " " + dtOpcEnc.Rows[0]["CANTIDAD-LETRA"]);
                        htDatosEtiquetas.Add("NOTA3", ", A LA FECHA DE VENCIMIENTO, ESTA FACTURA GENERARA INTERES MORATORIO A RAZON DEL 3% MENSUAL SOBRE EL MONTO ANTES INDICADO");
                    }
                    else
                    {
                        htDatosEtiquetas.Add("NOTA1", "");
                        htDatosEtiquetas.Add("NOTA2", "");
                        htDatosEtiquetas.Add("NOTA3", "");
                    }

                    htDatosEtiquetas.Add("CANTIDAD-LETRA", "CANTIDAD CON LETRA: ");

                    htDatosEtiquetas.Add("TOTAL-TARIMAS", "TOTAL TARIMAS");
                    htDatosEtiquetas.Add("NUMERO-CAJAS", "NUMERO DE CAJAS");
                    htDatosEtiquetas.Add("NUMERO-PIEZAS", "NUMERO DE PIEZAS");
                    htDatosEtiquetas.Add("PESO-BRUTO", "PESO BRUTO");
                    htDatosEtiquetas.Add("PESO-NETO", "PESO NETO");
                    htDatosEtiquetas.Add("COMENTARIOS", "COMENTARIOS: ");
                    htDatosEtiquetas.Add("LUGAR-PAGO", "LUGAR DE PAGO: " + dtOpcEnc.Rows[0]["LUGAR-PAGO"]);

                    htDatosEtiquetas.Add("VENCIMIENTO-PAGO", "AL VENCIMIENTO AGRADECEREMOS SU PAGO EN:");
                    htDatosEtiquetas.Add("FORMA-PAGO", "FORMA DE PAGO:");
                    htDatosEtiquetas.Add("BANCO", "BANCO:");
                    htDatosEtiquetas.Add("CUENTA-CIE", "CUENTA/CIE");
                    htDatosEtiquetas.Add("REFERENCIA2", "REFERENCIA:");

                    htDatosEtiquetas.Add("SELLO-DIGITA-CFDI", "SELLO DIGITAL DEL CFDI");
                    htDatosEtiquetas.Add("TIMBRE-FISCAL", "TIMBRE FISCAL DIGITAL");
                    htDatosEtiquetas.Add("FOLIO-FISCAL", "FOLIO FISCAL:");
                    htDatosEtiquetas.Add("FECHA-CERTIFICACION", "FECHA Y HORA DE CERTIFICACION:");
                    htDatosEtiquetas.Add("SERIE-CERTIFICADO-SAT", "No. DE SERIE DEL CERTIFICADO DEL SAT:");
                    htDatosEtiquetas.Add("SERIE-CERTIFICADO-EMISOR", "No. DE SERIE DEL CERTIFICADO DEL EMISOR:");
                    htDatosEtiquetas.Add("CADENA-ORIGINAL", "CADENA ORIGINAL DEL COMPLEMENTO DE CERTIFICACIÓN DIGITAL DEL SAT");
                    htDatosEtiquetas.Add("SELLO-DIGITA-SAT", "SELLO DIGITAL DEL SAT");
                    htDatosEtiquetas.Add("LEYENDA", "\"ESTE DOCUMENTO ES UNA REPRESENTACION IMPRESA DE UN COMPROBANTE FISCAL DIGITAL A TRAVES DE INTERNET\"");
                    htDatosEtiquetas.Add("PAGINA", "PÁGINA");
                    htDatosEtiquetas.Add("CP", "\nC.P. ");
                    htDatosEtiquetas.Add("CP1", "C.P. ");
                    htDatosEtiquetas.Add("TEL", "TEL. ");
                    htDatosEtiquetas.Add("FAX", "FAX. ");
                    htDatosEtiquetas.Add("REGIMEN-EMISOR", "REGIMEN FISCAL APLICABLE: ");
                    htDatosEtiquetas.Add("REGIMEN-DATO", "REGIMEN GENERAL DE LEY PERSONAS MORALES");
                    htDatosEtiquetas.Add("METODO-PAGO", "METODO DE PAGO");
                    htDatosEtiquetas.Add("CUENTA", "CUENTA");
                }
                else
                {
                    htDatosEtiquetas.Add("REFERENCIA", "REFERENCE");
                    htDatosEtiquetas.Add("FECHA-HORA-EMISION", "DATE & TIME EXPEDITION");
                    htDatosEtiquetas.Add("PEDIDO", "ORDER:");
                    htDatosEtiquetas.Add("REMISION", "RELEASED:");
                    htDatosEtiquetas.Add("NO-DE-PEDIDO-CLIENTE", "CUSTOMER ORDER");
                    htDatosEtiquetas.Add("VENCIMIENTO", "EXPIRATION DATE");

                    switch (tipoDoc)
                    {
                        case "00003":
                        case "1811":
                        case "21":
                            htDatosEtiquetas.Add("FACTURA", "INVOICE");
                            break;
                        case "01906":
                        case "1812":
                        case "22":
                            htDatosEtiquetas.Add("FACTURA", "CREDIT MEMO");
                            break;
                        case "01902":
                        case "1813":
                        case "23":
                            htDatosEtiquetas.Add("FACTURA", "DEBIT MEMO");
                            break;
                        default:
                            htDatosEtiquetas.Add("FACTURA", "INVOICE");
                            break;
                    }

                    htDatosEtiquetas.Add("EXPEDIDO-EN", "ISSUED BY:");
                    htDatosEtiquetas.Add("TRANSPORTE", "CARRIER");
                    htDatosEtiquetas.Add("CARTA-PORTE", "BILL OF LADING");
                    htDatosEtiquetas.Add("CONTENEDOR", "CONTAINER");
                    htDatosEtiquetas.Add("INCOTERMS", "INCOTERM");
                    htDatosEtiquetas.Add("PAIS-ORIGEN", "COUNTRY OF ORIGIN");
                    htDatosEtiquetas.Add("TERMINOS-CREDITO", "CREDIT TERM");
                    htDatosEtiquetas.Add("PUERTO-ORIGEN", "ORIGIN PORT");
                    htDatosEtiquetas.Add("PUERTO-DESTINO", "DESTINATION PORT");
                    htDatosEtiquetas.Add("FACTURADO-A", "BILL TO:");
                    htDatosEtiquetas.Add("ENTREGADO-A", "SHIP TO:");
                    htDatosEtiquetas.Add("CLAVE", "CODE");
                    htDatosEtiquetas.Add("DESCRIPCION", "DESCRIPTION");
                    htDatosEtiquetas.Add("PS", "PS");
                    htDatosEtiquetas.Add("UNIDAD-MEDIDA-EMPAQUE", "PACKAGE UNIT");
                    htDatosEtiquetas.Add("CANTIDAD", "QUANTITY");
                    htDatosEtiquetas.Add("UNIDAD-MEDIDA", "UNIT");
                    htDatosEtiquetas.Add("PRECIO-UNITARIO", "UNIT PRICE");
                    htDatosEtiquetas.Add("IMPORTE", "AMOUNT");
                    htDatosEtiquetas.Add("SUMA", "SUM");
                    htDatosEtiquetas.Add("DESCUENTO", "DISCOUNT");
                    htDatosEtiquetas.Add("SUB-TOTAL", "SUB TOTAL");
                    htDatosEtiquetas.Add("ADUANA", "CUSTOMS");
                    htDatosEtiquetas.Add("FLETE", "FREIGHT");
                    htDatosEtiquetas.Add("SEGURO", "INSURANCE");
                    htDatosEtiquetas.Add("IVA", "TAX: " + dtOpcEnc.Rows[0]["IVA"] + "% ");
                    htDatosEtiquetas.Add("TOTAL", "TOTAL");

                    htDatosEtiquetas.Add("NOTA1", "");
                    htDatosEtiquetas.Add("NOTA3", "");
                    htDatosEtiquetas.Add("NOTA2", "");

                    htDatosEtiquetas.Add("CANTIDAD-LETRA", "WRITTEN AMOUNT: ");

                    htDatosEtiquetas.Add("TOTAL-TARIMAS", "TOTAL PALLETS");
                    htDatosEtiquetas.Add("NUMERO-CAJAS", "TOTAL BOXES");
                    htDatosEtiquetas.Add("NUMERO-PIEZAS", "TOTAL PIECES");
                    htDatosEtiquetas.Add("PESO-BRUTO", "GROSS WEIGHT");
                    htDatosEtiquetas.Add("PESO-NETO", "NET WEIGHT");

                    htDatosEtiquetas.Add("COMENTARIOS", "COMMENTS: ");
                    htDatosEtiquetas.Add("LUGAR-PAGO", "PLACE OF PAYMENT: " + dtOpcEnc.Rows[0]["LUGAR-PAGO"]);

                    htDatosEtiquetas.Add("VENCIMIENTO-PAGO", "WHEN EXPIRES WE THANK YOU FOR YOUR PAYMENT AT:");
                    htDatosEtiquetas.Add("FORMA-PAGO", "PAYMENT FORM:");
                    htDatosEtiquetas.Add("BANCO", "BANK");
                    htDatosEtiquetas.Add("CUENTA-CIE", "ACCOUNT");
                    htDatosEtiquetas.Add("REFERENCIA2", "REFERENCE:");

                    htDatosEtiquetas.Add("SELLO-DIGITA-CFDI", "DIGITAL STAMP OF THE CFDI");
                    htDatosEtiquetas.Add("TIMBRE-FISCAL", "DIGITAL REVENUE STAMP:");
                    htDatosEtiquetas.Add("FOLIO-FISCAL", "FISCAL FOLIO:");
                    htDatosEtiquetas.Add("FECHA-CERTIFICACION", "CERTIFICATION DATE AND TIME:");
                    htDatosEtiquetas.Add("SERIE-CERTIFICADO-SAT", "SERIAL NUMBER OF SAT CERTIFICATE:");
                    htDatosEtiquetas.Add("SERIE-CERTIFICADO-EMISOR", "SERIAL NUMBER OF TRANSMITTER CERTIFICATE:");
                    htDatosEtiquetas.Add("CADENA-ORIGINAL", "ORIGINAL CHAIN COMPLEMENT OF THE DIGITAL SAT CERTIFICATION");
                    htDatosEtiquetas.Add("SELLO-DIGITA-SAT", "DIGITAL STAMP OF SAT");
                    htDatosEtiquetas.Add("LEYENDA", "\"THIS DOCUMENT IS A PRINTED REPRESENTATION  OF FISCAL DIGITAL CERTIFICATE BY INTERNET\"");
                    htDatosEtiquetas.Add("PAGINA", "PAGE");
                    htDatosEtiquetas.Add("CP", "\nZP. ");
                    htDatosEtiquetas.Add("CP1", "ZP. ");
                    htDatosEtiquetas.Add("TEL", "PH. ");
                    htDatosEtiquetas.Add("FAX", "FAX. ");
                    htDatosEtiquetas.Add("REGIMEN-EMISOR", "FISCAL CLASSIFICATION APPLIED:");
                    htDatosEtiquetas.Add("REGIMEN-DATO", "GENERAL TAX LAW FOR CORPORATIONS");
                    htDatosEtiquetas.Add("METODO-PAGO", "PAYMENT METHOD");
                    htDatosEtiquetas.Add("CUENTA", "ACCOUNT");

                }
                #endregion

                #region "Extraemos los datos del CFDI"

                Hashtable htDatosCfdi = new Hashtable();
                htDatosCfdi.Add("nombreEmisor", electronicDocument.Data.Emisor.Nombre.Value);
                htDatosCfdi.Add("rfcEmisor", rfcEmisor);
                htDatosCfdi.Add("nombreReceptor", electronicDocument.Data.Receptor.Nombre.Value);
                htDatosCfdi.Add("rfcReceptor", electronicDocument.Data.Receptor.Rfc.Value);
                htDatosCfdi.Add("sucursal", dtOpcEnc.Rows[0]["nombreSucursal"].ToString().ToUpper());
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
                    sbDirEmisor3.Append(htDatosEtiquetas["C.P. "]).Append(electronicDocument.Data.Emisor.Domicilio.CodigoPostal.Value).Append(", ");
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
                    sbDirExpedido3.Append(htDatosEtiquetas["C.P. "]).Append(electronicDocument.Data.Emisor.ExpedidoEn.CodigoPostal.Value).Append(", ");
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
                    sbDirReceptor3.Append(htDatosEtiquetas["C.P. "]).Append(electronicDocument.Data.Receptor.Domicilio.CodigoPostal.Value).Append(", ");
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
                htDatosCfdi.Add("paginaWeb", "www.galgo.com.mx");

                #endregion

                #region "Creamos el Objeto Documento y Tipos de Letra"

                Document document = new Document(PageSize.LETTER, 15, 15, 15, 40);
                document.AddAuthor("Facturaxion");
                document.AddCreator("r3Take");
                document.AddCreationDate();

                FileStream fs = new FileStream(pathPdf, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                pdfPageEventHandlerGalgo pageEventHandler = new pdfPageEventHandlerGalgo();
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.SetFullCompression();
                writer.ViewerPreferences = PdfWriter.PageModeUseNone;
                writer.PageEvent = pageEventHandler;
                writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

                Chunk dSaltoLinea = new Chunk("\n\n ");
                Color azul;
                Image imgISO;
                Image imgESR;
                Image imgPremio;
                int R;
                int G;
                int B;
                string pathIMGISO = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\images\galgo_ISO-9001.jpg";
                string pathIMGESR = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\images\ESR.jpg";
                string pathIMGPRM = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\images\galgo_PNE-2009.jpg";
                string pathIMGBNK = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\images\blanck.jpg";

                switch (rfcEmisor)
                {
                    case "IHG8212239Z4":
                        R = 1;
                        G = 43;
                        B = 127;
                        imgISO = Image.GetInstance(pathIMGISO);
                        imgESR = Image.GetInstance(pathIMGESR);
                        imgPremio = Image.GetInstance(pathIMGPRM);
                        break;
                    case "CGA821223QA2":
                        R = 1;
                        G = 43;
                        B = 127;
                        imgISO = Image.GetInstance(pathIMGISO);
                        imgESR = Image.GetInstance(pathIMGBNK);
                        imgPremio = Image.GetInstance(pathIMGPRM);
                        break;
                    case "CGA770614UX5":
                    case "SUA970910HP4":
                        R = 1;
                        G = 43;
                        B = 127;
                        imgISO = Image.GetInstance(pathIMGBNK);
                        imgESR = Image.GetInstance(pathIMGBNK);
                        imgPremio = Image.GetInstance(pathIMGBNK);
                        break;
                    case "HOM861209TN5":
                        R = 235;
                        G = 2;
                        B = 55;
                        imgISO = Image.GetInstance(pathIMGBNK);
                        imgESR = Image.GetInstance(pathIMGBNK);
                        imgPremio = Image.GetInstance(pathIMGBNK);
                        break;
                    case "MAG791115965":
                        R = 223;
                        G = 117;
                        B = 23;
                        imgISO = Image.GetInstance(pathIMGISO);
                        imgESR = Image.GetInstance(pathIMGBNK);
                        imgPremio = Image.GetInstance(pathIMGBNK);
                        break;
                    case "EOM910430QL4":
                        R = 235;
                        G = 2;
                        B = 55;
                        imgISO = Image.GetInstance(pathIMGISO);
                        imgESR = Image.GetInstance(pathIMGBNK);
                        imgPremio = Image.GetInstance(pathIMGBNK);
                        break;

                    default:
                        R = 0;
                        G = 44;
                        B = 122;
                        imgISO = Image.GetInstance(pathIMGBNK);
                        imgESR = Image.GetInstance(pathIMGBNK);
                        imgPremio = Image.GetInstance(pathIMGBNK);
                        break;
                }

                azul = new Color(R, G, B);
                Color blanco = new Color(255, 255, 255);
                Color Link = new Color(7, 73, 208);
                Color gris = new Color(236, 236, 236);
                Color grisOX = new Color(220, 215, 220);//233, 230, 233
                Color rojo = new Color(230, 7, 7);

                BaseFont EM = BaseFont.CreateFont(@"C:\Windows\Fonts\VERDANA.TTF", BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);

                Font f5 = new Font(EM, 5);
                Font f5B = new Font(EM, 5, Font.BOLD);
                Font f5BBI = new Font(EM, 5, Font.BOLDITALIC);
                Font f6 = new Font(EM, 5);
                Font f6B = new Font(EM, 5, Font.BOLD);
                Font f6L = new Font(EM, 5, Font.BOLD, Link);
                Font titulo = new Font(EM, 5, Font.BOLD, blanco);
                Font folio = new Font(EM, 6, Font.BOLD, rojo);
                PdfPCell cell;
                Paragraph par;

                #endregion

                #region "Construimos el Documento"

                #region "Construimos el Encabezado"

                PdfPTable encabezado = new PdfPTable(3);
                encabezado.WidthPercentage = 100;
                encabezado.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                encabezado.SetWidths(new int[3] { 30, 10, 60 });
                encabezado.DefaultCell.Border = 0;
                encabezado.LockedWidth = true;

                //Agregando Imagen de Logotipo
                string pathLogo = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\" + electronicDocument.Data.Emisor.Rfc.Value + @"\logo.jpg";
                string pathCedula = @"C:\Inetpub\repositorioFacturaxion\imagesFacturaEspecial\" + electronicDocument.Data.Emisor.Rfc.Value + @"\cedula.jpg";
                Image imgLogo = Image.GetInstance(pathLogo);
                imgLogo.ScalePercent(47f);

                //Agregando Imagen de Cédula de Identificación Fiscal
                Image imgCedula = Image.GetInstance(pathCedula);
                imgCedula.ScalePercent(47f);
                imgISO.ScalePercent(47f);
                imgESR.ScalePercent(47f);
                imgPremio.ScalePercent(47f);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(8f, 9f);
                par.Add(new Chunk(imgLogo, 0, 0));
                par.Add(new Chunk("\n" + htDatosCfdi["nombreEmisor"].ToString().ToUpper(), f6B));
                par.Add(new Chunk("\nRFC: " + htDatosCfdi["rfcEmisor"].ToString().ToUpper(), f6B));
                par.Add(new Chunk("\n" + htDatosCfdi["direccionEmisor1"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htDatosCfdi["direccionEmisor2"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htDatosCfdi["direccionEmisor3"].ToString().ToUpper(), f6));
                par.Add(new Chunk("\n" + htDatosEtiquetas["TEL"] + dtOpcEnc.Rows[0]["TEL-EMISOR"] + " ", f6));
                par.Add(new Chunk("\n" + htDatosEtiquetas["FAX"] + dtOpcEnc.Rows[0]["FAX-EMISOR"], f6));
                par.Add(new Chunk("\nPágina Web: ", f6));
                par.Add(new Chunk(htDatosCfdi["paginaWeb"].ToString(), f6L));
                cell = new PdfPCell(par);
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                encabezado.AddCell(cell);

                cell = new PdfPCell(imgCedula);
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                encabezado.AddCell(cell);

                #region "Detalle de la Celda de Encabezado"

                PdfPTable encabDet = new PdfPTable(3);
                encabDet.WidthPercentage = 100;

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["REFERENCIA"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["FECHA-HORA-EMISION"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["FACTURA"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = (float).5;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                #region "Detalle de la Celda Referencia"

                PdfPTable encabDet1 = new PdfPTable(2);
                encabDet1.WidthPercentage = 100;

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["PEDIDO"].ToString(), f6B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encabDet1.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["PEDIDO"].ToString(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encabDet1.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["REMISION"].ToString(), f6B));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                encabDet1.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["REMISION"].ToString(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                encabDet1.AddCell(cell);

                #endregion

                cell = new PdfPCell(encabDet1);
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);
                string[] fechaCFDI = Convert.ToDateTime(htDatosCfdi["fechaCfdi"].ToString()).GetDateTimeFormats('s');

                cell = new PdfPCell(new Phrase(fechaCFDI[0], f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["NUM-FOLIO"].ToString(), folio));
                cell.VerticalAlignment = Element.ALIGN_BOTTOM;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = (float).5;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["NO-DE-PEDIDO-CLIENTE"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["VENCIMIENTO"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(new Phrase("", f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = (float).5;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["NO-DE-PEDIDO-CLIENTE"].ToString(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                int anio = Convert.ToInt32(dtOpcEnc.Rows[0]["VENCIMIENTO"].ToString().Substring(0, 4));
                int month = Convert.ToInt32(dtOpcEnc.Rows[0]["VENCIMIENTO"].ToString().Substring(4, 2));
                int day = Convert.ToInt32(dtOpcEnc.Rows[0]["VENCIMIENTO"].ToString().Substring(6, 2));
                DateTime fechaVencimiento = new DateTime(anio, month, day);

                string vencimiento = string.Empty;

                if (IDIOMA == "S")
                {
                    vencimiento = day + "-" + fechaVencimiento.ToString("Y", _ci).Replace(", ", "-");
                }
                else
                {
                    vencimiento = day + "-" + fechaVencimiento.ToString("Y", _ce).Replace(", ", "-");
                }

                cell = new PdfPCell(new Phrase(vencimiento.ToUpper(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(new Phrase("", f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = (float).5;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(imgISO);
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(imgESR);
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                cell = new PdfPCell(imgPremio);
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                encabDet.AddCell(cell);

                #endregion

                cell = new PdfPCell(encabDet);
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                encabezado.AddCell(cell);

                #endregion

                #region "Construimos Tabla de Detalles Especiales"

                PdfPTable detalle = new PdfPTable(3);
                detalle.WidthPercentage = 100;
                detalle.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                detalle.SetWidths(new int[3] { 60, 20, 20 });
                detalle.DefaultCell.Border = 0;
                detalle.LockedWidth = true;

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["PAGO"].ToString().ToUpper(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderColor = grisOX;
                cell.BorderWidthTop = (float).5;
                cell.BorderWidthRight = (float).5;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthBottom = (float).5;
                detalle.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["TRANSPORTE"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                detalle.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["CARTA-PORTE"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                detalle.AddCell(cell);


                cell = new PdfPCell(new Phrase(htDatosEtiquetas["LUGAR-PAGO"].ToString().ToUpper(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BorderColor = grisOX;
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = (float).5;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthBottom = 0;
                detalle.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["TRANSPORTE"].ToString(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detalle.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["CARTA-PORTE"].ToString(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detalle.AddCell(cell);

                DateTime fecha = Convert.ToDateTime(htDatosCfdi["fechaCfdi"]);
                string fechaExpedido = string.Empty;

                if (IDIOMA == "S")
                {
                    fechaExpedido = fecha.Day + "-" + fecha.ToString("Y", _ci).Replace(", ", "-");
                }
                else
                {
                    fechaExpedido = fecha.Day + "-" + fecha.ToString("Y", _ce).Replace(", ", "-");
                }

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["EXPEDIDO-EN"] + " " +
                    electronicDocument.Data.Emisor.ExpedidoEn.Municipio.Value.ToString().ToUpper() + " " +
                    electronicDocument.Data.Emisor.ExpedidoEn.Estado.Value.ToString().ToUpper() + " " + fechaExpedido.ToUpper(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = (float).5;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detalle.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["CONTENEDOR"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detalle.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["INCOTERMS"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detalle.AddCell(cell);

                cell = new PdfPCell(new Phrase(nombreSucursal + "\n" +
                                               htDatosCfdi["direccionExpedido1"].ToString().ToUpper() + "\n" +
                                               htDatosCfdi["direccionExpedido2"].ToString().ToUpper() + "\n" +
                                               htDatosCfdi["direccionExpedido3"].ToString().ToUpper() + "\n" +
                                               htDatosEtiquetas["TEL"] + dtOpcEnc.Rows[0]["TEL-EXPEDIDO"] + "\n" +
                                               htDatosEtiquetas["FAX"] + dtOpcEnc.Rows[0]["FAX-EXPEDIDO"], f6));
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = (float).5;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detalle.AddCell(cell);

                #region "Detalle de la Celda 2 Detalles Especiales"

                PdfPTable detEsp1 = new PdfPTable(2);
                encabDet1.WidthPercentage = 100;

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["CONTENEDOR"].ToString(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detEsp1.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["INCOTERMS"].ToString(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detEsp1.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["PAIS-ORIGEN"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detEsp1.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["TERMINOS-CREDITO"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detEsp1.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["PAIS-ORIGEN"].ToString(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detEsp1.AddCell(cell);

                cell = new PdfPCell(new Phrase(CondicionesPago, f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detEsp1.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["PUERTO-ORIGEN"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detEsp1.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["PUERTO-DESTINO"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detEsp1.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["PUERTO-ORIGEN"].ToString(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detEsp1.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtOpcEnc.Rows[0]["PUERTO-DESTINO"].ToString(), f6));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detEsp1.AddCell(cell);

                #endregion

                cell = new PdfPCell(detEsp1);
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = (float).5;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                cell.Colspan = 2;
                detalle.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["FACTURADO-A"] + " " + dtOpcEnc.Rows[0]["FACTURADO-A"], titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                detalle.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["ENTREGADO-A"] + " " + dtOpcEnc.Rows[0]["ENTREGADO-CODE"], titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = gris;
                cell.Colspan = 2;
                detalle.AddCell(cell);

                StringBuilder direccionReceptor = new StringBuilder();
                direccionReceptor.
                    Append("\n").
                    Append(htDatosCfdi["direccionReceptor1"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["direccionReceptor2"].ToString().ToUpper()).Append("\n").
                    Append(htDatosCfdi["direccionReceptor3"].ToString().ToUpper()).Append("\n").
                    Append(htDatosEtiquetas["TEL"].ToString()).
                    //Append(dtOpcEnc.Rows[0]["TEL-FACTURADO"]).Append("\n").
                    Append(dtOpcEnc.Rows[0]["TEL-FACTURADO"]).Append(", ").
                    Append(dtOpcEnc.Rows[0]["TEL2-FACTURADO"]).Append("\n").
                    Append(htDatosEtiquetas["FAX"].ToString()).
                    Append(dtOpcEnc.Rows[0]["FAX-FACTURADO"].ToString());

                par = new Paragraph();
                par.SetLeading(7f, 1f);
                par.Add(new Chunk(electronicDocument.Data.Receptor.Nombre.Value + "\nRFC: " + electronicDocument.Data.Receptor.Rfc.Value + direccionReceptor.ToString(), f6));
                cell = new PdfPCell(par);
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                detalle.AddCell(cell);

                if (tipoIDoc.StartsWith("INVOIC"))
                {
                    string nombreEntregado = dtOpcEnc.Rows[0]["ENTREGADO-A-NOMBRE"].ToString().ToUpper();
                    string calleEntregado = dtOpcEnc.Rows[0]["ENTREGADO-A-CALLE"].ToString().ToUpper();
                    string coloniaEntregado = dtOpcEnc.Rows[0]["ENTREGADO-A-COLONIA"].ToString().ToUpper();
                    string municEntregado = dtOpcEnc.Rows[0]["ENTREGADO-A-MUNIC"].ToString().ToUpper();
                    string estadoEntregado = dtOpcEnc.Rows[0]["ENTREGADO-A-ESTADO"].ToString().ToUpper();
                    string paisEntregado = dtOpcEnc.Rows[0]["ENTREGADO-A-PAIS"].ToString().ToUpper();
                    string cpEntregado = dtOpcEnc.Rows[0]["ENTREGADO-A-CP"].ToString().ToUpper();
                    string tel1Entregado = dtOpcEnc.Rows[0]["ENTREGADO-A-TEL1"].ToString().ToUpper();
                    string tel2Entregado = dtOpcEnc.Rows[0]["ENTREGADO-A-TEL2"].ToString().ToUpper();
                    string faxEntregado = dtOpcEnc.Rows[0]["ENTREGADO-FAX"].ToString().ToUpper();
                    string saltoEntregado = "\n";
                    string separador = ", ";

                    if (nombreEntregado.Length > 0)
                    {
                        nombreEntregado = nombreEntregado + saltoEntregado;
                    }
                    else
                    {
                        nombreEntregado = "" + saltoEntregado;
                    }

                    if (calleEntregado.Length > 0)
                    {
                        calleEntregado = calleEntregado + saltoEntregado;
                    }
                    else
                    {
                        calleEntregado = "" + saltoEntregado;
                    }

                    if (coloniaEntregado.Length > 0)
                    {
                        coloniaEntregado = coloniaEntregado + separador + saltoEntregado;
                    }
                    else
                    {
                        coloniaEntregado = "" + saltoEntregado;
                    }

                    if (municEntregado.Length > 0)
                    {
                        municEntregado = municEntregado + separador;
                    }
                    else
                    {
                        municEntregado = "";
                    }

                    if (estadoEntregado.Length > 0)
                    {
                        estadoEntregado = estadoEntregado + saltoEntregado;
                    }
                    else
                    {
                        estadoEntregado = "" + saltoEntregado;
                    }

                    if (cpEntregado.Length > 0)
                    {
                        cpEntregado = htDatosEtiquetas["CP1"] + cpEntregado + separador;
                    }
                    else
                    {
                        cpEntregado = "" + separador;
                    }

                    if (paisEntregado.Length > 0)
                    {
                        paisEntregado = paisEntregado + saltoEntregado;
                    }
                    else
                    {
                        paisEntregado = "" + saltoEntregado;
                    }

                    if (tel1Entregado.Length > 0)
                    {
                        tel1Entregado = htDatosEtiquetas["TEL"] + tel1Entregado;
                    }
                    else
                    {
                        tel1Entregado = htDatosEtiquetas["TEL"].ToString();
                    }

                    if (tel2Entregado.Length > 0)
                    {
                        tel2Entregado = separador + tel2Entregado + saltoEntregado;
                    }
                    else
                    {
                        tel2Entregado = separador + saltoEntregado;
                    }

                    if (faxEntregado.Length > 0)
                    {
                        faxEntregado = htDatosEtiquetas["FAX"] + faxEntregado;
                    }
                    else
                    {
                        faxEntregado = htDatosEtiquetas["FAX"].ToString();
                    }

                    par = new Paragraph();
                    par.SetLeading(7f, 1f);
                    par.Add(new Chunk(nombreEntregado, f6));
                    par.Add(new Chunk(calleEntregado, f6));
                    par.Add(new Chunk(coloniaEntregado, f6));
                    par.Add(new Chunk(municEntregado, f6));
                    par.Add(new Chunk(estadoEntregado, f6));
                    par.Add(new Chunk(cpEntregado, f6));
                    par.Add(new Chunk(paisEntregado, f6));
                    par.Add(new Chunk(tel1Entregado, f6));
                    par.Add(new Chunk(tel2Entregado, f6));
                    par.Add(new Chunk(faxEntregado, f6));
                    cell = new PdfPCell(par);
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = (float).5;
                    cell.BorderWidthRight = (float).5;
                    cell.BorderWidthBottom = (float).5;
                    cell.BorderColor = gris;
                    cell.Colspan = 2;
                    detalle.AddCell(cell);
                }
                else if (tipoIDoc.StartsWith("FIDCCP02"))
                {
                    par = new Paragraph();
                    par.SetLeading(7f, 1f);
                    par.Add(new Chunk(electronicDocument.Data.Receptor.Nombre.Value + direccionReceptor.ToString(), f6));
                    cell = new PdfPCell(par);
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = (float).5;
                    cell.BorderWidthRight = 0;
                    cell.BorderWidthBottom = (float).5;
                    cell.BorderColor = gris;
                    cell.Colspan = 2;
                    detalle.AddCell(cell);
                }

                #endregion

                #region "Encabezado Partidas"

                PdfPTable encPartidas = new PdfPTable(7);
                encPartidas.WidthPercentage = 100;
                encPartidas.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                encPartidas.SetWidths(new int[7] { 10, 48, 6, 7, 10, 9, 10 });
                encPartidas.DefaultCell.Border = 0;
                encPartidas.LockedWidth = true;

                #region "Titulos de Partidas"

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["CLAVE"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encPartidas.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["DESCRIPCION"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encPartidas.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["PS"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encPartidas.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["CANTIDAD"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encPartidas.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["UNIDAD-MEDIDA"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encPartidas.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["PRECIO-UNITARIO"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encPartidas.AddCell(cell);

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["IMPORTE"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = (float).5;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = (float).5;
                cell.BorderColor = gris;
                encPartidas.AddCell(cell);

                #endregion

                #endregion

                #region "Construimos Tabla de Detalles partidas"

                PdfPTable partidas = new PdfPTable(7);
                partidas.WidthPercentage = 100;
                partidas.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                partidas.SetWidths(new int[7] { 10, 48, 6, 7, 10, 9, 10 });
                partidas.DefaultCell.Border = 0;
                partidas.LockedWidth = true;

                #region "contenido partidas"

                if (dtOpcDet.Rows.Count > 0)
                {
                    for (int i = 0; i < electronicDocument.Data.Conceptos.Count; i++)
                    {
                        string Descripcion = electronicDocument.Data.Conceptos[i].Descripcion.Value.ToUpper().TrimEnd().TrimStart();
                        string clave = electronicDocument.Data.Conceptos[i].NumeroIdentificacion.Value.ToUpper().TrimEnd().TrimStart();

                        if (Descripcion != "SEGURO")
                        {
                            if (Descripcion != "FLETE")
                            {
                                if (Descripcion != "ADUANA")
                                {
                                    switch (clave)
                                    {
                                        case "ZL2":
                                        case "ZG2":
                                        case "ZBON":
                                            #region "Descripción e Importe"
                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("N", _ci), f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = (float).5;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);
                                            #endregion
                                            break;
                                        case "ZDEV":
                                            #region "Solo descripción"
                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase("", f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = (float).5;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);
                                            #endregion
                                            break;
                                        default:
                                            #region "Con Clave"
                                            cell = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].NumeroIdentificacion.Value, f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            if (tipoIDoc.StartsWith("INVOIC"))
                                            {
                                                cell = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, f5));
                                            }
                                            else
                                            {
                                                cell = new PdfPCell(new Phrase(dtOpcDet.Rows[i]["DESCRIP"].ToString().Replace("*", "\n"), f5));
                                            }
                                            //cell = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Descripcion.Value, f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase(dtOpcDet.Rows[i]["PS"].ToString(), f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Cantidad.Value.ToString("N", _ci), f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Unidad.Value, f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].ValorUnitario.Value.ToString("N", _ci), f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = 0;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);

                                            cell = new PdfPCell(new Phrase(electronicDocument.Data.Conceptos[i].Importe.Value.ToString("N", _ci), f5));
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BorderWidthTop = 0;
                                            cell.BorderWidthLeft = (float).5;
                                            cell.BorderWidthRight = (float).5;
                                            cell.BorderWidthBottom = (float).5;
                                            cell.BorderColor = gris;
                                            partidas.AddCell(cell);
                                            #endregion
                                            break;
                                    }
                                }
                            }
                        }
                    }
                }

                #endregion

                #endregion

                #region "Construimos Tabla de Impuestos"

                Table impuestos = new Table(3);
                float[] headerwidthscontenido = { 75, 11, 15 };
                impuestos.Widths = headerwidthscontenido;
                impuestos.WidthPercentage = 100;
                impuestos.Padding = 1;
                impuestos.Spacing = 1;
                impuestos.BorderWidth = 0;
                impuestos.DefaultCellBorder = 0;
                Cell cel;

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(8f, 9f);
                par.Add(new Chunk(dtOpcEnc.Rows[0]["FECHA-PAGO-ANT"].ToString().ToUpper(), f5));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                cel.Rowspan = 3;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["SUMA"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(Convert.ToDouble(dtOpcEnc.Rows[0]["SUMA"]).ToString("N", _ci), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["DESCUENTO"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Descuento.Value.ToString("N", _ci), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["ADUANA"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(Convert.ToDouble(dtOpcEnc.Rows[0]["ADUANA"]).ToString("N", _ci), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 9f);
                par.Add(new Chunk(htDatosEtiquetas["NOTA1"].ToString(), f5BBI));
                par.Add(new Chunk(htDatosEtiquetas["NOTA2"].ToString(), f5BBI));
                par.Add(new Chunk(htDatosEtiquetas["NOTA3"].ToString(), f5BBI));
                cel = new Cell(par);
                cel.Rowspan = 3;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["FLETE"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(Convert.ToDouble(dtOpcEnc.Rows[0]["FLETE"]).ToString("N", _ci), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["SEGURO"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(Convert.ToDouble(dtOpcEnc.Rows[0]["SEGURO"]).ToString("N", _ci), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["SUB-TOTAL"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(Convert.ToDouble(dtOpcEnc.Rows[0]["SUB-TOTAL"]).ToString("N", _ci), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["CANTIDAD-LETRA"].ToString() + dtOpcEnc.Rows[0]["CANTIDAD-LETRA"], f5BBI));
                cel.Rowspan = 2;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["IVA"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);


                cel = new Cell(new Phrase(electronicDocument.Data.Impuestos.TotalTraslados.Value.ToString("N", _ci), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["TOTAL"].ToString(), f6B));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                cel = new Cell(new Phrase(electronicDocument.Data.Total.Value.ToString("N", _ci), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                cel.BorderColor = grisOX;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                impuestos.AddCell(cel);

                #endregion

                #region "Construimos Tabla de Adicionales"

                DefaultSplitCharacter split = new DefaultSplitCharacter();
                Table adicional = new Table(8);
                float[] headerwidthsAdicional = { 20, 20, 20, 20, 20, 20, 20, 20 };
                adicional.Widths = headerwidthsAdicional;
                adicional.WidthPercentage = 100;
                adicional.Padding = 1;
                adicional.Spacing = 1;
                adicional.BorderWidth = (float).5;
                adicional.DefaultCellBorder = 1;
                adicional.BorderColor = gris;

                cel = new Cell(new Phrase(htDatosEtiquetas["TOTAL-TARIMAS"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["NUMERO-CAJAS"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["NUMERO-PIEZAS"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 4;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["PESO-BRUTO"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["PESO-NETO"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEnc.Rows[0]["TOTAL-TARIMAS"].ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEnc.Rows[0]["NUMERO-CAJAS"].ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEnc.Rows[0]["NUMERO-PIEZAS"].ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 4;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEnc.Rows[0]["PESO-BRUTO"].ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEnc.Rows[0]["PESO-NETO"].ToString(), f5));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                par = new Paragraph();
                par.KeepTogether = true;
                par.SetLeading(7f, 9f);
                par.Add(new Chunk(htDatosEtiquetas["COMENTARIOS"] + "\n", f5B));
                par.Add(new Chunk(dtOpcEnc.Rows[0]["COMENTARIOS"].ToString().Replace("*", "\n"), f5));
                cel = new Cell(par);
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Rowspan = 2;
                cel.Colspan = 8;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["METODO-PAGO"].ToString(), f5B));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = (float).5;
                cel.Colspan = 8;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["METODO-PAGO"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                string medotoPago = electronicDocument.Data.MetodoPago.IsAssigned
                                          ? electronicDocument.Data.MetodoPago.Value
                                         : "";

                cel = new Cell(new Phrase(medotoPago, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["BANCO"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEnc.Rows[0]["BANCO-DEPOSITO"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["CUENTA"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                string cuenta = electronicDocument.Data.NumeroCuentaPago.IsAssigned
                                          ? electronicDocument.Data.NumeroCuentaPago.Value
                                          : "";

                cel = new Cell(new Phrase(cuenta, f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = 0;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = 0;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 3;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["VENCIMIENTO-PAGO"].ToString(), f5BBI));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = (float).5;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                cel.Colspan = 8;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["FORMA-PAGO"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEnc.Rows[0]["FORMA-PAGO"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = (float).5;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["BANCO"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEnc.Rows[0]["BANCO"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["CUENTA-CIE"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEnc.Rows[0]["CUENTA-CIE"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(htDatosEtiquetas["REFERENCIA2"].ToString(), titulo));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BackgroundColor = azul;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
                adicional.AddCell(cel);

                cel = new Cell(new Phrase(dtOpcEnc.Rows[0]["REFERENCIA2"].ToString(), f6));
                cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                cel.BorderColor = gris;
                cel.BorderWidthTop = (float).5;
                cel.BorderWidthRight = 0;
                cel.BorderWidthLeft = (float).5;
                cel.BorderWidthBottom = 0;
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
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = 0;
                    cel.BorderWidthBottom = 0;
                    cel.Rowspan = 7;
                    cel.Colspan = 2;
                    adicional.AddCell(cel);

                    par = new Paragraph();
                    par.SetLeading(7f, 0f);

                    par.Add(new Chunk(htDatosEtiquetas["SELLO-DIGITA-CFDI"] + "\n", f5B));
                    par.Add(new Chunk(electronicDocument.Data.Sello.Value, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.Colspan = 6;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(htDatosEtiquetas["TIMBRE-FISCAL"].ToString(), titulo));
                    cel.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cel.HorizontalAlignment = Element.ALIGN_CENTER;
                    cel.BorderColor = gris;
                    cel.BackgroundColor = azul;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 6;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(htDatosEtiquetas["FOLIO-FISCAL"].ToString(), f5B));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(objTimbre.Uuid.Value, f5));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(htDatosEtiquetas["FECHA-CERTIFICACION"].ToString(), f5B));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    string[] fechaTimbrado = Convert.ToDateTime(objTimbre.FechaTimbrado.Value).GetDateTimeFormats('s');

                    cel = new Cell(new Phrase(fechaTimbrado[0], f5));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(htDatosEtiquetas["SERIE-CERTIFICADO-SAT"].ToString(), f5B));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(objTimbre.NumeroCertificadoSat.Value, f5));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(htDatosEtiquetas["SERIE-CERTIFICADO-EMISOR"].ToString(), f5B));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(electronicDocument.Data.NumeroCertificado.Value, f5));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    cel = new Cell(new Phrase(htDatosEtiquetas["REGIMEN-EMISOR"].ToString(), f5B));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = 0;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    StringBuilder regimenes = new StringBuilder();
                    if (electronicDocument.Data.Emisor.Regimenes.IsAssigned)
                    {
                        for (int i = 0; i < electronicDocument.Data.Emisor.Regimenes.Count; i++)
                        {
                            regimenes.Append(electronicDocument.Data.Emisor.Regimenes[i].Regimen.Value).Append("\n");
                        }
                        if (IDIOMA == "E")
                        {
                            regimenes = new StringBuilder();
                            regimenes.Append(htDatosEtiquetas["REGIMEN-DATO"]);
                        }
                    }
                    else
                    {
                        regimenes.Append(htDatosEtiquetas["REGIMEN-DATO"]);
                    }

                    cel = new Cell(new Phrase(regimenes.ToString(), f5));
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.Colspan = 3;
                    adicional.AddCell(cel);

                    par = new Paragraph();
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk(htDatosEtiquetas["CADENA-ORIGINAL"] + "\n", f5B));
                    par.Add(new Chunk(electronicDocument.FingerPrintPac, f5).SetSplitCharacter(split));

                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = (float).5;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = (float).5;
                    cel.Colspan = 8;
                    adicional.AddCell(cel);

                    par = new Paragraph();
                    par.KeepTogether = true;
                    par.SetLeading(7f, 0f);
                    par.Add(new Chunk(htDatosEtiquetas["SELLO-DIGITA-SAT"] + "\n", f5B));
                    par.Add(new Chunk(objTimbre.SelloSat.Value, f5).SetSplitCharacter(split));
                    cel = new Cell(par);
                    cel.BorderColor = gris;
                    cel.BorderWidthTop = 0;
                    cel.BorderWidthRight = (float).5;
                    cel.BorderWidthLeft = (float).5;
                    cel.BorderWidthBottom = 0;
                    cel.Colspan = 8;
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

                cell = new PdfPCell(new Phrase(htDatosEtiquetas["LEYENDA"].ToString(), titulo));
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = azul;
                cell.BorderWidthTop = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                footer.AddCell(cell);

                #endregion

                #endregion

                pageEventHandler.encabezado = encabezado;
                pageEventHandler.dSaltoLinea = dSaltoLinea;
                pageEventHandler.detalle = detalle;
                pageEventHandler.encPartidas = encPartidas;
                pageEventHandler.piePaginaIdioma = IDIOMA;
                pageEventHandler.adicional = adicional;
                pageEventHandler.footer = footer;
                document.Open();
                document.Add(partidas);
                document.Add(impuestos);
                document.Add(adicional);
                fs.Flush();
                document.Close();
                writer.Close();
                fs.Close();

                string filePdfExt = pathPdf.Replace(_rutaDocs, _rutaDocsExt);
                string urlPathFilePdf = filePdfExt.Replace(@"\", "/");

                return "1#" + urlPathFilePdf;
            }
            catch (Exception ex)
            {
                return "0#" + ex.Message;
            }
        }
        #endregion
    }

    public class pdfPageEventHandlerGalgo : PdfPageEventHelper
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
        public PdfPTable encabezado { get; set; }
        public PdfPTable encPartidas { get; set; }
        public Chunk dSaltoLinea { get; set; }
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
            document.Add(dSaltoLinea);
            document.Add(detalle);
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

    public class DefaultSplitCharacter : ISplitCharacter
    {
        #region "ISplitCharacter"

        /**
         * An instance of the default SplitCharacter.
         */
        public static readonly ISplitCharacter DEFAULT = new DefaultSplitCharacter();

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