#region "using"

using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using System.Web.Services;
using iTextSharp.text;
using iTextSharp.text.pdf;
using HyperSoft.ElectronicDocumentLibrary.Document;
using r3TakeCore.Core;
using r3TakeCore.Data;
using Data = HyperSoft.ElectronicDocumentLibrary.Complemento.TimbreFiscalDigital.Data;
using wsRepresentacionImpresa.App_Code.com.r3Take.Utils;
using Microsoft.WindowsAzure;
using Microsoft.WindowsAzure.StorageClient;

#endregion

namespace wsRepresentacionImpresa
{
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]

    public class wsRepresentacionImpresa : WebService
    {
        private static readonly CultureInfo _ci = new CultureInfo("es-mx");
        private static readonly string _rutaDocs = ConfigurationManager.AppSettings["rutaDocs"];
        private static readonly string _rutaDocsExt = ConfigurationManager.AppSettings["rutaDocsExterna"];

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

        #region "generaRepresentacionImpresaCFDI"

        [WebMethod(MessageName = "generaRepresentacionImpresaCFDIporidCfdi", Description = "Genera Representación Impresa desde idCfdi", EnableSession = true)] 
        public string generaRepresentacionImpresaCFDI(Int64 idCfdi, string para, string cc, string cco)
        {
            try
            {
                //inpersonalizar inPersonate = new inpersonalizar();
                //bool inPersonificado = inPersonate.impersonateValidUser(ConfigurationManager.AppSettings["UsrSharePoint"], ConfigurationManager.AppSettings["Dominio"], ConfigurationManager.AppSettings["PwdSharePoint"]);
                string respuestaWS = "0#Impersonaliza";

                //if (inPersonificado)
                //{
                    Assembly objExecutingAssemblies = Assembly.GetExecutingAssembly();
                    HttpContext.Current.Application.Add("Ass", objExecutingAssemblies.Location);
                    DataTable dtCfdi = buscarCfdi(idCfdi);

                    if (string.IsNullOrEmpty(para))
                        para = string.Empty;

                    if (string.IsNullOrEmpty(cc))
                        cc = string.Empty;

                    if (string.IsNullOrEmpty(cco))
                        cco = string.Empty;

                    respuestaWS = cargarCfdi(dtCfdi, para, cc, cco);

                    //Para quitar la inpersonificación
                //    inPersonate.undoImpersonation();
                //}

                return respuestaWS;
            }
            catch (Exception ex)
            {
                insertErrorStr(ex, "WebService Representación Impresa", HttpContext.Current);
                return "0#" + ex.Message;
            }
        }

        #endregion

        #region "generaRepresentacionImpresaUUID"

        [WebMethod(MessageName = "generaRepresentacionImpresaCFDIporUUID", Description = "Genera Representación Impresa desde UUID", EnableSession = true)]
        public string generaRepresentacionImpresaUUID(string UUID, string para, string cc, string cco)
        {
            try
            {
                //inpersonalizar inPersonate = new inpersonalizar();
                //bool inPersonificado = inPersonate.impersonateValidUser(ConfigurationManager.AppSettings["UsrSharePoint"], ConfigurationManager.AppSettings["Dominio"], ConfigurationManager.AppSettings["PwdSharePoint"]);
                string respuestaWS = "0#Impersonaliza";

                //if (inPersonificado)
                //{
                    Assembly objExecutingAssemblies = Assembly.GetExecutingAssembly();
                    HttpContext.Current.Application.Add("Ass", objExecutingAssemblies.Location);
                    DataTable dtCfdi = buscarCfdi(UUID);

                    if (string.IsNullOrEmpty(para))
                        para = string.Empty;

                    if (string.IsNullOrEmpty(cc))
                        cc = string.Empty;

                    if (string.IsNullOrEmpty(cco))
                        cco = string.Empty;

                    respuestaWS = cargarCfdi(dtCfdi, para, cc, cco);

                //    //Para quitar la inpersonificación
                //    inPersonate.undoImpersonation();
                //}

                return respuestaWS;
            }
            catch (Exception ex)
            {
                insertErrorStr(ex, "WebService Representación Impresa", HttpContext.Current);
                return "0#" + ex.Message;
            }
        }

        #endregion

        #region "buscarCfdi"

        private static DataTable buscarCfdi(Int64 idCfdi)
        {
            DAL dal = new DAL();
            DataTable dtCfdi;
            StringBuilder sbCfdi = new StringBuilder();
            StringBuilder sbCfdiParams = new StringBuilder();

            sbCfdiParams.Append("F:I:").Append(idCfdi);
            sbCfdi.
                Append("SELECT idCfdi, idEmisor, idSucursalEmisor, idMoneda, tipoCFDI, rutaXML, xmlCFDI, observaciones, noOrdenCompra, xmlCFDI, idTipoComp, idRol ").
                Append("FROM cfdi ").
                Append("WHERE idCfdi = @0 ");

            dtCfdi = dal.QueryDT("DS_FE", sbCfdi.ToString(), sbCfdiParams.ToString(), HttpContext.Current);

            return dtCfdi;
        }

        private static DataTable buscarCfdi(string UUID)
        {
            DAL dal = new DAL();
            DataTable dtCfdi;
            StringBuilder sbCfdi = new StringBuilder();
            StringBuilder sbCfdiParams = new StringBuilder();

            sbCfdiParams.Append("F:S:").Append(UUID.ToUpper());
            sbCfdi.
                Append("SELECT idCfdi, idEmisor, idSucursalEmisor, idMoneda, tipoCFDI, rutaXML, xmlCFDI, observaciones, noOrdenCompra, xmlCFDI, idTipoComp ").
                Append("FROM cfdi ").
                Append("WHERE UUID = @0 ");

            dtCfdi = dal.QueryDT("DS_FE", sbCfdi.ToString(), sbCfdiParams.ToString(), HttpContext.Current);

            return dtCfdi;
        }

        #endregion

        #region "cargarCfdi"

        private static string cargarCfdi(DataTable dtCfdi, string para, string cc, string cco)
        {
            DAL dal = new DAL();

            if (dtCfdi.Rows.Count > 0)
            {
                Int64 idCfdi = Convert.ToInt64(dtCfdi.Rows[0]["idCfdi"]);
                Int64 idEmisor = Convert.ToInt64(dtCfdi.Rows[0]["idEmisor"]);
                Int32 idTipoComp = Convert.ToInt32(dtCfdi.Rows[0]["idTipoComp"]);
                int idRol = Convert.ToInt16(dtCfdi.Rows[0]["idRol"].ToString());

                Hashtable htParams = new Hashtable();
                htParams.Add("idCfdi", idCfdi);
                htParams.Add("xmlCFDI", dtCfdi.Rows[0]["xmlCFDI"].ToString());
                htParams.Add("idEmisor", idEmisor);
                htParams.Add("idSucursalEmisor", dtCfdi.Rows[0]["idSucursalEmisor"].ToString());
                htParams.Add("idMoneda", dtCfdi.Rows[0]["idMoneda"].ToString());
                htParams.Add("tipoComprobante", dtCfdi.Rows[0]["tipoCFDI"].ToString());
                htParams.Add("urlPathFileXml", dtCfdi.Rows[0]["rutaXML"].ToString());
                htParams.Add("observaciones", dtCfdi.Rows[0]["observaciones"].ToString());
                htParams.Add("noOrdenCompra", dtCfdi.Rows[0]["noOrdenCompra"].ToString());
                htParams.Add("idTipoComp", idTipoComp);

                // Validamos si el usuario tiene créidtos para realizar la generación del PDF

                StringBuilder sbValidaCreditos = new StringBuilder();
                DataTable dtValidaCreditos;

                if (idRol == 15)
                {
                    sbValidaCreditos.
                        Append("SELECT 1 AS cantidadCreditos");
                }

                else
                {
                    sbValidaCreditos.
                        Append("SELECT cantidadCreditos FROM creditos WHERE idUsuario = ").
                        Append("(SELECT TOP 1 U.idUser ").
                        Append("FROM empresas E ").
                        Append("LEFT OUTER JOIN vSucursales S ON E.idEmpresa = S.idEmpresa ").
                        Append("LEFT OUTER JOIN sucursalesXUsuario SU ON S.idSucursal = SU.idSucursal ").
                        Append("LEFT OUTER JOIN r3TakeCore.dbo.SYS_User U ON SU.idUser = U.idUser ").
                        Append("WHERE E.idEmpresa =  @0 ) ");
                }

                dtValidaCreditos = dal.QueryDT("DS_FE", sbValidaCreditos.ToString(), "F:I:" + idEmisor, HttpContext.Current);

                if (dtValidaCreditos.Rows.Count > 0)
                {
                    if (Convert.ToInt64(dtValidaCreditos.Rows[0]["cantidadCreditos"]) > 0)
                    {
                        // Cargamos el Xml del Cfdi.

                        Hashtable htFacturaxion = new Hashtable();
                        htFacturaxion = cargarXml(htParams);

                        if (htFacturaxion.Count == 1)
                        {
                            return htFacturaxion["logErr"].ToString();
                        }

                        if (htFacturaxion.Count > 1)
                        {
                            StringBuilder sbFacturaEspecial = new StringBuilder();
                            StringBuilder sbParamsFactEsp = new StringBuilder();
                            string urlPathFilePdf;

                            sbFacturaEspecial.
                                Append("SELECT espacioNombres, metodo, enviaMail ").
                                Append("FROM facturaEspecial ").
                                Append("WHERE idEmpresa = @0 AND tipoComprobante = @1 AND tipoPortal = 3 AND ST = 1");

                            sbParamsFactEsp.
                                Append("F:I:").Append(idEmisor).
                                Append(";").
                                Append("F:I:").Append(idTipoComp);

                            // Validamos si se trata de un CFDI con configuración especial

                            DataTable dtFacturaEspecial = dal.QueryDT("DS_FE", sbFacturaEspecial.ToString(), sbParamsFactEsp.ToString(), HttpContext.Current);

                            if (dtFacturaEspecial.Rows.Count > 0)
                            {
                                //Enlaza la operación de Carga de una clase y un método al Vuelo para la generación de la factura especial
                                urlPathFilePdf = loadGenerarPdf(dtFacturaEspecial.Rows[0]["espacioNombres"].ToString(), dtFacturaEspecial.Rows[0]["metodo"].ToString(), htFacturaxion, HttpContext.Current);
                                //urlPathFilePdf = loadGenerarPdf("wsRepresentacionImpresa.App_Code.com.Facturaxion.facturaEspecial.opamComprobante", "generarPdf", htFacturaxion, HttpContext.Current);
                            }
                            else
                            {
                                //Genera el PDF de forma convencional, mediante la configuración de la factura
                                urlPathFilePdf = generarPdf(htFacturaxion, HttpContext.Current);
                            }

                            if (urlPathFilePdf.StartsWith("1"))
                            {
                                String[] arrUrlPathFilePdf = urlPathFilePdf.Split(new Char[] { '#' });

                                #region "Actualiza Ruta de Pdf en Base de Datos"

                                StringBuilder sbActualizaRutaPdf = new StringBuilder();
                                StringBuilder sbParamsActRutaPdf = new StringBuilder();

                                sbParamsActRutaPdf.
                                    Append("F:I:").Append(Convert.ToInt64(htFacturaxion["idCfdi"])).
                                    Append(";").
                                    Append("F:S:").Append(arrUrlPathFilePdf[1]);

                                sbActualizaRutaPdf.Append("UPDATE cfdi SET rutaPDF = @1 WHERE idCfdi = @0 ");

                                dal.ExecuteNonQuery("DS_FE", sbActualizaRutaPdf.ToString(), sbParamsActRutaPdf.ToString(), HttpContext.Current);

                                #endregion

                                #region "Envía email"

                                string urlPathFileXml = htFacturaxion["urlPathFileXml"].ToString();

                                if (dtFacturaEspecial.Rows.Count == 0 || dtFacturaEspecial.Rows[0]["enviaMail"].ToString() == "True")
                                {

                                    Hashtable htEnvioMail = new Hashtable();
                                    htEnvioMail.Add("rutaPDF", arrUrlPathFilePdf[1]);
                                    htEnvioMail.Add("rutaXML", urlPathFileXml);
                                    htEnvioMail.Add("correoPara", para);
                                    htEnvioMail.Add("correoCC", cc);
                                    htEnvioMail.Add("correoCCO", cco);

                                        if (para.Length > 0 || cc.Length > 0 || cco.Length > 0)
                                        {
                                            string envioMail = enviaCfdEmail(htEnvioMail, HttpContext.Current);

                                            if (envioMail.StartsWith("0"))
                                                return "0#Error en envío de correo " + envioMail;
                                        }
                                }

                                #endregion

                                #region "Descontamos Créditos"

                                StringBuilder sbDescuentaCreditos = new StringBuilder();

                                sbDescuentaCreditos.
                                    Append("UPDATE creditos SET cantidadCreditos = cantidadCreditos - 1 WHERE idUsuario = ").
                                    Append("(SELECT TOP 1 U.idUser ").
                                    Append("FROM empresas E ").
                                    Append("LEFT OUTER JOIN vSucursales S ON E.idEmpresa = S.idEmpresa ").
                                    Append("LEFT OUTER JOIN sucursalesXUsuario SU ON S.idSucursal = SU.idSucursal ").
                                    Append("LEFT OUTER JOIN r3TakeCore.dbo.SYS_User U ON SU.idUser = U.idUser ").
                                    Append("WHERE E.idEmpresa =  @0 ) ");


                                if (idRol != 15)
                                    dal.ExecuteNonQuery("DS_FE", sbDescuentaCreditos.ToString(), "F:I:" + idEmisor, HttpContext.Current);

                                #endregion

                                #region "Lanzamos Hilo para envío de invitación al portal de facturaxion"

                                htFacturaxion.Add("hc", HttpContext.Current);
                                htFacturaxion.Add("para", para);
                                Thread threadEnvioInvitacion = new Thread((enviaInvitacion));
                                threadEnvioInvitacion.Start(htFacturaxion);

                                #endregion

                                return urlPathFilePdf;
                            }
                            else
                            {
                                return "0#No se pudo generar el Pdf del Cfdi";
                            }
                        }
                        else
                        {
                            return "0#No se pudo cargar el Xml del Cfdi";
                        }
                    }
                    else
                    {
                        return "0#Usuario No Cuenta Con Créditos Suficientes para Realizar la Operación";
                    }
                }
                else
                {
                    return "0#Usuario No Existente en Tabla de Créditos";
                }
            }
            else
            {
                return "0#No se Encontró el Cfdi";
            }
        }

        #endregion

        #region "cargarXml"

        private static Hashtable cargarXml(Hashtable htParams)
        {
            Int64 idCfdi = Convert.ToInt64(htParams["idCfdi"]);
            string urlPathFileXml = htParams["urlPathFileXml"].ToString();
            Int64 idEmisor = Convert.ToInt64(htParams["idEmisor"]);
            Int64 idSucursalEmisor = Convert.ToInt64(htParams["idSucursalEmisor"]);
            Int32 idMoneda = Convert.ToInt32(htParams["idMoneda"]);
            Int32 tipoComprobante = 1;
            string observaciones = htParams["observaciones"].ToString();
            string noOrdenCompra = htParams["noOrdenCompra"].ToString();

            string rutaDocumentoXml = urlPathFileXml.Replace(_rutaDocsExt, _rutaDocs).Replace("\\", "/");

            ElectronicDocument electronicDocument = ElectronicDocument.NewEntity();
            bool cfdiCargado = electronicDocument.LoadFromFile(rutaDocumentoXml);

            if (cfdiCargado)
            {
                string directorioPdf = Path.GetDirectoryName(rutaDocumentoXml);
                string archivoPdf = Path.GetFileNameWithoutExtension(rutaDocumentoXml);
                string rutaDocumentoPdf = Path.Combine(directorioPdf, archivoPdf + ".pdf");
                bool timbrar = true;
                
                // Generamos el objeto timbre

                Data objTimbre = electronicDocument.GetTimbre();
                                
                #region "Armamos el Hashtable con Parámetros Facturaxion"

                Hashtable htFacturaxion = new Hashtable();

                htFacturaxion.Add("idCfdi", idCfdi);
                htFacturaxion.Add("urlPathFileXml", urlPathFileXml);
                htFacturaxion.Add("rutaDocumentoPdf", rutaDocumentoPdf);
                htFacturaxion.Add("electronicDocument", electronicDocument);
                htFacturaxion.Add("objTimbre", objTimbre);
                htFacturaxion.Add("idEmisor", idEmisor);
                htFacturaxion.Add("idSucursalEmisor", idSucursalEmisor);
                htFacturaxion.Add("timbrar", timbrar);
                htFacturaxion.Add("observaciones", observaciones);
                htFacturaxion.Add("noOrdenCompra", noOrdenCompra);
                htFacturaxion.Add("idMoneda", idMoneda);
                htFacturaxion.Add("tipoComprobante", tipoComprobante);

                #endregion

                return htFacturaxion;
            }
            else
            {
                Hashtable htErr = new Hashtable();
                htErr.Add("logErr", "0#Error al leer XML - " + electronicDocument.ErrorText);
                return new Hashtable(htErr);
            }
        }

        #endregion

        #region "generarPdf"

        private static string generarPdf(Hashtable htFacturaxion, HttpContext hc)
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
                           Append("logoPosX, logoPosY, headerPosX, headerPosY, footerPosX, footerPosY, conceptosColWidth, desgloseColWidth, S.nombreSucursal ").
                           Append("FROM configuracionFacturas CF ").
                           Append("LEFT OUTER JOIN sucursales S ON S.idSucursal = @0 ").
                           Append("LEFT OUTER JOIN configuracionFactDet CFD ON CF.idConFact = CFD.idConFact ").
                           Append("WHERE CF.ST = 1 AND CF.idEmpresa = -1 AND CF.idTipoComp = @1 AND idCFDProcedencia = 1 AND objDesc NOT LIKE 'nuevoLbl%' ");
                   }
                   else
                   {
                       sbConfigFact.
                           Append("SELECT rutaTemplateHeader, rutaTemplateFooter, @4 AS rutaLogo, objDesc, posX, posY, fontSize, dbo.convertNumToTextFunction( @2, @3) AS cantidadLetra, ").
                           Append("logoPosX, logoPosY, headerPosX, headerPosY, footerPosX, footerPosY, conceptosColWidth, desgloseColWidth, S.nombreSucursal ").
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
               Document document = new Document(PageSize.LETTER, 25, 25, 25, 40);
               document.AddAuthor("Facturaxion");
               document.AddCreator("r3Take");
               document.AddCreationDate();
               //FileStream fs = new FileStream(pathPdf, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
               pdfPageEventHandler pageEventHandler = new pdfPageEventHandler();
               PdfWriter writer = PdfWriter.GetInstance(document, fs);
               writer.SetFullCompression();
               writer.ViewerPreferences = PdfWriter.PageModeUseNone;
               writer.PageEvent = pageEventHandler;
               writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);

               pageEventHandler.rutaImgFooter = dtConfigFact.Rows[0]["rutaTemplateFooter"].ToString();

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

               Font fontLbl = new Font(Font.HELVETICA, 6, Font.BOLD, new Color(43, 145, 175));
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

               document.Close();
               writer.Close();

               string filePdfExt = pathPdf.Replace(_rutaDocs, _rutaDocsExt);
               string urlPathFilePdf = filePdfExt.Replace(@"\", "/");

               //Subimos Archivo al Azure
               string res = App_Code.com.Facturaxion.facturaEspecial.wAzure.azureUpDownLoad(1, pathPdf);
               //res = App_Code.com.Facturaxion.facturaEspecial.wAzure.azureUpDownLoad(2, pathPdf);

               return "1#" + urlPathFilePdf;
            }
           catch (Exception ex)
           {
               fs.Flush();
               fs.Close();
               File.Delete(pathPdf);

               insertErrorStr(ex, "WebService Representación Impresa", HttpContext.Current);
               return "0#" + ex.Message;
           }
        }

        #endregion

        #region "enviaCFDEmail"

        public static string enviaCfdEmail(Hashtable h, HttpContext hc)
        {
            try
            {
                const string matchEmailPattern =
                                                @"^(([\w-]+\.)+[\w-]+|([a-zA-Z]{1}|[\w-]{2,}))@"
                                                + @"((([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?
				                                        [0-9]{1,2}|25[0-5]|2[0-4][0-9])\."
                                                + @"([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?
				                                        [0-9]{1,2}|25[0-5]|2[0-4][0-9])){1}|"
                                                + @"([a-zA-Z]+[\w-]+\.)+[a-zA-Z]{2,4})$";

                Regex emailValidator = new Regex(matchEmailPattern);

                StringBuilder sbCuerpo = new StringBuilder();
                string rutaPdf = h["rutaPDF"].ToString().Trim();
                string rutaXml = h["rutaXML"].ToString().Trim();
                string para = h["correoPara"].ToString().Trim();
                string cc = h["correoCC"].ToString().Trim();
                string cco = h["correoCCO"].ToString().Trim();

                if (para.Length > 0)
                {
                    #region "Validamos correos Para"

                    string paraAux = string.Empty;
                    String[] arrEmailPara = para.Split(new Char[] { ';' });

                    int numEmailsPara = arrEmailPara.Length;

                    for (int i = 0; i < numEmailsPara; i++)
                    {
                        if (emailValidator.IsMatch(arrEmailPara[i], 0))
                        {
                            paraAux += arrEmailPara[i];
                            paraAux += ";";
                        }
                    }

                    if (paraAux.Length > 0)
                    {
                        para = paraAux.Remove(paraAux.Length - 1, 1);
                    }
                    else
                    {
                        Exception ex = new Exception("|Para " + para + "|ParaAux" + paraAux);
                        insertErrorStr(ex, "WebService Representación Impresa", HttpContext.Current);
                    }

                    #endregion
                }
                
                if (cc.Length > 0)
                {
                    #region "Validamos correos CC"

                    string ccAux = string.Empty;
                    String[] arrEmailCc = cc.Split(new Char[] { ';' });

                    int numEmailsCc = arrEmailCc.Length;

                    for (int i = 0; i < numEmailsCc; i++)
                    {
                        if (emailValidator.IsMatch(arrEmailCc[i], 0))
                        {
                            ccAux += arrEmailCc[i];
                            ccAux += ";";
                        }
                    }

                    if (ccAux.Length > 0)
                    {
                        cc = ccAux.Remove(ccAux.Length - 1, 1);
                    }
                    else
                    {
                        Exception ex = new Exception("|CC " + cc + "|CCAux" + ccAux);
                        insertErrorStr(ex, "WebService Representación Impresa", HttpContext.Current);
                    }

                    #endregion
                }

                if (cco.Length > 0)
                {
                    #region "Validamos correos CCO"

                    string ccoAux = string.Empty;
                    String[] arrEmailCco = cco.Split(new Char[] { ';' });

                    int numEmailsCco = arrEmailCco.Length;

                    for (int i = 0; i < numEmailsCco; i++)
                    {
                        if (emailValidator.IsMatch(arrEmailCco[i], 0))
                        {
                            ccoAux += arrEmailCco[i];
                            ccoAux += ";";
                        }
                    }

                    if (ccoAux.Length > 0)
                    {
                        cco = ccoAux.Remove(ccoAux.Length - 1, 1);
                    }
                    else
                    {
                        Exception ex = new Exception("|CCO " + cco + "|CCOAux" + ccoAux);
                        insertErrorStr(ex, "WebService Representación Impresa", HttpContext.Current);
                    }

                    #endregion
                }

                string serieFolio = Path.GetFileNameWithoutExtension(rutaXml);
                string fileXml = Path.GetFileName(rutaXml);
                string filePdf = Path.GetFileName(rutaPdf);

                string rutaPdfLocal = rutaPdf.Replace(_rutaDocsExt, _rutaDocs).Replace("/", @"\");
                string rutaXmlLocal = rutaXml.Replace(_rutaDocsExt, _rutaDocs).Replace("/", @"\");

                byte[] fileBufferPdf = readByteArrayFromFile(rutaPdfLocal);
                byte[] fileBufferXml = readByteArrayFromFile(rutaXmlLocal);

                Hashtable htArchivosAdjuntos = new Hashtable();

                htArchivosAdjuntos.Add(filePdf, fileBufferPdf);
                htArchivosAdjuntos.Add(fileXml, fileBufferXml);

                sbCuerpo.
                        Append("<b>Comprobante Fiscal Digital " + serieFolio + " generado exitosamente</b> <br><br>").
                        Append("Facturaxion le da las gracias por facturar con nosotros: <br><br>").
                        Append("Adjunto a este correo se envia el PDF de la Factura y el XML generado. <br><br> ");

                StringBuilder mensajeCompleto = Mailing.envuelvePlantilla(@"\mailing\fx.html", "#CONTENT#", sbCuerpo.ToString());

                return Mailing.main( "21.21.104.25", 25, "Facturaxion <facturaxion@freightideas.net>", para, mensajeCompleto.ToString(), "Factura Electrónica", cco, cc, htArchivosAdjuntos, string.Empty);
            }
            catch (Exception ex)
            {
                insertErrorStr(ex, "WebService Representación Impresa", HttpContext.Current);
                return "0#" + ex.Message;
            }
        }

        #endregion

        #region "readByteArrayFromFile"

        private static byte[] readByteArrayFromFile(string fileName)
        {
            byte[] buff = null;
            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
            BinaryReader br = new BinaryReader(fs);
            long numBytes = new FileInfo(fileName).Length;
            buff = br.ReadBytes((int)numBytes);
            return buff;
        }

        #endregion

        #region "loadGenerarPdf"

        public static string loadGenerarPdf(String className, String methodName, Hashtable htFacturaxion, HttpContext hc)
        {
            Type theType;
            Object[] parameters;
            ConstructorInfo constructor;
            MethodInfo method;
            Object returnVal;
            try
            {
                theType = Type.GetType(className);
                constructor = theType.GetConstructor(Type.EmptyTypes);
                parameters = new Object[2] { htFacturaxion, hc };
                object classObject = constructor.Invoke(new object[] { });
                
                method = theType.GetMethod(methodName);
                returnVal = method.Invoke(classObject, parameters);

                return returnVal.ToString();
            }
            catch (Exception ex)
            {
                insertErrorStr(ex, "WebService Representación Impresa", HttpContext.Current);
                throw ex;
            }
        }

        #endregion

        #region "errorHandler"

        public static void insertErrorStr(Exception ex, string orden, HttpContext hc)
        {
            DAL dal = new DAL();
            string insError = "";
            StringBuilder sbParamsReg = new StringBuilder("");
            string cuenta = "WSRepImp";

            string innerExceptionMSG = "";

            if (ex.InnerException != null)
            {
                innerExceptionMSG = ex.InnerException.ToString();
            }
            
            try
            {
                StackTrace st = new StackTrace(ex, true);
                StackFrame[] frames = st.GetFrames();
                StringBuilder sbLineas = new StringBuilder();
               
                // Iterate over the frames extracting the information you need         
                foreach (StackFrame frame in frames)
                {
                    sbLineas.Append(frame.GetFileName()).Append("-")
                            .Append(frame.GetMethod().Name).Append("-")
                            .Append(frame.GetFileLineNumber()).Append("-")
                            .Append(frame.GetFileColumnNumber()).Append("|");
                }

                insError += "INSERT INTO SYS_errHandler (account,fechaHoraErr,InnerException,message,source,orden,idAMFOper,lineAndColumn, UG,FG) ";
                insError += "VALUES (@0,GETDATE(),@1,@2,@3,@4,@5,@6,@0,GETDATE());";
                insError += "SELECT SCOPE_IDENTITY() AS idError;";

                sbParamsReg.
                        Append("F:S:").Append(cuenta).
                        Append(";").
                        Append("F:S:").Append(innerExceptionMSG).
                        Append(";").
                        Append("F:S:").Append(ex.Message).
                        Append(";").
                        Append("F:S:").Append(ex.Source).
                        Append(";").
                        Append("F:I:1").
                        Append(";").
                        Append("F:S:0").
                        Append(";").
                        Append("F:S:").Append(sbLineas);

                dal.ExecuteNonQuery("DS_r3", insError, sbParamsReg.ToString(), hc);
            }
            catch (Exception)
            {
                sbParamsReg = new StringBuilder("");
                
                sbParamsReg.
                        Append("F:S:").Append(cuenta).
                        Append(";").
                        Append("F:S:").Append(innerExceptionMSG).
                        Append(";").
                        Append("F:S:").Append(ex.Message).
                        Append(";").
                        Append("F:S:").Append(ex.Source).
                        Append(";").
                        Append("F:I:1").
                        Append(";").
                        Append("F:S:0").
                        Append(";").
                        Append("F:S:wsRepImp");

                dal.ExecuteNonQuery("DS_r3", insError, sbParamsReg.ToString(), hc);
            }
        }

        #endregion

        #region "enviaInvitacion"
        private static void enviaInvitacion(object objAnalytics)
        {
            try
            {
                #region "Variables"

                Hashtable h = (Hashtable)objAnalytics;
                HttpContext hc = (HttpContext)h["hc"];
                ElectronicDocument ElectronicDocument = (ElectronicDocument)h["electronicDocument"];
                DAL dal = new DAL();
                StringBuilder sbParametros = new StringBuilder();
                StringBuilder sbAgregaInvitado = new StringBuilder();
                #endregion
                
                #region "parametros"

                Hashtable htParametros = new Hashtable();
                int idTipoPersona = ElectronicDocument.Data.Receptor.Rfc.Value.Length == 12 ? 1 : 2;
                
                htParametros.Add("idEmisor", h["idEmisor"]);
                htParametros.Add("guidInvitacion", Guid.NewGuid());
                htParametros.Add("razonSocialInvitado", ElectronicDocument.Data.Receptor.Nombre.Value);
                htParametros.Add("rfcInvitado", ElectronicDocument.Data.Receptor.Rfc.Value);
                htParametros.Add("idTipoPersona", idTipoPersona);
                htParametros.Add("emailInvitado", h["para"]);
                htParametros.Add("telefonoInvitado", "");
                htParametros.Add("UG", h["idEmisor"]);

                sbParametros.
                   Append("H:I:idEmisor").//0
                   Append(";H:S:guidInvitacion").//1
                   Append(";H:S:razonSocialInvitado").//2
                   Append(";H:S:rfcInvitado").//3
                   Append(";H:S:idTipoPersona").//4
                   Append(";H:S:emailInvitado").//5
                   Append(";H:S:telefonoInvitado").//6
                   Append(";H:S:UG");//7

                #endregion

                #region "Agrega Invitados"

                sbAgregaInvitado.
                    Append("DECLARE @ENVIAR AS INT ").
                    Append("SELECT @ENVIAR = COUNT(idEmpresa) FROM empresas ").
                    Append("WHERE idEmpresa = @0 AND ST = 1 AND enviaInvitacion = 1 ").
                    Append("IF(@ENVIAR = 1)  ").
                    Append("BEGIN  ").
                    Append("    DECLARE @REGISTRADO AS INT DECLARE @INVITADO AS INT ").
                    Append("    DECLARE @RFC AS VARCHAR(15) SET @RFC = UPPER(@3) ").
                    Append("    SELECT @REGISTRADO = COUNT(rfc) FROM empresas ").
                    Append("    WHERE ST = 1 AND validado = 0 AND ").
                    Append("    DECRYPTBYPASSPHRASE(dbo.seed(), rfc) = @RFC ").
                    Append("    SELECT @INVITADO = COUNT(rfcInvitado) FROM invitacionFX ").
                    Append("    WHERE ST = 1 AND rfcInvitado = @RFC ").
                    Append("    IF(@REGISTRADO = 0 AND @INVITADO = 0) ").
                    Append("    BEGIN ").
                    Append("        INSERT INTO invitacionFX (idEmpresa, guidInvitacion, razonSocialInvitado, ").
                    Append("        rfcInvitado, idTipoPersona, emailInvitado, telefonoInvitado, idProcedencia, enviaCorreo, FG, UG) ").
                    Append("        VALUES (@0, @1, @2, @RFC, @4, @5, @6, 2, 1, GETDATE(), @7) ").
                    Append("    END ").
                    Append("END ");

                dal.ExecuteNonQuery("DS_FE", sbAgregaInvitado.ToString(), sbParametros.ToString(), htParametros, hc);

                #endregion
            }
            catch (Exception ex)
            {
                insertErrorStr(ex, "1", HttpContext.Current);
            }
        }
        #endregion
    }

    public class pdfPageEventHandler : PdfPageEventHelper
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

        public int lblPaginaIdioma { get; set; }
        public string rutaImgFooter { get; set; }
        public Font FooterFont { get; set; }

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
            cb.EndText();
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            #region "Footer"

            PdfPTable tblFooter = new PdfPTable(1);

            tblFooter.TotalWidth = document.PageSize.Width;
            Image imgFooter = Image.GetInstance(rutaImgFooter);
            float imgFooterWidth = document.PageSize.Width - 70;
            float imgFooterHeight = imgFooter.Height / (imgFooter.Width / imgFooterWidth);
            imgFooter.ScaleAbsolute(imgFooterWidth, imgFooterHeight);

            PdfPCell cell = new PdfPCell(imgFooter);
            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
            cell.PaddingRight = 20;
            cell.Border = 0;

            tblFooter.AddCell(cell);
            tblFooter.WriteSelectedRows(0, -1, 0, (document.BottomMargin + 15), writer.DirectContent);

            #endregion

            base.OnEndPage(writer, document);
            
            string lblPagina;
            string lblDe;
            string lblFechaImpresion;
            
            switch (lblPaginaIdioma)
            {
                case 1:
                    lblPagina = "Página ";
                    lblDe = " de ";
                    lblFechaImpresion = "Fecha de Impresión ";
                    break;

                case 2:
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
            cb.SetTextMatrix(pageSize.GetLeft(30), pageSize.GetBottom(20));
            cb.ShowText(text);
            cb.EndText();

            cb.AddTemplate(template, pageSize.GetLeft(30) + len, pageSize.GetBottom(20));

            cb.BeginText();
            cb.SetFontAndSize(bf, 8);
            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, lblFechaImpresion + PrintTime, pageSize.GetRight(30), pageSize.GetBottom(20), 0);
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
}

        