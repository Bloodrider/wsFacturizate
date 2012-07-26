using System;
using System.Security.Principal;
using System.Runtime.InteropServices;

/// <summary>
/// Clase encargada de Impersonalizar un intervalo de código con otro usuario.
/// </summary>
namespace wsRepresentacionImpresa.App_Code.com.r3Take.Utils
{
    public class inpersonalizar
    {
        public const int LOGON32_LOGON_INTERACTIVE = 2;
        public const int LOGON32_PROVIDER_DEFAULT = 0;

        WindowsImpersonationContext impersonationContext;

        /// <summary>
        /// Objeto COM que se encarga de hacer la impersonalización de nombre advapi32.dll
        /// </summary>
        /// <param name="lpszUserName">Usuario con el que se sumplantará la Identidad</param>
        /// <param name="lpszDomain">Dominio en el que se suplantará la identidad</param>
        /// <param name="lpszPassword">Passwor del usuario con el que se sumplantará la Identidad</param>
        /// <param name="dwLogonType">Logon Interactive inicializada en 2</param>
        /// <param name="dwLogonProvider">Logon provider inicializado en 0</param>
        /// <param name="phToken">Token del Logon User</param>
        /// <returns>Regresa un valor entero con las entidades duplicadas</returns>
        [DllImport("advapi32.dll")]
        public static extern int LogonUserA(String lpszUserName, String lpszDomain, String lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        /// <summary>
        /// Objeto COM que se encarga de hacer la impersonalización de nombre advapi32.dll
        /// </summary>
        /// <param name="hToken">Token del Logon User</param>
        /// <param name="impersonationLevel">Nivel de Impersonalización</param>
        /// <param name="hNewToken">Nuevo Token generado del Logon User/param>
        /// <returns>Regresa un objeto booleano que indica si fue exitosa la suplantación del usuario</returns>
        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int DuplicateToken(IntPtr hToken, int impersonationLevel, ref IntPtr hNewToken);

        /// <summary>
        /// Objeto COM que se ncarga de revertir la suplantación y autenticarse
        /// </summary>
        /// <returns>Regresa un objeto bool que indica si se pudo revertir la suplantación y autenticarse </returns>
        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool RevertToSelf();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="handle">Objeto COM que se encarga de Cerrar el manejador de Tokens </param>
        /// <returns>Regresa un objeto de tipo bool indicando si fue exitoso el cierre del manejador de Tokens</returns>
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern bool CloseHandle(IntPtr handle);

        /// <summary>
        /// Método encargado de hacer impersonal cierto código de la aplicación, utilizando otro usuario suplantando la identidad momentaneamente.
        /// </summary>
        /// <param name="userName">Usuario con el que se sumplantará la Identidad</param>
        /// <param name="domain">Dominio en el que se suplantará la identidad</param>
        /// <param name="password">Passwor del usuario con el que se sumplantará la Identidad</param>
        /// <returns>Regresa un objeto de tipo bool que indica si fue extiso el proceso de impersonalización o suplantación del usuario.</returns>
        public bool impersonateValidUser(String userName, String domain, String password)
        {
            WindowsIdentity tempWindowsIdentity;
            IntPtr token = IntPtr.Zero;
            IntPtr tokenDuplicate = IntPtr.Zero;

            if (RevertToSelf())
            {
                if (LogonUserA(userName, domain, password, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, ref token) != 0)
                {
                    if (DuplicateToken(token, 2, ref tokenDuplicate) != 0)
                    {
                        tempWindowsIdentity = new WindowsIdentity(tokenDuplicate);
                        impersonationContext = tempWindowsIdentity.Impersonate();
                        if (impersonationContext != null)
                        {
                            CloseHandle(token);
                            CloseHandle(tokenDuplicate);
                            return true;
                        }
                    }
                }
            }
            if (token != IntPtr.Zero)
                CloseHandle(token);
            if (tokenDuplicate != IntPtr.Zero)
                CloseHandle(tokenDuplicate);
            return false;
        }

        /// <summary>
        /// Método que hace el reverse de la Impersonalización, regresando al usuario con el que se ejecuta la Aplicación de forma normal.
        /// </summary>
        public void undoImpersonation()
        {
            impersonationContext.Undo();
        }
    }
}