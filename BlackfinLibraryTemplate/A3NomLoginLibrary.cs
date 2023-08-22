using Cartes;
using KeyiberboardException.Exceptions;
using KeyiberboardModels.Models;
using MiTools;
using RPABaseAPI;
using System;
using System.Diagnostics;
using System.Threading;

namespace A3NomLibrary
{
    public class A3NomLoginLibrary : MyCartesAPIBase
    {
        protected string workingFile = Environment.CurrentDirectory;
        private static bool loaded = false;
        RPAWin32Component aceptarA3nom = null, cancelarA3nom = null, passwA3nom = null, usuarioA3nom = null;
        RPAWin32Component alertaInicioSesion = null, alertaAceptar = null, botonDT = null, errorCobol = null, aceptarErrorCobol = null;
        RPAWin32Component selectUsuario = null, buscarUsuario = null, siguiente = null, aceptarBuscar = null, errorBusqueda = null, aceptarError = null, cancelarBuscar = null;

        public A3NomLoginLibrary(MyCartesProcess owner) : base(owner)
        {
        }

        public override void Close()
        {
            //throw new NotImplementedException();
        }

        /// <summary>
        /// Metodo que se encarga de instanciar el objeto RPA y cada uno de los controles hijos
        /// </summary>
        protected override void MergeLibrariesAndLoadVariables()
        {
            //example
            loaded = cartes.merge(CurrentPath + "\\Cartes\\A3nomLogin.cartes.rpa") == 1;
            aceptarA3nom = cartes.GetComponent<RPAWin32Component>("$AceptarA3nom");
            cancelarA3nom = cartes.GetComponent<RPAWin32Component>("$CancelarA3nom");
            passwA3nom = cartes.GetComponent<RPAWin32Component>("$PasswA3nom");
            usuarioA3nom = cartes.GetComponent<RPAWin32Component>("$UsuarioA3nom");
            alertaInicioSesion = cartes.GetComponent<RPAWin32Component>("$AlertaInicioSesion");
            alertaAceptar = cartes.GetComponent<RPAWin32Component>("$AlertaAceptar");
            botonDT = cartes.GetComponent<RPAWin32Component>("$BotonDT");

            selectUsuario = cartes.GetComponent<RPAWin32Component>("$SelectUsuario");
            buscarUsuario = cartes.GetComponent<RPAWin32Component>("$BuscarUsuario");
            siguiente = cartes.GetComponent<RPAWin32Component>("$Siguiente");
            aceptarBuscar = cartes.GetComponent<RPAWin32Component>("$AceptarBuscar");
            errorBusqueda = cartes.GetComponent<RPAWin32Component>("$ErrorBusqueda");
            aceptarError = cartes.GetComponent<RPAWin32Component>("$AceptarError");
            cancelarBuscar = cartes.GetComponent<RPAWin32Component>("$CancelarBuscar");
            errorCobol = cartes.GetComponent<RPAWin32Component>("$ErrorCobol");
            aceptarErrorCobol = cartes.GetComponent<RPAWin32Component>("$AceptarErrorCobol");
        }

        /// <summary>
        /// Metodo donde se encuentra toda la logica del Bot
        /// </summary>
        /// <param name="workFolders">Objeto con parametros generales</param>
        /// <param name="employee">Objeto con los datos del empleado cargados desde el excel de insumo</param>
        /// <returns></returns>
        /// <exception cref="A3NomLoginException">Se dispara una excepcion del tipo login</exception>
        public string ExecuteMainProcess(ProcessWorkFolders workFolders, EmployeeWork employee)
        {
            string respuesta = string.Empty;
            try
            {
                //Se valida si existe alguna instancia de A3-Nom esta en ejecucion y se detiene la tarea
                try
                {
                    foreach (var process in Process.GetProcessesByName(workFolders.a3NomExe))
                    {
                        process.Kill();
                    }
                }
                catch { }
                Thread.Sleep((workFolders.tEspera*2));
                int contador = 0;
                //Se valida antes de ingresar si quedo algun mensaje de error y se cierra
                if (errorCobol.ComponentExist(workFolders.tEsperaComp))
                {
                    aceptarErrorCobol.waitforcomponent(workFolders.tEsperaComp);
                    aceptarErrorCobol.focus();
                    aceptarErrorCobol.doubleclick(1);
                    Thread.Sleep(workFolders.tEspera);
                }

                //Se inicia un while para intentar ingresar al aplicativo hasta 3 veces
                while (true)
                {
                    //Se buscan y completan cada uno de los campos para hacer login
                    cartes.run(workFolders.rutaA3nom);
                    cartes.reset("win32");
                    Thread.Sleep(30000);

                    usuarioA3nom.waitforcomponent(workFolders.tEsperaComp);
                    usuarioA3nom.focus();
                    //Si el usuario que ingreso antes es distinto que esta en el archivo de configuracion se debe completar
                    if (Convert.ToString(usuarioA3nom.Value) != workFolders.a3nomUser)
                    {
                        usuarioA3nom.Press(8, 1);
                        usuarioA3nom.click();

                        //Se identifico que muchas veces no permite pegar el usuario y se debe buscar en una nueva ventana
                        selectUsuario.waitforcomponent(workFolders.tEsperaComp);
                        selectUsuario.click();
                        Thread.Sleep(workFolders.tEspera);

                        buscarUsuario.waitforcomponent(workFolders.tEsperaComp);
                        buscarUsuario.focus();
                        buscarUsuario.Value = workFolders.a3nomUser;
                        Thread.Sleep(workFolders.tEspera);

                        siguiente.waitforcomponent(workFolders.tEsperaComp);
                        siguiente.click();
                        Thread.Sleep(workFolders.tEspera);

                        if (errorBusqueda.ComponentExist(workFolders.tEsperaComp))
                        {
                            aceptarError.waitforcomponent(workFolders.tEsperaComp);
                            aceptarError.click();
                            Thread.Sleep(workFolders.tEspera);

                            cancelarBuscar.waitforcomponent(workFolders.tEsperaComp);
                            cancelarBuscar.click();
                            Thread.Sleep(workFolders.tEspera);

                            cancelarA3nom.waitforcomponent(workFolders.tEsperaComp);
                            cancelarA3nom.click();
                            Thread.Sleep(workFolders.tEspera);
                            return $"Error no existe el usuario: {workFolders.a3nomUser}";
                        }

                        aceptarBuscar.waitforcomponent(workFolders.tEsperaComp);
                        aceptarBuscar.click();
                        Thread.Sleep(workFolders.tEspera);
                    }

                    passwA3nom.waitforcomponent(workFolders.tEsperaComp);
                    passwA3nom.focus();
                    passwA3nom.click();
                    for (int i = 0; i < workFolders.a3nomPassword.Length; i++)
                    {
                        char caracter = char.Parse(workFolders.a3nomPassword.Substring(i, 1));
                        if (char.IsNumber(caracter))
                        {
                            passwA3nom.Press((int)caracter, 1);
                            passwA3nom.Press(8, 1);
                        }
                        if (char.IsUpper(caracter))
                            passwA3nom.TypeKey(caracter.ToString(), "Shift", null);
                        else
                            passwA3nom.TypeKey(caracter.ToString());
                    }
                    Thread.Sleep(workFolders.tEspera);

                    aceptarA3nom.waitforcomponent(workFolders.tEsperaComp);
                    aceptarA3nom.click();
                    Thread.Sleep(5000);

                    if (botonDT.ComponentExist())
                        break;

                    //Si supera el maximo de intentos se sale del sistema
                    if (contador > 1)
                    {
                        cancelarA3nom.waitforcomponent(workFolders.tEsperaComp);
                        cancelarA3nom.click();
                        return "Error Maximo intentos de login A3 nom, no toma las credenciales";
                    }

                    if (alertaInicioSesion.ComponentExist())
                    {
                        alertaAceptar.waitforcomponent(workFolders.tEsperaComp);
                        alertaAceptar.click();
                        Thread.Sleep(workFolders.tEspera);

                        cancelarA3nom.waitforcomponent(workFolders.tEsperaComp);
                        cancelarA3nom.click();
                    }
                    contador++;
                }

                if (errorCobol.ComponentExist(workFolders.tEsperaComp))
                {
                    aceptarErrorCobol.waitforcomponent(workFolders.tEsperaComp);
                    aceptarErrorCobol.focus();
                    aceptarErrorCobol.doubleclick(1);
                    Thread.Sleep(workFolders.tEspera);
                }

            }
            catch (Exception ex)
            {
                cartes.forensic("A3NomLoginLibrary - ExecuteMainProcess - Exception: " + ex.ToString());
                throw new A3NomLoginException(ex.StackTrace);
            }

            return respuesta;
        }
    }
}
