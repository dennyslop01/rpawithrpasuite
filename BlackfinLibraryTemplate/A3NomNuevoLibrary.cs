using Cartes;
using KeyiberboardException.Exceptions;
using KeyiberboardModels.Models;
using MiTools;
using RPABaseAPI;
using System;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.Windows.Forms;

namespace A3NomLibrary
{
    public class A3NomNuevoLibrary : MyCartesAPIBase
    {
        protected string workingFile = Environment.CurrentDirectory;
        private static bool loaded = false;
        RPAWin32Component aceptarNuevo = null, cancelarNuevo = null, centroCosto = null, codTrabajador = null;
        RPAWin32Component nombreTrabajador = null, nroDocumento = null, tipoDocumento = null, ventanaNuevo = null, nuevoTrabajador = null;
        RPAWin32Component codCentroCosto = null, listaCentroCosto = null, aceptarCentroCosto = null, cancelarCentroCosto = null, selectCentroCosto = null;
        RPAWin32Component informaError = null, aceptarError = null, cancelarTrabajador = null, salirSistema = null;
        RPAWin32Component ventanaCopiar = null, tablaTrabajadores = null, copiarTrabajador = null, cancelarCopia = null;

        public A3NomNuevoLibrary(MyCartesProcess owner) : base(owner)
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
            loaded = cartes.merge(CurrentPath + "\\Cartes\\A3NomCrearBaseTrabajador.cartes.rpa") == 1;
            aceptarNuevo = cartes.GetComponent<RPAWin32Component>("$AceptarNew");
            cancelarNuevo = cartes.GetComponent<RPAWin32Component>("$CancelarNew");
            centroCosto = cartes.GetComponent<RPAWin32Component>("$CentroCostoNew");
            codTrabajador = cartes.GetComponent<RPAWin32Component>("$CodTrabajadorNew");
            nombreTrabajador = cartes.GetComponent<RPAWin32Component>("$NombreTrabajadorNew");
            nroDocumento = cartes.GetComponent<RPAWin32Component>("$NroDocumentoNew");
            tipoDocumento = cartes.GetComponent<RPAWin32Component>("$TipoDocumentoNew");
            ventanaNuevo = cartes.GetComponent<RPAWin32Component>("$VentanaNuevoNew");
            informaError = cartes.GetComponent<RPAWin32Component>("$InformacionErrorNew");
            aceptarError = cartes.GetComponent<RPAWin32Component>("$AceptarErrorNew");
            nuevoTrabajador = cartes.GetComponent<RPAWin32Component>("$NuevoTrabajadorNew");

            codCentroCosto = cartes.GetComponent<RPAWin32Component>("$CodCentroCostoNew");
            listaCentroCosto = cartes.GetComponent<RPAWin32Component>("$ListaCentroCostoNew");
            aceptarCentroCosto = cartes.GetComponent<RPAWin32Component>("$AceptarCentroCostoNew");
            cancelarCentroCosto = cartes.GetComponent<RPAWin32Component>("$CancelarCentroCostoNew");
            selectCentroCosto = cartes.GetComponent<RPAWin32Component>("$SelectCentroCostoNew");

            cancelarTrabajador = cartes.GetComponent<RPAWin32Component>("$CancelarTrabajadorNew");
            salirSistema = cartes.GetComponent<RPAWin32Component>("$SalirSistemaNew");

            ventanaCopiar = cartes.GetComponent<RPAWin32Component>("$VentanaCopiarNew"); 
            tablaTrabajadores = cartes.GetComponent<RPAWin32Component>("$TablaTrabajadoresNew");
            copiarTrabajador = cartes.GetComponent<RPAWin32Component>("$CopiarTrabajadorNew");
            cancelarCopia = cartes.GetComponent<RPAWin32Component>("$CancelarCopiaNew");
        }

        /// <summary>
        /// Metodo donde se encuentra toda la logica del Bot
        /// </summary>
        /// <param name="workFolders">Objeto con parametros generales</param>
        /// <param name="employee">Objeto con los datos del empleado cargados desde el excel de insumo</param>
        /// <returns></returns>
        /// <exception cref="A3NomNuevoException">Se retorna un tipo de excepcion nuevo</exception>
        public string ExecuteMainProcess(ProcessWorkFolders workFolders, EmployeeWork employee)
        {
            string respuesta = string.Empty;
            bool bCreado = false;
            int contador = 0;

            try
            {
                while (!bCreado)
                {
                    cartes.reset("win32");
                    Thread.Sleep(workFolders.tEspera);
                    //Se valida que la ventana de nuevo trabajador se encuetre cargada y activa
                    if (ventanaNuevo.ComponentExist(workFolders.tEsperaComp))
                    {
                        codTrabajador.waitforcomponent(workFolders.tEsperaComp);
                        codTrabajador.focus();
                        if(codTrabajador.Value == null)
                        {
                            for (int i = 0; i < employee.CodigoAux.Length; i++)
                            {
                                codTrabajador.TypeKey(employee.CodigoAux.Substring(i, 1));
                            }
                        }
                        //Se toma el correlativo que sugiere el sistema y se espera a ver si se tiene que hacer un cambio
                        string codTrabajaAux = codTrabajador.Value.Trim();
                        //Se valida si el campo codigo trabajador de la plantilla viene con algun valor
                        if (employee.Codigo.Length == 0)
                        {
                            employee.Codigo = codTrabajaAux;
                        }
                        else
                        {
                            //Si el campo viene con alguna letra en la plantilla se toma el campo sugerido y se le coloca la letra en la primera posicion
                            //employee.Codigo += codTrabajaAux.Substring(1, (codTrabajaAux.Length - 1));
                            codTrabajador.Press(8, 10);
                            for (int i = 0; i < employee.Codigo.Length; i++)
                            {
                                codTrabajador.TypeKey(employee.Codigo.Substring(i, 1));
                            }
                            codTrabajador.Press(9, 1);
                            Thread.Sleep(workFolders.tEspera);
                        }

                        //Se completan los demas campos del formulario
                        nombreTrabajador.waitforcomponent(workFolders.tEsperaComp);
                        nombreTrabajador.focus();
                        nombreTrabajador.click();
                        for (int i = 0; i < employee.Name.Length; i++)
                        {
                            if(employee.Name.Substring(i, 1)!= "Ñ")
                                nombreTrabajador.TypeKey(employee.Name.Substring(i, 1));
                            else
                            {
                                Clipboard.SetText("Ñ");
                                nombreTrabajador.TypeKey("86", "Control", null);
                                Clipboard.SetText(" ");
                            }
                        }
                        nombreTrabajador.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        tipoDocumento.waitforcomponent(workFolders.tEsperaComp);
                        tipoDocumento.Value = employee.DNIType;
                        Thread.Sleep(workFolders.tEspera);

                        nroDocumento.waitforcomponent(workFolders.tEsperaComp);
                        nroDocumento.focus();
                        nroDocumento.click();
                        for (int i = 0; i < employee.DNI.Length - 1; i++)
                        {
                            nroDocumento.TypeKey(employee.DNI.Substring(i, 1));
                        }
                        nroDocumento.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        //Luego que se coloca el nro del DNI se valida si el digito validador coincide con el que sugiere el sistema
                        string newNumero = nroDocumento.Value;
                        if (newNumero != employee.DNI)
                        {
                            ManejarSalida(false, workFolders);
                            return $"Error no coincide el digito validador del número de documento, para el trabajador de Nombre: {employee.Name}, Nro Doc: {employee.DNI} para el CLiente: {employee.Client}"; //Error Letra Validadora DNI
                        }

                        //Se completa el campo centro de costo buscando dentro de la ventana de codigos
                        selectCentroCosto.waitforcomponent(30);
                        selectCentroCosto.doubleclick(1);
                        Thread.Sleep(workFolders.tEspera);

                        codCentroCosto.waitforcomponent(workFolders.tEsperaComp);
                        codCentroCosto.click();
                        for (int i = 0; i <= employee.CodCenterWork.Length - 1; i++)
                        {
                            codCentroCosto.TypeKey(employee.CodCenterWork.Substring(i, 1));
                        }
                        Thread.Sleep(workFolders.tEspera);

                        bool existCC = false;
                        listaCentroCosto.waitforcomponent(workFolders.tEsperaComp);
                        for (int i = 0; i < listaCentroCosto.descendants; i++)
                        {
                            listaCentroCosto.dochild($@"\{i}", "click", "");
                            string[] cadena = listaCentroCosto.dochild($@"\{i}", "name", "").Split(Convert.ToChar(" "));
                            if (cadena.Length > 0)
                            {
                                if (int.Parse(cadena[0].Trim()) == int.Parse(employee.CodCenterWork))
                                {
                                    existCC = true;
                                    break;
                                }
                            }
                        }
                        if (existCC)
                        {
                            aceptarCentroCosto.waitforcomponent(workFolders.tEsperaComp);
                            aceptarCentroCosto.click();
                        }
                        else
                        {
                            ManejarSalida(true, workFolders);
                            return $"Error no existe Centro de Trabajo: {employee.CodCenterWork}, para el trabajador de Nombre: {employee.Name}, Nro Doc: {employee.DNI} para el CLiente: {employee.Client}";
                        }
                        Thread.Sleep(workFolders.tEspera);

                        //Si todo esta OK se completa el guardado
                        aceptarNuevo.waitforcomponent(workFolders.tEsperaComp);
                        aceptarNuevo.click();
                        Thread.Sleep(workFolders.tEspera);

                        //Si existe algun error al guardar se reporta y se termina el proceso con ese trabajador
                        if (informaError.ComponentExist(workFolders.tEsperaComp))
                        {
                            aceptarError.waitforcomponent(workFolders.tEsperaComp);
                            aceptarError.click();
                            Thread.Sleep(workFolders.tEspera);

                            contador++;
                            if (contador > 2)
                            {
                                ManejarSalida(false, workFolders);
                                return $"Error al intertar crear el trabajador de Nombre: {employee.Name}, Nro Doc: {employee.DNI} para el CLiente: {employee.Client}";
                            }
                            cancelarNuevo.waitforcomponent(workFolders.tEsperaComp);
                            cancelarNuevo.click();
                            Thread.Sleep(workFolders.tEspera);

                            nuevoTrabajador.waitforcomponent(workFolders.tEsperaComp);
                            nuevoTrabajador.click();
                            Thread.Sleep(workFolders.tEspera);
                        }
                        else
                        {
                            //Si no existe ningun error pero el trabajador pertenecio a la empresa en el pasado se revisa la lista que se presenta
                            if (ventanaCopiar.ComponentExist(workFolders.tEsperaComp))
                            {
                                bool existTra = false;
                                tablaTrabajadores.waitforcomponent(workFolders.tEsperaComp);
                                for (int i = 0; i < tablaTrabajadores.descendants; i++)
                                {
                                    tablaTrabajadores.dochild($@"\{i}", "click", "");
                                    string[] cadena = tablaTrabajadores.dochild($@"\{i}", "name", "").Split(Convert.ToChar(" "));
                                    if (cadena.Length > 0)
                                    {
                                        string nombreTrabaja = string.Empty;
                                        for (int j = 6; j < 11; j++)
                                        {
                                            nombreTrabaja += cadena[j] + " ";
                                        }

                                        //Si coincide alguno en cuanto al DNI y el Nombre se toma y se copian los datos
                                        //En caso contrario se reporta el error se cancela la copia y se pasa a terminar de cargar los datos
                                        if (cadena[0].Trim() == employee.DNI.Trim() && nombreTrabaja.Trim() == employee.Name.ToUpper().Trim())
                                        {
                                            existTra = true;
                                            break;
                                        }
                                    }
                                }
                                if (existTra)
                                {
                                    copiarTrabajador.waitforcomponent(workFolders.tEsperaComp);
                                    copiarTrabajador.click();
                                }
                                else
                                {
                                    cancelarCopia.waitforcomponent(workFolders.tEsperaComp);
                                    cancelarCopia.click();

                                    respuesta = $"Mantenimiento-Error existen trabajadores con el mismo DNI pero distinto nombre, al intentar crear el trabajador de Nombre: {employee.Name}, Nro Doc: {employee.DNI} para el CLiente: {employee.Client}";
                                }
                                Thread.Sleep(10000);
                            }
                         
                            bCreado = true;
                            respuesta = "Mantenimiento";
                        }
                    }
                }
            }
            catch (Exception e)
            {
                cartes.forensic("A3NomNuevoLibrary - ExecuteMainProcess - Exception: " + e.ToString());
                throw new A3NomNuevoException(e.StackTrace);
            }

            return respuesta;
        }

        /// <summary>
        /// Metodo usado para centralizar la salida del sistema
        /// </summary>
        void ManejarSalida(bool centroCosto, ProcessWorkFolders workFolders)
        {
            if(centroCosto)
            {
                cancelarCentroCosto.waitforcomponent(workFolders.tEsperaComp);
                cancelarCentroCosto.click();
                Thread.Sleep(workFolders.tEspera);
            }

            cancelarNuevo.waitforcomponent(workFolders.tEsperaComp);
            cancelarNuevo.click();
            Thread.Sleep(workFolders.tEspera);

            cancelarTrabajador.waitforcomponent(workFolders.tEsperaComp);
            cancelarTrabajador.click();
            Thread.Sleep(workFolders.tEspera);

            //salirSistema.waitforcomponent(workFolders.tEsperaComp);
            //salirSistema.click();
            //Thread.Sleep(workFolders.tEspera);
        }
    }
}