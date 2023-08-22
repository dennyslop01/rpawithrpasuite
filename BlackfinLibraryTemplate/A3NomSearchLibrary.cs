using Cartes;
using KeyiberboardException.Exceptions;
using KeyiberboardModels.Models;
using MiTools;
using RPABaseAPI;
using System;
using System.Threading;

namespace A3NomLibrary
{
    public class A3NomSearchLibrary : MyCartesAPIBase
    {
        protected string workingFile = Environment.CurrentDirectory;
        private static bool loaded = false;
        RPAWin32Component aceptarCliente = null, aceptarTrabajador = null, cancelarCliente = null, cancelarTrabajador = null;
        RPAWin32Component datosTrabajador = null, ingresarCliente = null, ingresarTrabajador = null, nifTrabajador = null;
        RPAWin32Component alertaNoExisteTrabajador = null, aceptarNoExisteTrabajador = null, tablaTrabajadores = null, codigoTrabajador = null;
        RPAWin32Component nombreCliente = null, tablaClientes = null, codCliente = null, codClienteDisplay = null, ultimoRegistro = null;
        RPAWin32Component nuevoTrabajador = null, nuevoVentana = null, mantenimientoVentana = null, salirSistema = null, subirBusqueda = null;

        public A3NomSearchLibrary(MyCartesProcess owner) : base(owner)
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
            loaded = cartes.merge(CurrentPath + "\\Cartes\\A3NomBuscarClienteTrabajador.cartes.rpa") == 1;
            datosTrabajador = cartes.GetComponent<RPAWin32Component>("$DatosTrabajador");

            codCliente = cartes.GetComponent<RPAWin32Component>("$CodCliente");
            ingresarCliente = cartes.GetComponent<RPAWin32Component>("$IngresarCliente");
            tablaClientes = cartes.GetComponent<RPAWin32Component>("$TablaClientes");
            aceptarCliente = cartes.GetComponent<RPAWin32Component>("$AceptarCliente");
            cancelarCliente = cartes.GetComponent<RPAWin32Component>("$CancelarCliente");
            codClienteDisplay = cartes.GetComponent<RPAWin32Component>("$CodClienteDisplay");
            nombreCliente = cartes.GetComponent<RPAWin32Component>("$NombreCliente");

            codigoTrabajador = cartes.GetComponent<RPAWin32Component>("$CodigoTrabajador");
            ingresarTrabajador = cartes.GetComponent<RPAWin32Component>("$IngresarTrabajador");
            nifTrabajador = cartes.GetComponent<RPAWin32Component>("$NifTrabajador");
            aceptarTrabajador = cartes.GetComponent<RPAWin32Component>("$AceptarTrabajador");
            cancelarTrabajador = cartes.GetComponent<RPAWin32Component>("$CancelarTrabajador");
            alertaNoExisteTrabajador = cartes.GetComponent<RPAWin32Component>("$AlertaNoExisteTrabajador");
            aceptarNoExisteTrabajador = cartes.GetComponent<RPAWin32Component>("$AceptarNoExisteTrabajador");
            tablaTrabajadores = cartes.GetComponent<RPAWin32Component>("$TablaTrabajadores");

            nuevoTrabajador = cartes.GetComponent<RPAWin32Component>("$NuevoTrabajador");
            nuevoVentana = cartes.GetComponent<RPAWin32Component>("$NuevoTrabajadorVentana");
            mantenimientoVentana = cartes.GetComponent<RPAWin32Component>("$MantenimientoTrabajadorVentana");
            salirSistema = cartes.GetComponent<RPAWin32Component>("$SalirSistema");
            ultimoRegistro = cartes.GetComponent<RPAWin32Component>("$UltimoRegistro");    
            subirBusqueda = cartes.GetComponent<RPAWin32Component>("$SubirBusqueda");
        }

        /// <summary>
        /// Metodo donde se encuentra toda la logica del Bot
        /// </summary>
        /// <param name="workFolders">Objeto con parametros generales</param>
        /// <param name="employee">Objeto con los datos del empleado cargados desde el excel de insumo</param>
        /// <returns></returns>
        /// <exception cref="A3NomSearchException">Retorna una excepcion del tipo buscar</exception>
        public string ExecuteMainProcess(ProcessWorkFolders workFolders, EmployeeWork employee)
        {
            string respuesta = string.Empty;
            try
            {
                cartes.reset("win32");
                Thread.Sleep(workFolders.tEspera);
                //Se debe validar que el cliete que viene en la plantilla exista en el sistema
                //En caso de no existir se reporta el error y se concluye con el trabajador
                datosTrabajador.waitforcomponent(workFolders.tEsperaComp);
                datosTrabajador.click();
                Thread.Sleep(workFolders.tEspera);

                codCliente.waitforcomponent(workFolders.tEsperaComp);
                codCliente.doubleclick(1);
                for (int i = 0; i < employee.CodClient.Length; i++)
                {
                    codCliente.TypeKey(employee.CodClient.Substring(i, 1));
                }
                Thread.Sleep(workFolders.tEspera);

                tablaClientes.waitforcomponent(workFolders.tEsperaComp);
                int nroClientes = tablaClientes.descendants;
                bool existCliente = false;
                for (int i = 0; i < nroClientes; i++)
                {
                    tablaClientes.dochild($@"\{i}", "click", "");
                    string[] cadena = tablaClientes.dochild($@"\{i}", "name", "").Split(Convert.ToChar(" "));
                    if (cadena.Length > 0)
                    {
                        if (int.Parse(cadena[0].Trim()) == int.Parse(employee.CodClient))
                        {
                            existCliente = true;
                            break;
                        }
                    }
                }

                if (existCliente)
                {
                    aceptarCliente.waitforcomponent(workFolders.tEsperaComp);
                    aceptarCliente.click();
                    Thread.Sleep(workFolders.tEspera);

                    codigoTrabajador.waitforcomponent(workFolders.tEsperaComp);
                    codigoTrabajador.doubleclick(1);
                    bool existTrabaja = false;

                    //Pensando en futuro se realizara busqueda del trabajador si la longitud del codigo es mayor a 3
                    if (employee.Codigo.Length > 3)
                    {
                        for (int i = 0; i < employee.Codigo.Length; i++)
                        {
                            codigoTrabajador.TypeKey(employee.Codigo.Substring(i, 1));
                        }
                        Thread.Sleep(workFolders.tEspera);

                        int nroTrabaja = tablaTrabajadores.descendants;

                        for (int i = 0; i < nroTrabaja; i++)
                        {
                            tablaTrabajadores.dochild($@"\{i}", "click", "");
                            string[] cadena = tablaTrabajadores.dochild($@"\{i}", "name", "").Split(Convert.ToChar(" "));
                            if (cadena.Length > 0)
                            {
                                string nombreTrabaja = string.Empty;
                                for (int j = 1; j < 6; j++)
                                {
                                    nombreTrabaja += cadena[j] + " ";
                                }

                                if (cadena[0].Trim() == employee.Codigo)
                                {
                                    if (nombreTrabaja.Trim().ToUpper() == employee.Name.ToUpper())
                                        existTrabaja = true;
                                    else
                                    {
                                        cancelarTrabajador.waitforcomponent(workFolders.tEsperaComp);
                                        cancelarTrabajador.focus();
                                        cancelarTrabajador.click();
                                        Thread.Sleep(workFolders.tEspera);

                                        //salirSistema.waitforcomponent(workFolders.tEsperaComp);
                                        //salirSistema.click();
                                        return $"Error el codigo trabajador {employee.Codigo}, existe con el nombre {nombreTrabaja.Trim()} y no con el nombre {employee.Name}";
                                    }
                                    break;
                                }
                            }
                        }
                    }

                    if (existTrabaja)
                    {                        
                        aceptarTrabajador.waitforcomponent(workFolders.tEsperaComp);
                        aceptarTrabajador.click();
                        Thread.Sleep(workFolders.tEspera);
                    }
                    else
                    {
                        string scodigo = "99999";
                        for (int i = 0; i < scodigo.Length; i++)
                        {
                            codigoTrabajador.TypeKey(scodigo.Substring(i, 1));
                        }
                        Thread.Sleep(workFolders.tEspera);

                        ultimoRegistro.waitforcomponent(workFolders.tEsperaComp);
                        ultimoRegistro.focus();
                        string registro = ultimoRegistro.name().Trim();
                        string[] cadena = registro.Split(' ');
                        int numero = 0;
                        int cont = 0;
                        while (numero == 0)
                        {
                            cont++;
                            try
                            {
                                numero = int.Parse(cadena[0].Trim());
                            }
                            catch
                            {
                                if(cont>1)
                                {
                                    subirBusqueda.waitforcomponent(workFolders.tEsperaComp);
                                    subirBusqueda.focus();
                                    subirBusqueda.click(4);
                                }
                                int nroTrabaja = tablaTrabajadores.descendants;

                                for (int i = nroTrabaja - 1; i > 0; i--)
                                {
                                    tablaTrabajadores.dochild($@"\{i}", "click", "");
                                    string[] cadena1 = tablaTrabajadores.dochild($@"\{i}", "name", "").Split(Convert.ToChar(" "));
                                    if (cadena1.Length > 0)
                                    {
                                        try
                                        {
                                            numero = int.Parse(cadena1[0].Trim());
                                            break;
                                        }
                                        catch { }
                                    }
                                }
                            }
                        }
                        employee.CodigoAux = (numero + 1).ToString("000000");

                        Thread.Sleep(workFolders.tEspera);

                        nuevoTrabajador.waitforcomponent(workFolders.tEsperaComp);
                        nuevoTrabajador.click();
                        Thread.Sleep(workFolders.tEspera);

                        return "Nuevo";
                    }

                    if (alertaNoExisteTrabajador.ComponentExist())
                    {
                        aceptarNoExisteTrabajador.waitforcomponent(workFolders.tEsperaComp);
                        aceptarNoExisteTrabajador.click();
                        Thread.Sleep(workFolders.tEspera);
                        return string.Empty;
                    }

                    if(mantenimientoVentana.ComponentExist())
                    {
                        return "Mantenimiento";
                    }                    
                }
                else
                {
                    cancelarCliente.waitforcomponent(workFolders.tEsperaComp);
                    cancelarCliente.click();

                    //salirSistema.waitforcomponent(workFolders.tEsperaComp);
                    //salirSistema.click();
                    return $"Error Cliente {employee.Client}, no existe";
                }
            }
            catch (Exception e)
            {
                cartes.forensic("A3NomSearchLibrary - ExecuteMainProcess - Exception: " + e.ToString());
                throw new A3NomSearchException(e.StackTrace);
            }

            return respuesta;
        }
    }
}
