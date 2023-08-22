using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Cartes;
using KeyiberboardException.Exceptions;
using KeyiberboardModels.Models;
using MiTools;
using RPABaseAPI;

namespace Blackfin_LibraryTemplate
{
    public class A3NomDocumentLibrary : MyCartesAPIBase
    {
        protected string workingFile = Environment.CurrentDirectory;
        private static bool loaded = false;
        RPAWin32Component salirUpdate = null, cancelarUpdate = null, centroCosto = null, ventanaUpdate = null, salirSistema = null;
        RPAWin32Component nroDocumento = null, salirCartas = null, rutaWindows = null, guardarPdf = null, cerrarPdf = null;         
        RPAWin32Component generarPdf = null, opcionCartas = null, salirComunicacion = null, buscarCarta = null, codCarta = null, aceptarCarta = null, generarCartas = null;
        RPAWin32Component datosPagador = null, checkListar = null, generarComunicado = null, ventanaPDF = null, ventanaGuardarComoPdf = null, rutaPdf = null;
        RPAWin32Component ventanaRetenciones = null, cancelarRetenciones = null, ventanaAfiliacion = null, noprocesar = null, aceptarAfiliacion = null;

        public A3NomDocumentLibrary(MyCartesProcess owner) : base(owner)
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
            loaded = cartes.merge(CurrentPath + "\\Cartes\\A3NomDocumentsTrabajador.cartes.rpa") == 1;
            ventanaUpdate = cartes.GetComponent<RPAWin32Component>("$VentanaUpdate");
            salirUpdate = cartes.GetComponent<RPAWin32Component>("$Salir");
            cancelarUpdate = cartes.GetComponent<RPAWin32Component>("$Cancelar");
            centroCosto = cartes.GetComponent<RPAWin32Component>("$CentroTrabajo");
            nroDocumento = cartes.GetComponent<RPAWin32Component>("$NroDocumento");
            salirSistema = cartes.GetComponent<RPAWin32Component>("$SalirSistema");
            datosPagador = cartes.GetComponent<RPAWin32Component>("$DatosPagador");
            checkListar = cartes.GetComponent<RPAWin32Component>("$CheckListar");
            generarComunicado = cartes.GetComponent<RPAWin32Component>("$GenerarComunicado");
            ventanaPDF = cartes.GetComponent<RPAWin32Component>("$VentanaPDF");
            ventanaGuardarComoPdf = cartes.GetComponent<RPAWin32Component>("$VentanaGuardarComoPdf");
            rutaPdf = cartes.GetComponent<RPAWin32Component>("$RutaPdf");
            guardarPdf = cartes.GetComponent<RPAWin32Component>("$GuardarPdf");
            cerrarPdf = cartes.GetComponent<RPAWin32Component>("$CerrarPdf");
            salirComunicacion = cartes.GetComponent<RPAWin32Component>("$SalirComunicacion");
            buscarCarta = cartes.GetComponent<RPAWin32Component>("$BuscarCarta");
            codCarta = cartes.GetComponent<RPAWin32Component>("$CodCarta");
            aceptarCarta = cartes.GetComponent<RPAWin32Component>("$AceptarCarta");
            generarPdf = cartes.GetComponent<RPAWin32Component>("$GenerarPdf");
            salirCartas = cartes.GetComponent<RPAWin32Component>("$SalirCartas");
            rutaWindows = cartes.GetComponent<RPAWin32Component>("$RutaWindows");
            generarCartas = cartes.GetComponent<RPAWin32Component>("$GenerarCartas");
            opcionCartas = cartes.GetComponent<RPAWin32Component>("$OpcionCartas");
            ventanaRetenciones = cartes.GetComponent<RPAWin32Component>("$VentanaRetenciones");
            cancelarRetenciones = cartes.GetComponent<RPAWin32Component>("$CancelarRetenciones");
            ventanaAfiliacion = cartes.GetComponent<RPAWin32Component>("$VentanaAfiliacion");
            noprocesar = cartes.GetComponent<RPAWin32Component>("$Noprocesar");
            aceptarAfiliacion = cartes.GetComponent<RPAWin32Component>("$AceptarAfiliacion");
        }

        /// <summary>
        /// Metodo donde se encuentra toda la logica del Bot
        /// </summary>
        /// <param name="workFolders">Objeto con parametros generales</param>
        /// <param name="employee">Objeto con los datos del empleado cargados desde el excel de insumo</param>
        /// <returns></returns>
        /// <exception cref="A3NomDocumentException">Se dispara una excepcion del tipo documento</exception>
        public string ExecuteMainProcess(ProcessWorkFolders workFolders, EmployeeWork employee)
        {
            string respuesta = string.Empty;

            try
            {
                cartes.reset("win32");
                Thread.Sleep(workFolders.tEspera);
                //Se valida que la ventana con los datos del trabajador este disponible
                if (ventanaUpdate.ComponentExist(workFolders.tEsperaComp))
                {
                    centroCosto.waitforcomponent(workFolders.tEsperaComp);
                    nroDocumento.waitforcomponent(workFolders.tEsperaComp);
                    //Se valida que los campos no esten disponibles para saber que guardo bien
                    if (centroCosto.enabled() == 0 && nroDocumento.enabled() == 0)
                    {
                        try
                        {
                            //Se valida y cierra cualquier proceso de chrome que este en memoria
                            foreach (var process in Process.GetProcessesByName("chrome"))
                            {
                                process.Kill();
                            }
                        }
                        catch { }
                        Thread.Sleep(workFolders.tEspera);

                        //Se valida si alguno de los reportes es solicitado en la plantilla de insumo
                        if (employee.PayerData != string.Empty || employee.ListContract != string.Empty ||
                            employee.ResidenceDeclare != string.Empty || employee.IncomeIRPF != string.Empty)
                        {
                            string filename = string.Empty;
                            //Se hace un recorrido hasta 4 ya que son el numero de documentos acordados
                            //Se realiza la impresion y guardado de los mismos
                            for (int i = 0; i < 4; i++)
                            {
                                switch (i)
                                {
                                    case 0://Comunic Datos Pagador - 145
                                        if (employee.PayerData != string.Empty)
                                        {
                                            OpenGenerateLetter(workFolders);
                                            datosPagador.waitforcomponent(workFolders.tEsperaComp);
                                            datosPagador.click();
                                            Thread.Sleep(workFolders.tEspera);
                                            if (checkListar.ComponentExist(workFolders.tEsperaComp))
                                            {
                                                checkListar.focus();
                                                //if (cartes.Execute("$CheckListar.Checked()").Trim() == "0")
                                                //    checkListar.click();
                                                Thread.Sleep(workFolders.tEspera);

                                                generarComunicado.waitforcomponent(workFolders.tEsperaComp);
                                                generarComunicado.focus();
                                                generarComunicado.click();

                                                filename = employee.DNI + "-" + "Comunic Datos Pagador - 145.pdf";
                                                if (!WaitDocument(filename, employee.Client, employee.Codigo + "-" + employee.Name.Replace(",", string.Empty), workFolders.rutaClientes, workFolders))
                                                {

                                                }
                                                salirComunicacion.waitforcomponent(workFolders.tEsperaComp);
                                                salirComunicacion.focus();
                                                salirComunicacion.click();
                                                Thread.Sleep(workFolders.tEspera);
                                            }
                                        }
                                        break;
                                    case 1:
                                        break;
                                    case 2://Declaracion de residencia fiscal
                                    case 3://Solicitud de IRPF voluntario
                                        if (employee.ResidenceDeclare != string.Empty || employee.IncomeIRPF != string.Empty)
                                        {
                                            string codigo = string.Empty;
                                            if (i == 2)
                                            {
                                                if (employee.ResidenceDeclare != string.Empty)
                                                {
                                                    codigo = "41";
                                                    filename = employee.DNI + "-" + "Declaracion de residencia fiscal.pdf";
                                                }
                                            }
                                            if (i == 3)
                                            {
                                                if (employee.IncomeIRPF != string.Empty)
                                                {
                                                    codigo = "34";
                                                    filename = employee.DNI + "-" + "Solicitud de IRPF voluntario.pdf";
                                                }
                                            }

                                            if (!string.IsNullOrEmpty(codigo))
                                            {
                                                OpenGenerateLetter(workFolders);
                                                opcionCartas.waitforcomponent(workFolders.tEsperaComp);
                                                opcionCartas.focus();
                                                opcionCartas.click();
                                                Thread.Sleep(workFolders.tEspera);
                                                if (codCarta.ComponentNotExist(10))
                                                {
                                                    buscarCarta.waitforcomponent(workFolders.tEsperaComp);
                                                    buscarCarta.focus();
                                                    buscarCarta.click();
                                                    Thread.Sleep(workFolders.tEspera);

                                                    codCarta.waitforcomponent(workFolders.tEsperaComp);
                                                    codCarta.focus();
                                                    codCarta.Press(46, 3);
                                                    Thread.Sleep(workFolders.tEspera);
                                                }

                                                for (int j = 0; j < codigo.Length; j++)
                                                {
                                                    codCarta.TypeKey(codigo.Substring(j, 1));
                                                }
                                                Thread.Sleep(workFolders.tEspera);
                                                aceptarCarta.waitforcomponent(workFolders.tEsperaComp);
                                                aceptarCarta.click();
                                                Thread.Sleep(workFolders.tEspera);

                                                generarPdf.waitforcomponent(workFolders.tEsperaComp);
                                                generarPdf.focus();
                                                generarPdf.click();

                                                if (!WaitDocument(filename, employee.Client, employee.Codigo + "-" + employee.Name.Replace(",", string.Empty), workFolders.rutaClientes, workFolders))
                                                {

                                                }
                                                salirCartas.waitforcomponent(workFolders.tEsperaComp);
                                                salirCartas.focus();
                                                salirCartas.click();
                                                Thread.Sleep(workFolders.tEspera);
                                            }
                                        }
                                        break;
                                }
                            }
                        }
                    }
                }
                CompletarSalida(workFolders);

            }
            catch (Exception e)
            {
                cartes.forensic("A3NomDocumentLibrary - ExecuteMainProcess - Exception: " + e.ToString());
                throw new A3NomDocumentException(e.StackTrace);
            }

            return respuesta;
        }

        /// <summary>
        /// Se centraliza la apertura de la ventana de cartas o reportes
        /// </summary>
        void OpenGenerateLetter(ProcessWorkFolders workFolders)
        {
            generarCartas.waitforcomponent(workFolders.tEsperaComp);
            generarCartas.focus();
            generarCartas.click();
            Thread.Sleep(workFolders.tEspera);
        }

        /// <summary>
        /// Metodo usado para centralizar la salida del sistema
        /// </summary>
        void CompletarSalida(ProcessWorkFolders workFolders)
        {
            salirUpdate.waitforcomponent(workFolders.tEsperaComp);
            salirUpdate.click();
            Thread.Sleep(workFolders.tEspera);

            while (ventanaRetenciones.ComponentExist(workFolders.tEsperaComp))
            {
                cancelarRetenciones.waitforcomponent(workFolders.tEsperaComp);
                cancelarRetenciones.click();
                Thread.Sleep(workFolders.tEspera);
            }
            while (ventanaAfiliacion.ComponentExist(workFolders.tEsperaComp))
            {
                noprocesar.waitforcomponent(workFolders.tEsperaComp);
                noprocesar.click();
                Thread.Sleep(workFolders.tEspera);

                aceptarAfiliacion.waitforcomponent(workFolders.tEsperaComp);
                aceptarAfiliacion.click();
                Thread.Sleep(workFolders.tEspera);
            }
            //salirSistema.waitforcomponent(workFolders.tEsperaComp);
            //salirSistema.click();
            //Thread.Sleep(workFolders.tEspera);
        }

        /// <summary>
        /// Metodo que se encarga de esperar que se cargue cualquier reporte en pdf y se encarga de guardarlo en la ruta del empleado
        /// </summary>
        /// <param name="filename">Nomnre del archivo</param>
        /// <param name="cliente">Cliente al que pertenece</param>
        /// <param name="trabajador">Trabajador al que pertenece</param>
        /// <param name="ruta">Ruta compartida</param>
        /// <returns></returns>
        bool WaitDocument(string filename, string cliente, string trabajador, string ruta, ProcessWorkFolders workFolders)
        {
            int contador = 0;
            bool bGenerado = false;
            string cadenaRuta = string.Empty;
            Thread.Sleep(10000);

            //Se inicia un contador en caso de que supere un tiempo muy largo, se estima unos 5 minutos
            while (contador < 1000)
            {
                //Se valida que la ventana de Chrome con el pdf esta cargada 
                if (ventanaPDF.ComponentExist(workFolders.tEsperaComp))
                {
                    //Se jecuta el Ctrl + S para guardar el documento
                    ventanaPDF.TypeKey("83", "Control", null);
                    Thread.Sleep(workFolders.tEspera);

                    //Se valida que la ventana de guardar como este disponible
                    if (ventanaGuardarComoPdf.ComponentExist(workFolders.tEsperaComp))
                    {
                        //Se valida que el campo para colocar la ruta de guardado este disponible
                        Thread.Sleep(workFolders.tEspera);
                        if (rutaPdf.ComponentExist(workFolders.tEsperaComp))
                        {
                            cadenaRuta = $@"{ruta}\{cliente}\ROBOT\{trabajador}\{filename}";
                            //Se guarda en el clipboard la ruta ya que el campo no permitio la escritura
                            Clipboard.SetText(cadenaRuta);
                            //Se valida si el archivo existe, y si es asi se elimina
                            if (File.Exists(cadenaRuta))
                                File.Delete(cadenaRuta);

                            rutaPdf.focus();
                            rutaPdf.Press(46, 10);
                            //Se ejecuta el ctrl + V para pegar lo que esta almacenado en el clipboard
                            rutaPdf.TypeKey("86", "Control", null);
                            rutaPdf.click();
                            Thread.Sleep(workFolders.tEspera);
                            //Se limpia el clipboard
                            Clipboard.SetText(" ");

                            //Se preciona enter para que guarde el documento
                            rutaPdf.Press(13, 1);
                            Thread.Sleep(workFolders.tEspera);

                            //Se cierra la ventana de Chrome
                            cerrarPdf.waitforcomponent(workFolders.tEsperaComp);
                            cerrarPdf.click();
                            Thread.Sleep(workFolders.tEspera);
                            bGenerado = true;
                            break;
                        }
                    }
                }
                Thread.Sleep(1000);
                contador++;
            }
            return bGenerado;
        }
    }
}