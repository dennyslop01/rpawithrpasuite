using A3NomLibrary;
using Blackfin_LibraryTemplate;
using BlackfinUtilesLib;
using BlackfinUtilesLib.Controllers;
using BlackfinUtilesLib.Models;
using DocumentFormat.OpenXml.Wordprocessing;
using KeyiberboardException.Exceptions;
using KeyiberboardModels.Models;
using RPABaseAPI;
using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using static System.Data.Entity.Infrastructure.Design.Executor;

namespace KeyiberboardAltaEmpleados
{
    public class MainProcess : MyCartesProcessBase
    {
        private enum estadoProceso
        {
            Devolver,
            Finalizado,
            Error,
            Proceso
        }
        private A3NomMainLibrary frpA3mainLibrary;
        private A3NomLoginLibrary frpA3loginLibrary;
        private A3NomSearchLibrary frpA3searchLibrary;
        private A3NomNuevoLibrary frpA3nuevoLibrary;
        private A3NomUpdateLibrary frpA3updateLibrary;
        private A3NomDocumentLibrary frpA3documentLibrary;

        private string appName = string.Empty;
        private int appId = 0;
        private string machine = Environment.MachineName;
        private const string fuente = "MainProcess";
        private string rowInit = string.Empty;
        private string lastProcess = string.Empty;
        private string lastStatus = string.Empty;
        private string nowYear = DateTime.Now.Year.ToString();
        private string nowMonth = DateTime.Now.ToString("MM");
        private string monthName = DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("es")).Substring(0,3).ToUpper();
        private ProcessWorkFolders workFolders = new ProcessWorkFolders();
        private RpaProcessController controller = new RpaProcessController();
        private int contador = 0;
        private int sustraDays = 0;
        private string usermail = null, passUserMail = null, hostMail = null, portMail = null, listMailNotification = null;
        private string htmlString = $@"<html>
                      <body>
                      <p>RPA Key Iberboard - Proceso ALta Empleados</p>
                      <p>%mesaje%</p>
                      <p>Saludos,<br>Blackfin Corporación</br></p>
                      </body>
                      </html>
                     ";
        private bool bEnvioCorreo;
        private string emailConsultor = string.Empty;
 
        protected override void DoExecute(ref DateTime start)
        {
            bool bLogin = false;

            //Se incia el proceso del Bot
            RegisterIteration(start, "Begin", $"RPA {appName} - Inicia Proceso", true);
            cartes.balloon("Iniciando el Proceso");

            Mails.SendMail(usermail, passUserMail, hostMail,
                    portMail, listMailNotification, $"{appName} - Alta Empleados - Inicio!",
                    htmlString.Replace("%mesaje%", "Ha iniciado el proceso de Alta Empleados, con fecha: " + DateTime.Now.ToString("dd.MM.yyyy")));

            WriteLogsProcess((int)estadoProceso.Proceso, "Ha iniciado el proceso de Alta Empleados");
            //El proceso se detendra si se producen 3 errores criticos consecutivos
            while (contador < 3)
            {
                try
                {
                    WriteLogsProcess((int)estadoProceso.Proceso, "Se valida en que estado se encuentra el proceso a nuvel de base de datos");
                    RpaProcess process = controller.GetByMachine(appName, machine);
                    if (process == null)
                    {
                        RpaProcess processCreate = new RpaProcess();
                        processCreate.Name = appName;
                        processCreate.Machine = machine;
                        processCreate.Process = fuente;
                        processCreate.Status = "Begin";
                        processCreate.ProcessDate = DateTime.Now;

                        controller.Create(processCreate);
                        process = controller.GetByMachine(appName, machine);
                        WriteLogsProcess((int)estadoProceso.Proceso, "El proceso no existia a nivel de base de datos y se crea");
                    }
                    else
                    {
                        appId = process.Id;
                        if (string.IsNullOrEmpty(process.RowInit))
                        {
                            rowInit = process.RowInit;
                            lastProcess = process.Process;
                            lastStatus = process.Status;
                        }

                        process.Process = fuente;
                        process.Status = "Begin";
                        process.ProcessDate = DateTime.Now;
                        process.RowInit = null;
                        process.Observation = null;

                        controller.Update(process);
                        WriteLogsProcess((int)estadoProceso.Proceso, "Se actualiza el proceso a nivel de base de datos");
                    }

                    EmployeeWork employeeWork = new EmployeeWork();
                    emailConsultor = string.Empty;
                    if (string.IsNullOrEmpty(rowInit))
                    {
                        string resultado = NewEmployee.Read(workFolders, ref employeeWork);
                        if (employeeWork.EmailConsultor != null)
                        {
                            emailConsultor = employeeWork.EmailConsultor.Trim();
                            WriteLogsProcess((int)estadoProceso.Proceso, $"Se realiza lectura del primer archivo encontrado para procesar, {employeeWork.rutaDoc}");

                            if (resultado.Contains("Error"))
                            {
                                EjecutarAccionesErrorControlado(resultado, start, false, emailConsultor);
                                bEnvioCorreo = true;
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(employeeWork.rutaDoc))
                                {
                                    if (!bLogin)
                                    {
                                        resultado = A3loginLib.ExecuteMainProcess(workFolders, employeeWork);
                                        WriteLogsProcess((int)estadoProceso.Proceso, "Se hace login en la aplicacion A3-nom");
                                        if (resultado.Contains("Error"))
                                        {
                                            EjecutarAccionesErrorControlado(resultado, start, true, emailConsultor);
                                            bEnvioCorreo = true;
                                            break;
                                        }
                                        else
                                            bLogin = true;
                                    }

                                    if (bLogin)
                                    {
                                        resultado = A3searchLib.ExecuteMainProcess(workFolders, employeeWork);
                                        WriteLogsProcess((int)estadoProceso.Proceso, $"Se realiza busqueda del trabajador: {employeeWork.DNIType}-{employeeWork.DNI} {employeeWork.Name}, para el cliente: {employeeWork.CodClient} - {employeeWork.Client}");
                                        if (resultado.Contains("Error"))
                                        {
                                            EjecutarAccionesErrorControlado(resultado, start, false, emailConsultor);
                                            bEnvioCorreo = true;
                                        }
                                        else if (resultado == "Nuevo")
                                        {
                                            resultado = A3nuevoLib.ExecuteMainProcess(workFolders, employeeWork);
                                            WriteLogsProcess((int)estadoProceso.Proceso, $"Se crea el trabajador: {employeeWork.DNIType}-{employeeWork.DNI} {employeeWork.Name}, para el cliente: {employeeWork.CodClient} - {employeeWork.Client}");

                                            if (resultado.Contains("Error"))
                                            {
                                                EjecutarAccionesErrorControlado(resultado.Replace("Mantenimiento-", " "), start, false, emailConsultor);
                                                bEnvioCorreo = true;

                                                if (resultado.Contains("Mantenimiento-"))
                                                    resultado = "Mantenimiento";
                                            }
                                        }

                                        if (resultado == "Mantenimiento")
                                        {
                                            resultado = A3updateLib.ExecuteMainProcess(workFolders, employeeWork);
                                            WriteLogsProcess((int)estadoProceso.Proceso, $"Se realiza mantenimiento de datos del trabajador: {employeeWork.DNIType}-{employeeWork.DNI} {employeeWork.Name}, para el cliente: {employeeWork.CodClient} - {employeeWork.Client}");
                                            if (resultado.Contains("Error"))
                                            {
                                                EjecutarAccionesErrorControlado(resultado, start, false, emailConsultor);
                                                bEnvioCorreo = true;
                                            }
                                            else
                                            {
                                                CreateFolderClient(employeeWork.Client, employeeWork.Codigo + "-" + employeeWork.Name.Replace(",", string.Empty));
                                                resultado = A3documentLib.ExecuteMainProcess(workFolders, employeeWork);
                                                if (resultado.Contains("Error"))
                                                {
                                                    EjecutarAccionesErrorControlado(resultado, start, false, emailConsultor);
                                                    bEnvioCorreo = true;
                                                }
                                                else
                                                {
                                                    NewEmployee.Write(workFolders, employeeWork);
                                                    MoverFileEmployee((int)estadoProceso.Finalizado, employeeWork.Codigo + "-" + employeeWork.Name.Replace(",", string.Empty), employeeWork.Client);
                                                    resultado = string.Empty;                                                    
                                                }                                                
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    process.Status = "EndOK";
                    process.ProcessDate = DateTime.Now;
                    process.RowInit = null;
                    process.Observation = null;

                    controller.Update(process);

                    if (!bEnvioCorreo)
                    {
                        RegisterIteration(start, "EndOK", $"RPA {appName} - Fin Proceso", true);
                        if (!string.IsNullOrEmpty(employeeWork.Name))
                        {
                            string listaNotificacionCompleta = listMailNotification;
                            if (!string.IsNullOrEmpty(emailConsultor))
                                listaNotificacionCompleta += "," + emailConsultor;

                            WriteLogsProcess((int)estadoProceso.Proceso, $"Finaliza el proceso de Alta Empleados, para el empleado: {employeeWork.Name}, {employeeWork.DNIType}: {employeeWork.DNI}, " +
                                        $"con fecha de inicio de contrato: {employeeWork.StartContract}");
                            
                            Mails.SendMail(usermail, passUserMail, hostMail,
                                    portMail, listaNotificacionCompleta, $"{appName} - Alta Empleados - Fin Proceso!",
                                    htmlString.Replace("%mesaje%", $"Finaliza el proceso de Alta Empleados, para el empleado: {employeeWork.Name}, {employeeWork.DNIType}: {employeeWork.DNI}, " +
                                        $"con fecha de inicio de contrato: {employeeWork.StartContract}"));

                            employeeWork.Name = string.Empty;
                        }
                    }
                    bEnvioCorreo = false;
                    contador = 0;

                    if (string.IsNullOrEmpty(employeeWork.Name))
                    {
                        WriteLogsProcess((int)estadoProceso.Proceso, "Finaliza el proceso de Alta Empleados, por horario");
                        break;
                    }
                }
                catch (ValidateDirectoriesException rexcep)
                {
                    bLogin = false;
                    contador++;
                    SaveControllerError("NewEmployee", rexcep.Message, start, false, emailConsultor, contador);
                }
                catch (ReadDataFileEmloyeeException redaex)
                {
                    bLogin = false;
                    contador++;
                    SaveControllerError("NewEmployee", redaex.Message, start, true, emailConsultor, contador);
                }
                catch (A3NomLoginException exa3)
                {
                    bLogin = false;
                    contador++;
                    SaveControllerError("A3NomLogin", exa3.Message, start, true, emailConsultor, contador);
                }
                catch (A3NomSearchException exSe)
                {
                    bLogin = false;
                    contador++;
                    SaveControllerError("A3NomSearch", exSe.Message, start, true, emailConsultor, contador);
                }
                catch (A3NomNuevoException exNue)
                {
                    bLogin = false;
                    contador++;
                    SaveControllerError("A3NomNuevo", exNue.Message, start, true, emailConsultor, contador);
                }
                catch(A3NomUpdateException exUpdate)
                {
                    bLogin = false;
                    contador++;
                    SaveControllerError("A3NomUpdate", exUpdate.Message, start, false, emailConsultor, contador);
                }
                catch(A3NomDocumentException exDoc)
                {
                    bLogin = false;
                    contador++;
                    SaveControllerError("A3NomDocument", exDoc.Message, start, false, emailConsultor, contador);
                }
                catch (Exception e)
                {
                    bLogin = false;
                    SaveControllerError(fuente, e.StackTrace, start, true, emailConsultor, contador);
                    break;
                }
            }

            try
            {
                foreach (var process in Process.GetProcessesByName(workFolders.a3NomExe))
                {
                    process.Kill();
                }
            }
            catch { }
            Thread.Sleep((workFolders.tEspera * 2));

            WriteLogsProcess((int)estadoProceso.Proceso, "Finaliza el proceso de Alta Empleados, no habia archivo para procesar");
            Mails.SendMail(usermail, passUserMail, hostMail,
                portMail, listMailNotification, $"{appName} - Alta Empleados - Fin Proceso!",
                htmlString.Replace("%mesaje%", "Finaliza el proceso de Alta Empleados, no habia archivo para procesar"));

            cartes.close();
        }

        void SaveControllerError(string source, string messageError, DateTime start, bool returnfile, string emailConsultor, int intento)
        {
            RpaProcess processError = new RpaProcess();
            processError.Id = appId;
            processError.Name = appName;
            processError.Machine = machine;
            processError.Process = source;
            processError.Status = "EndNoOK";
            processError.RowInit = rowInit;
            processError.ProcessDate = DateTime.Now;
            processError.Observation = messageError.ToString();
            controller.Update(processError);

            RegisterIteration(start, "EndNoOK", $"RPA {appName} - Fin Proceso con Error: " + messageError.ToString(), true);

            if(returnfile)
                MoverFileEmployee((int)estadoProceso.Devolver);
            else
                MoverFileEmployee((int)estadoProceso.Error);

            string listaNotificacionCompleta = listMailNotification;
            if (!string.IsNullOrEmpty(emailConsultor))
                listaNotificacionCompleta += "," + emailConsultor;

            WriteLogsProcess((int)estadoProceso.Error, messageError);

            if (intento == 3)
            {
                Mails.SendMail(usermail, passUserMail, hostMail,
                    portMail, listaNotificacionCompleta, $"{appName} - Alta Empleados - Error Proceso!",
                    htmlString.Replace("%mesaje%", $"Error al leer información del nuevo empleado. Detalle del Error: {messageError}"));
            }
        }

        void MoverFileEmployee(int estado, string empleado = null, string cliente = null)
        {
            string filesource = $@"{workFolders.rutaTemporal}\{workFolders.machineWork}\";
            string[] files = Directory.GetFiles(filesource);
            foreach (var file in files)
            {
                string filename = Path.GetFileName(file);
                string filedestino = string.Empty;
                switch(estado)
                {
                    case (int)estadoProceso.Devolver:
                        filedestino = $@"{workFolders.rutaBaseDocs}\{filename}";
                        break;
                    case (int)estadoProceso.Error:
                        filedestino = $@"{workFolders.rutaErrores}\{filename}";
                        break;
                    case (int)estadoProceso.Finalizado:
                        filedestino = $@"{workFolders.rutaFinalizados}\{filename}";
                        break;
                }

                if (File.Exists(filedestino))
                    File.Delete(filedestino);

                File.Copy(file, filedestino);

                if (estado == (int)estadoProceso.Finalizado)
                {
                    filedestino = $@"{workFolders.rutaClientes}\{cliente}\ROBOT\{empleado}\{filename}";
                    if (File.Exists(filedestino))
                        File.Delete(filedestino);

                    File.Copy(file, filedestino);
                }
                File.Delete(file);
            }
        }

        void CreateFolderClient(string cliente, string empleado)
        {
            string filedestino = $@"{workFolders.rutaClientes}\{cliente}";
            if (!Directory.Exists(filedestino))
                Directory.CreateDirectory(filedestino);

            filedestino = $@"{workFolders.rutaClientes}\{cliente}\ROBOT";
            if (!Directory.Exists(filedestino))
                Directory.CreateDirectory(filedestino);

            filedestino = $@"{workFolders.rutaClientes}\{cliente}\ROBOT\{empleado}";
            if (!Directory.Exists(filedestino))
                Directory.CreateDirectory(filedestino);
        }

        void WriteLogsProcess(int estado, string message)
        {
            if (!Directory.Exists($@"{workFolders.rutaLog}\{workFolders.machineWork}"))
                Directory.CreateDirectory($@"{workFolders.rutaLog}\{workFolders.machineWork}");

            string fileLog = $@"{workFolders.rutaLog}\{workFolders.machineWork}\Log_{((estado == (int)estadoProceso.Error) ? "Error" : "Proceso")}_{DateTime.Now.ToString("yyyyMMdd")}.log";

            StreamWriter stream = null;
            try
            {
                stream = File.AppendText(fileLog);
                stream.WriteLine(string.Format("{0} - {1}.", DateTime.Now, message));
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
        }

        void EjecutarAccionesErrorControlado(string mensaje, DateTime start, bool inicio, string emailConsultor)
        {
            RegisterIteration(start, "EndNoOK", $"RPA {appName} - Fin Proceso con Error: " + mensaje, true);

            string listaNotificacionCompleta = listMailNotification;
            if (!string.IsNullOrEmpty(emailConsultor))
                listaNotificacionCompleta += "," + emailConsultor;

            Mails.SendMail(usermail, passUserMail, hostMail,
                portMail, listaNotificacionCompleta, $"{appName} - Alta Empleados - Error Proceso!",
                htmlString.Replace("%mesaje%", mensaje));

            if(inicio)
                MoverFileEmployee((int)estadoProceso.Devolver);
            else
                MoverFileEmployee((int)estadoProceso.Error);

            WriteLogsProcess((int)estadoProceso.Error, mensaje);
        }

        protected override string getRPAMainFile()
        {
            return CurrentPath + "\\Cartes\\RPAMainProcess.cartes.rpa";
        }

        protected override void LoadConfiguration(XmlDocument XmlCfg)
        {
            //example
            appName = ToString(XmlCfg.SelectSingleNode("//record/appName"));
            workFolders.rutaBaseDocs = ToString(XmlCfg.SelectSingleNode("//record/rutaBaseDocs"));
            workFolders.rutaTemporal = ToString(XmlCfg.SelectSingleNode("//record/rutaTemporal"));
            workFolders.rutaFinalizados = ToString(XmlCfg.SelectSingleNode("//record/rutaFinalizados"));
            workFolders.rutaErrores = ToString(XmlCfg.SelectSingleNode("//record/rutaErrores"));
            workFolders.rutaClientes = ToString(XmlCfg.SelectSingleNode("//record/rutaClientes"));
            workFolders.rutaLog = ToString(XmlCfg.SelectSingleNode("//record/rutaLogProceso"));
            workFolders.rutaComplemento = $@"\LABORAL\GESTION DEL PERSONAL\{nowYear}\{nowMonth}.{monthName}\ALTAS\";
            workFolders.extensionWork = ToString(XmlCfg.SelectSingleNode("//record/extensionWork"));
            workFolders.nameFileWork = ToString(XmlCfg.SelectSingleNode("//record/nameFileWork"));
            workFolders.machineWork = machine;
            sustraDays = int.Parse(ToString(XmlCfg.SelectSingleNode("//record/sustraDays")));
            workFolders.rutaA3nom = ToString(XmlCfg.SelectSingleNode("//record/rutaA3nom"));
            workFolders.a3nomUser = ToString(XmlCfg.SelectSingleNode("//record/a3nomuser"));
            workFolders.a3nomPassword = ToString(XmlCfg.SelectSingleNode("//record/a3nompassword"));
            workFolders.a3NomExe = ToString(XmlCfg.SelectSingleNode("//record/a3nomexe"));
            workFolders.tEspera = int.Parse(ToString(XmlCfg.SelectSingleNode("//record/tEspera")));
            workFolders.tEsperaComp = int.Parse(ToString(XmlCfg.SelectSingleNode("//record/tEsperaComp")));

            usermail = ToString(XmlCfg.SelectSingleNode("//record/email"));
            passUserMail = ToString(XmlCfg.SelectSingleNode("//record/password"));
            hostMail = ToString(XmlCfg.SelectSingleNode("//record/smtphost"));
            portMail = ToString(XmlCfg.SelectSingleNode("//record/smtpport"));
            listMailNotification = ToString(XmlCfg.SelectSingleNode("//record/toEmail"));
        }

        public A3NomMainLibrary A3processLib
        {
            get
            {
                if (frpA3mainLibrary == null)
                    frpA3mainLibrary = new A3NomMainLibrary(this);

                return frpA3mainLibrary;
            }
        }

        public A3NomLoginLibrary A3loginLib
        {
            get
            {
                if (frpA3loginLibrary == null)
                    frpA3loginLibrary = new A3NomLoginLibrary(this);

                return frpA3loginLibrary;
            }
        }

        public A3NomSearchLibrary A3searchLib
        {
            get
            {
                if (frpA3searchLibrary == null)
                    frpA3searchLibrary = new A3NomSearchLibrary(this);

                return frpA3searchLibrary;
            }
        }

        public A3NomNuevoLibrary A3nuevoLib
        {
            get
            {
                if (frpA3nuevoLibrary == null)
                    frpA3nuevoLibrary = new A3NomNuevoLibrary(this);

                return frpA3nuevoLibrary;
            }
        }

        public A3NomUpdateLibrary A3updateLib
        {
            get
            {
                if (frpA3updateLibrary == null)
                    frpA3updateLibrary = new A3NomUpdateLibrary(this);

                return frpA3updateLibrary;
            }
        }

        public A3NomDocumentLibrary A3documentLib
        {
            get
            {
                if (frpA3documentLibrary == null)
                    frpA3documentLibrary = new A3NomDocumentLibrary(this);

                return frpA3documentLibrary;
            }
        }
    }
}