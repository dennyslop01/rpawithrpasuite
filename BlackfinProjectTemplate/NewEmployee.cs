using ClosedXML.Excel;
using KeyiberboardException.Exceptions;
using KeyiberboardModels.Models;
using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace KeyiberboardAltaEmpleados
{
    public static class NewEmployee
    {
        /// <summary>
        /// Metodo publico que llama a la lectura y validacion de directorios
        /// </summary>
        /// <param name="workFolders"></param>
        /// <param name="employeeWork"></param>
        /// <returns></returns>
        public static string Read(ProcessWorkFolders workFolders, ref EmployeeWork employeeWork)
        {
            if (ValidateDirectoriesExists(workFolders))
            {
                string[] files = Directory.GetFiles(workFolders.rutaBaseDocs);
                foreach (var file in files)
                {
                    string extension = Path.GetExtension(file);
                    string filename = Path.GetFileName(file);

                    if (extension.Contains(workFolders.extensionWork))
                    {
                        if (filename.Contains(workFolders.nameFileWork))
                        {
                            if (!Directory.Exists($@"{workFolders.rutaTemporal}\{workFolders.machineWork}"))
                                Directory.CreateDirectory($@"{workFolders.rutaTemporal}\{workFolders.machineWork}");

                            string filedestino = $@"{workFolders.rutaTemporal}\{workFolders.machineWork}\{filename}";

                            if (File.Exists(filedestino))
                                File.Delete(filedestino);

                            File.Copy(file, filedestino);
                            File.Delete(file);

                            string resultado = ReadDataFileEmloyee(ref employeeWork, filedestino);
                            if (resultado.Contains("Error"))
                                return resultado;

                            if (string.IsNullOrEmpty(employeeWork.Client) || string.IsNullOrEmpty(employeeWork.Name))
                            {
                                string filedestinoError = $@"{workFolders.rutaErrores}\{filename}";
                                File.Copy(filedestino, filedestinoError);
                                File.Delete(filedestino);
                            }
                            else
                            {
                                employeeWork.rutaDoc = filedestino;
                                break;
                            }
                        }
                        else
                            File.Delete(file);
                    }
                    else
                        File.Delete(file);
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// Metodo publico que llama a la escritura
        /// </summary>
        /// <param name="workFolders"></param>
        /// <param name="employee"></param>
        public static void Write(ProcessWorkFolders workFolders, EmployeeWork employee)
        {
            WriteFileEmployee(employee, $@"{workFolders.rutaTemporal}\{workFolders.machineWork}");
        }

        /// <summary>
        /// Metodo que valida que las carpetas compartidas y de respaldo existan
        /// </summary>
        /// <param name="workFolders"></param>
        /// <returns></returns>
        /// <exception cref="ValidateDirectoriesException"></exception>
        static bool ValidateDirectoriesExists(ProcessWorkFolders workFolders)
        {
            if (!Directory.Exists(workFolders.rutaBaseDocs))
                throw new ValidateDirectoriesException("Ruta Base compartida no existe");

            if (!Directory.Exists(workFolders.rutaTemporal))
                throw new ValidateDirectoriesException("Ruta Temporal no existe");

            if (!Directory.Exists(workFolders.rutaFinalizados))
                throw new ValidateDirectoriesException("Ruta Finalizados no existe");

            if (!Directory.Exists(workFolders.rutaErrores))
                throw new ValidateDirectoriesException("Ruta Errores no existe");

            if (!Directory.Exists(workFolders.rutaClientes))
                throw new ValidateDirectoriesException("Ruta Clientes compartida no existe");

            return true;
        }

        /// <summary>
        /// Metodo que lee la plantilla de Excel para los datos de entrada del trabajador
        /// </summary>
        /// <param name="employeeWork"></param>
        /// <param name="rutaFile"></param>
        /// <returns></returns>
        /// <exception cref="ReadDataFileEmloyeeException"></exception>
        static string ReadDataFileEmloyee(ref EmployeeWork employeeWork, string rutaFile)
        {
            try
            {
                var workRead = new XLWorkbook(rutaFile);
                var sheetRead = workRead.Worksheets.Where(x => x.Name == "Ficha").First();
                employeeWork.EmailConsultor = sheetRead.Cell("B1").GetString().Trim();

                bool bFaltaCampos = false;
                string campos = string.Empty;

                for(int i = 1; i < 47; i++)
                {
                    string key = sheetRead.Cell("A" + i).GetString().Trim();
                    string value = sheetRead.Cell("B" + i).GetString().Trim();

                    if (key.Contains('*'))
                    {
                        if (string.IsNullOrEmpty(value))
                        {
                            campos += $" {key.Replace("*", string.Empty)}, ";
                            bFaltaCampos = true;
                        }
                    }
                }

                if (bFaltaCampos)
                    return $"Error, faltan los siguientes campos: {campos.Trim()}, para poder procesar el archivo {rutaFile}";                                    
                
                employeeWork.Client = sheetRead.Cell("B2").GetString().Trim();
                employeeWork.CodClient = sheetRead.Cell("B3").GetString().Trim();
                employeeWork.Name = sheetRead.Cell("B6").GetString().Trim();
                employeeWork.DNIType = sheetRead.Cell("B7").GetString().Trim();
                employeeWork.DNI = sheetRead.Cell("B8").GetString().Trim();
                employeeWork.Codigo = sheetRead.Cell("D8").GetString().Trim();
                employeeWork.CodCenterWork = sheetRead.Cell("B9").GetString().Trim();
                employeeWork.BirthDate = DateTime.Parse(sheetRead.Cell("B10").GetString().Substring(0,10).Trim()).ToString("dd/MM/yyyy");
                if (sheetRead.Cell("B11").GetString().Length > 0)
                    employeeWork.NroMembership = int.Parse(sheetRead.Cell("B11").GetString().Trim()).ToString("00");
                
                if (sheetRead.Cell("C11").GetString().Length > 0)
                    employeeWork.NroMembership2 = Int64.Parse(sheetRead.Cell("C11").GetString().Trim()).ToString("00000000");

                if (sheetRead.Cell("D11").GetString().Length > 0)
                    employeeWork.NroMembership3 = int.Parse(sheetRead.Cell("D11").GetString().Trim()).ToString("00");

                employeeWork.Gender = ((sheetRead.Cell("B12").GetString().Trim()=="H")? "Hombre": "Mujer");
                employeeWork.CountryNationality = sheetRead.Cell("B13").GetString().Trim();
                employeeWork.Initials = sheetRead.Cell("C15").GetString().Trim();
                
                string viaPublica = ((sheetRead.Cell("B16").GetString().Trim()).ToLower()).ToUpper();
                employeeWork.PublicRoad = viaPublica;

                employeeWork.Number = sheetRead.Cell("B17").GetString().Trim();
                employeeWork.Floor = sheetRead.Cell("D17").GetString().Trim();
                employeeWork.Door = sheetRead.Cell("F17").GetString().Trim();
                employeeWork.Municipality = sheetRead.Cell("B18").GetString().Trim();
                employeeWork.PostalCode = sheetRead.Cell("B19").GetString().Trim();
                employeeWork.Country = sheetRead.Cell("B20").GetString().Trim();
                employeeWork.Phone = sheetRead.Cell("B21").GetString().Trim();
                employeeWork.Email = sheetRead.Cell("B22").GetString().Trim();
                employeeWork.PaymentBak = sheetRead.Cell("B24").GetString().Trim();
                employeeWork.IBAN = sheetRead.Cell("B26").GetString().Trim();
                if (sheetRead.Cell("C26").GetString().Length > 0)
                    employeeWork.Entity = int.Parse(sheetRead.Cell("C26").GetString().Trim()).ToString("0000");

                if (sheetRead.Cell("D26").GetString().Length > 0)
                    employeeWork.Agency = int.Parse(sheetRead.Cell("D26").GetString().Trim()).ToString("0000");

                if (sheetRead.Cell("E26").GetString().Length > 0) 
                    employeeWork.DC = int.Parse(sheetRead.Cell("E26").GetString().Trim()).ToString("00");

                if (sheetRead.Cell("F26").GetString().Length > 0)
                    employeeWork.Account = Int64.Parse(sheetRead.Cell("F26").GetString().Trim()).ToString("0000000000");

                employeeWork.OccupationCode = sheetRead.Cell("B28").GetString().Trim();
                employeeWork.CotiGroup = sheetRead.Cell("B30").GetString().Trim();
                employeeWork.OccupaCodeTGSS = sheetRead.Cell("B31").GetString().Trim();
                if (sheetRead.Cell("B33").GetString().Length > 0)
                    employeeWork.IncomeDay = DateTime.Parse(sheetRead.Cell("B33").GetString().Substring(0, 10).Trim()).ToString("dd/MM/yyyy");

                string tpartial = sheetRead.Cell("B35").GetString().Trim();
                if (tpartial.Contains("SI"))
                {
                    employeeWork.PartTime = sheetRead.Cell("C35").GetString().Trim();
                    employeeWork.WorkingDay = sheetRead.Cell("B36").GetString().Trim();
                    if (sheetRead.Cell("B37").GetString().Length > 0)
                        employeeWork.CoefficientPartTime = int.Parse(sheetRead.Cell("B37").GetString().Trim()).ToString("000");
                }
                else
                {
                    employeeWork.PartTime = "No";
                    employeeWork.WorkingDay = string.Empty;
                    employeeWork.CoefficientPartTime = string.Empty;
                }
                employeeWork.ContractType = sheetRead.Cell("B40").GetString().Trim();
                employeeWork.CodColectivo = sheetRead.Cell("D40").GetString().Trim();

                if (sheetRead.Cell("B41").GetString().Length > 0)
                    employeeWork.HoursDays = decimal.Parse(sheetRead.Cell("B41").GetString().Trim()).ToString("00.00");

                employeeWork.StartContract = DateTime.Parse(sheetRead.Cell("B42").GetString().Substring(0, 10).Trim()).ToString("dd/MM/yyyy");
                if (sheetRead.Cell("B43").GetString().Length > 0)
                    employeeWork.EndContract = DateTime.Parse(sheetRead.Cell("B43").GetString().Substring(0, 10).Trim()).ToString("dd/MM/yyyy");

                employeeWork.TrainingLevel = sheetRead.Cell("B44").GetString().Trim();
                employeeWork.GrossSalary = sheetRead.Cell("B45").GetString().Trim();
                employeeWork.PeriodSalary = sheetRead.Cell("B46").GetString().Trim();
                employeeWork.PayerData = sheetRead.Cell("B48").GetString().Trim();
                employeeWork.ListContract = sheetRead.Cell("B49").GetString().Trim();
                employeeWork.ResidenceDeclare = sheetRead.Cell("B50").GetString().Trim();
                employeeWork.IncomeIRPF = sheetRead.Cell("B51").GetString().Trim();

                return string.Empty;
            }
            catch (Exception ex)
            {
                throw new ReadDataFileEmloyeeException(ex.StackTrace);
            }
        }

        /// <summary>
        /// Metodo que actualiza el campo codigo depues que termina el proceso
        /// </summary>
        /// <param name="employee"></param>
        /// <param name="rutaFile"></param>
        /// <exception cref="ReadDataFileEmloyeeException"></exception>
        static void WriteFileEmployee(EmployeeWork employee, string rutaFile)
        {
            try
            {
                string[] files = Directory.GetFiles(rutaFile);
                foreach (var file in files)
                {
                    string extension = Path.GetExtension(file);
                    string filename = Path.GetFileName(file);

                    string filewrite = $@"{rutaFile}\{filename}";
                    var workRead = new XLWorkbook(filewrite);
                    var sheetRead = workRead.Worksheets.Where(x => x.Name == "Ficha").First();
                    sheetRead.Cell("C8").Value = employee.Codigo;

                    workRead.Save();
                }
            }
            catch (Exception ex)
            {
                throw new ReadDataFileEmloyeeException(ex.StackTrace);
            }
        }
    }
}
