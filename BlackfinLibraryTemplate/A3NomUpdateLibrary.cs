using Cartes;
using KeyiberboardException.Exceptions;
using KeyiberboardModels.Models;
using MiTools;
using RPABaseAPI;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;

namespace A3NomLibrary
{
    public class A3NomUpdateLibrary : MyCartesAPIBase
    {
        protected string workingFile = Environment.CurrentDirectory;
        private static bool loaded = false;
        RPAWin32Component salirUpdate = null, cancelarUpdate = null, centroCosto = null, ventanaUpdate = null, codMunicipio = null, codNacionalidad = null, salirSistema = null;
        RPAWin32Component codPaisDireccion = null, codProvincia = null, edoCivil = null, email = null, escDireccion = null, fechaNacimiento = null, nifLegal = null, nombreTrabajador = null, nroAfiliacion1 = null;
        RPAWin32Component nroAfiliacion2 = null, nroAfiliacion3 = null, nroDireccion = null, nroDocumento = null, nroMatricula = null, pisoDireccion = null, puertaDireccion = null, sexo = null, sigla = null;
        RPAWin32Component telefono = null, viaPublica = null, tipoDocumento = null, extension = null, bancoPago = null, mantenimientoBanco = null, ventanaBanco = null;
        RPAWin32Component entidadBank = null, agenciaBank = null, dcBank = null, cuentaBank = null, iban = null, aceptarBank = null, cancelarBank = null, modificar = null;
        RPAWin32Component ventanaRetenciones = null, cancelarRetenciones = null, ventanaAfiliacion = null, noprocesar = null, aceptarAfiliacion = null, nivelFormativo = null;
        RPAWin32Component finContrato = null, inicioContrato = null, contratoOpt = null, horasDias = null, tipoContrato = null, codOcupacion = null, validacionBanco = null, cancelarValidacion = null;
        RPAWin32Component cotizacionOpt = null, fechaIngreso = null, ventanaDatosAfiliacion = null, salirDatosVentanaAfilia = null, filiacionOpt = null, grupoTarifa = null, codTGSS = null;
        RPAWin32Component tParcial = null, coeTParcial = null, porcJornada = null, alertaColectivo = null, aceptarAlertaColectivo = null, codColectivo = null;
        public A3NomUpdateLibrary(MyCartesProcess owner) : base(owner)
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
            loaded = cartes.merge(CurrentPath + "\\Cartes\\A3NomUpdateTrabajador.cartes.rpa") == 1;
            ventanaUpdate = cartes.GetComponent<RPAWin32Component>("$VentanaUpdate");
            salirUpdate = cartes.GetComponent<RPAWin32Component>("$Salir");
            cancelarUpdate = cartes.GetComponent<RPAWin32Component>("$Cancelar");
            centroCosto = cartes.GetComponent<RPAWin32Component>("$CentroTrabajo");
            codMunicipio = cartes.GetComponent<RPAWin32Component>("$CodMunicipio");
            codNacionalidad = cartes.GetComponent<RPAWin32Component>("$CodNacionalidad");
            codPaisDireccion = cartes.GetComponent<RPAWin32Component>("$CodPaisDireccion");
            codProvincia = cartes.GetComponent<RPAWin32Component>("$CodProvincia");
            edoCivil = cartes.GetComponent<RPAWin32Component>("$EdoCivil");
            email = cartes.GetComponent<RPAWin32Component>("$Email");
            escDireccion = cartes.GetComponent<RPAWin32Component>("$EscDireccion");
            fechaNacimiento = cartes.GetComponent<RPAWin32Component>("$FechaNacimiento");
            nifLegal = cartes.GetComponent<RPAWin32Component>("$NifLegal");
            nombreTrabajador = cartes.GetComponent<RPAWin32Component>("$NombreTrabajador");
            nroAfiliacion1 = cartes.GetComponent<RPAWin32Component>("$NroAfiliacion1");
            nroAfiliacion2 = cartes.GetComponent<RPAWin32Component>("$NroAfiliacion2");
            nroAfiliacion3 = cartes.GetComponent<RPAWin32Component>("$NroAfiliacion3");
            nroDireccion = cartes.GetComponent<RPAWin32Component>("$NroDireccion");
            nroDocumento = cartes.GetComponent<RPAWin32Component>("$NroDocumento");
            nroMatricula = cartes.GetComponent<RPAWin32Component>("$NroMatricula");
            pisoDireccion = cartes.GetComponent<RPAWin32Component>("$PisoDireccion");
            puertaDireccion = cartes.GetComponent<RPAWin32Component>("$PuertaDireccion");
            sexo = cartes.GetComponent<RPAWin32Component>("$Sexo");
            sigla = cartes.GetComponent<RPAWin32Component>("$Sigla");
            telefono = cartes.GetComponent<RPAWin32Component>("$Telefono");
            tipoDocumento = cartes.GetComponent<RPAWin32Component>("$TipoDocumento");
            viaPublica = cartes.GetComponent<RPAWin32Component>("$ViaPublica");
            extension = cartes.GetComponent<RPAWin32Component>("$Extension");
            bancoPago = cartes.GetComponent<RPAWin32Component>("$BancoPago");
            mantenimientoBanco = cartes.GetComponent<RPAWin32Component>("$MantenimientoBanco");
            ventanaBanco = cartes.GetComponent<RPAWin32Component>("$VentanaBanco");
            entidadBank = cartes.GetComponent<RPAWin32Component>("$EntidadBank");
            agenciaBank = cartes.GetComponent<RPAWin32Component>("$AgenciaBank");
            dcBank = cartes.GetComponent<RPAWin32Component>("$DCBank");
            cuentaBank = cartes.GetComponent<RPAWin32Component>("$CuentaBank");
            iban = cartes.GetComponent<RPAWin32Component>("$IBAN");
            aceptarBank = cartes.GetComponent<RPAWin32Component>("$AceptarBank");
            cancelarBank = cartes.GetComponent<RPAWin32Component>("$CancelarBank");
            codOcupacion = cartes.GetComponent<RPAWin32Component>("$CodOcupacion");
            modificar = cartes.GetComponent<RPAWin32Component>("$Modificar");
            ventanaRetenciones = cartes.GetComponent<RPAWin32Component>("$VentanaRetenciones");
            cancelarRetenciones = cartes.GetComponent<RPAWin32Component>("$CancelarRetenciones");
            ventanaAfiliacion = cartes.GetComponent<RPAWin32Component>("$VentanaAfiliacion");
            noprocesar = cartes.GetComponent<RPAWin32Component>("$Noprocesar");
            aceptarAfiliacion = cartes.GetComponent<RPAWin32Component>("$AceptarAfiliacion");
            contratoOpt = cartes.GetComponent<RPAWin32Component>("$ContratoOpt");
            tipoContrato = cartes.GetComponent<RPAWin32Component>("$TipoContrato");
            inicioContrato = cartes.GetComponent<RPAWin32Component>("$InicioContrato");
            finContrato = cartes.GetComponent<RPAWin32Component>("$FinContrato");
            horasDias = cartes.GetComponent<RPAWin32Component>("$HorasDias");
            nivelFormativo = cartes.GetComponent<RPAWin32Component>("$NivelFormativo");
            validacionBanco = cartes.GetComponent<RPAWin32Component>("$ValidacionBanco");
            cancelarValidacion = cartes.GetComponent<RPAWin32Component>("$CancelarValidacion");
            salirSistema = cartes.GetComponent<RPAWin32Component>("$SalirSistema");
            cotizacionOpt = cartes.GetComponent<RPAWin32Component>("$CotizacionOpt");
            fechaIngreso = cartes.GetComponent<RPAWin32Component>("$FechaIngreso");
            ventanaDatosAfiliacion = cartes.GetComponent<RPAWin32Component>("$VentanaDatosAfiliacion");
            salirDatosVentanaAfilia = cartes.GetComponent<RPAWin32Component>("$SalirDatosVentanaAfilia");
            filiacionOpt = cartes.GetComponent<RPAWin32Component>("$FiliacionOpt");
            grupoTarifa = cartes.GetComponent<RPAWin32Component>("$GrupoTarifa");
            codTGSS = cartes.GetComponent<RPAWin32Component>("$CodTGSS");
            tParcial = cartes.GetComponent<RPAWin32Component>("$TParcial");
            coeTParcial = cartes.GetComponent<RPAWin32Component>("$CoeTParcial");
            porcJornada = cartes.GetComponent<RPAWin32Component>("$PorcJornada");
            alertaColectivo = cartes.GetComponent<RPAWin32Component>("$AlertaColectivo");
            aceptarAlertaColectivo = cartes.GetComponent<RPAWin32Component>("$AceptarAlertaColectivo");
            codColectivo = cartes.GetComponent<RPAWin32Component>("$CodColectivo");

        }

        /// <summary>
        /// Metodo donde se encuentra toda la logica del Bot
        /// </summary>
        /// <param name="workFolders">Objeto con parametros generales</param>
        /// <param name="employee">Objeto con los datos del empleado cargados desde el excel de insumo</param>
        /// <returns></returns>
        /// <exception cref="A3NomUpdateException">Retorna una excepcion del tipo actualizar</exception>
        public string ExecuteMainProcess(ProcessWorkFolders workFolders, EmployeeWork employee)
        {
            string respuesta = string.Empty;
            bool bCreado = false;

            try
            {
                while (!bCreado)
                {
                    cartes.reset("win32");
                    Thread.Sleep(workFolders.tEspera);
                    //Se valida que la ventana de mantenimiento exista
                    if (ventanaUpdate.ComponentExist(workFolders.tEsperaComp))
                    {
                        centroCosto.waitforcomponent(workFolders.tEsperaComp);
                        nroDocumento.waitforcomponent(workFolders.tEsperaComp);
                        //Se valida si los campos estas inhabilitados en caso de ser asi se hace clic en el boton de modificar
                        if (centroCosto.enabled() == 0 && nroDocumento.enabled() == 0)
                        {
                            modificar.waitforcomponent(workFolders.tEsperaComp);
                            modificar.click();
                            Thread.Sleep(workFolders.tEspera);
                        }

                        //Se inicia la carga de todos los datos que vienen en la platilla
                        nombreTrabajador.waitforcomponent(workFolders.tEsperaComp);
                        nombreTrabajador.focus();
                        if (nombreTrabajador.Value != null)
                        {
                            if (nombreTrabajador.Value != employee.Name)
                            {
                                nombreTrabajador.TypeKey("36");
                                nombreTrabajador.TypeKey("35", "Shift", null);
                                nombreTrabajador.Press(8, 120);
                                //nombreTrabajador.click();
                                for (int i = 0; i < employee.Name.Length; i++)
                                {
                                    nombreTrabajador.TypeKey(employee.Name.Substring(i, 1));
                                }
                                nombreTrabajador.Press(9, 1);
                                Thread.Sleep(workFolders.tEspera);
                            }
                        }

                        fechaNacimiento.waitforcomponent(workFolders.tEsperaComp);
                        fechaNacimiento.focus();
                        //fechaNacimiento.click();
                        for (int i = 0; i < employee.BirthDate.Length; i++)
                        {
                            fechaNacimiento.TypeKey(employee.BirthDate.Substring(i, 1));
                        }
                        fechaNacimiento.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        nroAfiliacion1.waitforcomponent(workFolders.tEsperaComp);
                        nroAfiliacion1.focus();
                        //nroAfiliacion1.click();
                        for (int i = 0; i < employee.NroMembership.Length; i++)
                        {
                            nroAfiliacion1.TypeKey(employee.NroMembership.Substring(i, 1));
                        }
                        nroAfiliacion1.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        nroAfiliacion2.waitforcomponent(workFolders.tEsperaComp);
                        nroAfiliacion2.focus();
                        //nroAfiliacion2.click();
                        for (int i = 0; i < employee.NroMembership2.Length; i++)
                        {
                            nroAfiliacion2.TypeKey(employee.NroMembership2.Substring(i, 1));
                        }
                        nroAfiliacion2.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        nroAfiliacion3.waitforcomponent(workFolders.tEsperaComp);
                        nroAfiliacion3.focus();
                        //nroAfiliacion3.click();
                        for (int i = 0; i < employee.NroMembership3.Length; i++)
                        {
                            nroAfiliacion3.TypeKey(employee.NroMembership3.Substring(i, 1));
                        }
                        nroAfiliacion3.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        sexo.waitforcomponent(workFolders.tEsperaComp);
                        sexo.focus();
                        //sexo.click();
                        sexo.Value = employee.Gender;
                        Thread.Sleep(workFolders.tEspera);

                        codNacionalidad.waitforcomponent(workFolders.tEsperaComp);
                        codNacionalidad.focus();
                        //codNacionalidad.click();
                        for (int i = 0; i < employee.CountryNationality.Length; i++)
                        {
                            codNacionalidad.TypeKey(employee.CountryNationality.Substring(i, 1));
                        }
                        codNacionalidad.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        sigla.waitforcomponent(workFolders.tEsperaComp);
                        sigla.focus();
                        //sigla.click();
                        sigla.Value = employee.Initials;
                        Thread.Sleep(workFolders.tEspera);

                        viaPublica.waitforcomponent(workFolders.tEsperaComp);
                        viaPublica.focus();
                        viaPublica.Value = employee.PublicRoad;
                        viaPublica.click();
                        //for (int i = 0; i < employee.PublicRoad.Length; i++)
                        //{
                        //    viaPublica.TypeKey(employee.PublicRoad.Substring(i, 1));
                        //}
                        viaPublica.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        nroDireccion.waitforcomponent(workFolders.tEsperaComp);
                        nroDireccion.focus();
                        //nroDireccion.click();
                        for (int i = 0; i < employee.Number.Length; i++)
                        {
                            nroDireccion.TypeKey(employee.Number.Substring(i, 1));
                        }
                        nroDireccion.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        pisoDireccion.waitforcomponent(workFolders.tEsperaComp);
                        pisoDireccion.focus();
                        //pisoDireccion.click();
                        for (int i = 0; i < employee.Floor.Length; i++)
                        {
                            pisoDireccion.TypeKey(employee.Floor.Substring(i, 1));
                        }
                        pisoDireccion.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        puertaDireccion.waitforcomponent(workFolders.tEsperaComp);
                        puertaDireccion.focus();
                        //puertaDireccion.click();
                        for (int i = 0; i < employee.Door.Length; i++)
                        {
                            puertaDireccion.TypeKey(employee.Door.Substring(i, 1));
                        }
                        puertaDireccion.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        codMunicipio.waitforcomponent(workFolders.tEsperaComp);
                        //codMunicipio.focus();
                        codMunicipio.click();
                        for (int i = 0; i < employee.Municipality.Length; i++)
                        {
                            codMunicipio.TypeKey(employee.Municipality.Substring(i, 1));
                        }
                        codMunicipio.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        codProvincia.waitforcomponent(workFolders.tEsperaComp);
                        //codProvincia.focus();
                        codProvincia.click();
                        for (int i = 0; i < employee.PostalCode.Length; i++)
                        {
                            codProvincia.TypeKey(employee.PostalCode.Substring(i, 1));
                        }
                        codProvincia.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        codPaisDireccion.waitforcomponent(workFolders.tEsperaComp);
                        codPaisDireccion.focus();
                        //codPaisDireccion.click();
                        for (int i = 0; i < employee.Country.Length; i++)
                        {
                            codPaisDireccion.TypeKey(employee.Country.Substring(i, 1));
                        }
                        codPaisDireccion.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        telefono.waitforcomponent(workFolders.tEsperaComp);
                        telefono.focus();
                        //telefono.click();
                        telefono.Value = employee.Phone;
                        Thread.Sleep(workFolders.tEspera);

                        email.waitforcomponent(workFolders.tEsperaComp);
                        email.focus();
                        //email.click();
                        for (int i = 0; i < employee.Email.Length; i++)
                        {
                            if (employee.Email.Substring(i, 1) != "@")
                            {
                                email.TypeKey(employee.Email.Substring(i, 1));
                            }
                            else
                            {
                                Clipboard.SetText("@");
                                email.TypeKey("86", "Control", null);
                                Clipboard.SetText(" ");
                            }
                        }
                        email.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        codOcupacion.waitforcomponent(workFolders.tEsperaComp);
                        codOcupacion.focus();
                        //codOcupacion.click();
                        for (int i = 0; i < employee.OccupationCode.Length; i++)
                        {
                            codOcupacion.TypeKey(employee.OccupationCode.Substring(i, 1));
                        }
                        codOcupacion.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        bancoPago.waitforcomponent(workFolders.tEsperaComp);
                        bancoPago.focus();
                        //bancoPago.click();
                        for (int i = 0; i < employee.PaymentBak.Length; i++)
                        {
                            bancoPago.TypeKey(employee.PaymentBak.Substring(i, 1));
                        }
                        bancoPago.Press(9, 1);
                        Thread.Sleep(workFolders.tEspera);

                        bool cargarBanco = false;
                        mantenimientoBanco.waitforcomponent(workFolders.tEsperaComp);
                        mantenimientoBanco.focus();
                        mantenimientoBanco.click();
                        Thread.Sleep(10000);

                        int contBanco = 0;
                        while (!cargarBanco)
                        {
                            contBanco++;
                            if (ventanaBanco.ComponentExist(workFolders.tEsperaComp))
                            {
                                entidadBank.waitforcomponent(workFolders.tEsperaComp);
                                entidadBank.focus();
                                //entidadBank.click();
                                for (int i = 0; i < employee.Entity.Length; i++)
                                {
                                    entidadBank.TypeKey(employee.Entity.Substring(i, 1));
                                }
                                entidadBank.Press(9, 1);
                                Thread.Sleep(workFolders.tEspera);

                                agenciaBank.waitforcomponent(workFolders.tEsperaComp);
                                agenciaBank.focus();
                                //agenciaBank.click();
                                for (int i = 0; i < employee.Agency.Length; i++)
                                {
                                    agenciaBank.TypeKey(employee.Agency.Substring(i, 1));
                                }
                                agenciaBank.Press(9, 1);
                                Thread.Sleep(workFolders.tEspera);

                                dcBank.waitforcomponent(workFolders.tEsperaComp);
                                dcBank.focus();
                                //dcBank.click();
                                for (int i = 0; i < employee.DC.Length; i++)
                                {
                                    dcBank.TypeKey(employee.DC.Substring(i, 1));
                                }
                                dcBank.Press(9, 1);
                                Thread.Sleep(workFolders.tEspera);

                                cuentaBank.waitforcomponent(workFolders.tEsperaComp);
                                cuentaBank.focus();
                                //cuentaBank.click();
                                for (int i = 0; i < employee.Account.Length; i++)
                                {
                                    cuentaBank.TypeKey(employee.Account.Substring(i, 1));
                                }
                                cuentaBank.Press(9, 1);
                                Thread.Sleep(workFolders.tEspera);

                                if (validacionBanco.ComponentExist(workFolders.tEsperaComp))
                                {
                                    cancelarValidacion.waitforcomponent(workFolders.tEsperaComp);
                                    cancelarValidacion.click();
                                    Thread.Sleep(workFolders.tEspera);
                                    if (contBanco >= 3)
                                    {
                                        cancelarBank.waitforcomponent(workFolders.tEsperaComp);
                                        cancelarBank.click();
                                        break;
                                    }
                                }
                                else
                                {
                                    iban.waitforcomponent(workFolders.tEsperaComp);
                                    iban.focus();
                                    //iban.click();
                                    Thread.Sleep(workFolders.tEspera);
                                    string ibanTemp = $"{employee.IBAN}{employee.Entity}{employee.Agency}{employee.DC}{employee.Account}";
                                    if (ibanTemp.Contains(iban.Value))
                                    {
                                        aceptarBank.waitforcomponent(workFolders.tEsperaComp);
                                        aceptarBank.focus();
                                        aceptarBank.click();
                                        Thread.Sleep(workFolders.tEspera);
                                        cargarBanco = true;
                                    }
                                    else
                                        break;

                                }
                            }
                        }

                        if (!cargarBanco)
                        {
                            cancelarUpdate.waitforcomponent(workFolders.tEsperaComp);
                            cancelarUpdate.focus();
                            cancelarUpdate.click();
                            Thread.Sleep(workFolders.tEspera);
                            CompletarSalida(workFolders);
                            return $"Error al intentar realizar la carga de los datos del banco, para el trabajador de Nombre: {employee.Name}, Nro Doc: {employee.DNI} para el CLiente: {employee.Client}";
                        }
                        else
                        {
                            cotizacionOpt.waitforcomponent(workFolders.tEsperaComp);
                            cotizacionOpt.focus();
                            cotizacionOpt.click();
                            Thread.Sleep(workFolders.tEspera);

                            grupoTarifa.waitforcomponent(workFolders.tEsperaComp);
                            grupoTarifa.focus();
                            //grupoTarifa.click();
                            for (int i = 0; i < employee.CotiGroup.Length; i++)
                            {
                                grupoTarifa.TypeKey(employee.CotiGroup.Substring(i, 1));
                            }
                            grupoTarifa.Press(9, 1);
                            Thread.Sleep(workFolders.tEspera);

                            codTGSS.waitforcomponent(workFolders.tEsperaComp);
                            codTGSS.focus();
                            //codTGSS.click();
                            for (int i = 0; i < employee.OccupaCodeTGSS.Length; i++)
                            {
                                codTGSS.TypeKey(employee.OccupaCodeTGSS.Substring(i, 1));
                            }
                            codTGSS.Press(9, 1);
                            Thread.Sleep(workFolders.tEspera);

                            fechaIngreso.waitforcomponent(workFolders.tEsperaComp);
                            fechaIngreso.focus();
                            //fechaIngreso.click();
                            for (int i = 0; i < employee.IncomeDay.Length; i++)
                            {
                                fechaIngreso.TypeKey(employee.IncomeDay.Substring(i, 1));
                            }
                            fechaIngreso.Press(9, 1);
                            Thread.Sleep(workFolders.tEspera);

                            tParcial.waitforcomponent(workFolders.tEsperaComp);
                            tParcial.focus();
                            tParcial.Value = employee.PartTime;
                            tParcial.Press(9, 1);
                            Thread.Sleep(workFolders.tEspera);

                            coeTParcial.waitforcomponent(workFolders.tEsperaComp);
                            coeTParcial.focus();
                            //coeTParcial.click();
                            for (int i = 0; i < employee.CoefficientPartTime.Length; i++)
                            {
                                coeTParcial.TypeKey(employee.CoefficientPartTime.Substring(i, 1));
                            }
                            coeTParcial.Press(9, 1);
                            Thread.Sleep(workFolders.tEspera);

                            if(porcJornada.ComponentExist(workFolders.tEsperaComp))
                            {
                                porcJornada.waitforcomponent(workFolders.tEsperaComp);
                                porcJornada.focus();
                                //porcJornada.click();
                                for (int i = 0; i < employee.WorkingDay.Length; i++)
                                {
                                    porcJornada.TypeKey(employee.WorkingDay.Substring(i, 1));
                                }
                                porcJornada.Press(9, 1);
                                Thread.Sleep(workFolders.tEspera);
                            }

                            contratoOpt.waitforcomponent(workFolders.tEsperaComp);
                            contratoOpt.focus();
                            contratoOpt.click();
                            Thread.Sleep(workFolders.tEspera);

                            tipoContrato.waitforcomponent(workFolders.tEsperaComp);
                            tipoContrato.focus();
                            //tipoContrato.click();
                            for (int i = 0; i < employee.ContractType.Length; i++)
                            {
                                tipoContrato.TypeKey(employee.ContractType.Substring(i, 1));
                            }
                            tipoContrato.Press(9, 1);
                            Thread.Sleep(workFolders.tEspera);

                            if (ventanaDatosAfiliacion.ComponentExist(workFolders.tEsperaComp))
                            {                                                                
                                string codColectivoAux = "";
                                switch (employee.ContractType)
                                {
                                    case "402":
                                    case "502":
                                        codColectivoAux = "967";
                                        break;
                                }

                                if (!String.IsNullOrEmpty(employee.CodColectivo))
                                    codColectivoAux = employee.CodColectivo;

                                if (codColectivoAux != "")
                                {
                                    codColectivo.waitforcomponent(workFolders.tEsperaComp);
                                    codColectivo.focus();
                                    codColectivo.click();

                                    for (int i = 0; i < 30; i++)
                                    {
                                        if (string.IsNullOrEmpty(codColectivo.Value))
                                            codColectivo.Press(40, 1);
                                        else
                                        {
                                            if (Convert.ToString(codColectivo.Value).StartsWith(codColectivoAux))
                                                break;
                                        }
                                        codColectivo.Press(40, 1);
                                    }
                                    if (!Convert.ToString(codColectivo.Value).StartsWith(codColectivoAux))
                                    {
                                        throw new Exception("No se encuentra coincidencia con el Codigo de Colectivo");
                                    }

                                    codColectivo.Press(9, 1);
                                    Thread.Sleep(workFolders.tEspera);
                                }
                                salirDatosVentanaAfilia.waitforcomponent(workFolders.tEsperaComp);
                                salirDatosVentanaAfilia.focus();
                                salirDatosVentanaAfilia.click();
                                Thread.Sleep(workFolders.tEspera);
                                //}
                            }

                            inicioContrato.waitforcomponent(workFolders.tEsperaComp);
                            inicioContrato.focus();
                            //inicioContrato.click();
                            for (int i = 0; i < employee.StartContract.Length; i++)
                            {
                                inicioContrato.TypeKey(employee.StartContract.Substring(i, 1));
                            }
                            inicioContrato.Press(9, 1);
                            Thread.Sleep(workFolders.tEspera);

                            if (!string.IsNullOrEmpty(employee.EndContract))
                            {
                                finContrato.waitforcomponent(workFolders.tEsperaComp);
                                finContrato.focus();
                                //finContrato.click();
                                for (int i = 0; i < employee.EndContract.Length; i++)
                                {
                                    finContrato.TypeKey(employee.EndContract.Substring(i, 1));
                                }
                                finContrato.Press(9, 1);
                                Thread.Sleep(workFolders.tEspera);
                            }

                            nivelFormativo.waitforcomponent(workFolders.tEsperaComp);
                            nivelFormativo.focus();
                            //nivelFormativo.click();
                            for (int i = 0; i < employee.TrainingLevel.Length; i++)
                            {
                                nivelFormativo.TypeKey(employee.TrainingLevel.Substring(i, 1));
                            }
                            nivelFormativo.Press(9, 1);
                            Thread.Sleep(workFolders.tEspera);

                            horasDias.waitforcomponent(workFolders.tEsperaComp);
                            horasDias.focus();
                            //horasDias.click();
                            if (!string.IsNullOrEmpty(employee.HoursDays))
                            {
                                for (int i = 0; i < employee.HoursDays.Length; i++)
                                {
                                    horasDias.TypeKey(employee.HoursDays.Substring(i, 1));
                                }
                                horasDias.Press(9, 1);
                                Thread.Sleep(workFolders.tEspera);
                            }
                            else
                                horasDias.Value = "";

                            modificar.waitforcomponent(workFolders.tEsperaComp);
                            modificar.click();
                            Thread.Sleep(workFolders.tEspera);

                            //filiacionOpt.waitforcomponent(workFolders.tEsperaComp);
                            //filiacionOpt.focus();
                            //filiacionOpt.click();
                            Thread.Sleep(workFolders.tEspera);
                            if (nivelFormativo.enabled() == 0 && inicioContrato.enabled() == 0)
                            {
                                bCreado = true;
                                return respuesta;
                            }
                        }
                        CompletarSalida(workFolders);
                    }
                }
            }
            catch (Exception e)
            {
                cartes.forensic("A3NomUpdateLibrary - ExecuteMainProcess - Exception: " + e.ToString());
                throw new A3NomUpdateException(e.StackTrace);
            }

            return respuesta;
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
            while(ventanaAfiliacion.ComponentExist(workFolders.tEsperaComp))
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
    }
}