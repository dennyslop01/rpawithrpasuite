using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KeyiberboardModels.Models
{
    public class ProcessWorkFolders
    {
        public string rutaBaseDocs { get; set; }
        public string rutaTemporal { get; set; }
        public string rutaFinalizados { get; set; }
        public string rutaErrores { get; set; }
        public string rutaClientes { get; set; }
        public string rutaComplemento { get; set; }
        public string rutaLog { get; set; }
        public string extensionWork { get; set; }
        public string nameFileWork { get; set; }
        public string machineWork { get; set; }
        public string rutaA3nom { get; set; }
        public string a3nomUser { get; set; }
        public string a3nomPassword { get; set; }
        public string a3NomExe { get; set; }
        public int tEspera { get; set; }
        public int tEsperaComp { get; set; }

    }
}
