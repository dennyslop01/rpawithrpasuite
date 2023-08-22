using KeyiberboardModels.Models;
using MiTools;
using RPABaseAPI;
using System;

namespace A3NomLibrary
{
    public class A3NomMainLibrary : MyCartesAPIBase
    {
        protected string workingFile = Environment.CurrentDirectory;
        //private static bool loaded = false;

        public A3NomMainLibrary(MyCartesProcess owner) : base(owner)
        {
        }

        public override void Close()
        {
            //throw new NotImplementedException();
        }

        protected override void MergeLibrariesAndLoadVariables()
        {
            //example
            //loaded = cartes.merge(CurrentPath + "\\Cartes\\A3nomLogin.cartes.rpa") == 1;
        }

        public string ExecuteMainProcess(ProcessWorkFolders workFolders, EmployeeWork employee)
        {
            string respuesta = string.Empty;
            try
            {



            }
            catch (Exception e)
            {
                cartes.forensic("A3NomMainLibrary - ExecuteMainProcess - Exception: " + e.ToString());
                throw;
            }

            return respuesta;
        }
    }
}