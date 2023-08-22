using System;

namespace KeyiberboardAltaEmpleados
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            MainProcess mainProcess = new MainProcess();
            mainProcess.Execute();
        }
    }
}
