using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using renombradorbeco;
using renombradorbeco.app;

namespace RenombradorBeco
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string dir = Directory.GetCurrentDirectory();
                if (args.Length > 0) dir = args[0];

                Console.WriteLine($"- Leyendo carpeta {dir}");

                ContainerInfo<DirectoryInfo> containerInfo = dir;

                var excelProcess = new ExcelProcess(containerInfo);

                excelProcess.Process();
            }

            catch (Exception ex)
            {
                Console.WriteLine($"\n\n${ex.Message}");
            }

            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Console.WriteLine("\n\nPresione enter para cerrar esta ventana");
                Console.ReadLine();
            }
        }

    }
}
