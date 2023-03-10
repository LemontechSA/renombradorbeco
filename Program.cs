using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace RenombradorBeco
{
    class Program
    {
        const string PROC_COL = "N° Procedimiento";
        const string OP_COL = "Operación";

        static void Main(string[] args)
        {
            try
            {
                string dir = Directory.GetCurrentDirectory();
                if (args.Length > 0) dir = args[0];

                Console.WriteLine("- Leyendo carpeta " + dir);
                var d = new DirectoryInfo(dir);

                var xlsFile = d.GetFiles("*.xlsx").FirstOrDefault();
                if (xlsFile == null) throw new Exception("No se encontraron archivos excel en la carpeta.\n\n" +
                    "Debe ejecutar este programa en la misma carpeta donde estén los documentos PDF y el excel con los números de procedimiento, " +
                    "o bien arrastrar dicha carpeta encima de este programa");

                ProcessExcel(d, xlsFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n\n" + ex.Message);
            } 
            finally 
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Console.WriteLine("\n\nPresione enter para cerrar esta ventana");
                Console.ReadLine();
            }
        }

        static Dictionary<string, int> ColIndexes(Excel.Range xls)
        {
            Dictionary<string, int> indexes = new Dictionary<string, int>();
            int colCount = xls.Columns.Count;

            for (int i = 1; i <= colCount; i++)
            {
                indexes[xls.Cells[1, i].Value2.ToString()] = i;
            }

            if (!indexes.ContainsKey(PROC_COL) || !indexes.ContainsKey(OP_COL))
                throw new Exception("No se encontraron las columnas '" + PROC_COL + "' y '" + OP_COL + "'");

            return indexes;
        }

        static void ProcessExcel(DirectoryInfo d, FileInfo xlsFile)
        {
            Console.WriteLine("- Extrayendo números de procedimiento de " + xlsFile.Name);

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xls = null;

            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(xlsFile.FullName);
                xlWorksheet = xlWorkbook.Sheets[1];
                xls = xlWorksheet.UsedRange;

                var indexes = ColIndexes(xls);
                CopyFiles(d, xls, indexes);

                xlWorkbook.Close();
                xlApp.Quit();
            }
            finally
            {
                if(xlApp != null) Marshal.ReleaseComObject(xlApp);
                if(xlWorkbook != null) Marshal.ReleaseComObject(xlWorkbook);
                if(xlWorksheet != null) Marshal.ReleaseComObject(xlWorksheet);
                if(xls != null) Marshal.ReleaseComObject(xls);
            }
        }

        static void CopyFiles(DirectoryInfo d, Excel.Range xls, Dictionary<string, int> indexes)
        {
            var procCol = indexes[PROC_COL];
            var opCol = indexes[OP_COL];

            var pendingPdfs = d.GetFiles("*.pdf").Select(f => f.Name).ToList();
            var totalPdfs = pendingPdfs.Count;
            var rowCount = xls.Rows.Count;

            Console.WriteLine("- Encontrados " + totalPdfs + " documentos pdf, y " + (rowCount - 1) + " filas en el excel\n\n");

            for (int i = 2; i <= rowCount; i++)
            {
                string opNum = xls.Cells[i, opCol].Value2.ToString();
                string procNum = xls.Cells[i, procCol].Value2.ToString();
                CopyProcFiles(d, opNum, procNum, pendingPdfs);
            }

            Console.WriteLine("\n\n- Copiados " + (totalPdfs - pendingPdfs.Count) + " de " + totalPdfs);
            if(pendingPdfs.Count > 0)
            {
                var regex = new Regex(@"^\d+");
                var nums = pendingPdfs.GroupBy(n => regex.Match(n).ToString())
                                      .Select(kv => "'" + kv.Key + "' (" + kv.Count() + " documentos):\n" + 
                                                    string.Join("\n", kv.Select(v => "    " + v)));
                Console.WriteLine("\n- N°s de operación no encontrados en el excel:\n" + string.Join("\n", nums));
            }
        }

        static void CopyProcFiles(DirectoryInfo d, string opNum, string procNum, List<string> pendingPdfs)
        {
            var pdfs = d.GetFiles(opNum + "-*.pdf");
            Console.WriteLine(procNum + " -> " + opNum + " (" + pdfs.Length + " archivos)");
            if (pdfs.Length == 0) return;

            d.CreateSubdirectory(procNum);
            foreach (FileInfo pdf in pdfs)
            {
                string newPath = pdf.Name.Replace(opNum + "-", procNum + "\\");
                Console.WriteLine(pdf.Name + " -> " + newPath);
                pdf.CopyTo(d.FullName + "\\" + newPath, true);
                pendingPdfs.Remove(pdf.Name);
            }
        }
    }
}
