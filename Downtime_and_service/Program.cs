global using Excel = Microsoft.Office.Interop.Excel;
using ActionModule;

namespace Downtime_and_service
{
    class Program
    {
        static void Main(string[] args)
        {
            var ExcelObj = new Excel.Application();
            ExcelObj.Visible = true;
            ExcelObj.WindowState = Excel.XlWindowState.xlMaximized;

            Console.WriteLine("Введите дату отчета в формате ГГГГ.ММ.ДД");
            string? d_reverse = Console.ReadLine();
            
            var date = new Dictionary<string, string>()
            {
                ["day"] = d_reverse!.Substring(8, 2),
                ["month"] = d_reverse.Substring(5, 2),
                ["year"] = d_reverse.Substring(0, 4)
            };

            string d_full = date["day"] + "." + date["month"] + "." + date["year"];
            string d_briefly = $"{date["day"]}.{date["month"]}";

            Excel.Workbook? ExcelWorkBook_report = null;
            var ExcelWorkBook_sources = new Dictionary<string, Excel.Workbook>();
            Excel.Workbook? ExcelWorkBook_rating = null;

            while (true) {
                Console.WriteLine("Выберите функцию:");
                Console.WriteLine("1. Открыть файлы отчета");
                Console.WriteLine("2. Копировать сведения");
                Console.WriteLine("3. Сохраниить и закрыть все файлы");
                Console.WriteLine("4. Завершить скрипт");

                string? v = Console.ReadLine();

                if (v == "1")
                {
                    ExcelWorkBook_report = FileAction.FuncOpen1(ExcelObj, date);
                    FileAction.FuncOpen2(ExcelObj, date, d_briefly, ExcelWorkBook_sources);
                    ExcelWorkBook_rating = FileAction.FuncOpen3(ExcelObj, date, d_briefly, d_full);
                }
                else if(v == "2")
                {
                    FileAction.FuncCopy(ExcelObj, date, d_briefly, ExcelWorkBook_report!, ExcelWorkBook_sources, ExcelWorkBook_rating!);
                }
                else if(v == "3")
                {
                    FileAction.FuncClose(ExcelObj, date, d_full, ExcelWorkBook_report!, ExcelWorkBook_sources, ExcelWorkBook_rating!);
                }
                else if(v == "4")
                {
                    //System.Diagnostics.Process ExcelProcess = new System.Diagnostics.Process();

                    var Ex = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                    Ex[0].Kill();
                    break;
                }
            }
        }
    }
}