namespace Downtime_and_service;

class ClassWorkbook
{
    public static void Report_create(Date.ClassDate date, Excel.Application excel, File.ClassFile report_config, Excel.Workbook excelWorkBook_report)
    {
        var excelWorkSheet_rating = report_config.Activate_sheet(excelWorkBook_report, "Источник Рейтинг");
        excelWorkSheet_rating.Visible = Excel.XlSheetVisibility.xlSheetVisible;
        excel.Run("Выкладки");

        var excelWorkSheet_install = report_config.Activate_sheet(excelWorkBook_report, "Установочные");
        string? amount_line_old_start = (string)excelWorkSheet_install.Range["B2"].Text;
        string? amount_line_old_end = Convert.ToString(Int32.Parse(amount_line_old_start) - 13);
        string? amount_line_next_start = Convert.ToString(Convert.ToInt32(amount_line_old_start) + 1);
        string? amount_line_next_end = Convert.ToString(Convert.ToInt32(amount_line_old_start) + 14);
            
        excelWorkSheet_rating.Range["A" + amount_line_old_start + ":AC" + amount_line_old_end].Copy();
        excelWorkSheet_rating.Paste(excelWorkSheet_rating.Range["A" + amount_line_next_start]);
        //DateTime date1 = new DateTime(Int32.Parse(config["year"]), Int32.Parse(config["month"]), Int32.Parse(config["day"]));
        //DateTime date2 = DateTime.Parse(config["date_current_report"]);
        DateTime date_val = Convert.ToDateTime(date.d_full);
        excelWorkSheet_rating.Range["B" + amount_line_next_start + ":B" + amount_line_next_end].Value = date_val;
    }

    public static void Report_copy(Date.ClassDate date, Excel.Application excel, File.ClassFile report_config, Excel.Workbook excelWorkBook_report, Excel.Workbook excelWorkBook_rating)
    {
        if (excel != null & report_config != null & excelWorkBook_report != null & excelWorkBook_rating != null)
        {
            var sheet_rating = report_config!.Activate_sheet(excelWorkBook_report!, "Рейтинг");
                sheet_rating.Activate();
                sheet_rating.Range["R2"].Value = Convert.ToDateTime(date.d_full);
                excel!.Run("Сортировка_рейтинга");
                sheet_rating.Range["A22", "B34"].Copy();

            var sort_rating = report_config.Activate_sheet(excelWorkBook_rating!, date.d_briefly);
                var paste = Excel.XlPasteType.xlPasteValues;
                sort_rating.Range["A3"].PasteSpecial(paste);
                var sort = Excel.XlSortOrder.xlAscending;
                dynamic range2 = sort_rating.Range["A3", "B15"];
                range2.Sort(range2.Columns[1], sort);
                //range2.Sort(range2.Columns.Item[1], sort);
        }
    }

    public static void Report_save(Excel.Application excel, File.ClassFile report_config, Excel.Workbook excelWorkBook_report)
    {
        if (excel != null & report_config != null & excelWorkBook_report != null)
        {
            var excelWorkSheet = report_config!.Activate_sheet(excelWorkBook_report!, "Источник Рейтинг");
                excelWorkSheet.Activate();
                excel!.Run("Выкладки");
                excelWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;

            var excelWorkSheet_active = report_config.Activate_sheet(excelWorkBook_report!, "Рейтинг");
                excelWorkSheet_active.Activate();
                excel.DisplayAlerts = false;

            excelWorkBook_report!.SaveAs(report_config.save_link, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
        }
    }

    public static void Sources_create(Date.ClassDate date, File.ClassFile sources_config, Excel.Workbook excelWorkBook_sources)
    {
        int last_sheet = excelWorkBook_sources.Worksheets.Count;

        int checker = 0;
        foreach (Excel.Worksheet sheet_existing in excelWorkBook_sources.Sheets)
        {
            if (sheet_existing.Name == date.d_full)
            {
                checker = 1;
            }
        }
        if (checker == 1)
        {
            var ExcelWorkSheet_sheet1 = (Excel.Worksheet)excelWorkBook_sources.Worksheets.Add(System.Reflection.Missing.Value, excelWorkBook_sources.Worksheets[last_sheet]);
            ExcelWorkSheet_sheet1.Name = "Лист1";
        }
        else
        {
            var ExcelWorkSheet_brefly = (Excel.Worksheet)excelWorkBook_sources.Worksheets.Add(System.Reflection.Missing.Value, excelWorkBook_sources.Worksheets[last_sheet]);
            ExcelWorkSheet_brefly.Name = date.d_briefly;
        }

        var ExcelWorkSheet_install = sources_config.Activate_sheet(excelWorkBook_sources, "Установочные");
        ExcelWorkSheet_install.Range["B1"].Value = date.d_briefly;
    }

    public static void Sources_copy(Date.ClassDate date, File.ClassFile sources_config, Excel.Workbook excelWorkBook_sources, Excel.Workbook excelWorkBook_report, int row)
    {
        if (sources_config != null & excelWorkBook_sources != null & excelWorkBook_report != null)
        {
            var value1 = sources_config!.Activate_sheet(excelWorkBook_sources!, "Установочные");
                string range = (string)value1.Range["B2"].Text;
                value1.Range[range].Copy();
            
            var sheet_installation = sources_config.Activate_sheet(excelWorkBook_report!, "Установочные");
                var q = sources_config.Activate_range(sheet_installation, row, 2);
                int amount_line_old = Convert.ToInt32(q.Text);
                var worksheet = sources_config.Activate_sheet(excelWorkBook_report, sources_config.name_rus);
                worksheet.Range["C" + (amount_line_old + 1)].PasteSpecial(Excel.XlPasteType.xlPasteValues);

                var q2 = sources_config.Activate_range(sheet_installation, row, 2);
                int amount_line_new = Convert.ToInt32(q2.Text);
                worksheet.Range["B" + (amount_line_old + 1), "B" + amount_line_new].Value = Convert.ToDateTime(date.d_full);
                worksheet.Range["A" + amount_line_old].Copy();
                worksheet.Range["A" + (amount_line_old + 1), "A" + amount_line_new].PasteSpecial();
        }
    }

    public static void Rating_create(Date.ClassDate date, Excel.Application excel, File.ClassFile rating_config, Excel.Workbook excelWorkBook_rating)
    {
        int last_sheet = excelWorkBook_rating.Worksheets.Count;
        var ExcelWorkSheet_new_list = (Excel.Worksheet)excelWorkBook_rating.Worksheets.Add(System.Reflection.Missing.Value, excelWorkBook_rating.Worksheets[last_sheet]);
        ExcelWorkSheet_new_list.Name = date.d_briefly;
        ExcelWorkSheet_new_list.Range["A1"].Value = "Рейтинг подразделений по 4 показателям за " + date.d_full;
    }

    public static void Source_and_rating_save(Excel.Workbook excelWorkBook_sources_and_rating)
    {
        if (excelWorkBook_sources_and_rating != null)
        {
            excelWorkBook_sources_and_rating.Save();
        }
    }
}