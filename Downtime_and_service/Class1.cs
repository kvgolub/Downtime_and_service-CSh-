namespace PathModule
{
    class FilePath
    {
        public string? path_month;
        public string? path_yaer;
        public string? path_month_name;
        public string? path_directory;
        public string? path_source;

        public FilePath(Dictionary<string, string> date)
        {
            var month_name = new Dictionary<string, string>()
            {
                ["01"] = "Январь",
                ["02"] = "Февраль",
                ["03"] = "Март",
                ["04"] = "Апрель",
                ["05"] = "Май",
                ["06"] = "Июнь",
                ["07"] = "Июль",
                ["08"] = "Август",
                ["09"] = "Сентябрь",
                ["10"] = "Октябрь",
                ["11"] = "Ноябрь",
                ["12"] = "Декабрь"
            };

            this.path_month = date["month"];
            path_month_name = month_name[this.path_month];
            this.path_yaer = date["year"];
            path_directory = "C:\\Работа\\2. Отчеты\\1. Ежедневный\\4. Простои и сервис\\" + this.path_yaer + "\\" + this.path_month + ". " + this.path_month_name + "\\";
            path_source = $"Исходники из 1С_{ this.path_month_name} { this.path_yaer}\\";
        }

        public string[] source_name_eng = {
            "ExcelWorkBook_operators_CA",
            "ExcelWorkBook_operators_CC",
            "ExcelWorkBook_technician",
            "ExcelWorkBook_revenue",
            "ExcelWorkBook_amount",
            "ExcelWorkBook_not_connection",
            "ExcelWorkBook_not_work"
        };

        public string[] source_name_rus = {
            "1.1. Операторы (1С, Все ТА)",
            "1.2. Операторы (1С, КУ)",
            "2. Техники (1С)",
            "3. % потерянной выручки (1С)",
            "4. Кол-во ТА (1С)",
            "5. ТА без связи (Вин)",
            "6. Простои (Вин)"
        };
    }
}


/*
    [hashtable]
    file()
    {
        $source = @{ }
        $index = 0
        foreach ($node in $this.source_name_rus) {
            $source.Add($this.source_name_eng[$index], $node + ".xlsx")
            $index += 1
        }
        return ($source)
    }
*/