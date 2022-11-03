namespace Date;
public class ClassDate
{
    private string date;
    public string day;
    public string month;
    public string yaer;
    public string d_full;
    public string d_briefly;
    public Dictionary<string, string> month_name = new Dictionary<string, string>()
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

    public ClassDate(string date)
    {
        this.date = date;
        day = date.Substring(0, 2);
        month = date.Substring(3, 2);
        yaer = date.Substring(6, 4);
        d_full = day + "." + month + "." + yaer;
        d_briefly = day + "." + month;
    }

    public string Day()
    {
        return date.Substring(8, 2);
    }
    
    public string Month()
    {
        return date.Substring(5, 2);
    }

    public string Year()
    {
        return date.Substring(0, 4);
    }

    public string D_full()
    {
        return day + "." + month + "." + yaer;
    }

    public string D_briefly()
    {
        return day + "." + month;
    }
}
