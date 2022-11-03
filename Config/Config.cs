namespace Config;

public class ClassConfig
{
    public static Dictionary<string, string> Get_Config()
    {
        var config_dict = new Dictionary<string, string>();
        String line;

        StreamReader sr = new StreamReader(Directory.GetCurrentDirectory() + "\\Config.txt");
        line = sr.ReadLine()!;
        while (line != null)
        {
            
            line = sr.ReadLine()!;
            while (line != "" & line != null)
            {
                var massiv = line!.Split("=");
                config_dict[massiv[0]] = massiv[1];
                line = sr.ReadLine()!;
            }
            line = sr.ReadLine()!;
        }
        sr.Close();

        return config_dict;
    }
}