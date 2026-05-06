using SalesCalculator;
using Newtonsoft.Json;


Console.WriteLine("Введите путь к Excel-файлу:");
string xlsPath = Console.ReadLine().Replace("\"", "");

Console.WriteLine("Введите путь к JSON-словарю:");
string configPath = Console.ReadLine().Replace("\"", "");

var jsonTable = File.ReadAllText("table.json");

var configTable = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(jsonTable);

ProcessExcelService.ProcessExcel(xlsPath, configPath, configTable);



public static class Config
{
    public static Dictionary<string, string> LoadConfig(string path)
    {
        if (!File.Exists(path)) return new Dictionary<string, string>();

        return JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(path));
    }
}
