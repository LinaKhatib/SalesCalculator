using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace SalesCalculator
{
    public static class PositionCheckingService
    {
        private static bool _configChanged = false;
        public static void CheckPosition(string configPath, Dictionary<string, string> config, IXLWorksheet worksheet, int lastRow)
        {
            List<string> listOfCategory = new List<string>{ "напитки", "напитки СТРД", "добавки", "Зерно", "Капсулы",
                                                   "Дрипы", "Чай", "Аксессуары", "Какао", "Шоколад, цветы, шоколад в зернах",
                                                   "Сандо", "Сэндвичи", "Сырники", "Пирожные", "Выпечка",
                                                   "конфетка", "мастер-класс", "сертификат" };
            for (int row = 2; row <= lastRow; row++)
            {
                string position = worksheet.Cell(row, 2).GetValue<string>();
                if (string.IsNullOrWhiteSpace(position)) continue;

                if (!config.ContainsKey(position))
                {
                    Console.WriteLine($"\n[!] Новая позиция: {position}");

                    Console.WriteLine("Список категорий: ");
                    for (int i = 0; i < listOfCategory.Count; i++)
                    {
                        Console.WriteLine($"{i + 1}  - {listOfCategory[i]}");
                    }

                    Console.Write("Введите номер категории для этого товара: ");
                    bool isValid = false;
                    int category;
                    while (!isValid)
                    {
                        string input = Console.ReadLine();
                        if (int.TryParse(input, out category) && category >= 1 && category <= listOfCategory.Count)
                        {
                            isValid = true;
                            config.Add(position, listOfCategory[category - 1]);
                            _configChanged = true;
                        }
                        else
                        {
                            Console.Write("Ошибка. Попробуйте повторить: ");
                        }
                    }
                    _configChanged = true;
                }
            }

            SavingNewProducts(configPath, config);
        }

        private static void SavingNewProducts(string configPath, Dictionary<string, string> config)
        {
            if (_configChanged)
            {
                File.WriteAllText(configPath, JsonConvert.SerializeObject(config, Formatting.Indented));
                Console.WriteLine("Словарь обновлен.");
            }
        }
    }
}
