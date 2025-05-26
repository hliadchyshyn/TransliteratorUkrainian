using ClosedXML.Excel;
using System.Text;

class Program
{
    static void Main()
    {
        var inputFile = "input.xlsx";  // Вхідний Excel файл
        var outputFile = "output.csv"; // Вихідний CSV файл

        // Перевірка наявності файлу перед тим, як спробувати його відкрити
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Файл {inputFile} не знайдено!");
            return;
        }

        var result = ReadExcelAndTransliterate(inputFile);

        // Запис результату у CSV файл
        File.WriteAllLines(outputFile, result);
        Console.WriteLine($"Транслітерація завершена. Результат збережено в {outputFile}");
    }

    static string[] ReadExcelAndTransliterate(string inputFile)
    {
        var transliteratedRows = new List<string>();

        // Відкриваємо Excel файл за допомогою ClosedXML
        using (var workbook = new XLWorkbook(inputFile))
        {
            var worksheet = workbook.Worksheet(1); // Перше робоче лист
            int rowCount = worksheet.RowsUsed().Count();
            int colCount = worksheet.ColumnsUsed().Count();

            // Перебір всіх рядків та стовпців
            for (int row = 1; row <= rowCount; row++)
            {
                StringBuilder rowData = new StringBuilder();

                // Перебір всіх стовпців
                for (int col = 1; col <= colCount; col++)
                {
                    var cellValue = worksheet.Cell(row, col).GetValue<string>();
                    string transliteratedValue = TransliterateUkrainian(cellValue);
                    rowData.Append(transliteratedValue);

                    if (col < colCount)
                        rowData.Append(",");
                }

                transliteratedRows.Add(rowData.ToString());
            }
        }

        return transliteratedRows.ToArray();
    }

    // Функція для транслітерації українського тексту
    static string TransliterateUkrainian(string text)
    {
        string[,] cyrillicToLatin = new string[,]
        {
            { "А", "A" }, { "Б", "B" }, { "В", "V" }, { "Г", "H" }, { "Ґ", "G" }, { "Д", "D" },
            { "Е", "E" }, { "Є", "Ye" }, { "Ж", "Zh" }, { "З", "Z" }, { "И", "Y" }, { "І", "I" },
            { "Ї", "Yi" }, { "Й", "Y" }, { "К", "K" }, { "Л", "L" }, { "М", "M" }, { "Н", "N" },
            { "О", "O" }, { "П", "P" }, { "Р", "R" }, { "С", "S" }, { "Т", "T" }, { "У", "U" },
            { "Ф", "F" }, { "Х", "Kh" }, { "Ц", "Ts" }, { "Ч", "Ch" }, { "Ш", "Sh" }, { "Щ", "Shch" },
            { "Ь", "" }, { "Ю", "Yu" }, { "Я", "Ya" },
            { "а", "a" }, { "б", "b" }, { "в", "v" }, { "г", "h" }, { "ґ", "g" }, { "д", "d" },
            { "е", "e" }, { "є", "ie" }, { "ж", "zh" }, { "з", "z" }, { "и", "y" }, { "і", "i" },
            { "ї", "i" }, { "й", "y" }, { "к", "k" }, { "л", "l" }, { "м", "m" }, { "н", "n" },
            { "о", "o" }, { "п", "p" }, { "р", "r" }, { "с", "s" }, { "т", "t" }, { "у", "u" },
            { "ф", "f" }, { "х", "kh" }, { "ц", "ts" }, { "ч", "ch" }, { "ш", "sh" }, { "щ", "shch" },
            { "ь", "" }, { "ю", "iu" }, { "я", "ia" }
        };

        for (int i = 0; i < cyrillicToLatin.GetLength(0); i++)
        {
            text = text.Replace(cyrillicToLatin[i, 0], cyrillicToLatin[i, 1]);
        }

        return text;
    }
}
