using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Xceed.Words.NET;

class Program
{
    static void Main()
    {
        string filePath = "C:\\Users\\Rog\\Desktop\\ПРОЕКТ\\shevchenko-taras-hryhorovych-kobzar3545.docx"; 
        string text;

        using (var document = DocX.Load(filePath))
        {
            text = document.Text;
        }

        var words = Regex.Replace(text, @"[^\w\s]", "").Split(new[] { ' ', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

        var wordCounts = words.Where(word => word.Length >= 5).GroupBy(word => word).Select(group => new { Word = group.Key, Count = group.Count() })
            .OrderByDescending(x => x.Count).Take(50).ToList();

        Console.WriteLine(" №  | Слово          | Встретилось раз");
        Console.WriteLine(new string('-', 35));

        int index = 1;
        foreach (var entry in wordCounts)
        {
            Console.WriteLine($" {index++,-2} | {entry.Word,-15} | {entry.Count}");
        }
    }
}
