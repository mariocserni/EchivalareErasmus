using System;
using ClosedXML.Excel;

using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ProiectEchivalareErasmus;

class Program
{
    static void Main(string[] args)
    {
        var subjects = Utilities.ReadFromFile();
        bool exit = false;

        while (!exit)
        {
            Console.WriteLine("\nAlegeti o optiune:");
            Console.WriteLine("1. Editeaza discipline");
            Console.WriteLine("2. Verifica creditele");
            Console.WriteLine("3. Exporta disciplinele in Excel si creeaza tabelul");
            Console.WriteLine("4. Iesire");
            Console.Write("Optiunea aleasa: ");
            string option = Console.ReadLine().ToUpper();

            switch (option)
            {
                case "1":
                    Utilities.EditSubjects(subjects);
                    break;
                case "2":
                    if (!Utilities.VerifyCredits(subjects))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Creditele de la universitatea gazda nu sunt destule si nu vor putea fi echivalate!");
                        Console.ResetColor();
                    }
                    break;
                case "3":
                    Utilities.ExportToExcel(subjects);
                    Utilities.GenerateDocx(subjects, Config.DocxFilePath);
                    Console.WriteLine("Datele au fost exportate in Excel si DOCX cu succes.");
                    break;
                case "4":
                    exit = true;
                    break;
                default:
                    Console.WriteLine("Optiune invalida. Incercati din nou.");
                    break;
            }
        }
    }
}



