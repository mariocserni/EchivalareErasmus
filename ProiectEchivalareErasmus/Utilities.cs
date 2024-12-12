using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProiectEchivalareErasmus
{
    internal class Utilities
    {
        public static List<Subject> ReadFromFile()
        {
            var subjects = new List<Subject>();

            using (var workbook = new XLWorkbook(Config.FilePath))
            {
                var worksheet = workbook.Worksheet(1);
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    if (string.IsNullOrEmpty(row.Cell(1).GetValue<string>()))
                    {
                        break;
                    }
                    int number, year, semester, creditsUPT, hostCredits, grade;
                    int.TryParse(row.Cell(1).GetValue<string>(), out number);
                    int.TryParse(row.Cell(3).GetValue<string>(), out year);
                    int.TryParse(row.Cell(4).GetValue<string>(), out semester);
                    int.TryParse(row.Cell(5).GetValue<string>(), out creditsUPT);
                    int.TryParse(row.Cell(7).GetValue<string>(), out hostCredits);
                    int.TryParse(row.Cell(9).GetValue<string>(), out grade);

                    var subject = new Subject(number, row.Cell(2).GetValue<string>(), year, semester, creditsUPT, row.Cell(6).GetValue<string>(), hostCredits, grade);
                    subjects.Add(subject);
                }
            }

            foreach (var subject in subjects)
            {
                Console.WriteLine($"{subject.Number} : {subject.DisciplineUPT} - {subject.HostDiscipline} : {subject.CreditsUPT} - {subject.HostCredits}, Nota : {subject.Grade}");
            }
            Console.WriteLine("\n");
            return subjects;
        }

        public static bool VerifyCredits(List<Subject> subjects)
        {
            int totalCreditsUPT = 0;
            int totalHostCredits = 0;

            foreach (var subject in subjects)
            {
                totalCreditsUPT += subject.CreditsUPT;
                totalHostCredits += subject.HostCredits;
            }

            Console.WriteLine();
            Console.WriteLine($"Total UPT Credits: {totalCreditsUPT}");
            Console.WriteLine($"Total Host Credits: {totalHostCredits}\n");

            return totalHostCredits >= totalCreditsUPT;
        }

        public static void DisplaySubjects(List<Subject> subjects)
        {
            Console.WriteLine("Lista de discipline:");
            for (int i = 0; i < subjects.Count; i++)
            {
                var subject = subjects[i];
                Console.WriteLine($"{i + 1}. {subject.DisciplineUPT} - {subject.HostDiscipline}");
            }
        }

        public static void EditSubject(List<Subject> subjects, int index)
        {
            Subject subject = subjects[index];
            Console.WriteLine();
            Console.WriteLine("Alegeti ce doriti sa editati:");
            Console.WriteLine($"1. Discipline UPT (Actual: {subject.DisciplineUPT})");
            Console.WriteLine($"2. Discipline gazda (Actual: {subject.HostDiscipline})");
            Console.Write("Optiunea aleasa: ");
            string choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    Console.Write("Introduceti noua disciplina UPT: ");
                    subject.DisciplineUPT = Console.ReadLine();
                    break;
                case "2":
                    Console.Write("Introduceti noua disciplina gazda: ");
                    subject.HostDiscipline = Console.ReadLine();
                    break;
                default:
                    Console.WriteLine("Optiune invalida.");
                    break;
            }
        }

        public static void EditSubjects(List<Subject> subjects)
        {
            bool returnToMenu = false;
            while (!returnToMenu)
            {
                Console.Clear();
                Utilities.DisplaySubjects(subjects);
                Console.WriteLine();
                Console.WriteLine("Introduceti numarul disciplinei pentru editare, 'R' pentru a reveni la meniul principal:");

                string input = Console.ReadLine();
                if (input.ToUpper() == "R")
                {
                    returnToMenu = true;
                    continue;
                }

                int index;
                if (int.TryParse(input, out index) && index >= 1 && index <= subjects.Count)
                {
                    Utilities.EditSubject(subjects, index - 1);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("\nDisciplina a fost actualizata cu succes.");
                    Console.ResetColor();
                    Console.WriteLine("\nApasati orice tastă pentru a continua...");
                    Console.ReadKey();
                }
                else
                {
                    Console.WriteLine("Numarul disciplinei este invalid.");
                    Console.WriteLine("\nApasati orice tasta pentru a încerca din nou...");
                    Console.ReadKey();
                }
            }
        }

        public static void ExportToExcel(List<Subject> subjects)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Discipline");
                worksheet.Cell("A1").Value = "Nr. crt.";
                worksheet.Cell("B1").Value = "Disciplina UPT";
                worksheet.Cell("C1").Value = "An";
                worksheet.Cell("D1").Value = "Semestru";
                worksheet.Cell("E1").Value = "Nr. credite";
                worksheet.Cell("F1").Value = "Disciplina univ. gazda";
                worksheet.Cell("G1").Value = "Nr. credite";

                int row = 2;
                foreach (var subject in subjects)
                {
                    worksheet.Cell(row, 1).Value = subject.Number;
                    worksheet.Cell(row, 2).Value = subject.DisciplineUPT;
                    worksheet.Cell(row, 3).Value = subject.Year;
                    worksheet.Cell(row, 4).Value = subject.Semester;
                    worksheet.Cell(row, 5).Value = subject.CreditsUPT;
                    worksheet.Cell(row, 6).Value = subject.HostDiscipline;
                    worksheet.Cell(row, 7).Value = subject.HostCredits;
                    worksheet.Cell(row, 8).Value = "";  
                    worksheet.Cell(row, 9).Value = subject.Grade; 
                    row++;
                }

                row++;
                worksheet.Cell(row, 2).Value = "total";
                worksheet.Cell(row, 3).Value = subjects.Sum(s => s.CreditsUPT);
                worksheet.Cell(row, 6).Value = "total";  
                worksheet.Cell(row, 7).Value = subjects.Sum(s => s.HostCredits);

                string newFilePath = Config.FilePath.Replace(".xlsx", " - modificat.xlsx");
                workbook.SaveAs(newFilePath);
                Console.WriteLine($"Datele au fost exportate cu succes in {newFilePath}");
            }
        }

        public static void GenerateDocx(List<Subject> subjects, string filePath)
        {
            using (var doc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                Table table = new Table();

                // Define the table properties and width
                TableProperties tblProps = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                        new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                        new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                        new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                        new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                        new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 }
                    )
                );
                table.Append(tblProps);

                // Adding the header row
                TableRow headerRow = new TableRow();
                string[] headers = new string[] { "Nr. crt.", "Disciplina urmată la universitatea gazdă", "Disciplina UPT cu care se echivalează", "Credite UPT", "Nota" };
                foreach (var header in headers)
                {
                    TableCell headerCell = CreateTableCell(header, true);
                    headerRow.Append(headerCell);
                }
                table.Append(headerRow);

                HashSet<string> displayedHostDisciplines = new HashSet<string>();

                int counter = 1;
                foreach (var subject in subjects)
                {
                    TableRow row = new TableRow();

                    row.Append(CreateTableCell(counter.ToString(), false)); // Nr. crt.
                    if (!displayedHostDisciplines.Contains(subject.HostDiscipline))
                    {
                        row.Append(CreateTableCell(subject.HostDiscipline, false)); // Disciplina gazdă
                        displayedHostDisciplines.Add(subject.HostDiscipline);
                    }
                    else
                    {
                        row.Append(CreateTableCell("", false)); // Empty for duplicate host discipline
                    }
                    row.Append(CreateTableCell(subject.DisciplineUPT, false)); // Disciplina UPT
                    row.Append(CreateTableCell(subject.CreditsUPT.ToString(), false)); // Credite UPT
                    row.Append(CreateTableCell(subject.Grade.ToString(), false)); // Nota

                    table.Append(row);
                    counter++;
                }

                body.Append(table);
                mainPart.Document.Save();
            }
        }

        private static TableCell CreateTableCell(string text, bool isHeader)
        {
            TableCell cell = new TableCell();

            // Set the cell properties for borders
            TableCellProperties cellProps = new TableCellProperties();
            TableCellBorders borders = new TableCellBorders(
                new TopBorder { Val = BorderValues.Single, Size = 8 },
                new BottomBorder { Val = BorderValues.Single, Size = 8 },
                new LeftBorder { Val = BorderValues.Single, Size = 8 },
                new RightBorder { Val = BorderValues.Single, Size = 8 }
            );
            cellProps.Append(borders);

            if (isHeader)
            {
                // Apply shading for header cells
                cellProps.Append(new Shading { Fill = "D9D9D9" });
            }

            cell.Append(cellProps);
            cell.Append(new Paragraph(new Run(new Text(text))));

            return cell;
        }




    }
}
