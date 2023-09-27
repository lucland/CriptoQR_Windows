using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using LiteDB;

namespace criptoqrcode
{
    internal class PopularCadastro
    {

        public static void ParseAndInsertToLiteDB(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File {filePath} does not exist.");
                return;
            }

            var lines = File.ReadAllLines(filePath);

            using (var db = new LiteDatabase(@"C:\compartilhamento\dados\banco.db"))
            {
                var collection = db.GetCollection<Cadastro>("cadastro");

                foreach (var line in lines)
                {
                    string number = ExtractValueFromLine(line, "Number");
                    if (string.IsNullOrEmpty(number))
                    {
                        Console.WriteLine("Skipping invalid line.");
                        continue;
                    }

                    // Check if this number already exists in the database
                    var existingRecord = collection.FindOne(x => x.Number == number);

                    if (existingRecord == null)
                    {
                        var cadastro = new Cadastro(
                            number,
                            ExtractValueFromLine(line, "Name"),
                            ExtractValueFromLine(line, "Company", alternativeLabel: "Compay"),
                            ExtractValueFromLine(line, "Function", alternativeLabel: "Funcition"),
                            ExtractValueFromLine(line, "Id", alternativeLabel: "Identidade"),
                            ExtractValueFromLine(line, "E-mail"),
                            ExtractValueFromLine(line, "Vessel"),
                            ExtractValueFromLine(line, "Project", alternativeLabel: "Projec"),
                            ExtractValueFromLine(line, "ASO"),
                            ExtractValueFromLine(line, "NR34", alternativeLabel: "NR-34"),
                            ExtractValueFromLine(line, "NR10", alternativeLabel: "NR-10"),
                            ExtractValueFromLine(line, "NR33", alternativeLabel: "NR-33"),
                            ExtractValueFromLine(line, "NR35", alternativeLabel: "NR-35"),
                            // Assuming "Estado", "User", "Motivo", "Local", "Data", "Data2", and "Document" 
                            // are elsewhere in your file structure. If they aren't, you'll have to adjust accordingly.
                            // For this example, I'll use placeholders:
                            "placeholder_for_Estado", // adjust accordingly
                            "placeholder_for_User", // adjust accordingly
                            "placeholder_for_Motivo", // adjust accordingly
                            "placeholder_for_Local", // adjust accordingly
                            "placeholder_for_Data", // adjust accordingly
                            "placeholder_for_Data2", // adjust accordingly
                            "placeholder_for_Document" // adjust accordingly
                        );

                        collection.Insert(cadastro);
                        Console.WriteLine($"Inserted {cadastro.Number} into the database.");
                    }
                    else
                    {
                        Console.WriteLine($"Number {number} already exists in the database.");
                    }
                }
            }
        }

        private static string ExtractValueFromLine(string line, string label, string alternativeLabel = null)
        {
            string pattern = $@"{label}\s*:\s*([^:]+)";
            var match = Regex.Match(line, pattern);
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }

            // Try the alternative label if provided
            if (!string.IsNullOrEmpty(alternativeLabel))
            {
                pattern = $@"{alternativeLabel}\s*:\s*([^:]+)";
                match = Regex.Match(line, pattern);
                if (match.Success)
                {
                    return match.Groups[1].Value.Trim();
                }
            }

            return null; // Return null if neither primary nor alternative label is found
        }

    }
}
