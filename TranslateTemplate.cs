using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;

namespace TranslateTemplateProject
{
    public class TranslateTemplate
    {
        public static void Process(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: dotnet run <dotx_path> <language>");
                return;
            }
            string dotxPath = args[0];
            string newLanguage = args[1];
            if (newLanguage != "nn" && newLanguage != "en")
            {
                Console.WriteLine("Language must be 'nn' or 'en'");
                return;
            }
            if (!File.Exists(dotxPath))
            {
                Console.WriteLine($"The file {dotxPath} does not exist.");
                return;
            }
            Console.WriteLine($"Processing file: {dotxPath}");
            var translations = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(File.ReadAllText("utrykk.json"));
            try
            {
                string newDotxPath = dotxPath.Replace("_nb", $"_{newLanguage}");
                File.Copy(dotxPath, newDotxPath, true);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(newDotxPath, true))
                {
                    ReplaceTextInDocument(wordDoc, translations, newLanguage);
                }
                Console.WriteLine($"File saved as: {newDotxPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file {dotxPath}: {ex.Message}");
            }
            Console.WriteLine("Process completed successfully.");
        }

        static void ReplaceTextInDocument(WordprocessingDocument wordDoc, Dictionary<string, Dictionary<string, string>> translations, string newLanguage)
        {
            foreach (var text in wordDoc.MainDocumentPart.Document.Body.Descendants<Text>())
            {
                if (translations.ContainsKey(text.Text) && translations[text.Text].ContainsKey(newLanguage))
                {
                    text.Text = translations[text.Text][newLanguage];
                }
            }

            foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
            {
                foreach (var text in headerPart.Header.Descendants<Text>())
                {
                    if (translations.ContainsKey(text.Text) && translations[text.Text].ContainsKey(newLanguage))
                    {
                        text.Text = translations[text.Text][newLanguage];
                    }
                }
            }

            foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
            {
                foreach (var text in footerPart.Footer.Descendants<Text>())
                {
                    if (translations.ContainsKey(text.Text) && translations[text.Text].ContainsKey(newLanguage))
                    {
                        text.Text = translations[text.Text][newLanguage];
                    }
                }
            }
        }
    }
}