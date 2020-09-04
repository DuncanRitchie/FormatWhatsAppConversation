using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace FormatWhatsAppConversation
{
    class Program
    {
        static Application Application;
        static Document Document;

        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to the app!");
            string readFilePath = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, @"..\..\..\_chat.txt"));
            Console.WriteLine($"Reading file {readFilePath}");

            if (File.Exists(readFilePath))
            {
                Application = new Application();
                Document = Application.Documents.Add(Visible: true);

                File.ReadAllLines(readFilePath)
                    .ConvertToDocumentText()
                    .WriteToDocument(Document)
                    .ApplyHeadingStyle();

                Document.SaveAs2("hello.docx", ReadOnlyRecommended: false);
                Console.WriteLine("The document has been saved.");
            }
            else
            {
                Console.WriteLine("_chat.txt does not exist ☹️");
            }

            Console.WriteLine("The app has finished.");
            Console.ReadLine();
            CloseWord();
        }

        static void CloseWord()
        {
            Application?.Quit();
            Console.WriteLine("Word has been closed.");
        }
    }

    public static class ExtensionMethods
    {
        public static List<string> ConvertToDocumentText(this string[] textLines)
        {
            List<string> formattedText = new List<string> { };

            Regex regex = new Regex(@"\[(?<datetime>\d\d\/\d\d\/\d\d\d\d, \d\d:\d\d:\d\d)\] (?<author>[^:]+): (?<content>.*$)");

            foreach (string line in textLines)
            {
                if (line.Trim().Equals("")) { continue; }

                Match match = regex.Match(line);
                if (match.Success)
                {
                    //Console.WriteLine($"{match.Groups["author"]} at {match.Groups["datetime"]} said: {match.Groups["content"]}");

                    formattedText.Add($"{match.Groups["author"].Value} — {match.Groups["datetime"].Value}");
                    formattedText.Add(match.Groups["content"].Value);
                }
                else
                {
                    formattedText.Add(line);
                }
            }

            return formattedText;
        }

        public static Document WriteToDocument(this List<string> text, Document document)
        {
            document.Content.Text = Join(text);
            return document;
        }

        static string Join(List<string> strings)
        {
            string output = "";
            foreach (string line in strings)
            {
                output = $"{output}{line}\n";
            }
            return output;
        }

        public static void ApplyHeadingStyle(this Document document)
        {
            Regex headingRegex = new Regex(@"(?<author>[^—]+) — (?<datetime>\d\d\/\d\d\/\d\d\d\d, \d\d:\d\d:\d\d)$");
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                Match match = headingRegex.Match(paragraph.Range.Text.Trim());
                if (match.Success)
                {
                    paragraph.set_Style(WdBuiltinStyle.wdStyleHeading2);
                }
            }
        }
    }
}
