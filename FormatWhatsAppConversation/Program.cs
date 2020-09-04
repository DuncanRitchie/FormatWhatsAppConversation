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
            Console.WriteLine("Re-arranging the text for the document...");
            List<string> formattedText = new List<string> { };

            Regex mainRegex = new Regex(@"\[(?<datetime>\d\d\/\d\d\/\d\d\d\d, \d\d:\d\d:\d\d)\] (?<author>[^:]+): (?<content>.*$)");
            Regex titleRegex = new Regex(@"\[(?<datetime>\d\d\/\d\d\/\d\d\d\d, \d\d:\d\d:\d\d)\] (?<creator>.+) created group “(?<title>.+)”");
            string groupStartMessage = "Messages and calls are end-to-end encrypted. No one outside of this chat, not even WhatsApp, can read or listen to them.";

            foreach (string line in textLines)
            {
                if (line.Trim().Equals("")) { continue; }

                Match mainMatch = mainRegex.Match(line);
                if (mainMatch.Success)
                {
                    if (mainMatch.Groups["content"].Value.Contains(groupStartMessage))
                    {
                        Console.WriteLine("Ignoring WhatsApp’s message about encryption.");
                    }
                    else
                    {
                        formattedText.Add($"{mainMatch.Groups["author"].Value} — {mainMatch.Groups["datetime"].Value}");
                        formattedText.Add(mainMatch.Groups["content"].Value);
                    }
                    continue;
                }

                Match titleMatch = titleRegex.Match(line);
                if (titleMatch.Success)
                {
                    formattedText.Add(titleMatch.Groups["title"].Value);
                    formattedText.Add($"{titleMatch.Groups["creator"]} created this group at {titleMatch.Groups["datetime"]}");
                }
                else
                {
                    formattedText.Add(line);
                }
            }

            Console.WriteLine("The text has been arranged into the format for the document.");
            return formattedText;
        }

        public static Document WriteToDocument(this List<string> text, Document document)
        {
            Console.WriteLine("Writing the text to the document...");
            document.Content.Text = Join(text);
            Console.WriteLine("The text has been written to the document.");
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
            string titleText = document.Paragraphs[1].Range.Text.Trim();
            document.Paragraphs[1].set_Style(WdBuiltinStyle.wdStyleTitle);
            Console.WriteLine($"“{titleText}” has been styled as the document title.");

            Console.WriteLine("Applying heading styles...");
            Regex headingRegex = new Regex(@"(?<author>[^—]+) — (?<datetime>\d\d\/\d\d\/\d\d\d\d, \d\d:\d\d:\d\d)$");
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                Match match = headingRegex.Match(paragraph.Range.Text.Trim());
                if (match.Success)
                {
                    paragraph.set_Style(WdBuiltinStyle.wdStyleHeading2);
                }
            }
            Console.WriteLine("All heading styles have been applied.");
        }
    }
}
