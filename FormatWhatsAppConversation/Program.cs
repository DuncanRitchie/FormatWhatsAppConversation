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

        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to the app!");
            Console.WriteLine("Reading file...");
            string readFilePath = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, @"..\..\..\_chat.txt"));
            Console.WriteLine(readFilePath);

            if (File.Exists(readFilePath))
            {
                string[] readText = File.ReadAllLines(readFilePath);
                //Console.WriteLine(readText);

                List<string> formattedText = new List<string> { };

                Regex regex = new Regex(@"\[(?<datetime>\d\d\/\d\d\/\d\d\d\d, \d\d:\d\d:\d\d)\] (?<author>[^:]+): (?<content>.*$)");

                foreach (string line in readText)
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


                Application = new Application();
                Document doc = Application.Documents.Add(Visible: true);
                var missing = Type.Missing;
                Application.Visible = true;

                doc.Content.Text = Join(formattedText);

                Regex headingRegex = new Regex(@"(?<author>[^—]+) — (?<datetime>\d\d\/\d\d\/\d\d\d\d, \d\d:\d\d:\d\d)$");
                foreach (Paragraph paragraph in doc.Paragraphs)
                {
                    Match match = headingRegex.Match(paragraph.Range.Text.Trim());
                    if (match.Success)
                    {
                        paragraph.set_Style(WdBuiltinStyle.wdStyleHeading2);
                    }
                }

                doc.SaveAs2("hello.docx", ReadOnlyRecommended: false);
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

        static string Join(List<string> strings)
        {
            string output = "";
            foreach (string line in strings)
            {
                output = $"{output}{line}\n";
            }
            return output;
        }
    }
}
