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
        static string InputTextFilePath = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, @"..\..\..\_chat.txt"));
        static string OutputDocumentFileName = "formatted-chat.docx";

        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to the app!");

            PromptForInputFile();
            Console.WriteLine($"Reading file {InputTextFilePath}...");

            if (File.Exists(InputTextFilePath))
            {
                Application = new Application();
                Document = Application.Documents.Add(Visible: true);

                File.ReadAllLines(InputTextFilePath)
                    .ConvertToDocumentText()
                    .WriteToDocument(Document)
                    .ApplyHeadingStyle()
                    .InsertPictures(InputTextFilePath);

                Document.SaveAs2(OutputDocumentFileName, ReadOnlyRecommended: false);
                Console.WriteLine($"The document has been saved in your Documents folder as {OutputDocumentFileName}.");
            }
            else
            {
                Console.WriteLine("_chat.txt does not exist ☹️");
            }

            Console.WriteLine("The app has finished.");
            CloseWord();
            Console.ReadLine();
        }

        static void PromptForInputFile()
        {
            Console.WriteLine("Please enter the folder path containg _chat.txt.");
            string userEntry = Console.ReadLine();
            string inputTextFilePath = Path.GetFullPath(Path.Combine(userEntry, @"_chat.txt"));
            if (File.Exists(inputTextFilePath))
            {
                InputTextFilePath = inputTextFilePath;
            }
            else
            {
                Console.WriteLine($"{inputTextFilePath} does not exist.");
                PromptForInputFile();
            }
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
            List<string> paragraphTexts = new List<string> { };

            Regex mainRegex = new Regex(@"\[(?<datetime>\d\d\/\d\d\/\d\d\d\d, \d\d:\d\d:\d\d)\] (?<author>[^:]+): (?<content>.*$)");
            Regex titleRegex = new Regex(@"\[(?<datetime>\d\d\/\d\d\/\d\d\d\d, \d\d:\d\d:\d\d)\] (?<creator>.+) created group “(?<title>.+)”");
            string groupStartMessage = "Messages and calls are end-to-end encrypted. No one outside of this chat, not even WhatsApp, can read or listen to them.";

            foreach (string line in textLines)
            {
                //// Omit empty paragraphs.
                if (line.Trim().Equals("")) { continue; }

                //// If the paragraph contains a datetime, author, and content...
                Match mainMatch = mainRegex.Match(line);
                if (mainMatch.Success)
                {
                    //// The default message about WhatsApp encryption gets omitted.
                    if (mainMatch.Groups["content"].Value.Contains(groupStartMessage))
                    {
                        Console.WriteLine("Ignoring WhatsApp’s message about encryption.");
                    }
                    //// Non-WhatsApp messages.
                    else
                    {
                        //// This will match headingRegex and be formatted as Heading 2, in ApplyHeadingStyle.
                        paragraphTexts.Add($"{mainMatch.Groups["author"].Value} — {mainMatch.Groups["datetime"].Value}");
                        //// This is just the (first paragraph of) content.
                        paragraphTexts.Add(mainMatch.Groups["content"].Value);
                    }
                    continue;
                }

                //// If the line is the default message about who created the WhatsApp group...
                Match titleMatch = titleRegex.Match(line);
                if (titleMatch.Success)
                {
                    //// ApplyHeadingStyle will style this as Title if it is the first paragraph in the document.
                    paragraphTexts.Add(titleMatch.Groups["title"].Value);
                    //// This will probably be the second paragraph in the document.
                    paragraphTexts.Add($"{titleMatch.Groups["creator"]} created this group at {titleMatch.Groups["datetime"]}");
                }
                //// Lines that don’t match the regexes are paragraphs that aren’t first for their author and datetime.
                else
                {
                    paragraphTexts.Add(line);
                }
            }

            Console.WriteLine("The text has been arranged into the format for the document.");
            return paragraphTexts;
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

        public static Document ApplyHeadingStyle(this Document document)
        {
            //// Style the first paragraph as the title.
            string titleText = document.Paragraphs[1].Range.Text.Trim();
            document.Paragraphs[1].set_Style(WdBuiltinStyle.wdStyleTitle);
            Console.WriteLine($"“{titleText}” has been styled as the document title.");

            //// Apply style of Heading 2 to paragraphs containing an author and a datetime.
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
            return document;
        }

        public static Document InsertPictures(this Document document, string filepathInFolder)
        {
            Console.WriteLine("Looking for places to insert pictures...");
            Regex attachedRegex = new Regex(@"(?<attachedTag>‎<attached: (?<filename>[\w-]+\.(?<extension>\w{2,6})))>");
            float desiredSize = 360;

            foreach (Paragraph paragraph in document.Paragraphs)
            {
                Match attachedMatch = attachedRegex.Match(paragraph.Range.Text);
                if (attachedMatch.Success && attachedMatch.Groups["extension"].Value != "mp4" && attachedMatch.Groups["extension"].Value != "vcf")
                {
                    try
                    {
                        string pictureFilePath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(filepathInFolder), attachedMatch.Groups["filename"].Value));
                        paragraph.Range.Text = "\r";
                        Console.WriteLine($"Inserting a picture: {pictureFilePath} ...");
                        InlineShape picture = paragraph.Previous().Range.InlineShapes.AddPicture(pictureFilePath);
                        picture.Title = attachedMatch.Groups["filename"].Value;

                        //// Resize the image, preserving aspect ratio.
                        float aspectRatio = picture.Width / picture.Height;
                        if (aspectRatio > 1)
                        {
                            picture.Width = desiredSize;
                            picture.Height = desiredSize / aspectRatio;
                        }
                        else
                        {
                            picture.Height = desiredSize;
                            picture.Width = desiredSize * aspectRatio;
                        }

                        //// Delete the “<attached: ...>” text.
                        //paragraph.Range.Text = paragraph.Range.Text.Replace(attachedMatch.Groups["attachedTag"].Value, "");
                    }
                    catch (Exception error)
                    {
                        Console.WriteLine(error.Message);
                    }
                }
            }
            return document;
        }
    }
}
