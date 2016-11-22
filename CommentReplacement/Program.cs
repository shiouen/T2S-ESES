using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace CommentReplacement {
    public class Program {
        public static void Main(string[] args) {
            Stopwatch stopwatch = Stopwatch.StartNew();

            string path = Path.Combine(Environment.CurrentDirectory, @"Files\", "test.docx");

            using (WordprocessingDocument document = WordprocessingDocument.Open(path, true)) {
                IEnumerable<OpenXmlElement> elements = document.MainDocumentPart.Document.Body.ToList();
                HashSet<OpenXmlElement> elementsToRemove = new HashSet<OpenXmlElement>();

                bool removal = false;
                Regex translationModifiedRegex = new Regex("##(.)*translation(.)*#");
                Regex trailingHashtagRegex = new Regex("^##$");

                foreach (OpenXmlElement element in elements) {
                    // prepare an element for removal if enabled
                    if (removal) { elementsToRemove.Add(element); }

                    // enable removal for elements following a 'Commentaire'
                    if (element.InnerText.Contains("Attribute=\"Commentaire\"")) { removal = true; }

                    // disable removal for elements following 'Constraint modified ...'
                    if (translationModifiedRegex.IsMatch(element.InnerText)) { removal = false; }

                    // prepare trailing hashtags for removal
                    if (trailingHashtagRegex.IsMatch(element.InnerText)) {
                        // in some cases no 'Constraint modifed ...', but only trailing
                        removal = false;
                        elementsToRemove.Add(element);
                    }
                }

                Console.WriteLine(String.Format("Paragraphs to be removed: {0}", elementsToRemove.Count));

                foreach (OpenXmlElement elementToRemove in elementsToRemove) {
                    elementToRemove.RemoveAllChildren();
                    elementToRemove.Remove();
                }

                document.MainDocumentPart.Document.Save();
            }

            stopwatch.Stop();
            Console.WriteLine(String.Format("Stopwatch: {0}", stopwatch.ElapsedMilliseconds));
            Console.Read();
        }
    }
}
 