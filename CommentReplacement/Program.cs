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
                HashSet<OpenXmlElement> uniqueElementsToRemove = new HashSet<OpenXmlElement>();
                List<OpenXmlElement> elementsToRemove = new List<OpenXmlElement>();

                bool removal = false;
                Regex translationModifiedRegex = new Regex("##(.)*translation(.)*#");
                Regex trailingHashtagRegex = new Regex("^( )*##( )*$");
                Regex colorRegex = new Regex("0{4}F{2}");

                foreach (OpenXmlElement element in elements) {
                    // prepare an element for removal if enabled
                    if (removal) { elementsToRemove.Add(element); }

                    // enable removal for elements following a 'Commentaire'
                    // and clear the current to be removed elements
                    if (element.InnerText.Contains("Attribute=\"Commentaire\"")) {
                        removal = true;
                        elementsToRemove.Clear();
                    }

                    // disable removal for elements following 'Constraint modified ...'
                    // and add current elements to be removed to the hashset used for actual removal
                    if (translationModifiedRegex.IsMatch(element.InnerText)) {
                        removal = false;
                        uniqueElementsToRemove.UnionWith(elementsToRemove);
                    }

                    // prepare trailing hashtags for removal
                    if (trailingHashtagRegex.IsMatch(element.InnerText) && colorRegex.IsMatch(element.InnerXml)) {
                        uniqueElementsToRemove.Add(element);
                    }
                }

                Console.WriteLine(String.Format("Paragraphs to be removed: {0}", uniqueElementsToRemove.Count));

                foreach (OpenXmlElement elementToRemove in uniqueElementsToRemove) {
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
 