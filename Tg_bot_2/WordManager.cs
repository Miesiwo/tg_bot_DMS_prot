using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Tg_bot_2
{
    class WordManager
    {
        public static string origRaport;
        public static string copyRaport;

        public enum TypesOfDoc
        {
            Рапорт_майно,
            Рапорт_виплата,
            Рапорт_травма

        }

        public static void pathMaker(int type)
        {
            
            string sourceRaport = ((WordManager.TypesOfDoc)type).ToString();
            int numerator = int.Parse(File.ReadAllText(@"C:\Users\Cadet\OneDrive\Робочий стіл\proj\Копии\Numerator.txt"));
            origRaport = $@"C:\Users\Cadet\OneDrive\Робочий стіл\proj\Шаблоны\" + sourceRaport + ".docx";
            copyRaport = $@"C:\Users\Cadet\OneDrive\Робочий стіл\proj\Копии\" + sourceRaport + "_" + numerator.ToString() + ".docx";


            File.Delete(@"C:\Users\Cadet\OneDrive\Робочий стіл\proj\Копии\Numerator.txt");
            numerator++;
            File.WriteAllText(@"C:\Users\Cadet\OneDrive\Робочий стіл\proj\Копии\Numerator.txt", numerator.ToString());


        }

        public static void CopyMaker(string source, string copy)
        {
            using (WordprocessingDocument mainDoc = WordprocessingDocument.Open(source, true))
            {
                using (var resultDoc = WordprocessingDocument.Create(copy,
                  WordprocessingDocumentType.Document))
                {
                    foreach (var part in mainDoc.Parts)
                        resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId);
                }

            }
        }

        public static void Replacer(List<string> replaceStrings, int type)
        {
            pathMaker(type);
            CopyMaker(origRaport,copyRaport);

            string[] charsToReplace = { "!", "@", "#", "%", "^", "&", "\\*", "\\(", "\\)", "-" };
            string replaceText = null;
            using (WordprocessingDocument mainDoc = WordprocessingDocument.Open(copyRaport, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(mainDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }
                int length = charsToReplace.Length > replaceStrings.Count ? replaceStrings.Count : charsToReplace.Length;
                for (int i = 0; i < length; i++)
                {
                    if (isDate(replaceStrings[i]))
                    {
                        replaceText = DateMaker(replaceStrings[i]);
                    }
                    else
                    {
                        replaceText = replaceStrings[i];
                    }

                    Regex regexText = new Regex(charsToReplace[i]);
                    docText = regexText.Replace(docText, replaceText);
                }


                using (StreamWriter sw = new StreamWriter(mainDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }


            }
        }


        public static string DateMaker(string date)
        {
            Regex reg = new Regex(@"\s*(\d*)\s*(\D*)\s*(\d*)");
            Match mtch = reg.Match(date);
            return "«" + mtch.Groups[1].ToString() + "» " + mtch.Groups[2].ToString() + " " + mtch.Groups[3].ToString() + "р.";
        }

        public static bool isDate(string date) => new Regex(@"\s*(\d+)\s*(\D+)\s*(\d+)").IsMatch(date);


    }
}
