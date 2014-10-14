using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace llv.Samples.ValidateOpenXML.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Enter the path of your file : ");
                String pathFile = Console.ReadLine();
                String extension = Path.GetExtension(pathFile);
                switch (extension)
                {
                    case ".xlsx":
                        ValidateSpreadsheetDocument(pathFile);
                        break;
                    case ".docx":
                        ValidateWordDocument(pathFile);
                        break;
                    case ".pptx":
                        ValidatePresentationDocument(pathFile);
                        break;
                    default:
                        Console.WriteLine("The extension of the file is incorrect");
                        break;
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
            finally
            {
                Console.ReadKey();
            }

        }

        private static void ValidateWordDocument(String path)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, false))
            {
                OpenXmlValidator openXmlValidator = new OpenXmlValidator();
                int count = 0;
                foreach (ValidationErrorInfo error in openXmlValidator.Validate(wordDocument))
                {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }
            }
        }

        private static void ValidateSpreadsheetDocument(String path)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
            {
                OpenXmlValidator openXmlValidator = new OpenXmlValidator();
                int count = 0;
                foreach (ValidationErrorInfo error in openXmlValidator.Validate(spreadsheetDocument))
                {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }
            }
        }

        private static void ValidatePresentationDocument(String path)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(path, false))
            {
                OpenXmlValidator openXmlValidator = new OpenXmlValidator();
                int count = 0;
                foreach (ValidationErrorInfo error in openXmlValidator.Validate(presentationDocument))
                {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }
            }
        }
    }
}
