using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace llv.Samples.ValidateOpenXML.ConsoleApp
{
    class Program
    {
        static void Main()
        {
            try
            {
                Console.WriteLine("Enter the path of your file : ");
                String pathFile = Console.ReadLine();
                string extension = Path.GetExtension(pathFile);
                if (extension != null)
                {
                    extension = extension.ToLowerInvariant();
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
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.ToString());
            }
        }

        private static void ValidateWordDocument(String path)
        {
            using (OpenXmlPackage wordDocument = WordprocessingDocument.Open(path, false))
            {
                Validate(wordDocument);
            }
        }

        private static void ValidateSpreadsheetDocument(String path)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
            {
                Validate(spreadsheetDocument);
            }
        }

        private static void ValidatePresentationDocument(String path)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(path, false))
            {
                Validate(presentationDocument);
            }
        }

        private static void Validate(OpenXmlPackage wordDocument)
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
}
