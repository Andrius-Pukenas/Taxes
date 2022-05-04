using System;
using System.IO;
using System.Linq;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace Count
{
    public class ValueExtractor
    {
        private readonly string _customerNumber;

        public ValueExtractor(string customerNumber)
        {
            _customerNumber = customerNumber;
        }

        public ExtractedValues ExtractAll()
        {
            DateTime previousDate = DateTime.Now.AddMonths(-1);

            decimal vv = ExtractSingleDecimal($"{Configuration.DownloadsPath}\\saskaita.pdf", "VISO\n", "\n", 1);
            decimal vst = ExtractSingleDecimal($"{Configuration.DownloadsPath}\\{_customerNumber}_{previousDate.Year}_{previousDate.Month}.pdf", "Mokėti iš viso, EUR", "\n", 1);

            string[] pdfFiles = Directory.GetFiles($"{Configuration.DownloadsPath}\\", "*.pdf");
            string previousMonth = previousDate.Month.ToString().PadLeft(2, '0');
            string fileNamePattern = "-";
            string manoBustasFilePath = pdfFiles.Single(path => path.Contains(fileNamePattern));
            
            decimal manoBustas = ExtractSingleDecimal(manoBustasFilePath, "Mokėti iš viso, €", "\n", 1);
            //čia reikia pakeisti, jeigu kyla ginčų su nuomininku dėl to, kas turi mokėti už tam tikrus darbus, pvz., lifto remontą
            //manoBustas = manoBustas - decimal.Parse("21,93");
            
            decimal ignitisCommon;
            try
            {
                ignitisCommon = ExtractSingleDecimal(manoBustasFilePath, "VISO:"/*"Elektra bendram naudojimui"*/, "€", 2);
            }
            catch
            {
                ignitisCommon = ExtractSingleDecimal(manoBustasFilePath, "VISO:"/*"Elektra bendram naudojimui"*/, "€", 1);
            }

            return new ExtractedValues
            {
                Vv = vv,
                Vst = vst,
                ManoBustas = manoBustas,
                IgnitisCommon = ignitisCommon
            };
        }

        private decimal ExtractSingleDecimal(string filePath, string searchTextBeforeTarget, string searchTextAfterTarget, int pageNumber)
        {
            var extractedValue = ExtractSingleString(filePath, searchTextBeforeTarget, searchTextAfterTarget, pageNumber);

            return decimal.Parse(extractedValue, Configuration.NfiForInvoice);
        }

        private string ExtractSingleString(string filePath, string searchTextBeforeTarget, string searchTextAfterTarget, int pageNumber)
        {
            string strText;
            using (PdfReader reader = new PdfReader(filePath))
            {
                ITextExtractionStrategy its = new LocationTextExtractionStrategy();
                strText = PdfTextExtractor.GetTextFromPage(reader, pageNumber, its);
            }

            string customerNumberNormalized = _customerNumber.Substring(1, _customerNumber.Length - 1);
            if (strText.IndexOf(customerNumberNormalized) == -1 && pageNumber == 1)
            {
                throw new Exception("Wrong customer number!");
            }

            int indexOfSearchTextBeforeTarget = strText.IndexOf(searchTextBeforeTarget);
            if (indexOfSearchTextBeforeTarget == -1)
            {
                throw new Exception("Wrong template of invoice!");
            }

            int startPosition = indexOfSearchTextBeforeTarget + searchTextBeforeTarget.Length;

            int indexOfSearchTextAfter = strText.IndexOf(searchTextAfterTarget, startPosition);
            if (indexOfSearchTextAfter == -1)
            {
                throw new Exception("Wrong template of invoice!");
            }

            int length = indexOfSearchTextAfter - startPosition;

            string extractedValue = strText.Substring(startPosition, length);
            return extractedValue;
        }
    }
}