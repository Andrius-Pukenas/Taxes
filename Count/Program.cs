using System;
using System.IO;

namespace Count
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Choose your destiny:");
            int chosenConfiguration = int.Parse(Console.ReadLine());

            string customerNumber;
            decimal monthly;
            if (chosenConfiguration == 1)
            {
                customerNumber = Configuration.CustomerNumber_0446181;
                monthly = Configuration.Monthly_0446181;
            }
            else
            {
                customerNumber = Configuration.CustomerNumber_2464400;
                monthly = Configuration.Monthly_2464400;
            }

            var gatheredInformation = new GatheredInformation(customerNumber);
            ExtractedValues extractedValues = new ValueExtractor(customerNumber).ExtractAll();

            decimal ignitisTotal = CalculateIgnitisTotal(gatheredInformation, extractedValues);

            decimal sumWithoutTelia = extractedValues.Vv + extractedValues.Vst + extractedValues.ManoBustas + ignitisTotal;

            if (sumWithoutTelia != gatheredInformation.SumToValidate)
            {
                throw new Exception("Nesutampa sumos!");
            }

            byte[] file = new ExcelBuilder().BuildExcel(
                gatheredInformation,
                extractedValues, 
                ignitisTotal,
                sumWithoutTelia,
                monthly);

            DateTime previousDate = DateTime.Now.AddMonths(-1);

            File.WriteAllBytes($"{Configuration.DownloadsPath}\\{customerNumber}_{previousDate.Year}_{previousDate.Month}.xlsx", file);
        }

        private static decimal CalculateIgnitisTotal(GatheredInformation gatheredInformation, ExtractedValues extractedValues)
        {
            decimal ignitiskWh = (gatheredInformation.IgnitisTo - gatheredInformation.IgnitisFrom) * Configuration.CostOfkWh;
            decimal ignitiskWhRounded = Math.Round(ignitiskWh, 2);
            decimal ignitisTotal = ignitiskWhRounded + extractedValues.IgnitisCommon;
            return ignitisTotal;
        }
    }
}