using System;
using System.Globalization;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace Count
{
    public static class Configuration
    {
        static Configuration()
        {
            DownloadsPath = $"C:\\Users\\{Environment.UserName}\\Downloads";
            Nfi = new NumberFormatInfo { NumberDecimalSeparator = "." };
            NfiForInvoice = new NumberFormatInfo { NumberDecimalSeparator = "," };
            RoundToAmount = 10;

            string path = "appsettings.json";

            IConfigurationRoot configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetParent(AppContext.BaseDirectory).FullName)
                .AddJsonFile(path, false)
                .Build();

            CustomerNumber_0446181 = configuration["CustomerNumber_0446181"];
            Monthly_0446181 = decimal.Parse(configuration["Monthly_0446181"], Nfi);

            CustomerNumber_2464400 = configuration["CustomerNumber_2464400"];
            Monthly_2464400 = decimal.Parse(configuration["Monthly_2464400"], Nfi);

            CostOfkWh = decimal.Parse(configuration["CostOfkWh"], Nfi);
            ValueTelia = decimal.Parse(configuration["ValueTelia"], Nfi);
            Garbage = decimal.Parse(configuration["Garbage"], Nfi);
        }

        public static string DownloadsPath { get; set; }
        public static NumberFormatInfo Nfi { get; set; }
        public static NumberFormatInfo NfiForInvoice { get; set; }
        public static decimal RoundToAmount { get; set; }


        public static string CustomerNumber_0446181 { get; set; }
        public static decimal Monthly_0446181 { get; set; }

        public static string CustomerNumber_2464400 { get; set; }
        public static decimal Monthly_2464400 { get; set; }
        
        public static decimal CostOfkWh { get; set; }
        public static decimal ValueTelia { get; set; }
        public static decimal Garbage { get; set; }
    }
}