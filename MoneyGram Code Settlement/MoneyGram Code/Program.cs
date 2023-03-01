using System;

namespace MoneyGram_Code
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFile = @"C:\Users\HP\Downloads\SettlementDetailFX_100513823_02-10-2023.xlsx";
            string outputFile = @"C:\Users\HP\Downloads\SettlementDetailFX_100513823_02-10-2023.csv";
            new ExcelConverter().ConvertFile(inputFile, outputFile);
            Console.ReadLine();
        }
    }
}
