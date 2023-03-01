﻿using System;

namespace MoneyGram_Code
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFile = @"C:\Users\HP\Downloads\DailyActivityDetailFX_100513823_02-07-2023.xlsx";
            string outputFile = @"C:\Users\HP\Downloads\Converted DailyActivityDetailFX_100513823_02-07-2023 test.csv";
            new ExcelConverter().ConvertFile(inputFile, outputFile);
            Console.ReadLine();
        }
    }
}
