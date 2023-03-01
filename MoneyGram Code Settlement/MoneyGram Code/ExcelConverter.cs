﻿using CsvHelper;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

//Settlement Code
namespace MoneyGram_Code
{
    public class ExcelConverter
    {
        public ExcelConverter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public void ConvertFile(string inputFile, string outputFile)
        {
            var list = new List<ExcelCols>();

            using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int countHeader = 0;

                    double per = 0.2;

                    string ccy = "ccy";

                    string agentName = "Branch Name";

                    //string excise = "Excise duty";

                    //string computedbaseamnt = "Computed Base Amount";

                    string currentDate = string.Empty;
                    string branchName = string.Empty;
                    string tranId = string.Empty;

                    while (reader.Read())
                    {
                        var row = new ExcelCols();

                        var val = reader.GetValue(0)?.ToString();

                        var value = reader.GetValue(1)?.ToString();

                       
                        if (!string.IsNullOrEmpty(value) && (value.Contains("Settlement Currency :") || value.Contains("Settlement Id") || value.Contains("Business Date:") || value.Contains("Orient") || value.Contains("KAMPALA") || value.Contains("UG")))
                            
                        {
                            
                            continue;
                        }

                        var value3 = reader.GetValue(9)?.ToString();

                        if (!string.IsNullOrEmpty(value3) &&(value3.Contains("Receive") || value3.Contains("Total") || value3.Contains("Grand Total")))
                        {
                            continue;
                        }

                        var value4 = reader.GetValue(10)?.ToString(); 
                        
                        if(!string.IsNullOrEmpty(value4) && value4.Contains("USD"))
                        {
                            continue;
                        }

                        var value5 =  reader.GetValue(13)?.ToString();

                        if(!string.IsNullOrEmpty(value5) && value5.Contains("Transaction Currency")) 
                        {
                            continue;
                        }

                        var value6 = reader.GetValue(30)?.ToString();   

                        if(!string.IsNullOrEmpty(value6) && (value6.Contains("Account Net Due")))
                        {
                            continue;
                        }

                        var value7 = reader.GetValue(19)?.ToString();
                        if (!string.IsNullOrEmpty(value7) && (value7.Contains("stl")))
                        {
                            continue;
                        }
                        var value8 = reader.GetValue(31)?.ToString();
                        if(!string.IsNullOrEmpty(value8) && value8.Contains("Grand Total Net Due"))
                        {
                            continue;
                        }
                        var value9 = reader.GetValue(32)?.ToString();
                        if(!string.IsNullOrEmpty(value9) && value9.Contains("Total Net Due"))
                        {
                            continue;
                        }

                        //tran date
                        row.Col0 = reader.GetValue(1)?.ToString().Replace("\n", "");
                        if (!string.IsNullOrEmpty(row.Col0))
                        {
                            currentDate = row.Col0;
                        }
                        //tran id
                        row.Col1 = reader.GetValue(4)?.ToString().Replace("\n", "");
                        if(!string.IsNullOrEmpty(row.Col1)) 
                        {
                            tranId = row.Col1;   
                        }
                        //ref #
                        row.Col2 = reader.GetValue(8)?.ToString().Replace("\n", "");
                        
                        //adding values below a colunm
                        //agent name
                        row.Col3 = reader.GetValue(6)?.ToString().Replace("\n", "");
                        if(!string.IsNullOrEmpty(row.Col3) && string.IsNullOrEmpty(val))
                        {

                            if(!string.IsNullOrEmpty(row.Col0) && row.Col0.Contains("Account Number"))
                            {
                                branchName = "";
                            }
                            //branchName = branchName + row.Col3;
                            branchName += row.Col3;
                        }
                        //if (!string.IsNullOrEmpty(row.Col3))
                        //{
                        //    branchName += row.Col3;
                        //    //branchName = branchName + row.Col3;
                        //}
                        //prod
                        row.Col4 = reader.GetValue(12)?.ToString().Replace("\n", "");
                        //type
                        row.Col5 = reader.GetValue(13)?.ToString().Replace("\n", "");
                        //origin cntry
                        row.Col6 = reader.GetValue(15)?.ToString().Replace("\n", "");
                        //rev cntry
                        row.Col7 = reader.GetValue(16)?.ToString().Replace("\n", "");
                        //fx rate
                        row.Col8 = reader.GetValue(18)?.ToString().Replace("\n", "");

                        if (!string.IsNullOrEmpty(row.Col8) && string.IsNullOrEmpty(row.Col0))
                        {
                            row.Col0 = currentDate;
                        }
                        //if (!string.IsNullOrEmpty(row.Col8) && string.IsNullOrEmpty(row.Col1))
                        //{
                        //    row.Col1 = tranId;
                        //}
                        if (!string.IsNullOrEmpty(row.Col8) && string.IsNullOrEmpty(row.Col3))
                        {
                            row.Col3 = branchName;
                        }
                        
                        //fx date
                        row.Col9 = reader.GetValue(23)?.ToString().Replace("\n", "");
                        //fx margin
                        row.Col10 = reader.GetValue(24)?.ToString().Replace("\n", "") + reader.GetValue(21)?.ToString();
                        //ccy
                        row.Col11 = reader.GetValue(25)?.ToString();
                        //base amount 
                        row.Col12 = reader.GetValue(26)?.ToString().Replace("\n", "");
                        //fee amount
                        row.Col13 = reader.GetValue(27)?.ToString().Replace("\n", "");
                        //fx rev share amount
                        row.Col14 = reader.GetValue(29)?.ToString().Replace("\n", "") + reader.GetValue(30)?.ToString() + reader.GetValue(31)?.ToString();
                        //commission amount
                        row.Col15 = reader.GetValue(34)?.ToString().Replace("\n", "") + reader.GetValue(35)?.ToString();

                        //total
                        //row.Col16 = reader.GetValue(36)?.ToString();


                        if (countHeader == 1)
                        {
                            row.Col3 = agentName;
                            row.Col11 = ccy;
                            //row.Col14 = excise;
                            //row.Col15 = computedbaseamnt;
                        }

                        countHeader++;

                        try
                        {
                            double baseamnt = Convert.ToDouble(reader.GetValue(25));
                            double feeamnt = Convert.ToDouble(reader.GetValue(26));

                            if (reader.GetValue(11) != null && reader.GetValue(12) != null && reader.GetValue(11).ToString() == "MT" && reader.GetValue(12).ToString() == "SEN")
                            { 
                                row.Col15 = (baseamnt + feeamnt + (feeamnt * per)).ToString();
                            }
                            else
                            {
                                row.Col15 = reader.GetValue(25)?.ToString();
                            }
                            
                        }
                        catch (Exception)
                        {

                        }

                        //if(!string.IsNullOrEmpty(row.Col0))
                           list.Add(row);
                    }
                }
            }

            var list2 = ProduceSecondList(inputFile).Skip(1).ToList();

            var list3 = CombineTheTwoLists(list, list2);

            WriteToFile(list3, outputFile);
        }

        private List<ExcelCols> ProduceSecondList(string inputFile)
        {
            var list3 = new List<ExcelCols>();

            using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int countHeader1 = 0;

                    double per = 0.2;

                    while (reader.Read())
                    {
                        string excise = "Excise duty";

                        var row = new ExcelCols();

                        var value1 = reader.GetValue(1)?.ToString();

                        if (!string.IsNullOrEmpty(value1) && (value1.Contains("Account Number : ") || value1.Contains("Settlement Currency : ")))
                        {
                            continue;
                        }

                        var value2 = reader.GetValue(12)?.ToString();

                        if (string.IsNullOrEmpty(value2) || value2.Contains("REC"))
                        {
                            continue;
                        }

                        //tran date
                        row.Col0 = reader.GetValue(1)?.ToString().Replace("\n", "");
                        //tran id
                        row.Col1 = reader.GetValue(4)?.ToString().Replace("\n", "");
                        //ref #
                        row.Col2 = reader.GetValue(8)?.ToString().Replace("\n", "");
                        //prod
                        row.Col3 = "excise duty";
                        //type
                        row.Col4 = reader.GetValue(12)?.ToString().Replace("\n", "");
                        //origin cntry
                        row.Col5 = reader.GetValue(14)?.ToString().Replace("\n", "");
                        //rev cntry
                        row.Col6 = reader.GetValue(15)?.ToString().Replace("\n", "");
                        //fx rate
                        row.Col7 = reader.GetValue(17)?.ToString().Replace("\n", "");
                        //fx date
                        row.Col8 = reader.GetValue(22)?.ToString().Replace("\n", "");
                        //fx margin
                        row.Col9 = reader.GetValue(23)?.ToString().Replace("\n", "");
                        //base amount 
                        row.Col10 = reader.GetValue(25)?.ToString().Replace("\n", "");
                        //fee amount
                        row.Col11 = reader.GetValue(26)?.ToString().Replace("\n", "");
                        //fx rev share amount
                        row.Col12 = reader.GetValue(28)?.ToString().Replace("\n", "") + reader.GetValue(29)?.ToString().Replace("\n", "") + reader.GetValue(30)?.ToString().Replace("\n", "");
                        //commission amount
                        row.Col13 = reader.GetValue(33)?.ToString().Replace("\n", "") + reader.GetValue(34)?.ToString().Replace("\n", "");

                        //excise duty calculation (0.2% of amount)
                        try
                        {
                            double cost = Convert.ToDouble(reader.GetValue(26));
                            row.Col14 = (cost * per).ToString("0.##");
                        }
                        catch (Exception)
                        {

                        }

                        row.Col15 = row.Col14;

                        if (countHeader1 == 0)
                        {
                            row.Col14 = excise;
                            countHeader1++;
                        }

                        list3.Add(row);
                    }
                }
            }

            return list3;
        }
        private List<ExcelCols> CombineTheTwoLists(List<ExcelCols> list, List<ExcelCols> list2)
        {
            var combinedList = new List<ExcelCols>();

            combinedList.AddRange(list);
            combinedList.AddRange(list2);

            return combinedList;
        }

        private void WriteToFile(List<ExcelCols> rows, string outputFile)
        {
            using (var writer = new StreamWriter(outputFile))
            {
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    foreach (var row in rows)
                    {
                        csv.WriteRecord(row);
                        csv.NextRecord();
                    }
                }
            }
        }

        public class ExcelCols
        {
            public string Col0 { get; set; }
            public string Col1 { get; set; }
            public string Col2 { get; set; }
            public string Col3 { get; set; }
            public string Col4 { get; set; }
            public string Col5 { get; set; }
            public string Col6 { get; set; }
            public string Col7 { get; set; }
            public string Col8 { get; set; }
            public string Col9 { get; set; }
            public string Col10 { get; set; }
            public string Col11 { get; set; }
            public string Col12 { get; set; }
            public string Col13 { get; set; }
            public string Col14 { get; set; }
            public string Col15 { get; set; }
            public string Col16 { get; set; }   

        }
    }
}
