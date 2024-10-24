using Microsoft.Extensions.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace EspresaUpdate2WDEIB
{
    class QuotePair
    {
        public int start;
        public int end;
    }

    class Program
    {
        static void Main(string[] args)
        {
            string inputFile = string.Empty;
            string[] fields;
            string cleansedLine = string.Empty;

            if (args.Length == 2)
            {
                if (args[0].ToLower().Trim() == "-inputfile")
                {
                    inputFile = args[1].Trim();
                }
                else
                {
                    Console.WriteLine("Usage: EspresaUpdate2WDEIB -Inputfile <path>");
                }
            }
            else
            {
                Console.WriteLine("Usage: EspresaUpdate2WDEIB -Inputfile <path>");
            }

            //inputFile = "C:\\temp\\VAI-Espresa-Payroll-20240926.csv";  // for testing on local device

            try
            {
                if (File.Exists(inputFile))
                {
                    StreamReader sr = new StreamReader(inputFile);
                    string line=string.Empty;

                    using (sr)
                    {
                        string? EIB = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("AppSettings")["eibTemplate"];
                        string? eibOut = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("AppSettings")["eibOutput"];
                        string? taxcode = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("AppSettings")["TaxCode"];
                        string eibFile = string.Empty;
                        string outputFile = string.Empty;
                        int ix = 0;

                        FileStream fsT = new FileStream(EIB, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        XSSFWorkbook workbook1 = new XSSFWorkbook(fsT);

                        ISheet worksheet1 = workbook1.GetSheet("Import Payroll Input");
                        IRow w1Row;

                        ISheet worksheet2 = workbook1.GetSheet("Payroll Input Data");
                        IRow w2Row;

                        DateTime payrolldate = PayrollDate();

                        ix = 5;
                        var style = workbook1.CreateCellStyle();

                        w1Row = worksheet1.CreateRow(5);
                        w1Row.CreateCell(0);
                        w1Row.CreateCell(1);
                        w1Row.CreateCell(2);
                        w1Row.Cells[1].SetCellValue("1");
                        w1Row.Cells[2].SetCellValue("Espresa_Input_" + payrolldate.ToString("yyyyMMdd"));

                        while (sr.Peek() >= 0)
                        {
                            line = sr.ReadLine();

                            // remove comma's between the quotes to prevent errors parsing fields
                            cleansedLine = CleanseLine(line);
                            fields = cleansedLine.Split(',');

                            if (fields[0].Substring(0, 1) == "W") // Found Employee ID
                            {
                                w2Row = worksheet2.CreateRow(ix);
                                CreateRowCells(w2Row, style);
                                w2Row.Cells[1].SetCellValue("1");
                                w2Row.Cells[2].SetCellValue((ix - 4).ToString());
                                w2Row.Cells[8].SetCellValue(payrolldate.ToString("yyyy-MM-dd"));
                                w2Row.Cells[9].SetCellValue(payrolldate.ToString("yyyy-MM-dd"));
                                w2Row.Cells[11].SetCellValue(fields[0]);
                                w2Row.Cells[13].SetCellValue(taxcode);
                                w2Row.Cells[15].SetCellValue(fields[3]);
                                ix++;
                            }
                        }

                        var outPath = Path.GetDirectoryName(inputFile) + "\\";
                        outputFile = outPath + "INT106-Espresa-Payroll.xlsx";
                        using (FileStream stream = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                        {
                            workbook1.Write(stream);
                        }

                        Console.WriteLine("Text written to Excel successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private static void CreateRowCells(IRow row, ICellStyle style)
        {
            for (int j = 0; j < 16; j++)
            {
                if (j == 15)
                {
                    row.CreateCell(j, CellType.Numeric);
                    row.Cells[j].CellStyle = style;
                }
                else
                    row.CreateCell(j);
            }
        }
        private static DateTime PayrollDate()
        {
            int dow = (int)DateTime.Now.DayOfWeek;
            DateTime pdate = DateTime.Now;

            switch (dow)
            {
                case 0:
                    pdate = DateTime.Now.AddDays(-14);
                    break;
                case 1:
                    pdate = DateTime.Now.AddDays(-15);
                    break;
                case 2:
                    pdate = DateTime.Now.AddDays(-16);
                    break;
                case 3:
                    pdate = DateTime.Now.AddDays(-17);
                    break;
                case 4:
                    pdate = DateTime.Now.AddDays(-18);
                    break;
                case 5:
                    pdate = DateTime.Now.AddDays(-19);
                    break;
                case 6:
                    pdate = DateTime.Now.AddDays(-20);
                    break;

            }
            return pdate;
        }
        private static string CleanseLine(string line)
        {
            char[] lineChars = line.ToCharArray();
            List<QuotePair> quotePairs = new List<QuotePair>();
            int count = 0;
            int startPos = 0;
            int endPos = 0;
            int totalQuotes = 0;
            string result = line;
            bool foundFirstQuote = false;

            for (int i = 0; i < lineChars.Length; i++)
            {
                foundFirstQuote = false;
                if (lineChars[i] == '"')
                {
                    // The first quote in a balanced pair found
                    foundFirstQuote = true;

                    count++;
                    totalQuotes++;
                }

                if (foundFirstQuote && count == 1)
                    startPos = i;

                if (count == 2)
                {
                    // The closing quote has been found.
                    endPos = i;
                    quotePairs.Add(new QuotePair { start = startPos, end = endPos });
                    count = 0;
                }
            }

            // Make sure number of quotes equal balanced pairs
            if (totalQuotes % 2 == 0)
            {
                // Get rid of commas between quotes.
                foreach (QuotePair pair in quotePairs)
                {
                    for (int i = pair.start; i < pair.end; i++)
                    {
                        if (lineChars[i] == ',')
                            lineChars[i] = ' ';
                    }
                }
                result = new string(lineChars);
            }

            return result;
        }
    }
}
