using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace PdfToXls
{
    class Program
    {
        static void Main()
        {
            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            Console.Title = "PdfToXls";

            //Console.WriteLine("----------");
            //Console.WriteLine("Esse projeto utiliza PdfPig");
            //Console.WriteLine("https://github.com/UglyToad/PdfPig");
            //Console.WriteLine("----------");

            //Console.WriteLine("Iniciando");

            do
            {
                Console.Write(LineAsterisk());
                Console.WriteLine("Arquivos PDFs que estão no mesmo diretório da aplicação");
                Console.Write(LineAsterisk());

                var directoryInfo = new DirectoryInfo(baseDirectory);
                var files = directoryInfo.GetFiles();
                var listFiles = new Dictionary<int, string>();


                foreach (FileInfo file in files)
                    if (file.Extension == ".pdf")
                        listFiles.Add(listFiles.Count + 1, file.Name);

                if (listFiles.Count > 0)
                {
                    //Print the files to choose by number
                    foreach (var item in listFiles)
                        Console.WriteLine($"{item.Key} - {item.Value}");
                }
                Console.Write(LineAsterisk());
                Console.Write("Entre com o numero do arquivo para inciciar o processamento (ou 'q' para sair):");

                var val = Console.ReadLine();
                if (!int.TryParse(val, out var opt))
                {
                    if (string.Equals(val, "q", StringComparison.OrdinalIgnoreCase))
                        return;

                    Console.WriteLine($"Nenhuma opção encontrada para esse valor: {val}.");
                    continue;
                }

                if (listFiles.Count > 0)
                {
                    //Print the files to choose by number
                    foreach (var item in listFiles)
                        if (item.Key == opt)
                            OpenDocumentAndExtracLines.Run(Path.Combine(baseDirectory, item.Value));
                }

            } while (true);

        }

        public static string LineAsterisk()
        {
            string result = string.Empty;
            for (int i = 0; i < Console.WindowWidth; i++)
                result += "*";

            return result;
        }
        public static class OpenDocumentAndExtracLines
        {
            public static string fileDirectory { get; set; }

            public static void Run(string filePath)
            {
                if (!File.Exists(filePath))
                    return;

                fileDirectory = Path.GetDirectoryName(filePath);
                var informationLines = new List<InformationOfLine>();

                using (var document = PdfDocument.Open(filePath))
                {
                    foreach (var page in document.GetPages())
                    {
                        var startInformation = false;
                        var text = ContentOrderTextExtractor.GetText(page, true);
                        var lines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                        foreach (var line in lines)
                        {
                            if (!string.IsNullOrWhiteSpace(line))
                            {
                                var informationLine = new InformationOfLine();

                                if (line.Contains("DATA LANÇAMENTO VALOR TOTAL"))
                                    startInformation = true;
                                else if ((startInformation) && IsDate(line.Substring(0, 10)))
                                    informationLine = SplitInformation(line);

                                if (informationLine.HasValues)
                                    Console.WriteLine($"{informationLine.Date} - {informationLine.Description} - {informationLine.Value} - {informationLine.TotalValue}");

                                if (informationLine.HasValues)
                                    informationLines.Add(informationLine);
                            }
                        }
                    }
                }

                string fileCreated = string.Empty;
                if (informationLines.Count > 0)
                    fileCreated = CreateXlsxFile(informationLines);

                if (!string.IsNullOrWhiteSpace(fileCreated))
                    Console.WriteLine($"File created: {fileCreated}");
            }

            public static bool IsDate(string date)
            {
                DateTime dateTime;
                return DateTime.TryParse(date, out dateTime);
            }

            public static InformationOfLine SplitInformation(string line)
            {
                var result = new InformationOfLine();

                //The first 10 characters is the date
                var date = line.Substring(0, 10);

                //The next characters up to R$ is the description
                var description = line.Substring(11, line.IndexOf("R$") - 11);

                //The first value starts at the R$ position and goes to the next R$ position
                var start = line.IndexOf("R$");
                var end = line.IndexOf("R$", start + 1);
                var moneyValue = line.Substring(start, end - start);

                //The second value starts at the last R$ position and goes to final of the line
                var totalValue = line.Substring(end, line.Length - end);

                result.Date = date;
                result.Description = description;
                result.Value = moneyValue;
                result.TotalValue = totalValue;

                return result;
            }

            public static string CreateXlsxFile(List<InformationOfLine> informationLines)
            {
                string fileXlsxPath = Path.Combine(fileDirectory, Guid.NewGuid().ToString() + ".xlsx");

                var fs = File.Create(fileXlsxPath);

                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Extrato");

                var styleHeader = workbook.CreateCellStyle();
                styleHeader.FillForegroundColor = HSSFColor.Grey25Percent.Index;
                styleHeader.FillPattern = FillPattern.SolidForeground;

                var stylePayment = workbook.CreateCellStyle();
                stylePayment.FillForegroundColor = HSSFColor.LightBlue.Index;
                stylePayment.FillPattern = FillPattern.SolidForeground;

                var doubleCellFormate = workbook.CreateDataFormat();
                var doubleDataFormate = doubleCellFormate.GetFormat("#,##0.00");
                var doubleCellStyle = workbook.CreateCellStyle();
                doubleCellStyle.DataFormat = doubleDataFormate;

                IRow row;
                ICell cell;

                int nLine = 0;
                int nCol = 0;

                sheet.SetColumnWidth(0, 20 * 256);
                sheet.SetColumnWidth(1, 50 * 256);
                sheet.SetColumnWidth(2, 20 * 256);
                sheet.SetColumnWidth(3, 20 * 256);

                row = sheet.CreateRow(nLine);
                var headerInformation = new List<string> { "Date", "Description", "Value", "TotalValue" };
                foreach (var itemHeader in headerInformation)
                {
                    cell = row.CreateCell(nCol);
                    cell.SetCellValue(itemHeader);
                    cell.CellStyle = styleHeader;

                    nCol++;
                }
                nLine++;


                foreach (var itemLine in informationLines)
                {
                    var colInformation = new List<string> { itemLine.Date, itemLine.Description, itemLine.Value, itemLine.TotalValue };

                    row = sheet.CreateRow(nLine);

                    nCol = 0;
                    foreach (var itemCol in colInformation)
                    {
                        cell = row.CreateCell(nCol);
                        cell.SetCellValue(itemCol);

                        if (itemLine.Description.Contains("DEPOSITO"))
                            cell.CellStyle = stylePayment;

                        nCol++;
                    }
                    nLine++;
                }

                var unpaymentMonths = PaymentAnalisis(informationLines);
                if (unpaymentMonths.Count > 0)
                {
                    nLine = 0;
                    ISheet sheetUnpayment = workbook.CreateSheet("Nao Pagos");

                    row = sheetUnpayment.CreateRow(nLine);
                    cell = row.CreateCell(0);
                    cell.SetCellValue("Meses que não houve pagamento");
                    cell.CellStyle = styleHeader;

                    sheetUnpayment.SetColumnWidth(0, 60 * 256);

                    foreach (var item in unpaymentMonths)
                    {
                        nLine++;
                        row = sheetUnpayment.CreateRow(nLine);
                        cell = row.CreateCell(0);
                        cell.SetCellValue(item);
                    }
                }

                workbook.Write(fs);

                return fileXlsxPath;

            }

            public static List<string> PaymentAnalisis(List<InformationOfLine> informationLines)
            {
                string yearFirstPayment = string.Empty;
                string monthFirstPayment = string.Empty;

                string yearLastPayment = string.Empty;
                string monthLastPayment = string.Empty;

                foreach (var itemLine in informationLines)
                {
                    if (itemLine.Description.Contains("DEPOSITO"))
                    {
                        #region setup date
                        //Year is the lasts 4 characters of the description
                        var year = itemLine.Description.Substring(itemLine.Description.Length - 5, 4);
                        //Find in description the month
                        var month = string.Empty;
                        if (itemLine.Description.Contains("JANEIRO"))
                            month = "01";
                        else if (itemLine.Description.Contains("FEVEREIRO"))
                            month = "02";
                        else if (itemLine.Description.Contains("MARCO"))
                            month = "03";
                        else if (itemLine.Description.Contains("ABRIL"))
                            month = "04";
                        else if (itemLine.Description.Contains("MAIO"))
                            month = "05";
                        else if (itemLine.Description.Contains("JUNHO"))
                            month = "06";
                        else if (itemLine.Description.Contains("JULHO"))
                            month = "07";
                        else if (itemLine.Description.Contains("AGOSTO"))
                            month = "08";
                        else if (itemLine.Description.Contains("SETEMBRO"))
                            month = "09";
                        else if (itemLine.Description.Contains("OUTUBRO"))
                            month = "10";
                        else if (itemLine.Description.Contains("NOVEMBRO"))
                            month = "11";
                        else if (itemLine.Description.Contains("DEZEMBRO"))
                            month = "12";
                        #endregion

                        if (string.IsNullOrWhiteSpace(yearFirstPayment) && string.IsNullOrWhiteSpace(monthFirstPayment))
                        {
                            yearFirstPayment = year;
                            monthFirstPayment = month;
                        }
                        else
                        {
                            yearLastPayment = year;
                            monthLastPayment = month;
                        }
                    }
                }

                //Create a list of months between the first and last payment
                var months = new Dictionary<string, bool>();
                var dateFirst = new DateTime(int.Parse(yearFirstPayment), int.Parse(monthFirstPayment), 1);
                var dateLast = new DateTime(int.Parse(yearLastPayment), int.Parse(monthLastPayment), 1);

                while (dateFirst <= dateLast)
                {
                    bool hasPayment = false;
                    foreach (var itemLine in informationLines)
                    {
                        if (itemLine.Description.Contains("DEPOSITO"))
                        {
                            #region setup date
                            //Year is the lasts 4 characters of the description
                            var year = itemLine.Description.Substring(itemLine.Description.Length - 5, 4);
                            //Find in description the month
                            var month = string.Empty;
                            if (itemLine.Description.Contains("JANEIRO"))
                                month = "01";
                            else if (itemLine.Description.Contains("FEVEREIRO"))
                                month = "02";
                            else if (itemLine.Description.Contains("MARCO"))
                                month = "03";
                            else if (itemLine.Description.Contains("ABRIL"))
                                month = "04";
                            else if (itemLine.Description.Contains("MAIO"))
                                month = "05";
                            else if (itemLine.Description.Contains("JUNHO"))
                                month = "06";
                            else if (itemLine.Description.Contains("JULHO"))
                                month = "07";
                            else if (itemLine.Description.Contains("AGOSTO"))
                                month = "08";
                            else if (itemLine.Description.Contains("SETEMBRO"))
                                month = "09";
                            else if (itemLine.Description.Contains("OUTUBRO"))
                                month = "10";
                            else if (itemLine.Description.Contains("NOVEMBRO"))
                                month = "11";
                            else if (itemLine.Description.Contains("DEZEMBRO"))
                                month = "12";
                            #endregion

                            if ((year == dateFirst.ToString("yyyy")) && (month == dateFirst.ToString("MM")))
                            {
                                hasPayment = true;
                                break;
                            }
                        }
                    }

                    months.Add(dateFirst.ToString("MM/yyyy"), hasPayment);

                    dateFirst = dateFirst.AddMonths(1);
                }

                //Print the months and if not has payment
                var result = new List<string>();
                foreach (var item in months)
                {
                    if (!item.Value)
                        result.Add(item.Key);
                }

                return result;
            }

        }
    }
}