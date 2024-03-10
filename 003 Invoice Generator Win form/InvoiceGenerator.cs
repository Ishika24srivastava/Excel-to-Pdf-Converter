using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GenerateInvoice
{
    internal class Entry
    {
        private static bool isAlreadyValidated = false;
        public static void MainEntry(string inputPath, string outputPath, Label progressText, ProgressBar progress, bool wantToProcess = false)
        {
            progressText.Text = "Processing Initiated...";
            InvoiceGeneration invoices = new InvoiceGeneration();
            if (File.Exists(inputPath))
            {
                if (isAlreadyValidated)
                {
                    if (wantToProcess)
                    {
                        progress.Value = 50;
                        string templatePath = outputPath + "\\templateFile.xlsx";
                        CopyEmbeddedResource(templatePath);
                        invoices.GeneratePdf(inputPath, outputPath, templatePath, progressText, progress);
                    }
                    else
                    {
                        MessageBox.Show("Validated Sucessfully. Click Ok to Continue..");
                    }
                }
                else
                {
                    if (invoices.ValidateExcel(inputPath, progressText))
                    {
                        progressText.Text = "Validated Data Successfully";
                       
                        if (wantToProcess)
                        {
                            progress.Value = 50;
                            string templatePath = outputPath + "\\templateFile.xlsx";
                            CopyEmbeddedResource(templatePath);
                            invoices.GeneratePdf(inputPath, outputPath, templatePath, progressText, progress);
                        }
                        else
                        {
                            MessageBox.Show("File Validated Successfully. Click on Convert Button.");
                        }
                        isAlreadyValidated = true;
                    }
                    else
                    {
                        isAlreadyValidated = false;
                        MessageBox.Show("Please Validate the Input File....");
                        return;
                    }
                }
            }
        }

        private static void CopyEmbeddedResource(string destinationPath)
        {
            
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream resourceStream = assembly.GetManifestResourceStream("InvoiceGenerator_WinForm.Template_Invoice.xlsx"))
            {
                if (File.Exists(destinationPath))
                {
                    try
                    {
                        File.Delete(destinationPath);
                    }
                    catch
                    {

                    }
                }
                
                using (FileStream fileStream = new FileStream(destinationPath, FileMode.Create))
                {
                    
                    resourceStream.CopyTo(fileStream);
                }
            }
        }

    }



    internal class InvoiceGeneration
    {
        public bool ValidateExcel(string inputPath, Label progressText)
        {
            progressText.Text = " Start Validating..";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks inputWorkbooks = excelApp.Workbooks;
            Excel.Workbook inputWorkbook = inputWorkbooks.Open(inputPath);
            Excel.Worksheet inputWorksheet = inputWorkbook.Worksheets[1];

            Excel.Range inputUsedRange = inputWorksheet.UsedRange;
            object[,] allCellValues = (object[,])inputUsedRange.Value;

            bool isValidData = true;
            StringBuilder errors = new StringBuilder();
            int totalRows = allCellValues.GetLength(0);

            try
            {
                string performaInvoiceNumber = "";
                string invoiceDate = "";
                string POReference = "";
                string clientName = "";
                string clientAddress = "";
                string clientCountry = "";
                int[] LineNumber = new int[5];
                string[] Description = new string[5];
                string officeOrVessel = "";
                string officeOrVesselName = "";
                string[] hsnOrSoc = new string[5];
                string currency = "";
                string[] unitCount = new string[5];
                string[] months = new string[5];
                string[] rate = new string[5];
                string[] amount = new string[5];
                string bankName = "";
                string bankAddress = "";
                string bankSwiftCode = "";
                string benefeciaryAccountNumber = "";
                string Remarks = "";
                string creditDays = "";

                bool isSameInvoice = true;
                int indexOfLineNumber = 0;


                inputUsedRange = inputUsedRange.Resize[totalRows, 24];
                inputUsedRange.Range["A2:X" + totalRows].Interior.Color = Excel.XlRgbColor.rgbWhite;
                inputUsedRange.Range["X1"].Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;

                inputUsedRange.Range["X1"].Value = "Errors";


                for (int row = 2; row <= totalRows; row++)
                {
                    for (int column = 1; column <= 23; column++)
                    {
                        string currentCellValue = allCellValues[row, column]?.ToString();
                        if (currentCellValue != null && currentCellValue != "")
                        {
                            if (column == 1)
                            {
                                if (performaInvoiceNumber == currentCellValue)
                                {
                                    if (indexOfLineNumber == 4)
                                    {
                                        isValidData = false;
                                        errors.AppendLine("Line Numbers Exceeds Limit. please Check it.");
                                    }
                                    else
                                    {
                                        indexOfLineNumber++;
                                    }
                                    isSameInvoice = true;
                                }
                                else
                                {
                                    indexOfLineNumber = 0;
                                    isSameInvoice = false;
                                    performaInvoiceNumber = currentCellValue;
                                }

                            }
                            else if (!isSameInvoice)
                            {
                                switch (column)
                                {
                                    case 2:
                                        invoiceDate = currentCellValue;
                                        break;
                                    case 3:
                                        POReference = currentCellValue;
                                        break;
                                    case 4:
                                        clientName = currentCellValue;
                                        break;
                                    case 5:
                                        clientAddress = currentCellValue;
                                        break;
                                    case 6:
                                        clientCountry = currentCellValue;
                                        break;
                                    case 9:
                                        officeOrVessel = currentCellValue;
                                        break;
                                    case 10:
                                        officeOrVesselName = currentCellValue;
                                        break;
                                    case 12:
                                        currency = currentCellValue;
                                        break;
                                    case 17:
                                        bankName = currentCellValue;
                                        break;
                                    case 18:
                                        bankAddress = currentCellValue;
                                        break;
                                    case 19:
                                        bankSwiftCode = currentCellValue;
                                        break;
                                    case 21:
                                        benefeciaryAccountNumber = currentCellValue;
                                        break;
                                    case 22:
                                        creditDays = currentCellValue;
                                        break;
                                    case 23:
                                        Remarks = currentCellValue;
                                        break;
                                    default:
                                        break;
                                }
                            }
                            else
                            {
                                switch (column)
                                {
                                    case 2:
                                        {
                                            if (invoiceDate != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Invoice Date is Not Same");
                                            }
                                            break;
                                        }
                                    case 3:
                                        {
                                            if (POReference != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Po Reference Is Not Same");
                                            }
                                            break;
                                        }
                                    case 4:
                                        {
                                            if (clientName != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Client Name is Not Same");
                                            }
                                            break;
                                        }
                                    case 5:
                                        {
                                            if (clientAddress != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("CLient Address Is Not Same");
                                            }
                                            break;
                                        }
                                    case 6:
                                        {
                                            if (clientCountry != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Client Country Is Not Same");
                                            }
                                            break;
                                        }
                                    case 10:
                                        {
                                            if (officeOrVesselName != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine(officeOrVessel + " Name Is Not Same");
                                            }
                                            break;
                                        }
                                    case 12:
                                        {
                                            if (currency != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Currency Is Not Same");
                                            }
                                            break;
                                        }
                                    case 17:
                                        {
                                            if (bankName != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Bank Name Is Not Same");
                                            }
                                            break;
                                        }
                                    case 18:
                                        {
                                            if (bankAddress != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Bank Address Is Not Same");
                                            }
                                            break;
                                        }
                                    case 19:
                                        {
                                            if (bankSwiftCode != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Bank SWIFT Code Is Not Same");
                                            }
                                            break;
                                        }
                                    case 21:
                                        {
                                            if (benefeciaryAccountNumber != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Benefeciary Account Number Is Not Same");
                                            }
                                            break;
                                        }
                                    case 22:
                                        {
                                            if (creditDays != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Credit Days are Not Same");
                                            }
                                            break;
                                        }
                                    case 23:
                                        {
                                            if (Remarks != currentCellValue)
                                            {
                                                isValidData = false;
                                                errors.AppendLine("Remarks are Not Same");
                                            }
                                            break;
                                        }
                                    default:
                                        break;
                                }
                            }
                            if (column == 7)
                            {
                                int num;
                                if (!int.TryParse(currentCellValue, out num))
                                {
                                    isValidData = false;
                                    errors.AppendLine("Invalid data for Line Number");
                                }
                                else if ((indexOfLineNumber == 0 && num != 1) || (indexOfLineNumber > 0 && LineNumber[indexOfLineNumber - 1] + 1 != num))
                                {
                                    isValidData = false;
                                    errors.AppendLine("Invalid data for Line Number");
                                }
                                LineNumber[indexOfLineNumber] = num;
                            }
                            else if (column == 8)
                            {
                                Description[indexOfLineNumber] = currentCellValue;
                            }
                            else if (column == 11)
                            {
                                hsnOrSoc[indexOfLineNumber] = currentCellValue;
                            }
                            else if (column == 13)
                            {
                                unitCount[indexOfLineNumber] = currentCellValue;
                            }
                            else if (column == 14)
                            {
                                if (currentCellValue != "0")
                                    months[indexOfLineNumber] = currentCellValue;
                                else
                                {
                                    isValidData = false;
                                    errors.AppendLine("Invalid data for Month");
                                }
                            }
                            else if (column == 15)
                            {
                                rate[indexOfLineNumber] = currentCellValue;
                            }
                            else if (column == 16)
                            {
                                amount[indexOfLineNumber] = currentCellValue;
                            }
                        }
                        else if ((currentCellValue == null || currentCellValue == "") && column != 23)
                        {
                            isValidData = false;
                            errors.AppendLine("Column " + (char)(column + 64) + " is Empty");
                        }
                    }
                    if (errors.ToString() != "")
                    {
                        inputUsedRange.Cells[row, 24].Value = errors.ToString();
                        inputUsedRange.Range["A" + row + ":X" + row].Interior.Color = Excel.XlRgbColor.rgbRed;
                        errors.Clear();
                    }
                }
                inputUsedRange.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                inputWorkbook.Save();
                inputWorkbook.Close();
                inputWorkbooks.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(inputUsedRange);
                Marshal.ReleaseComObject(inputWorksheet);
                Marshal.ReleaseComObject(inputWorkbook);
                Marshal.ReleaseComObject(inputWorkbooks);
                Marshal.ReleaseComObject(excelApp);
            }

            if (isValidData)
                return true;
            else
                return false;
        }

        public void GeneratePdf(string inputPath, string outputPath, string templatePath, Label progressText, ProgressBar progress)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks inputWorkbooks = excelApp.Workbooks;
            Excel.Workbook inputWorkbook = inputWorkbooks.Open(inputPath);
            Excel.Worksheet inputWorksheet = inputWorkbook.Worksheets[1];

            Excel.Range inputUsedRange = inputWorksheet.UsedRange;
            object[,] allCellValues = (object[,])inputUsedRange.Value;

            inputWorkbook.Close(false);
            inputWorkbooks.Close();
            Marshal.ReleaseComObject(inputUsedRange);
            Marshal.ReleaseComObject(inputWorksheet);
            Marshal.ReleaseComObject(inputWorkbook);
            Marshal.ReleaseComObject(inputWorkbooks);
            progress.Value = 60;

            string performaInvoiceNumber = "";
            string invoiceDate = "";
            string POReference = "";
            string clientName = "";
            string clientAddress = "";
            string clientCountry = "";
            int[] LineNumber = new int[5];
            string[] Description = new string[5];
            string officeOrVessel = "";
            string officeOrVesselName = "";
            string[] hsnOrSoc = new string[5];
            string currency = "";
            string[] unitCount = new string[5];
            string[] months = new string[5];
            string[] rate = new string[5];
            string[] amount = new string[5];
            string bankName = "";
            string bankAddress = "";
            string bankSwiftCode = "";
            string benefeciaryAccountNumber = "";
            string Remarks = "";
            string creditDays = "";
            int indexOfLineNumber = 0;

            try
            {
                bool isSameInvoice = true;
                int totalRows = allCellValues.GetLength(0);
                int numberOfInvoicesCreated = 0;
                int IncreaseValue = 900 / totalRows;

                for (int row = 2; row <= totalRows; row++)
                {
                    for (int column = 1; column <= 23; column++)
                    {
                        string currentCellValue = allCellValues[row, column]?.ToString();
                        if (currentCellValue != null && currentCellValue != "")
                        {
                            if (column == 1)
                            {
                                if (performaInvoiceNumber == currentCellValue)
                                {
                                    indexOfLineNumber++;
                                    isSameInvoice = true;
                                }
                                else
                                {
                                    if (numberOfInvoicesCreated > 0)
                                        GenerateInvoiceFromTemplate();

                                    numberOfInvoicesCreated++;
                                    progress.Value += 1;
                                    //Console.WriteLine("Creating Invoice :" + numberOfInvoicesCreated);
                                    progressText.Text = "Invoices Generation Started..";
                                    progressText.Text = "Generating Invoice  :" + numberOfInvoicesCreated;
                                    indexOfLineNumber = 0;
                                    isSameInvoice = false;
                                    performaInvoiceNumber = currentCellValue;
                                }
                            }
                            else if (!isSameInvoice)
                            {
                                switch (column)
                                {
                                    case 2:
                                        invoiceDate = currentCellValue;
                                        break;
                                    case 3:
                                        POReference = currentCellValue;
                                        break;
                                    case 4:
                                        clientName = currentCellValue;
                                        break;
                                    case 5:
                                        clientAddress = currentCellValue;
                                        break;
                                    case 6:
                                        clientCountry = currentCellValue;
                                        break;
                                    case 9:
                                        officeOrVessel = currentCellValue;
                                        break;
                                    case 10:
                                        officeOrVesselName = currentCellValue;
                                        break;
                                    case 12:
                                        currency = currentCellValue;
                                        break;
                                    case 17:
                                        bankName = currentCellValue;
                                        break;
                                    case 18:
                                        bankAddress = currentCellValue;
                                        break;
                                    case 19:
                                        bankSwiftCode = currentCellValue;
                                        break;
                                    case 21:
                                        benefeciaryAccountNumber = currentCellValue;
                                        break;
                                    case 22:
                                        creditDays = currentCellValue;
                                        break;
                                    case 23:
                                        Remarks = currentCellValue;
                                        break;
                                    default:
                                        break;
                                }
                            }

                            if (column == 7)
                            {
                                LineNumber[indexOfLineNumber] = int.Parse(currentCellValue);
                            }
                            else if (column == 8)
                            {
                                Description[indexOfLineNumber] = currentCellValue;
                            }
                            else if (column == 11)
                            {
                                hsnOrSoc[indexOfLineNumber] = currentCellValue;
                            }
                            else if (column == 13)
                            {
                                unitCount[indexOfLineNumber] = currentCellValue;
                            }
                            else if (column == 14)
                            {
                                months[indexOfLineNumber] = currentCellValue;
                            }
                            else if (column == 15)
                            {
                                rate[indexOfLineNumber] = currentCellValue;
                            }
                            else if (column == 16)
                            {
                                amount[indexOfLineNumber] = currentCellValue;
                            }
                        }
                    }
                    progress.Value += IncreaseValue;
                }
                GenerateInvoiceFromTemplate();
                File.Delete(templatePath);
                progress.Value = 1000;
                progressText.Text = "Successfully Genereated Invoices";
                MessageBox.Show("All Invoices Generated Successfully.");
            }
            catch (Exception ex)
            {
                if (!(ex is IOException))
                {
                    progressText.Text = "Exception Occured";
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }


            //Function to Generate Pdf
            void GenerateInvoiceFromTemplate()
            {
                Excel.Workbooks templateWorkbooks = excelApp.Workbooks;
                Excel.Workbook templateWorkbook = templateWorkbooks.Open(templatePath);
                Excel.Worksheet templateWorksheet = templateWorkbook.Worksheets[1];

                Excel.Range templateWorksheetRange = templateWorksheet.Range["A1"];
                try
                {
                    for (int rowNumber = 1; rowNumber <= 69; rowNumber++)
                    {
                        
                        switch (rowNumber)
                        {
                            case 3:
                                {
                                    templateWorksheetRange.Range["K3:L3"].Value = performaInvoiceNumber;
                                    break;
                                }
                            case 5:
                                {

                                    templateWorksheetRange.Range["K5:L5"].Value = invoiceDate;
                                    break;
                                }
                            case 7:
                                {
                                    templateWorksheetRange.Range["K7:L8"].Value = POReference;
                                    break;
                                }
                            case 16:
                                {
                                    templateWorksheetRange.Range["A16:G16"].Value = clientName;
                                    templateWorksheetRange.Range["K16:L16"].Value = creditDays + " Days";
                                    break;
                                }
                            case 17:
                                {
                                    if (clientAddress.Split('\n').Length <= 4)
                                    {
                                        templateWorksheetRange.Range["A17:F20"].MergeCells = true;
                                        templateWorksheetRange.Range["A17:F20"].Value = clientAddress;
                                    }
                                    else
                                    {
                                        templateWorksheetRange.Range["A15"].Value = clientName;

                                        templateWorksheetRange.Range["A16:F20"].MergeCells = true;
                                        templateWorksheetRange.Range["A16:F20"].Value = clientAddress;
                                    }
                                    break;
                                }
                            case 18:
                                {
                                    templateWorksheetRange.Range["J18"].Value = officeOrVessel + ":";
                                    templateWorksheetRange.Range["K18:L18"].Value = officeOrVesselName;
                                    break;
                                }
                            case 24:
                                {
                                    int startIndex = 24;
                                    for (int currentIndexOfLineNumber = 0; currentIndexOfLineNumber <= indexOfLineNumber; currentIndexOfLineNumber++, startIndex += 3)
                                    {
                                        templateWorksheetRange.Range["A" + startIndex + ":A" + (startIndex + 1)].MergeCells = true;
                                        templateWorksheetRange.Range["B" + startIndex + ":G" + (startIndex + 1)].MergeCells = true;
                                        templateWorksheetRange.Range["H" + startIndex + ":H" + (startIndex + 1)].MergeCells = true;
                                        templateWorksheetRange.Range["I" + startIndex + ":I" + (startIndex + 1)].MergeCells = true;
                                        templateWorksheetRange.Range["J" + startIndex + ":J" + (startIndex + 1)].MergeCells = true;
                                        templateWorksheetRange.Range["K" + startIndex + ":K" + (startIndex + 1)].MergeCells = true;
                                        templateWorksheetRange.Range["L" + startIndex + ":L" + (startIndex + 1)].MergeCells = true;

                                        templateWorksheetRange.Range["A" + startIndex + ":A" + (startIndex + 1)].Value = LineNumber[currentIndexOfLineNumber] + ".";
                                        templateWorksheetRange.Range["B" + startIndex + ":G" + (startIndex + 1)].Value = Description[currentIndexOfLineNumber];
                                        templateWorksheetRange.Range["H" + startIndex + ":H" + (startIndex + 1)].Value = hsnOrSoc[currentIndexOfLineNumber];
                                        templateWorksheetRange.Range["I" + startIndex + ":I" + (startIndex + 1)].Value = unitCount[currentIndexOfLineNumber];
                                        templateWorksheetRange.Range["J" + startIndex + ":J" + (startIndex + 1)].Value = months[currentIndexOfLineNumber];
                                        templateWorksheetRange.Range["K" + startIndex + ":K" + (startIndex + 1)].Value = rate[currentIndexOfLineNumber];
                                        templateWorksheetRange.Range["L" + startIndex + ":L" + (startIndex + 1)].Value = amount[currentIndexOfLineNumber];
                                    }
                                    break;
                                }
                            case 43:
                                {
                                    if (Remarks != "")
                                    {
                                        templateWorksheetRange.Range["A43:G45"].MergeCells = true;
                                        templateWorksheetRange.Range["A43:G45"].Value = Remarks;
                                        templateWorksheetRange.Range["A43:G45"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                        templateWorksheetRange.Range["A43:G45"].IndentLevel = 1;
                                    }
                                    break;
                                }
                            case 61:
                                {
                                    templateWorksheetRange.Range["A61"].Value = bankName;
                                    break;
                                }
                            case 62:
                                {
                                    string[] bankAddressLines = bankAddress.Split('\n');
                                    templateWorksheetRange.Range["A62"].Value = bankAddressLines[0];
                                    templateWorksheetRange.Range["A63"].Value = bankAddressLines[1];
                                    break;
                                }
                            case 64:
                                {
                                    templateWorksheetRange.Range["A64"].Value = "SWIFT: " + bankSwiftCode;
                                    templateWorksheetRange.Range["L64"].Value = "A/c No. " + benefeciaryAccountNumber;
                                    templateWorksheetRange.Range["L64"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                    break;
                                }
                        }
                    }

                    templateWorksheet.PageSetup.PrintArea = templateWorksheetRange.Range["A1:L69"].Address;

                    string invoicePath = outputPath + "\\" + performaInvoiceNumber + "_";
                    string[] clientnameInWords = clientName.Split(' ');
                    if (clientnameInWords.Length >= 2)
                    {
                        invoicePath += clientnameInWords[0] + " " + clientnameInWords[1] + ".pdf";
                    }
                    else
                    {
                        invoicePath += clientnameInWords[0] + ".pdf";
                    }
                    templateWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, invoicePath,
                                                        Excel.XlFixedFormatQuality.xlQualityStandard, true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    //Console.WriteLine(ex.ToString());
                }
                finally
                {
                    templateWorkbook.Close(false);
                    templateWorkbooks.Close();
                    Marshal.ReleaseComObject(templateWorksheetRange);
                    Marshal.ReleaseComObject(templateWorksheet);
                    Marshal.ReleaseComObject(templateWorkbook);
                    Marshal.ReleaseComObject(templateWorkbooks);
                }
            }
        }

    }
}
