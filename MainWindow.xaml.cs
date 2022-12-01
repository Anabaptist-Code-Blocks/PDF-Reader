using ClosedXML.Excel;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;

namespace PDF_Reader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            var args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                //Get the file path argument if it exits. This handles instances where a file is opened with the app.
                var filePath = args[1];
                ExtractTextFromPDF(filePath);
                var open = args.Length >= 2 && args[2] == "open";
                ExportToExcel(Items, open);
                //If the app is being run from commandline no need to show window.
                if (args.Length > 1) App.Current.Shutdown();
            }
        }

        #region Fields
        private string outputPath = "";
        private List<Item> Items = new List<Item>();
        //regex to remove everything but numeric digits and decimal
        private Regex NumericRegex = new Regex("[^0-9.]");
        #endregion

        #region Methods

        private void ExtractTextFromPDF(string filePath)
        {
            //here we extract the text from the pdf file using iText7 https://itextpdf.com/ AGPLv3 license
            if (File.Exists(filePath))
            {
                //get the path for saving the Excel file. We'll put it in the same folder and simply change the extension to xlsx
                var baseName = Path.GetFileNameWithoutExtension(filePath);
                var path = Path.GetFullPath(filePath);
                var basePath = Path.GetDirectoryName(path);
                outputPath = Path.Combine(basePath, baseName + ".xlsx");
            }
            //time the parsing so we can tell how long it took
            var sw = new Stopwatch();
            sw.Start();
            try
            {
                Items = new List<Item>();
                PdfReader pdfReader = new PdfReader(filePath);
                PdfDocument pdfDoc = new PdfDocument(pdfReader);
                for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                {
                    ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();
                    string pageContent = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                    Items.AddRange(ParsePage(pageContent));
                }
                pdfDoc.Close();
                pdfReader.Close();
                //show the items on the datagrid
                itemDG.ItemsSource = Items;
                sw.Stop();
                status.Text = $"{sw.ElapsedMilliseconds} ms to parse pdf!";
            }
            catch (Exception ex)
            {
                //something went wrong - update the status label to display the error
                sw.Stop();
                status.Text = $"Error: {ex.Message}";
            }
        }

        private List<Item> ParsePage(string content)
        {
            var items = new List<Item>();
            var lines = content.Split('\n');
            var start = false;
            Item lastItem = null;
            foreach (var line in lines)
            {
                //ckeck for a text that we know will appear before the start of the items list
                if (!start && line.Contains("ORDER QTY"))
                {
                    start = true;
                    continue;
                }
                //check for some text that we know will appear at the end of the items list
                if (line.Contains("** Continued") || line.Contains("The pricing on this bid"))
                {
                    return items;
                }
                if (start)
                {
                    //we are past the point where the items list starts
                    //here we process the line which is a single string

                    //first we split the string on each space since we know there is a space between each field
                    var fields = line.Split(' ');
                    if (fields.Length < 5)
                    {
                        //if there are less than 5 fields we can be pretty sure that this line is only and addtional description line 
                        if (lastItem != null) lastItem.Description += $"\r\n{line}";
                        continue;
                    }
                    if (fields.Length >= 5)
                    {
                        var fieldQty = fields[0];
                        //we know that the price field is the second to last in the list.
                        var fieldPrice = fields[fields.Length - 2];
                        //we know that the total field will be the last one
                        var fieldTotal = fields[fields.Length - 1];
                        //we calculate the length of the description portion by deducting the length of the know fields
                        int descLen = line.Length - (fieldQty.Length + fieldPrice.Length + fieldTotal.Length + 3);
                        //we need to use substring to extract the description field because it probably contains multiple spaces
                        var fieldDescription = line.Substring(fieldQty.Length + 1, Math.Max(0, descLen));
                        //use regex to parse out non numeric characters and convert to int
                        int.TryParse(NumericRegex.Replace(fieldQty, ""), out int qty);
                        //the remaining portion is the unit of measure
                        var qtyUM = fieldQty.Substring(qty.ToString().Length);
                        //we know that the first part of the description is the part numer
                        var partNum = fieldDescription.Split(' ')[0];
                        //and the remaining portion is the description
                        var desc = fieldDescription.Substring(partNum.Length + 1);
                        //extract the price portion minus unit of measure and convert to double
                        var priceStr = NumericRegex.Replace(fieldPrice, "");
                        double.TryParse(priceStr, out double price);
                        //the remaining portion is the price unit of measure
                        var priceUM = fieldPrice.Substring(priceStr.Length + 1);
                        //the last field is the total and should already contain only a number to converted to double
                        double.TryParse(fieldTotal, out double total);
                        if (total == 0)
                        {
                            //if this really is a description line with more than five words the last word won't be convertable to a number
                            //in this case we'll add this as an additional description line
                            if (lastItem != null) lastItem.Description += $"\r\n{line}";
                            continue;
                        }
                        //create a new item and add it to the list
                        lastItem = new Item()
                        {
                            Qty = qty,
                            QtyUM = qtyUM,
                            PartNum = partNum,
                            Description = desc,
                            Price = price,
                            PriceUM = priceUM,
                            Total = total
                        };
                        items.Add(lastItem);
                    }

                }
            }
            //we should never get here but just in case...
            return items;
        }
        private void ExportToExcel(List<Item> items, bool openWorkBook)
        {
            //here we generate the excel file using ClosedXML https://github.com/ClosedXML/ClosedXML MIT Licensed
            try
            {

                IXLWorkbook workbook = new XLWorkbook();
                IXLWorksheet worksheet = workbook.Worksheets.Add("Items");
                worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
                //create the header
                worksheet.Cell("A1").Value = "Qty";
                worksheet.Cell("B1").Value = "UM";
                worksheet.Cell("C1").Value = "Part Number";
                worksheet.Cell("D1").Value = "Description";
                worksheet.Cell("E1").Value = "Price";
                worksheet.Cell("F1").Value = "UM";
                worksheet.Cell("G1").Value = "Ext Price";
                //style the header row
                worksheet.Range("A1:G1").Style.Font.Bold = true;
                worksheet.Range("A1:G1").Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                worksheet.Range("A1:A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Range("C1:C1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                worksheet.Range("E1:E1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Range("G1:G1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column(1).Width = 10;
                worksheet.Column(2).Width = 5;
                worksheet.Column(3).Width = 15;
                worksheet.Column(4).Width = 65;
                worksheet.Column(5).Width = 10;
                worksheet.Column(6).Width = 5;
                worksheet.Column(7).Width = 10;
                worksheet.SheetView.FreezeRows(1);
                var currentRow = 1;
                foreach (var item in items)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = item.Qty;
                    worksheet.Cell(currentRow, 2).Value = item.QtyUM;
                    worksheet.Cell(currentRow, 3).Value = item.PartNum;
                    worksheet.Cell(currentRow, 4).Value = item.Description;
                    worksheet.Cell(currentRow, 5).Value = item.Price;
                    worksheet.Cell(currentRow, 6).Value = item.PriceUM;
                    worksheet.Cell(currentRow, 7).Value = item.Total;

                }
                worksheet.Range(2, 1, currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Range(2, 3, currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                worksheet.Range(2, 5, currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Range(2, 7, currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Range(2, 1, currentRow, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range(2, 1, currentRow, 7).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range(2, 1, currentRow, 7).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

                workbook.SaveAs(outputPath);
                if (openWorkBook) Process.Start(outputPath);
            }
            catch (Exception ex)
            {
                status.Text = $"Error: {ex.Message}";
            }


        }

        #endregion

        #region Event Handlers

        private void Window_Drop(object sender, DragEventArgs e)
        {
            //an object was dropped onto the window.
            var data = e.Data;
            //check to see if the dropped object is a file
            if (data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] filesNames = data.GetData(DataFormats.FileDrop) as string[];
                if (filesNames != null && filesNames.Length > 0)
                {
                    if (filesNames.Length > 1)
                    {
                        status.Text = "Whoa! Only one file can be processed at a time.";
                    }
                    else
                    {
                        ExtractTextFromPDF(filesNames[0]);
                    }
                }
            }
            //dropping files from Outlook takes another 40 lines of code to handle
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            //allow dragging file onto window
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private void SaveToExcel_Click(object sender, RoutedEventArgs e)
        {
            //save to excel
            ExportToExcel(Items, true);

        }

        #endregion

    }

    public class Item
    {
        public int Qty { get; set; }
        public string QtyUM { get; set; }
        public string PartNum { get; set; }
        public string Description { get; set; }
        public double Price { get; set; }
        public string PriceUM { get; set; }
        public double Total { get; set; }
    }
}
