using ExcelExport.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Export
{
    public class Excel : ExportUtils
    {
        public static void CreateExcel(WorkbookModel workbook)
        {
            var file = Utils.GetFileInfo(workbook.FileName, workbook.FilePath);
                        
            using (ExcelPackage xlPackage = new ExcelPackage(file))
            {   
                foreach (var sheet in workbook.WorksheetModels)
                {
                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add(sheet.WorksheetName);
                    if (worksheet != null)
                    {
                        //Add Logo
                        AddImage(worksheet, workbook.ImageModel);
                        //Create Headers and format them 
                        int cellNumber = 0;
                        foreach (var worksheetHeaderModel in sheet.WorksheetHeaderModels)
                        {
                            if(cellNumber == 0)
                            cellNumber = Utils.GetNumberFromString(worksheetHeaderModel.HeaderCell);
                            string cellLetter = new String(worksheetHeaderModel.HeaderCell.Where(Char.IsLetter).ToArray());
                            worksheetHeaderModel.HeaderCellRange = string.Format("{0}{1}:{2}{3}", cellLetter, cellNumber, Utils.GetExcelColumnName(sheet.WorksheetData.Columns.Count), cellNumber);
                            worksheetHeaderModel.HeaderCell = string.Format("{0}{1}", cellLetter, cellNumber);
                            CreateWorksheetHeader(worksheetHeaderModel, worksheet);
                            cellNumber++;
                        }
                        //Set from which row Column header should start
                        sheet.DataColumnHeaderModel.StartRow = cellNumber + 1;
                        //Set from which row data writing should start
                        sheet.WorksheetDataStartRow = cellNumber + 2;
                        
                        //Create Column Header                        
                        CreateColumnHeader(sheet, worksheet);
                        //Add data to excel
                        int row  = AddDataToExcel(sheet, worksheet, sheet.WorksheetData);
                        //Create footer
                        CreateFooter(worksheet, workbook.Disclaimer, row, sheet.WorksheetData.Columns.Count);                        
                    }
                }
                // save the new spreadsheet
                xlPackage.Save();
            }

            //return file.FullName;
        }

        /// <summary>
        /// This method is only for reference. It will not be used
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="exportDataList"></param>
        [Obsolete("Do not use this method. It is just for reference")]
        private static void CreateExcel<T>(WorkbookModel workbook, List<T> exportDataList)
        {

            PropertyInfo[] piT = typeof(T).GetProperties();

            var file = Utils.GetFileInfo("Sample3.xlsx", "");
            // ok, we can run the real code of the sample now
            using (ExcelPackage xlPackage = new ExcelPackage(file))
            {
                // get handle to the existing worksheet
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add("Sales");
                //var namedStyle = xlPackage.Workbook.Styles.CreateNamedStyle("HyperLink");   //This one is language dependent
                //namedStyle.Style.Font.UnderLine = true;
                //namedStyle.Style.Font.Color.SetColor(Color.Blue);
                if (worksheet != null)
                {
                    const int startRow = 5;
                    int row = startRow;
                    //Create Headers and format them 
                    worksheet.Cells["A1"].Value = "Transaction Detailed Report.";
                    using (ExcelRange r = worksheet.Cells["A1:G1"])
                    {
                        r.Merge = true;
                        r.Style.Font.SetFromFont(new Font("Britannic Bold", 22));
                        r.Style.Font.Color.SetColor(Color.White);
                        r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                        r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
                    }
                    worksheet.Cells["A2"].Value = "From Date 01-April-2018 To 31-July-2018";
                    using (ExcelRange r = worksheet.Cells["A2:G2"])
                    {
                        r.Merge = true;
                        r.Style.Font.SetFromFont(new Font("Britannic Bold", 18));
                        r.Style.Font.Color.SetColor(Color.Black);
                        r.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                        r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                    }

                    worksheet.Cells["A4"].Value = "CustomerID";
                    worksheet.Cells["B4"].Value = "CompanyName";
                    worksheet.Cells["C4"].Value = "ContactName";
                    worksheet.Cells["D4"].Value = "ContactTitle";
                    worksheet.Cells["E4"].Value = "Address";
                    worksheet.Cells["F4"].Value = "City";
                    worksheet.Cells["G4"].Value = "Region";
                    worksheet.Cells["H4"].Value = "PostalCode";
                    worksheet.Cells["I4"].Value = "Country";
                    worksheet.Cells["J4"].Value = "Phone";
                    worksheet.Cells["K4"].Value = "Fax";

                    worksheet.Cells["A4:K4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells["A4:K4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                    worksheet.Cells["A4:K4"].Style.Font.Bold = true;
                    
                    foreach (var item in exportDataList)
                    {
                        int col = 1;
                        for (int property = 0; property < piT.Count(); property++)
                        {
                            worksheet.Cells[row, col].Value = piT[property].GetValue(item, null);
                            col++;
                        }
                        row++;
                    }

                    //Set column width
                    worksheet.Column(1).Width = 20;
                    worksheet.Column(2).Width = 50;
                    worksheet.Column(3).Width = 50;
                    worksheet.Column(4).Width = 50;
                    worksheet.Column(5).Width = 55;
                    worksheet.Column(6).Width = 40;
                    worksheet.Column(7).Width = 12;
                    worksheet.Column(8).Width = 15;
                    worksheet.Column(9).Width = 20;
                    worksheet.Column(10).Width = 25;
                    worksheet.Column(11).Width = 25;

                    // lets set the header text 
                    worksheet.HeaderFooter.OddHeader.CenteredText = "AdventureWorks Inc. Sales Report";
                    // add the page number to the footer plus the total number of pages
                    worksheet.HeaderFooter.OddFooter.RightAlignedText =
                        string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                    // add the sheet name to the footer
                    worksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                    // add the file path to the footer
                    worksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;
                }
                // we had better add some document properties to the spreadsheet 

                // set some core property values
                xlPackage.Workbook.Properties.Title = "Sample 3";
                xlPackage.Workbook.Properties.Author = "John Tunnicliffe";
                xlPackage.Workbook.Properties.Subject = "ExcelPackage Samples";
                xlPackage.Workbook.Properties.Keywords = "Office Open XML";
                xlPackage.Workbook.Properties.Category = "ExcelPackage Samples";
                xlPackage.Workbook.Properties.Comments = "This sample demonstrates how to create an Excel 2007 file from scratch using the Packaging API and Office Open XML";

                // set some extended property values
                xlPackage.Workbook.Properties.Company = "AdventureWorks Inc.";
                xlPackage.Workbook.Properties.HyperlinkBase = new Uri("http://www.codeplex.com/MSFTDBProdSamples");

                // set some custom property values
                xlPackage.Workbook.Properties.SetCustomPropertyValue("Checked by", "John Tunnicliffe");
                xlPackage.Workbook.Properties.SetCustomPropertyValue("EmployeeID", "1147");
                xlPackage.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "ExcelPackage");

                // save the new spreadsheet
                xlPackage.Save();
            }

            //return file.FullName;
        }
    }
}
