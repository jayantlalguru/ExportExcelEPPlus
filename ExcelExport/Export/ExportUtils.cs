using ExcelExport.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Export
{
    public abstract class ExportUtils
    {
        /// <summary>
        /// Adds an image(mainly logo) to worksheet
        /// </summary>
        /// <param name="excelWorksheet">ExcelWorksheet object</param>
        /// <param name="imageModel">ImageModel object</param>
        public static void AddImage(ExcelWorksheet excelWorksheet, ImageModel imageModel)
        {   
            //int PixelTop = 88;
            //int PixelLeft = 129;
            //int Height = 320;
            //int Width = 200;
            Image img = Image.FromFile(imageModel.ImagePath);
            ExcelPicture pic = excelWorksheet.Drawings.AddPicture("Logo", img);
            pic.SetPosition(imageModel.RowIndex, 0, imageModel.ColumnIndex, 0);
            //pic.SetPosition(PixelTop, PixelLeft);  
            //pic.SetSize(Height, Width);
            //pic.SetSize(40);  
            excelWorksheet.Protection.IsProtected = false;
            excelWorksheet.Protection.AllowSelectLockedCells = false;
        }
        /// <summary>
        /// Creates Header for the worksheet
        /// </summary>
        /// <param name="worksheetModel">WorksheetModel model</param>
        /// <param name="worksheet">ExcelWorksheet object</param>
        protected static void CreateWorksheetHeader(WorksheetHeaderModel worksheetHeaderModel, ExcelWorksheet worksheet)
        {
            worksheet.Cells[worksheetHeaderModel.HeaderCell].Value = worksheetHeaderModel.HeaderText;
            //string cellRange = string.Format("{0}:{1}{2}", worksheetHeaderModel.HeaderCell, Utils.GetExcelColumnName(worksheetModel.WorksheetData.Columns.Count), Utils.GetColumnNumber(Utils.GetExcelColumnName(worksheetModel.WorksheetData.Columns.Count)));
            using (ExcelRange r = worksheet.Cells[worksheetHeaderModel.HeaderCellRange])
            {
                r.Merge = worksheetHeaderModel.MergeHeader;
                r.Style.Font.SetFromFont(new Font(worksheetHeaderModel.HeaderFontFamily, worksheetHeaderModel.HeaderFontSize));
                r.Style.Font.Color.SetColor(worksheetHeaderModel.HeaderTextColor);
                r.Style.HorizontalAlignment = (OfficeOpenXml.Style.ExcelHorizontalAlignment)worksheetHeaderModel.HorizontalAlignment;
                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(worksheetHeaderModel.HeaderBackgroundColor);
            }
        }
        /// <summary>
        /// Creates column header and sets the style
        /// </summary>
        /// <param name="worksheetModel">WorksheetModel object</param>
        /// <param name="worksheet">ExcelWorksheet object</param>
        protected static void CreateColumnHeader(WorksheetModel worksheetModel, ExcelWorksheet worksheet)
        {
            int columnCount = 1;
            int rowNumberForColumnHeader = worksheetModel.DataColumnHeaderModel.StartRow;            
            //string cellLetter = new String(worksheetModel.DataColumnHeaderModel.ColumnStartCell.Where(Char.IsLetter).ToArray());
            int columnIndexValue = Utils.GetColumnNumber(worksheetModel.DataColumnHeaderModel.ColumnStartCell);
            string lastColumnCell = string.Empty;
            foreach (var column in worksheetModel.DataColumnHeaderModel.Columns)
            {
                string columnLetter = string.Empty;
                columnLetter = Utils.GetExcelColumnName(columnIndexValue);
                worksheet.Cells[columnLetter + rowNumberForColumnHeader.ToString()].Value = column.ColumnName;
                //Set column width
                worksheet.Column(columnCount).Width = column.ColumnWidth;
                lastColumnCell = columnLetter;
                columnCount++;
                columnIndexValue++;
            }
            
            string worksheetHeaderCellRange = string.Format("{0}{1}:{2}{3}", worksheetModel.DataColumnHeaderModel.ColumnStartCell, worksheetModel.DataColumnHeaderModel.StartRow, lastColumnCell, rowNumberForColumnHeader);
            worksheet.Cells[worksheetHeaderCellRange].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[worksheetHeaderCellRange].Style.Fill.BackgroundColor.SetColor(worksheetModel.DataColumnHeaderModel.BackgroundColor);
            worksheet.Cells[worksheetHeaderCellRange].Style.Font.Bold = worksheetModel.DataColumnHeaderModel.Bold;
            worksheet.Cells[worksheetHeaderCellRange].Style.Font.Color.SetColor(worksheetModel.DataColumnHeaderModel.FontColor);
        }
        /// <summary>
        /// Sets border around data
        /// </summary>
        /// <param name="worksheet">ExcelWorksheet object</param>
        /// <param name="startRow">row index to start setting border</param>
        /// <param name="totalRows">Total roes to set border</param>
        /// <param name="totalColumns">Total columns to set border</param>
        private static void SetBorder(ExcelWorksheet worksheet, int startRow, int totalRows, int totalColumns)
        {   
            //Border around the data range
            worksheet.Cells[startRow, 1, (totalRows + startRow - 1), totalColumns].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[startRow, 1, (totalRows + startRow - 1), totalColumns].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[startRow, 1, (totalRows + startRow - 1), totalColumns].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[startRow, 1, (totalRows + startRow - 1), totalColumns].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
        /// <summary>
        /// Primary functionality to set data to cells. Calls SetBorder to set all the borders
        /// </summary>
        /// <param name="worksheetModel">Model for worsheet</param>
        /// <param name="worksheet">ExcelWorksheet object</param>
        /// <param name="dataTable">Table containing data to export</param>
        /// <returns></returns>
        protected static int AddDataToExcel(WorksheetModel worksheetModel, ExcelWorksheet worksheet, DataTable dataTable)
        {   
            int row = worksheetModel.WorksheetDataStartRow;
            SetBorder(worksheet, row, dataTable.Rows.Count, dataTable.Columns.Count);
            //Assign data to cells
            foreach (DataRow dr in worksheetModel.WorksheetData.Rows)
            {
                int columIndex = 1;
                foreach (var item in worksheetModel.WorksheetData.Columns)
                {
                    worksheet.Cells[row, columIndex].Value = dr[item.ToString()];                    
                    columIndex++;
                }
                row++;
            }
            return row;
        }
        /// <summary>
        /// Creates footer disclaimer. 
        /// Styles are fixed as this won't change.
        /// Wraps text and merge rows
        /// </summary>
        /// <param name="worksheet">ExcelWorksheet object</param>
        /// <param name="disclaimer">Text for disclaimer</param>
        /// <param name="row">row where disclaimer to set</param>
        /// <param name="range">DataTable column count</param>
        protected static void CreateFooter(ExcelWorksheet worksheet, string disclaimer, int row, int range)
        {
            worksheet.Cells["A" + (row + 2).ToString()].Value = "Disclaimer:";
            worksheet.Cells["A" + (row + 2).ToString()].Style.Font.Bold = true;
            string footerCell = "A" + (row + 3).ToString();            
            worksheet.Cells[footerCell].Value = disclaimer;            
            worksheet.Row(row + 3).Height = 60;
            string cellRange = string.Format("{0}:{1}{2}", footerCell, Utils.GetExcelColumnName(range), Utils.GetNumberFromString(footerCell));
            using (ExcelRange r = worksheet.Cells[cellRange])
            {
                r.Merge = true;                
                r.Style.Font.SetFromFont(new Font("Arial", 10));
                r.Style.Font.Color.SetColor(Color.Black);
                r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                r.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                r.Style.WrapText = true;
            }
        }
    }
}
