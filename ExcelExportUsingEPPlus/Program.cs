using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelExport.Export;
using ExcelExport.Models;

namespace ExcelExportUsingEPPlus
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Customer> customers = new List<Customer>();
            ExcelExport.Export.Utils.OutputDir = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}SampleApp");
            using (SqlConnection sqlConn = new SqlConnection(@"Data Source=JAYANTGURU-PC\SQLSERVER2017;Initial Catalog=Practice;Persist Security Info=True;User ID=sa;Password=guru@000"))
            {
                sqlConn.Open();
                using (SqlCommand sqlCmd = new SqlCommand(@"SELECT [CustomerID]
                                                                  ,[CompanyName]
                                                                  ,[ContactName]
                                                                  ,[ContactTitle]
                                                                  ,[Address]
                                                                  ,[City]
                                                                  ,[Region]
                                                                  ,[PostalCode]
                                                                  ,[Country]
                                                                  ,[Phone]
                                                                  ,[Fax]
                                                                   FROM [Practice].[dbo].[Customers]", sqlConn))
                {
                    using (SqlDataReader sqlReader = sqlCmd.ExecuteReader())
                    {
                        // get the data and fill rows 5 onwards
                        while (sqlReader.Read())
                        {
                            Customer customer = new Customer();
                            customer.CustomerID = sqlReader["CustomerID"] != DBNull.Value ? Convert.ToInt32(sqlReader["CustomerID"]) : 0;
                            customer.CompanyName = sqlReader["CompanyName"] != DBNull.Value ? sqlReader["CompanyName"].ToString() : string.Empty;
                            customer.ContactName = sqlReader["ContactName"] != DBNull.Value ? sqlReader["ContactName"].ToString() : string.Empty;
                            customer.ContactTitle = sqlReader["ContactTitle"] != DBNull.Value ? sqlReader["ContactTitle"].ToString() : string.Empty;
                            customer.Address = sqlReader["Address"] != DBNull.Value ? sqlReader["Address"].ToString() : string.Empty;
                            customer.City = sqlReader["City"] != DBNull.Value ? sqlReader["City"].ToString() : string.Empty;
                            customer.Region = sqlReader["Region"] != DBNull.Value ? sqlReader["Region"].ToString() : string.Empty;
                            customer.PostalCode = sqlReader["PostalCode"] != DBNull.Value ? sqlReader["PostalCode"].ToString() : string.Empty;
                            customer.Country = sqlReader["Country"] != DBNull.Value ? sqlReader["Country"].ToString() : string.Empty;
                            customer.Phone = sqlReader["Phone"] != DBNull.Value ? sqlReader["Phone"].ToString() : string.Empty;
                            customer.Fax = sqlReader["Fax"] != DBNull.Value ? sqlReader["Fax"].ToString() : string.Empty;
                            customers.Add(customer);
                        }
                        sqlReader.Close();
                    }
                }
                sqlConn.Close();
            }
            DataTable dataTable = Utils.ToDataTable<Customer>(customers);

            WorkbookModel workbook = new WorkbookModel();
            workbook.FilePath = ConfigurationManager.AppSettings["ExcelFilePath"];
            workbook.Disclaimer = ConfigurationManager.AppSettings["Disclaimer"];
            workbook.ImageModel.ImagePath = ConfigurationManager.AppSettings["LogoPath"];
            WorksheetModel worksheetModel = new WorksheetModel();
            worksheetModel.WorksheetData = dataTable;            
            foreach (var item in dataTable.Columns)
            {
                ColumnModel columnModel = new ColumnModel();
                columnModel.ColumnName = item.ToString();
                worksheetModel.DataColumnHeaderModel.Columns.Add(columnModel);
            }
            WorksheetHeaderModel worksheetHeaderModel = new WorksheetHeaderModel();
            worksheetHeaderModel.HeaderFontSize = 15;
            worksheetHeaderModel.HeaderText = "Transaction Detailed Report";
            worksheetHeaderModel.MergeHeader = true;
            worksheetModel.WorksheetHeaderModels.Add(worksheetHeaderModel);
            WorksheetHeaderModel worksheetHeaderModel2 = new WorksheetHeaderModel();
            worksheetHeaderModel2.HeaderFontSize = 15;
            worksheetHeaderModel2.HeaderText = "Report Dated 01-May-2018 to 31-July-2018";
            worksheetHeaderModel2.MergeHeader = true;
            worksheetModel.WorksheetHeaderModels.Add(worksheetHeaderModel2);
            WorksheetHeaderModel worksheetHeaderModel3 = new WorksheetHeaderModel();
            worksheetHeaderModel3.HeaderFontSize = 15;
            worksheetHeaderModel3.HeaderText = "AMC: Aditya Birla Mutual Funds";
            worksheetHeaderModel3.MergeHeader = true;
            worksheetModel.WorksheetHeaderModels.Add(worksheetHeaderModel3);
            WorksheetHeaderModel worksheetHeaderModel4 = new WorksheetHeaderModel();
            worksheetHeaderModel4.HeaderFontSize = 15;
            worksheetHeaderModel4.HeaderText = "Asset Type: Equity";
            worksheetHeaderModel4.MergeHeader = true;
            worksheetHeaderModel4.HorizontalAlignment = HorizontalAlignment.Left;
            worksheetModel.WorksheetHeaderModels.Add(worksheetHeaderModel4);
            WorksheetHeaderModel worksheetHeaderModel5 = new WorksheetHeaderModel();
            worksheetHeaderModel5.HeaderFontSize = 10;
            worksheetHeaderModel5.HeaderText = "Sub Type: Mid Cap";
            worksheetHeaderModel5.MergeHeader = true;
            worksheetHeaderModel5.HorizontalAlignment = HorizontalAlignment.Left;
            worksheetModel.WorksheetHeaderModels.Add(worksheetHeaderModel5);
            workbook.WorksheetModels.Add(worksheetModel);
            Excel.CreateExcel(workbook);
        }
    }
}
