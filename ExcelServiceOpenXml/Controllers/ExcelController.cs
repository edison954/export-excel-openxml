using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Net;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Data;

namespace ExcelServiceOpenXml.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {

        [HttpPost]
        public FileResult ExportExcel(string fileName = "")
        {
            #region Init Data
            var datasource = new System.Data.DataTable();
            datasource.Columns.Add(new DataColumn("ID", typeof(Int32)));
            datasource.Columns.Add(new DataColumn("Name", typeof(string)));
            datasource.Columns.Add(new DataColumn("Score", typeof(Int32)));
            datasource.Columns.Add(new DataColumn("Team", typeof(string)));

            datasource.Rows.Add(10, "Bob", 12, "Xi'An");
            datasource.Rows.Add(11, "Tommy", 6, "Xi'An");
            datasource.Rows.Add(12, "Jaguar", 15, "Xi'An");
            datasource.Rows.Add(2, "Phillip", 9, "BeiJing");
            datasource.Rows.Add(3, "Hunter", 10, "BeiJing");
            datasource.Rows.Add(4, "Hellen", 8, "BeiJing");
            datasource.Rows.Add(5, "Jim", 9, "BeiJing");
            DataSet ds = new DataSet();
            ds.Tables.Add(datasource);
            #endregion

            MemoryStream stream = new MemoryStream();
            using (var workbook = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();
                foreach (System.Data.DataTable table in ds.Tables)
                {

                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                    sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                    DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    uint sheetId = 1;
                    if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                    {
                        sheetId =
                            sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                    List<String> columns = new List<string>();
                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }


                    sheetData.AppendChild(headerRow);

                    foreach (System.Data.DataRow dsrow in table.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        foreach (String col in columns)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }

                }

            }


            stream.Seek(0, SeekOrigin.Begin);

            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);

            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var donwloadFile = string.Format("attachment;filename={0}.xlsx;", string.IsNullOrEmpty(fileName) ? Guid.NewGuid().ToString() : WebUtility.UrlEncode(fileName));

            return File(bytes, contentType, donwloadFile);
        }
    }
}