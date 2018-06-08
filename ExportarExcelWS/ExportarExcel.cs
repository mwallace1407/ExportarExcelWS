using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportarExcelWS
{
    public class ExportarExcel
    {
        #region OpenXml

        static private int rowsPerSheet = 250000;
        private DataTable ResultsData = new DataTable();

        public string GenerarExcel(IDataReader reader, string RutaArchivos, string Archivo, int RegistrosPorHoja)
        {
            try
            {
                //string Archivo = "";
                DataTable dtSchema = reader.GetSchemaTable();
                int c = 0;
                bool firstTime = true;
                var listCols = new List<DataColumn>();

                rowsPerSheet = RegistrosPorHoja;

                if (dtSchema != null)
                {
                    foreach (DataRow drow in dtSchema.Rows)
                    {
                        string columnName = Convert.ToString(drow["ColumnName"]);
                        var column = new DataColumn(columnName, (Type)(drow["DataType"]));

                        column.Unique = (bool)drow["IsUnique"];
                        column.AllowDBNull = (bool)drow["AllowDBNull"];
                        column.AutoIncrement = (bool)drow["IsAutoIncrement"];
                        listCols.Add(column);
                        ResultsData.Columns.Add(column);
                    }
                }

                while (reader.Read())
                {
                    DataRow dataRow = ResultsData.NewRow();

                    for (int i = 0; i < listCols.Count; i++)
                    {
                        dataRow[(listCols[i])] = reader[i];
                    }

                    ResultsData.Rows.Add(dataRow);
                    c++;

                    if (c == rowsPerSheet)
                    {
                        c = 0;
                        Archivo = ExportToOxml(firstTime, Archivo, RutaArchivos);
                        ResultsData.Clear();
                        firstTime = false;
                    }
                }

                if (ResultsData.Rows.Count > 0)
                {
                    Archivo = ExportToOxml(firstTime, Archivo, RutaArchivos);
                    ResultsData.Clear();
                }

                reader.Close();

                return Archivo;
            }
            catch (Exception ex)
            {
                return "Error: Generar=>" + ex.Message + ex.StackTrace;
            }
        }

        private string ExportToOxml(bool firstTime, string fileName, string RutaArchivos)
        {
            try
            {
                //Check if the file exists.
                //if (firstTime)
                //{
                //    Random rnd = new Random();

                //    fileName = "Reporte_" + DateTime.Now.ToString("yyyyMMddHHmmss") + rnd.Next(1000).ToString().PadLeft(4, Convert.ToChar("0")) + ".xlsx";

                //    while (System.IO.File.Exists(fileName))
                //        fileName = "Reporte_" + DateTime.Now.ToString("yyyyMMddHHmmss") + rnd.Next(1000).ToString().PadLeft(4, Convert.ToChar("0")) + ".xlsx";
                //}

                uint sheetId = 1; //Start at the first sheet in the Excel workbook.

                if (firstTime)
                {
                    //This is the first time of creating the excel file and the first sheet.
                    // Create a spreadsheet document by supplying the filepath.
                    // By default, AutoSave = true, Editable = true, and Type = xlsx.
                    SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                        Create(System.IO.Path.Combine(RutaArchivos, fileName), SpreadsheetDocumentType.Workbook);

                    // Add a WorkbookPart to the document.
                    WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();

                    // Add a WorksheetPart to the WorkbookPart.
                    var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    var bold1 = new Bold();
                    CellFormat cf = new CellFormat();

                    // Add Sheets to the Workbook.
                    Sheets sheets;
                    sheets = spreadsheetDocument.WorkbookPart.Workbook.
                        AppendChild<Sheets>(new Sheets());

                    // Append a new worksheet and associate it with the workbook.
                    var sheet = new Sheet()
                    {
                        Id = spreadsheetDocument.WorkbookPart.
                            GetIdOfPart(worksheetPart),
                        SheetId = sheetId,
                        Name = "Hoja" + sheetId
                    };
                    sheets.Append(sheet);

                    //Add Header Row.
                    var headerRow = new Row();
                    foreach (DataColumn column in ResultsData.Columns)
                    {
                        var cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(column.ColumnName)
                        };
                        headerRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(headerRow);

                    foreach (DataRow row in ResultsData.Rows)
                    {
                        var newRow = new Row();
                        foreach (DataColumn col in ResultsData.Columns)
                        {
                            var cell = new Cell
                            {
                                DataType = CellValues.String,
                                CellValue = new CellValue(row[col].ToString())
                            };
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }
                    workbookpart.Workbook.Save();

                    spreadsheetDocument.Close();
                }
                else
                {
                    // Open the Excel file that we created before, and start to add sheets to it.
                    var spreadsheetDocument = SpreadsheetDocument.Open(System.IO.Path.Combine(RutaArchivos, fileName), true);

                    var workbookpart = spreadsheetDocument.WorkbookPart;
                    if (workbookpart.Workbook == null)
                        workbookpart.Workbook = new Workbook();

                    var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);
                    var sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets;

                    if (sheets.Elements<Sheet>().Any())
                    {
                        //Set the new sheet id
                        sheetId = sheets.Elements<Sheet>().Max(s => s.SheetId.Value) + 1;
                    }
                    else
                    {
                        sheetId = 1;
                    }

                    // Append a new worksheet and associate it with the workbook.
                    var sheet = new Sheet()
                    {
                        Id = spreadsheetDocument.WorkbookPart.
                            GetIdOfPart(worksheetPart),
                        SheetId = sheetId,
                        Name = "Hoja" + sheetId
                    };
                    sheets.Append(sheet);

                    //Add the header row here.
                    var headerRow = new Row();

                    foreach (DataColumn column in ResultsData.Columns)
                    {
                        var cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(column.ColumnName)
                        };
                        headerRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(headerRow);

                    foreach (DataRow row in ResultsData.Rows)
                    {
                        var newRow = new Row();

                        foreach (DataColumn col in ResultsData.Columns)
                        {
                            var cell = new Cell
                            {
                                DataType = CellValues.String,
                                CellValue = new CellValue(row[col].ToString())
                            };
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }

                    workbookpart.Workbook.Save();

                    // Close the document.
                    spreadsheetDocument.Close();
                }

                return fileName;
            }
            catch (Exception ex)
            {
                return "Error: Export=>" + ex.Message + ex.StackTrace;
            }
        }

        #endregion OpenXml
    }
}