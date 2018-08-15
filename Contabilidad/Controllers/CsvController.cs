using Contabilidad.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace Contabilidad.Controllers
{
    public class CsvController : Controller
    {
        // GET: Csv
        public ActionResult Index()
        {
            return View();
        }

        public FileStreamResult Exportcsv(List<HttpPostedFileBase> files)
        {
            XMLHelper.Init();

            MemoryStream stream = new MemoryStream();
            SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            document.AddWorkbookPart();
            document.WorkbookPart.Workbook = new Workbook();
            document.WorkbookPart.AddNewPart<WorksheetPart>();
            document.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet = new Worksheet();
            Worksheet worksheet = document.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet;
            SheetData sheetData = document.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet.AppendChild<SheetData>(new SheetData());

            var count = 0;

            foreach (HttpPostedFileBase base2 in files)
            {
                byte[] bytes = null;
                using (BinaryReader reader = new BinaryReader(base2.InputStream))
                {
                    bytes = reader.ReadBytes(base2.ContentLength);
                }

                Cell Title = XMLHelper.InsertCellInWorksheet("A", XMLHelper.RowIndex, sheetData, worksheet);
                Title.CellValue = new CellValue(string.Format("Archivo {0}", count));
                Title.DataType = CellValues.String;

                XMLHelper.RowIndex++;
                count++;

                Encoding.UTF8.GetString(bytes).Split('\n').ToList().ForEach(line =>
                {
                    Cell row = XMLHelper.InsertCellInWorksheet("A", XMLHelper.RowIndex, sheetData, worksheet);
                    row.CellValue = new CellValue(line);
                    row.DataType = CellValues.String;

                    XMLHelper.RowIndex++;
                });
            }

            document.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet.Save();
            document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Sheet newChild = new Sheet
            {
                Id = document.WorkbookPart.GetIdOfPart(document.WorkbookPart.WorksheetParts.First<WorksheetPart>()),
                SheetId = 1,
                Name = "Hoja 1"
            };
            document.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild<Sheet>(newChild);
            document.WorkbookPart.Workbook.Save();
            document.Close();
            stream.Position = 0;
            FileStreamResult result1 = new FileStreamResult(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "csv-" + DateTime.Now.Ticks + ".xlsx"
            };
            return result1;
        }
    }
}