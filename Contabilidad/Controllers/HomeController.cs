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
using System.Xml.Linq;

namespace Contabilidad.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }

        public FileStreamResult Exportexcel(List<HttpPostedFileBase> xml)
        {
            MemoryStream stream = new MemoryStream();
            SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            document.AddWorkbookPart();
            document.WorkbookPart.Workbook = new Workbook();
            document.WorkbookPart.AddNewPart<WorksheetPart>();
            document.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet = new Worksheet();
            document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = GenerateStyleSheet();
            Worksheet worksheet = document.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet;
            SheetData sheetData = document.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet.AppendChild<SheetData>(new SheetData());
            string[] strArray = new string[] {  "A",      "B",     "C",    "D",        "E",        "F",                    "G",                         "H",                "I",                        "J", "K", "L", "M", "N", "O", "P","Q", "R", "S", "T" };
            string[] strArray2 = new string[] {"Fecha", "Serie", "Folio", "UUID", "RFC (Emisor)", "Domicilio (Emisor)"  , "Razón Social (Emisor)", "RFC (Receptor)", "Domicilio (Receptor)", "Razón Social (Receptor)", "Desglose Conceptos e Impuestos", "", "", "", "", "", "Total Impuestos Retenidos", "Total Impuestos Trasladados","Total" };
            int index = 0;
            uint rowIndex = 2;
            string[] strArray3 = strArray2;

            MergeCells mergeCells = new MergeCells();

            // Insert a MergeCells object into the specified position.
            if (worksheet.Elements<CustomSheetView>().Count() > 0)
                worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
            else
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());

            // Create the merged cell and append it to the MergeCells collection.
            MergeCell mergeCell = new MergeCell()
            {
                Reference =
                new StringValue("K1:P1")
            };
            mergeCells.Append(mergeCell);

            for (int i = 0; i < strArray3.Length; i++)
            {
                string text1 = strArray3[i];
                Cell cell1 = InsertCellInWorksheet(strArray[index], 1, sheetData, worksheet);
                cell1.CellValue = new CellValue(strArray2[index]);
                cell1.DataType = CellValues.String;
                index++;
            }
            foreach (HttpPostedFileBase base2 in xml)
            {
                byte[] bytes = null;
                using (BinaryReader reader = new BinaryReader(base2.InputStream))
                {
                    bytes = reader.ReadBytes(base2.ContentLength);
                }
                char ch = (char)0xfeff;
                XDocument document2 = XDocument.Parse(Encoding.UTF8.GetString(bytes).Replace(ch.ToString(), ""));
                XNamespace namespace2 = document2.Root.Name.Namespace;
                XNamespace namespace3 = "http://www.sat.gob.mx/TimbreFiscalDigital";
                XNamespace namespace4 = "http://www.sat.gob.mx/implocal";
                DateTime time = DateTime.Parse(document2.Root.Attribute("fecha") != null ? document2.Root.Attribute("fecha").Value : document2.Root.Attribute("Fecha").Value);

                Cell cell2 = InsertCellInWorksheet("A", rowIndex, sheetData, worksheet);
                cell2.CellValue = new CellValue($"{time:G}");
                cell2.DataType = CellValues.String;

                Cell cell3 = InsertCellInWorksheet("B", rowIndex, sheetData, worksheet);
                cell3.CellValue = new CellValue((document2.Root.Attribute("serie") != null) ? document2.Root.Attribute("serie").Value : "");
                cell3.DataType = CellValues.String;
                Cell cell4 = InsertCellInWorksheet("C", rowIndex, sheetData, worksheet);
                cell4.CellValue = new CellValue((document2.Root.Attribute("folio") != null) ? document2.Root.Attribute("folio").Value : "");
                cell4.DataType = CellValues.Number;
                Cell cell5 = InsertCellInWorksheet("D", rowIndex, sheetData, worksheet);
                cell5.CellValue = new CellValue(document2.Root.Element((XName)(namespace2 + "Complemento")).Element((XName)(namespace3 + "TimbreFiscalDigital")).Attribute("UUID").Value);
                cell5.DataType = CellValues.String;
                Cell cell6 = InsertCellInWorksheet("E", rowIndex, sheetData, worksheet);

                var rfc_emisor = document2.Root.Element((XName)(namespace2 + "Emisor")).Attribute("rfc") != null
                    ? document2.Root.Element((XName)(namespace2 + "Emisor")).Attribute("rfc").Value
                    : document2.Root.Element((XName)(namespace2 + "Emisor")).Attribute("Rfc").Value;

                cell6.CellValue = new CellValue(rfc_emisor);
                cell6.DataType = CellValues.String;

                Cell CellDomicilioEmisor = InsertCellInWorksheet("F", rowIndex, sheetData, worksheet);

                var DomicilioEmisor = string.Empty;
                var currentDomicilioEmisor = (document2.Root.Element((XName)(namespace2 + "Emisor")).Element((XName)(namespace2 + "DomicilioFiscal")));
                if (currentDomicilioEmisor != null)
                {
                    var calleEmisor = currentDomicilioEmisor.Attribute("calle") != null ? currentDomicilioEmisor.Attribute("calle").Value : string.Empty;
                    var noExteriorEmisor = currentDomicilioEmisor.Attribute("noExterior") != null ? currentDomicilioEmisor.Attribute("noExterior").Value : string.Empty;
                    var noInteriorEmisor = currentDomicilioEmisor.Attribute("noInterior") != null ? currentDomicilioEmisor.Attribute("noInterior").Value : string.Empty;
                    var coloniaEmisor = currentDomicilioEmisor.Attribute("colonia") != null ? currentDomicilioEmisor.Attribute("colonia").Value : string.Empty;
                    var municipioEmisor = currentDomicilioEmisor.Attribute("municipio") != null ? currentDomicilioEmisor.Attribute("municipio").Value : string.Empty;
                    var estadoEmisor = currentDomicilioEmisor.Attribute("estado") != null ? currentDomicilioEmisor.Attribute("estado").Value : string.Empty;
                    var paisEmisor = currentDomicilioEmisor.Attribute("pais") != null ? currentDomicilioEmisor.Attribute("pais").Value : string.Empty;
                    var codigoPostalEmisor = currentDomicilioEmisor.Attribute("codigoPostal") != null ? currentDomicilioEmisor.Attribute("codigoPostal").Value : string.Empty;
                    DomicilioEmisor = calleEmisor + (string.IsNullOrEmpty(noExteriorEmisor) ? string.Empty : ", No. Ext: " + noExteriorEmisor)
                                                    + (string.IsNullOrEmpty(noInteriorEmisor) ? string.Empty : ", No. Int: " + noInteriorEmisor)
                                                    + (string.IsNullOrEmpty(coloniaEmisor) ? string.Empty : ", " + coloniaEmisor)
                                                    + (string.IsNullOrEmpty(municipioEmisor) ? string.Empty : ", " + municipioEmisor)
                                                    + (string.IsNullOrEmpty(estadoEmisor) ? string.Empty : ", " + estadoEmisor)
                                                    + (string.IsNullOrEmpty(paisEmisor) ? string.Empty : ", " + paisEmisor)
                                                    + (string.IsNullOrEmpty(codigoPostalEmisor) ? string.Empty : ", CP: " + codigoPostalEmisor);

                }

                CellDomicilioEmisor.CellValue = new CellValue(DomicilioEmisor);
                CellDomicilioEmisor.DataType = CellValues.String;

                Cell cell7 = InsertCellInWorksheet("G", rowIndex, sheetData, worksheet);
                cell7.CellValue = new CellValue((document2.Root.Element((XName)(namespace2 + "Emisor")).Attribute("nombre") != null) ? document2.Root.Element((XName)(namespace2 + "Emisor")).Attribute("nombre").Value : "");
                cell7.DataType = CellValues.String;
                Cell cell8 = InsertCellInWorksheet("H", rowIndex, sheetData, worksheet);

                var rfc_receptor = document2.Root.Element((XName)(namespace2 + "Receptor")).Attribute("rfc") != null
                    ? document2.Root.Element((XName)(namespace2 + "Receptor")).Attribute("rfc").Value
                    : document2.Root.Element((XName)(namespace2 + "Receptor")).Attribute("Rfc").Value;

                cell8.CellValue = new CellValue(rfc_receptor);
                cell8.DataType = CellValues.String;

                var DomicilioReceptor = string.Empty;
                var currentDomicilioReceptor = (document2.Root.Element((XName)(namespace2 + "Receptor")).Element((XName)(namespace2 + "Domicilio")));
                if (currentDomicilioReceptor != null)
                {
                    var calle = currentDomicilioReceptor.Attribute("calle") != null ? currentDomicilioReceptor.Attribute("calle").Value : string.Empty;
                    var noExterior = currentDomicilioReceptor.Attribute("noExterior") != null ? currentDomicilioReceptor.Attribute("noExterior").Value : string.Empty;
                    var noInterior = currentDomicilioReceptor.Attribute("noInterior") != null ? currentDomicilioReceptor.Attribute("noInterior").Value : string.Empty;
                    var colonia = currentDomicilioReceptor.Attribute("colonia") != null ? currentDomicilioReceptor.Attribute("colonia").Value : string.Empty;
                    var municipio = currentDomicilioReceptor.Attribute("municipio") != null ? currentDomicilioReceptor.Attribute("municipio").Value : string.Empty;
                    var estado = currentDomicilioReceptor.Attribute("estado") != null ? currentDomicilioReceptor.Attribute("estado").Value : string.Empty;
                    var pais = currentDomicilioReceptor.Attribute("pais") != null ? currentDomicilioReceptor.Attribute("pais").Value : string.Empty;
                    var codigoPostal = currentDomicilioReceptor.Attribute("codigoPostal") != null ? currentDomicilioReceptor.Attribute("codigoPostal").Value : string.Empty;
                    DomicilioReceptor = calle + (string.IsNullOrEmpty(noExterior) ? string.Empty : ", No. Ext: " + noExterior)
                                                    + (string.IsNullOrEmpty(noInterior) ? string.Empty : ", No. Int: " + noInterior)
                                                    + (string.IsNullOrEmpty(colonia) ? string.Empty : ", " + colonia)
                                                    + (string.IsNullOrEmpty(municipio) ? string.Empty : ", " + municipio)
                                                    + (string.IsNullOrEmpty(estado) ? string.Empty : ", " + estado)
                                                    + (string.IsNullOrEmpty(pais) ? string.Empty : ", " + pais)
                                                    + (string.IsNullOrEmpty(codigoPostal) ? string.Empty : ", CP: " + codigoPostal);

                }

                Cell CellDomicilioReceptor = InsertCellInWorksheet("I", rowIndex, sheetData, worksheet);
                CellDomicilioReceptor.CellValue = new CellValue(DomicilioReceptor);
                CellDomicilioReceptor.DataType = CellValues.String;

                if ((document2.Root.Element((XName)(namespace2 + "Impuestos")) != null))
                {
                    //rowIndex++;
                    Cell cell36 = InsertCellInWorksheet("Q", rowIndex, sheetData, worksheet);
                    cell36.CellValue = new CellValue((document2.Root.Element((XName)(namespace2 + "Impuestos")).Attribute("totalImpuestosRetenidos") != null) ? document2.Root.Element((XName)(namespace2 + "Impuestos")).Attribute("totalImpuestosRetenidos").Value : "");
                    cell36.DataType = CellValues.Number;
                    Cell cell37 = InsertCellInWorksheet("R", rowIndex, sheetData, worksheet);
                    cell37.CellValue = new CellValue((document2.Root.Element((XName)(namespace2 + "Impuestos")).Attribute("totalImpuestosTrasladados") != null) ? document2.Root.Element((XName)(namespace2 + "Impuestos")).Attribute("totalImpuestosTrasladados").Value : "");
                    cell37.DataType = CellValues.Number;
                    Cell cell38 = InsertCellInWorksheet("S", rowIndex, sheetData, worksheet);
                    cell38.CellValue = new CellValue();
                    cell38.DataType = CellValues.Number;
                }

                var total = document2.Root.Attribute("total") != null
                    ? document2.Root.Attribute("total").Value
                    : document2.Root.Attribute("Total").Value;

                Cell totalcell = InsertCellInWorksheet("S", rowIndex, sheetData, worksheet);
                totalcell.CellValue = new CellValue(total);
                totalcell.DataType = CellValues.Number;


                Cell cell9 = InsertCellInWorksheet("J", rowIndex, sheetData, worksheet);
                cell9.CellValue = new CellValue((document2.Root.Element((XName)(namespace2 + "Receptor")).Attribute("nombre") != null) ? document2.Root.Element((XName)(namespace2 + "Receptor")).Attribute("nombre").Value : "");
                cell9.DataType = CellValues.String;
                Cell cell10 = InsertCellInWorksheet("K", rowIndex, sheetData, worksheet);
                cell10.CellValue = new CellValue("Cantidad");
                cell10.DataType = CellValues.String;
                cell10.StyleIndex = 1;
                Cell cell11 = InsertCellInWorksheet("L", rowIndex, sheetData, worksheet);
                cell11.CellValue = new CellValue("Unidad");
                cell11.DataType = CellValues.String;
                cell11.StyleIndex = 1;
                Cell cell12 = InsertCellInWorksheet("M", rowIndex, sheetData, worksheet);
                cell12.CellValue = new CellValue("No. Identificaci\x00f3n");
                cell12.DataType = CellValues.String;
                cell12.StyleIndex = 1;
                Cell cell13 = InsertCellInWorksheet("N", rowIndex, sheetData, worksheet);
                cell13.CellValue = new CellValue("Descripci\x00f3n");
                cell13.DataType = CellValues.String;
                cell13.StyleIndex = 1;
                Cell cell14 = InsertCellInWorksheet("O", rowIndex, sheetData, worksheet);
                cell14.CellValue = new CellValue("Valor Unitario");
                cell14.DataType = CellValues.String;
                cell14.StyleIndex = 1;
                Cell cell15 = InsertCellInWorksheet("P", rowIndex, sheetData, worksheet);
                cell15.CellValue = new CellValue("Importe");
                cell15.DataType = CellValues.String;
                cell15.StyleIndex = 1;
                foreach (XElement element in document2.Root.Element((XName)(namespace2 + "Conceptos")).Descendants())
                {
                    if (element.Attribute("cantidad") != null)
                    {
                        rowIndex++;
                        Cell cell16 = InsertCellInWorksheet("K", rowIndex, sheetData, worksheet);
                        cell16.CellValue = new CellValue((element.Attribute("cantidad") != null) ? element.Attribute("cantidad").Value : "");
                        cell16.DataType = CellValues.Number;
                        Cell cell17 = InsertCellInWorksheet("L", rowIndex, sheetData, worksheet);
                        cell17.CellValue = new CellValue((element.Attribute("unidad") != null) ? element.Attribute("unidad").Value : "");
                        cell17.DataType = CellValues.String;
                        Cell cell18 = InsertCellInWorksheet("M", rowIndex, sheetData, worksheet);
                        cell18.CellValue = new CellValue((element.Attribute("noIdentificacion") != null) ? element.Attribute("noIdentificacion").Value : "");
                        cell18.DataType = CellValues.String;
                        Cell cell19 = InsertCellInWorksheet("N", rowIndex, sheetData, worksheet);
                        cell19.CellValue = new CellValue((element.Attribute("descripcion") != null) ? element.Attribute("descripcion").Value : "");
                        cell19.DataType = CellValues.String;
                        Cell cell20 = InsertCellInWorksheet("O", rowIndex, sheetData, worksheet);
                        cell20.CellValue = new CellValue((element.Attribute("valorUnitario") != null) ? element.Attribute("valorUnitario").Value : "");
                        cell20.DataType = CellValues.Number;
                        Cell cell21 = InsertCellInWorksheet("P", rowIndex, sheetData, worksheet);
                        cell21.CellValue = new CellValue((element.Attribute("importe") != null) ? element.Attribute("importe").Value : "");
                        cell21.DataType = CellValues.Number;
                    }
                }

                if (document2.Root.Element((XName)(namespace2 + "Impuestos")) != null)
                {

                    if (document2.Root.Element((XName)(namespace2 + "Impuestos")).Element((XName)(namespace2 + "Traslados")) != null)
                    {

                        rowIndex++;
                        Cell cell22 = InsertCellInWorksheet("M", rowIndex, sheetData, worksheet);
                        cell22.CellValue = new CellValue("Impuestos Trasladados");
                        cell22.DataType = CellValues.String;
                        cell22.StyleIndex = 1;
                        Cell cell23 = InsertCellInWorksheet("N", rowIndex, sheetData, worksheet);
                        cell23.CellValue = new CellValue("Impuesto");
                        cell23.DataType = CellValues.String;
                        cell23.StyleIndex = 1;
                        Cell cell24 = InsertCellInWorksheet("O", rowIndex, sheetData, worksheet);
                        cell24.CellValue = new CellValue("Tasa");
                        cell24.DataType = CellValues.String;
                        cell24.StyleIndex = 1;
                        Cell cell25 = InsertCellInWorksheet("P", rowIndex, sheetData, worksheet);
                        cell25.CellValue = new CellValue("Importe");
                        cell25.DataType = CellValues.String;
                        cell25.StyleIndex = 1;
                        foreach (XElement element2 in document2.Root.Element((XName)(namespace2 + "Impuestos")).Element((XName)(namespace2 + "Traslados")).Descendants())
                        {
                            rowIndex++;
                            Cell cell26 = InsertCellInWorksheet("N", rowIndex, sheetData, worksheet);

                            var impuesto = element2.Attribute("impuesto") != null
                                ? element2.Attribute("impuesto").Value
                                : element2.Attribute("Impuesto").Value;

                            cell26.CellValue = new CellValue(impuesto);
                            cell26.DataType = CellValues.String;
                            Cell cell27 = InsertCellInWorksheet("O", rowIndex, sheetData, worksheet);

                            var tasa = element2.Attribute("tasa") != null
                                ? element2.Attribute("tasa").Value
                                : element2.Attribute("TasaOCuota").Value;

                            cell27.CellValue = new CellValue();
                            cell27.DataType = CellValues.Number;
                            Cell cell28 = InsertCellInWorksheet("P", rowIndex, sheetData, worksheet);

                            var importe = element2.Attribute("importe") != null
                                ? element2.Attribute("importe").Value
                                : element2.Attribute("Importe").Value;

                            cell28.CellValue = new CellValue(importe);
                            cell28.DataType = CellValues.Number;
                        }
                    }

                    if (document2.Root.Element((XName)(namespace2 + "Impuestos")).Element((XName)(namespace2 + "Retenciones")) != null)
                    {
                        rowIndex++;
                        Cell cell29 = InsertCellInWorksheet("M", rowIndex, sheetData, worksheet);
                        cell29.CellValue = new CellValue("Impuestos Retenciones");
                        cell29.DataType = CellValues.String;
                        cell29.StyleIndex = 1;
                        Cell cell30 = InsertCellInWorksheet("N", rowIndex, sheetData, worksheet);
                        cell30.CellValue = new CellValue("Impuesto");
                        cell30.DataType = CellValues.String;
                        cell30.StyleIndex = 1;
                        Cell cell31 = InsertCellInWorksheet("O", rowIndex, sheetData, worksheet);
                        cell31.CellValue = new CellValue("Tasa");
                        cell31.DataType = CellValues.String;
                        cell31.StyleIndex = 1;
                        Cell cell32 = InsertCellInWorksheet("P", rowIndex, sheetData, worksheet);
                        cell32.CellValue = new CellValue("Importe");
                        cell32.DataType = CellValues.String;
                        cell32.StyleIndex = 1;
                        foreach (XElement element3 in document2.Root.Element((XName)(namespace2 + "Impuestos")).Element((XName)(namespace2 + "Retenciones")).Descendants())
                        {
                            rowIndex++;
                            Cell cell33 = InsertCellInWorksheet("N", rowIndex, sheetData, worksheet);
                            cell33.CellValue = new CellValue(element3.Attribute("impuesto").Value);
                            cell33.DataType = CellValues.String;
                            Cell cell34 = InsertCellInWorksheet("O", rowIndex, sheetData, worksheet);
                            cell34.CellValue = new CellValue((element3.Attribute("tasa") != null) ? element3.Attribute("tasa").Value : "");
                            cell34.DataType = CellValues.Number;
                            Cell cell35 = InsertCellInWorksheet("P", rowIndex, sheetData, worksheet);
                            cell35.CellValue = new CellValue(element3.Attribute("importe").Value);
                            cell35.DataType = CellValues.Number;
                        }
                    }
                }

                if(document2.Root.Element((XName)(namespace2 + "Complemento")).Element((XName)(namespace4 + "ImpuestosLocales")) != null)
                {

                    rowIndex++;
                    Cell cell29 = InsertCellInWorksheet("M", rowIndex, sheetData, worksheet);
                    cell29.CellValue = new CellValue("Impuestos Locales");
                    cell29.DataType = CellValues.String;
                    cell29.StyleIndex = 1;
                    Cell cell30 = InsertCellInWorksheet("N", rowIndex, sheetData, worksheet);
                    cell30.CellValue = new CellValue("Importe");
                    cell30.DataType = CellValues.String;
                    cell30.StyleIndex = 1;
                    Cell cell31 = InsertCellInWorksheet("O", rowIndex, sheetData, worksheet);
                    cell31.CellValue = new CellValue("Tasa de Traslado");
                    cell31.DataType = CellValues.String;
                    cell31.StyleIndex = 1;
                    Cell cell32 = InsertCellInWorksheet("P", rowIndex, sheetData, worksheet);
                    cell32.CellValue = new CellValue("Importe Local Trasladado");
                    cell32.DataType = CellValues.String;
                    cell32.StyleIndex = 1;

                    foreach (XElement element4 in document2.Root.Element((XName)(namespace2 + "Complemento")).Element((XName)(namespace4 + "ImpuestosLocales")).Descendants()){
                        rowIndex++;
                        Cell cell39 = InsertCellInWorksheet("N", rowIndex, sheetData, worksheet);
                        cell39.CellValue = new CellValue(element4.Attribute("Importe").Value);
                        cell39.DataType = CellValues.Number;
                        Cell cell40 = InsertCellInWorksheet("O", rowIndex, sheetData, worksheet);
                        cell40.CellValue = new CellValue(element4.Attribute("TasadeTraslado").Value);
                        cell40.DataType = CellValues.Number;
                        Cell cell41 = InsertCellInWorksheet("P", rowIndex, sheetData, worksheet);
                        cell41.CellValue = new CellValue(element4.Attribute("ImpLocTrasladado").Value);
                        cell41.DataType = CellValues.String;
                    }

                }

                rowIndex++;
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
                FileDownloadName = "factura-"+ DateTime.Now.Ticks +".xlsx"
            };
            return result1;
        }
        
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, SheetData sheetData, Worksheet worksheet)
        {
            Row row;
            string cellReference = columnName + rowIndex;
            if ((from r in sheetData.Elements<Row>()
                 where r.RowIndex == rowIndex
                 select r).Count<Row>() != 0)
            {
                row = (from r in sheetData.Elements<Row>()
                       where r.RowIndex == rowIndex
                       select r).First<Row>();
            }
            else
            {
                row = new Row
                {
                    RowIndex = rowIndex
                };
                OpenXmlElement[] newChildren = new OpenXmlElement[] { row };
                sheetData.Append(newChildren);
            }
            if ((from c in row.Elements<Cell>()
                 where c.CellReference.Value == (columnName + rowIndex)
                 select c).Count<Cell>() > 0)
            {
                return (from c in row.Elements<Cell>()
                        where c.CellReference.Value == cellReference
                        select c).First<Cell>();
            }
            Cell refChild = null;
            foreach (Cell cell3 in row.Elements<Cell>())
            {
                if (string.Compare(cell3.CellReference.Value, cellReference, true) > 0)
                {
                    refChild = cell3;
                    break;
                }
            }
            Cell newChild = new Cell
            {
                CellReference = cellReference
            };
            row.InsertBefore<Cell>(newChild, refChild);
            worksheet.Save();
            return newChild;
        }

        private Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
                new Fonts(
                    new Font(                                                               // Index 0 - The default font.
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 1 - The bold font.
                        new Bold(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 2 - The Italic font.
                        new Italic(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 2 - The Times Roman font. with 16 size
                        new FontSize() { Val = 16 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" })
                ),
                new Fills(
                    new Fill(                                                           // Index 0 - The default fill.
                        new PatternFill() { PatternType = PatternValues.None }),
                    new Fill(                                                           // Index 1 - The default fill of gray 125 (required)
                        new PatternFill() { PatternType = PatternValues.Gray125 }),
                    new Fill(                                                           // Index 2 - The yellow fill.
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }
                        )
                        { PatternType = PatternValues.Solid })
                ),
                new Borders(
                    new Border(                                                         // Index 0 - The default border.
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border(                                                         // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
                        new LeftBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Thin },
                        new RightBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Thin },
                        new TopBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Thin },
                        new BottomBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                ),
                new CellFormats(
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },                          // Index 0 - The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                    new CellFormat() { FontId = 1, FillId = 2, BorderId = 0, ApplyFont = true },       // Index 1 - Bold 
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 2 - Italic
                    new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 3 - Times Roman
                    new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Index 4 - Yellow Fill
                    new CellFormat(                                                                   // Index 5 - Alignment
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    )
                    { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }      // Index 6 - Border
                )
            ); // return
        }

    }
}
