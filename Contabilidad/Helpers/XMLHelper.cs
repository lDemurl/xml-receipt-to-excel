using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Contabilidad.Helpers
{
    public static class XMLHelper
    {
        public static uint RowIndex { get; set; }

        public static void Init()
        {
            RowIndex = 1;
        }

        public static Cell InsertCellInWorksheet(string columnName, uint rowIndex, SheetData sheetData, Worksheet worksheet)
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
    }
}