using System;
using static System.Console;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

public interface ICell
{

    public int RowSpan { get; set; }
    public int ColSpan { get; set; }
    public void OnInit(Row row, int column);
}
public class TextCell : ICell
{
    public string Reference { get; set; }
    public string Value { get; set; }
    public int RowSpan { get; set; } = 1;
    public int ColSpan { get; set; } = 1;
    public CellValues DataType { get; set; } = CellValues.String;
    public virtual void OnInit(Row row, int columnNumber)
    {
        Cell refCell = null;
        Cell newCell = new Cell() { CellReference = this.Reference = (columnNumber.ToString() + ":" + row.RowIndex.ToString()) };
        row.InsertBefore(newCell, refCell);

        // Устанавливает тип значения.
        newCell.CellValue = new CellValue(Value);
        newCell.DataType = new EnumValue<CellValues>(DataType);
    }
}
public class FormCell : TextCell
{
    private string formulaText;

    public FormCell(string formulaText= "SUM(A1,C5)")
    {
        this.formulaText = formulaText;
    }

    public override void OnInit(Row row, int columnNumber)
    {
        Cell refCell = null;
        Cell cell = new Cell() { CellReference = this.Reference = (columnNumber.ToString() + ":" + row.RowIndex.ToString()) };
        row.InsertBefore(cell, refCell);

        // Устанавливает тип значения.
      
        CellFormula cellformula = new CellFormula();
        cellformula.Text = this.formulaText;
        CellValue cellValue = new CellValue();
        cellValue.Text = "0";
        cell.Append(cellformula);
        cell.Append(cellValue);
        //newCell.DataType = new EnumValue<CellValues>(type);
    }
}