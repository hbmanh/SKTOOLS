using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using BorderStyle = NPOI.SS.UserModel.BorderStyle;
using CellType = NPOI.SS.UserModel.CellType;
using Form = System.Windows.Forms.Form;

namespace RevitAddin.ExcelImporter
{
    [Transaction(TransactionMode.Manual)]
    public class ImportExcelCommand : IExternalCommand
    {
        private const double DEFAULT_SPACING_BETWEEN_TABLES = 3.28; // reserve (not used now)

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            if (!(uiDoc.ActiveView is ViewDrafting draftingView))
            {
                TaskDialog.Show("Warning", "Please open a Drafting View before running this command.");
                return Result.Cancelled;
            }

            ImportOptionsForm form = new ImportOptionsForm();
            if (form.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            string filePath = form.SelectedFilePath;
            var fileType = form.SelectedFileType;

            XYZ basePoint;
            try
            {
                basePoint = uiDoc.Selection.PickPoint("Pick a base point:");
            }
            catch
            {
                return Result.Cancelled;
            }

            using (Transaction tx = new Transaction(doc, "Import Excel Full"))
            {
                tx.Start();

                if (fileType == FileType.Excel)
                {
                    var tableData = ReadExcel(filePath);
                    DrawTable(doc, draftingView, tableData, basePoint);
                }
                else if (fileType == FileType.PDF)
                {
                    TaskDialog.Show("Notice", "Import PDF is currently not supported. Please select an Excel file.");
                }

                tx.Commit();
            }

            return Result.Succeeded;
        }

        private TableData ReadExcel(string filePath)
        {
            var tableData = new TableData();

            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(stream);
                ISheet sheet = workbook.GetSheetAt(0);

                tableData.MergedRegions = new List<CellRangeAddress>();
                for (int i = 0; i < sheet.NumMergedRegions; i++)
                {
                    tableData.MergedRegions.Add(sheet.GetMergedRegion(i));
                }

                for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row == null) continue;

                    for (int colIndex = row.FirstCellNum; colIndex < row.LastCellNum; colIndex++)
                    {
                        ICell cell = row.GetCell(colIndex);
                        if (cell == null) continue;

                        var style = cell.CellStyle;
                        var font = workbook.GetFontAt(style.FontIndex);

                        CellInfo info = new CellInfo
                        {
                            Row = rowIndex,
                            Column = colIndex,
                            Text = GetCellValue(cell),
                            FontSize = font.FontHeightInPoints,
                            IsBold = font.IsBold,
                            IsItalic = font.IsItalic,
                            Alignment = style.Alignment.ToString(),
                            RowHeight = row.Height > 0 ? row.Height / 20.0 : 15,
                            ColumnWidth = sheet.GetColumnWidth(colIndex) / 256.0,
                            HasTopBorder = style.BorderTop != BorderStyle.None,
                            HasBottomBorder = style.BorderBottom != BorderStyle.None,
                            HasLeftBorder = style.BorderLeft != BorderStyle.None,
                            HasRightBorder = style.BorderRight != BorderStyle.None
                        };
                        tableData.Cells.Add(info);
                    }
                }
            }

            return tableData;
        }

        private string GetCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    return cell.ToString();
                default:
                    return string.Empty;
            }
        }

        private void DrawTable(Document doc, ViewDrafting view, TableData data, XYZ basePoint)
        {
            Dictionary<(int, int), CellInfo> cellMap = data.Cells.ToDictionary(c => (c.Row, c.Column));

            foreach (var cell in data.Cells)
            {
                if (IsPartOfMergedRegion(cell.Row, cell.Column, data.MergedRegions, out var region))
                {
                    if (cell.Row != region.FirstRow || cell.Column != region.FirstColumn)
                        continue;
                }

                double x1 = basePoint.X + GetOffsetX(data, cell.Column);
                double y1 = basePoint.Y - GetOffsetY(data, cell.Row);
                double width = GetWidth(data, cell.Column, region);
                double height = GetHeight(data, cell.Row, region);

                double x2 = x1 + width;
                double y2 = y1 - height;

                DrawBorders(doc, view, x1, y1, x2, y2, cell);

                if (!string.IsNullOrEmpty(cell.Text))
                {
                    XYZ textPos = new XYZ((x1 + x2) / 2, (y1 + y2) / 2, 0);
                    double textSize = (cell.FontSize / 72.0) / 12.0;
                    CreateText(doc, view, textPos, cell, textSize);
                }
            }
        }

        private bool IsPartOfMergedRegion(int row, int col, List<CellRangeAddress> regions, out CellRangeAddress regionFound)
        {
            foreach (var region in regions)
            {
                if (row >= region.FirstRow && row <= region.LastRow &&
                    col >= region.FirstColumn && col <= region.LastColumn)
                {
                    regionFound = region;
                    return true;
                }
            }
            regionFound = null;
            return false;
        }

        private double GetOffsetX(TableData data, int col)
        {
            double total = 0;
            for (int i = 0; i < col; i++)
            {
                total += (data.Cells.FirstOrDefault(c => c.Column == i)?.ColumnWidth ?? 7.5) * 7.0017 / 72.0 / 12.0;
            }
            return total;
        }

        private double GetOffsetY(TableData data, int row)
        {
            double total = 0;
            for (int i = 0; i < row; i++)
            {
                total += (data.Cells.FirstOrDefault(c => c.Row == i)?.RowHeight ?? 15) / 72.0 / 12.0;
            }
            return total;
        }

        private double GetWidth(TableData data, int startCol, CellRangeAddress region)
        {
            double width = 0;
            if (region == null)
            {
                width = (data.Cells.FirstOrDefault(c => c.Column == startCol)?.ColumnWidth ?? 7.5) * 7.0017 / 72.0 / 12.0;
            }
            else
            {
                for (int col = region.FirstColumn; col <= region.LastColumn; col++)
                {
                    width += (data.Cells.FirstOrDefault(c => c.Column == col)?.ColumnWidth ?? 7.5) * 7.0017 / 72.0 / 12.0;
                }
            }
            return width;
        }

        private double GetHeight(TableData data, int startRow, CellRangeAddress region)
        {
            double height = 0;
            if (region == null)
            {
                height = (data.Cells.FirstOrDefault(c => c.Row == startRow)?.RowHeight ?? 15) / 72.0 / 12.0;
            }
            else
            {
                for (int row = region.FirstRow; row <= region.LastRow; row++)
                {
                    height += (data.Cells.FirstOrDefault(c => c.Row == row)?.RowHeight ?? 15) / 72.0 / 12.0;
                }
            }
            return height;
        }

        private void DrawBorders(Document doc, ViewDrafting view, double x1, double y1, double x2, double y2, CellInfo cell)
        {
            if (cell.HasTopBorder)
                DrawLine(doc, view, new XYZ(x1, y1, 0), new XYZ(x2, y1, 0));
            if (cell.HasBottomBorder)
                DrawLine(doc, view, new XYZ(x1, y2, 0), new XYZ(x2, y2, 0));
            if (cell.HasLeftBorder)
                DrawLine(doc, view, new XYZ(x1, y1, 0), new XYZ(x1, y2, 0));
            if (cell.HasRightBorder)
                DrawLine(doc, view, new XYZ(x2, y1, 0), new XYZ(x2, y2, 0));
        }

        private void DrawLine(Document doc, ViewDrafting view, XYZ start, XYZ end)
        {
            Line line = Line.CreateBound(start, end);
            doc.Create.NewDetailCurve(view, line);
        }

        private void CreateText(Document doc, ViewDrafting view, XYZ position, CellInfo cell, double textSize)
        {
            var textTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);

            TextNoteOptions opts = new TextNoteOptions(textTypeId)
            {
                HorizontalAlignment = ParseAlignment(cell.Alignment)
            };

            TextNote note = TextNote.Create(doc, view.Id, position, cell.Text, opts);

            // ⚡ Fix lỗi NullReference:
            Parameter textSizeParam = note.get_Parameter(BuiltInParameter.TEXT_SIZE);
            if (textSizeParam != null)
            {
                textSizeParam.Set(textSize);
            }
        }

        private HorizontalTextAlignment ParseAlignment(string align)
        {
            if (align == "Center") return HorizontalTextAlignment.Center;
            if (align == "Right") return HorizontalTextAlignment.Right;
            return HorizontalTextAlignment.Left;
        }
    }

    public class TableData
    {
        public List<CellInfo> Cells = new List<CellInfo>();
        public List<CellRangeAddress> MergedRegions = new List<CellRangeAddress>();
    }

    public class CellInfo
    {
        public int Row;
        public int Column;
        public string Text;
        public double FontSize;
        public bool IsBold;
        public bool IsItalic;
        public string Alignment;
        public double RowHeight;
        public double ColumnWidth;
        public bool HasTopBorder;
        public bool HasBottomBorder;
        public bool HasLeftBorder;
        public bool HasRightBorder;
    }

    public class ImportOptionsForm : Form
    {
        public string SelectedFilePath { get; private set; }
        public FileType SelectedFileType { get; private set; }

        public ImportOptionsForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Import Options";
            this.Width = 400;
            this.Height = 200;

            Button excelButton = new Button() { Text = "Select Excel File", Left = 50, Width = 120, Top = 30 };
            Button pdfButton = new Button() { Text = "Select PDF File", Left = 200, Width = 120, Top = 30 };
            Button okButton = new Button() { Text = "OK", Left = 100, Width = 80, Top = 100 };
            Button cancelButton = new Button() { Text = "Cancel", Left = 200, Width = 80, Top = 100 };

            excelButton.Click += ExcelButton_Click;
            pdfButton.Click += PdfButton_Click;
            okButton.Click += OkButton_Click;
            cancelButton.Click += CancelButton_Click;

            this.Controls.Add(excelButton);
            this.Controls.Add(pdfButton);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);
        }

        private void ExcelButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                SelectedFilePath = openFileDialog.FileName;
                SelectedFileType = FileType.Excel;
            }
        }

        private void PdfButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF Files|*.pdf",
                Title = "Select a PDF File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                SelectedFilePath = openFileDialog.FileName;
                SelectedFileType = FileType.PDF;
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedFilePath))
            {
                MessageBox.Show("Please select a file before clicking OK.");
                return;
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }

    public enum FileType
    {
        Excel,
        PDF
    }
}
