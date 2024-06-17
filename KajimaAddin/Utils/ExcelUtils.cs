//using System;
//using System.Collections.Generic;
//using System.Diagnostics;
//using System.Text;
//using System.Threading.Tasks;
//using Excel = Microsoft.Office.Interop.Excel;
//using System.Reflection;
//using Microsoft.Office.Interop.Excel;
//using System.Linq;
//using TakeuchiAddins.Component;
//using TakeuchiAddins.Parameters;
//using System.IO;

//namespace TakeuchiAddins
//{
//     public static class ExcelUtils
//     {
//        public static void SaveToExcel(_Worksheet ws, CwGroup cwGroup, List<string> values)
//        {
//            if (cwGroup != null && cwGroup.CwDataGroup != null)
//            {
//                for (int i = 0; i < cwGroup.CwDataGroup.Count; i++)
//                {
//                    var cwData = cwGroup.CwDataGroup.ElementAt(i);
                    
//                    if (cwData != null)
//                    {
//                        int row = cwData.IniRow;
//                        int firstCol = cwData.IniCol;
//                        int checkRow = 1;
//                        foreach (var cwValue in cwData.CwDataVal)
//                        {
//                            if (!(string.IsNullOrWhiteSpace(cwGroup.Name)))
//                            {
//                                ws.Cells[row, cwGroup.CwCol] = cwGroup.Name;
//                            }
//                            foreach (var property in cwValue.GetType().GetProperties())
//                            {
//                                var pos = cwValue.ExcelMapping(property.Name);
//                                if (pos == null) continue;
//                                int rowOffset = pos.Value.Item1;
//                                int colOffset = pos.Value.Item2;
//                                if (rowOffset == 1)
//                                {
//                                    checkRow = 2;
//                                }
//                                int col = firstCol + colOffset;
//                                object valStr = property.GetValue(cwValue);
//                                if (valStr == null) continue;
//                                if (valStr.Equals("0"))
//                                {
//                                    ws.Cells[row + rowOffset, col] = "";
//                                }
//                                else
//                                {
//                                    int digit = 3;
//                                    if (values.Contains(property.Name))
//                                    {
//                                        digit = 2;
//                                    }
//                                    if (property.PropertyType == typeof(double))
//                                    {
//                                        if ((double)(property.GetValue(cwValue)) != 0)
//                                        {
//                                            ws.Cells[row + rowOffset, col] = Math.Round((double)property.GetValue(cwValue), digit);
//                                        }
//                                    }
//                                    else if (property.PropertyType == typeof(int))
//                                    {
//                                        if ((int)(property.GetValue(cwValue)) != 0)
//                                        {
//                                            ws.Cells[row + rowOffset, col] = (int)property.GetValue(cwValue);
//                                        }
//                                    }
//                                    else
//                                    {
//                                        ws.Cells[row + rowOffset, col] = property.GetValue(cwValue);
//                                    }
                                   
//                                }
                                
//                            }
//                            if (checkRow == 1)
//                            {
//                                row++;
//                            }
//                            else
//                            {
//                                row += 2;
//                            }
//                            checkRow = 1;
//                        }
//                    }
//                }
//            }
//            else
//            {
//                ws.Cells[cwGroup.CwRow, cwGroup.CwCol] = "";
//            }
//        }

//        public static void PrintHeading(_Worksheet ws, string[] heading, int iniRow, int iniCol)
//        {
//            int row = iniRow;
//            int col = iniCol;
//            for (int i = 0; i < heading.Length; i++)
//            {
                
//                ws.Cells[row, col] = heading[i];
//                col = col + 1;
//            }
//        }

//        public static void SaveListStringToExcel(_Worksheet ws, List<string> listStr)
//        {
//            int row = 1;
//            int col = 1;
//            foreach (var str in listStr)
//            {
//                var listOneRow = str.Split(';').ToList();
//                foreach (var oneRow in listOneRow)
//                {
//                    ws.Cells[row, col] = oneRow;
//                    col++;
//                }
//                row++;
//                col = 1;
//            }
//        }

//        public static void ColoringPlaxisExcel(_Worksheet ws, int iniRow, int iniCol, int step)
//        {
//            Excel.Range last = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
//            int lastUsedRow = last.Row;
//            int lastUsedCol = last.Column;
//            for (int i = iniRow; i <= lastUsedRow; i+=step)
//            {
//                for (int j = iniCol; j <= lastUsedCol; j+=3)
//                {
//                    Excel.Range curCell = ws.Cells[i, j];
//                    curCell.Interior.Color = System.Drawing.Color.FromArgb(253, 233, 217);
//                    curCell = ws.Cells[i, j + 1];
//                    curCell.Interior.Color = System.Drawing.Color.FromArgb(216, 228, 188);
//                    curCell = ws.Cells[i, j + 2];
//                    curCell.Interior.Color = System.Drawing.Color.FromArgb(183, 222, 232);
//                }
//            }
//        }

//        public static bool IsFileInUse(FileInfo file)
//        {
//            FileStream stream = null;
//            try
//            {
//                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
//            }
//            catch (IOException)
//            {
//                return true;
//            }
//            finally
//            {
//                if (stream != null)
//                    stream.Close();
//            }
//            return false;
//        }
//    }
//}
