using System;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace l1ExpData
{
    class Excel
    {
        string parh = "";
        int sheet;
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel()
        {

        }
        public Excel(string parh, int sheet)
        {
            this.parh = parh;
            wb = excel.Workbooks.Open(parh);
            ws = wb.Worksheets[sheet];
        }

        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        public void Close()
        {
            wb.Close();
        }
        public void CreateNewSheet()
        {
            Worksheet tempSheet = wb.Worksheets.Add(After: ws);
        }
        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
        }

        public void SelectWorksheet(int sheetNumber)
        {
            this.ws = wb.Worksheets[sheetNumber];
        }

        public void DeleatWorksheet(int sheetNumber)
        {
            wb.Worksheets[sheetNumber].Delete();
        }

        public double[,] ReadRange(int startx, int starty, int endx,int endy)
        {
            Range range = (Range)ws.Range[ws.Cells[startx, starty], ws.Cells[endx, endy]];
            object[,] holder = range.Value2;
            var arrdouble = new double[endx-startx+1,endy-starty+1];
            int xret = 0;
            for (int x = 1; x <= endx - startx+1; xret++, x++)
            {
                for (int y = 1; y <= endy - starty+1; y++)
                {
                    if ((Convert.ToString(holder[x, y]) == "")|| (Convert.ToString(holder[x, y]) == "-")|| (Convert.ToString(holder[x, y]) == "..."))
                    {
                        xret--;
                        break;
                    }
                    else
                    arrdouble[xret, y-1] = Convert.ToDouble( holder[x, y]);
                }
            }
            var returnDouble = new double[xret, endy - starty + 1];
            for(int i = 0; i < xret; i++)
            {
                for(int k = 0; k < endy - starty + 1; k++)
                {
                    returnDouble[i, k] = arrdouble[i, k];
                }
            }
            return returnDouble;
        }

        public double ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2;
            else
                return 0;
        }

        public void WriteToCellString(int i, int j, string s)
        {
            ws.Cells[i, j].Value2 = s;
        }

        public void WriteToCellString(int i, int j, double s)
        {
            ws.Cells[i, j].Value2 = s;
        }

        public void WriteRange(int startx, int starty, int endx, int endy, double [,]writeFloat)
        {
            Range range = (Range)ws.Range[ws.Cells[startx, starty], ws.Cells[endx, endy]];
            range.Value2 = writeFloat;
        }

        public void WriteRange(int startx, int starty, int endy, double[] writeFloat)
        {
            Range range = (Range)ws.Range[ws.Cells[startx, starty], ws.Cells[startx, endy]];
            range.Value2 = writeFloat;
        }
        public void WriteRange(int startx, int starty, int endy, int[] writeFloat)
        {
            Range range = (Range)ws.Range[ws.Cells[startx, starty], ws.Cells[startx, endy]];
            range.Value2 = writeFloat;
        }
    }
}
