using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace StudyStatistic
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel(string path,int Sheet) 
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }
        public _Excel.Workbook GetWorkbook() 
        {
            return wb;
        }
        //在路径上创造一个新的Excel文件
        public Excel(string path) 
        {
            wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            ws = wb.Worksheets[1];
            wb.SaveAs(path);
        }
        public _Excel.Worksheet GetWorksheet() 
        {
            return ws;
        }
        //对于csv来说，这种Sort会改变其数值
        public void Sort(string colCharacter) 
        {
            //用了几行
            int row = ws.UsedRange.Rows.Count;
            string srow = row.ToString();
            string range1 = colCharacter + "2";
            string range2 = colCharacter + srow;
            ws.Sort.SetRange(ws.Range[range1,range2]);
            ws.Sort.SortFields.Add(ws.Range[range1,range2], _Excel.XlSortOn.xlSortOnValues, _Excel.XlSortOrder.xlAscending,_Excel.XlSortMethod.xlPinYin);
            ws.Sort.Apply();
        }
        public string ReadCell(int row,int col) 
        {
            row++;
            col++;

            if (ws.Cells[row, col].Value2!= null) 
            {
                return ws.Cells[row, col].Value2;
            }
            else
            {
                return "";
            }
        }
        public void WriteCell(int row,int col,string words) 
        {
            row++;
            col++;

            ws.Cells[row, col].Value2 = words;
        }
        public void save() 
        {
            wb.Saved = true;
        }
        public void exit()
        {
            excel.DisplayAlerts = false;
            wb.Save();
            wb.Close(true);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);


        }
        public void addSheet(string name) 
        {
            _Excel.Worksheet newWorksheet;
            newWorksheet = wb.Worksheets.Add(After:ws);
            newWorksheet.Name = name;
        }

        //把当前excel的Id和Name读到字典里面
        internal void ReadIdAndName(Dictionary<string, string> excel1IdAndName)
        {
            int colnum = ws.UsedRange.Columns.Count;
            int rownum = ws.UsedRange.Rows.Count;
            int Idcolindex = 0;
            int namecolIndex = 0;
            //遍历第一行
            for(int i = 0; i < colnum; i++) 
            {
                string header = ws.Cells[1, i + 1].Value2;
                if (header.Equals("员工编号") || header.Equals("员工编码"))
                    Idcolindex = i + 1;//得到Id那一列的index
                else if (header.Equals("姓名"))
                    namecolIndex = i + 1;//得到姓名那一列的index
            }
            string id;
            string name;
            //由于给我的格式，员工编码和姓名一定是连着的，因此可以直接得到两列数据
            //但是有两种情况，一种是员工编码在左，姓名编码在右，一种是员工编码在右，姓名编码在左
            if (Idcolindex > namecolIndex) 
            {
                //员工编码在右
                _Excel.Range IdAndNameRange = ws.Range[ws.Cells[2, namecolIndex], ws.Cells[rownum, Idcolindex]];
                object[,] idAndName = IdAndNameRange.Value2;

                for (int i = 0; i < rownum - 1; i++)
                {
                    id = idAndName[i + 1, 2].ToString().Replace("\t", String.Empty);
                    name = idAndName[i + 1, 1].ToString().Replace("\t", String.Empty);
                    excel1IdAndName.Add(id,name);
                }
            }
            else
            {
                //员工编码在左
                _Excel.Range IdAndNameRange = ws.Range[ws.Cells[2, Idcolindex], ws.Cells[rownum, namecolIndex]];
                object[,] idAndName = IdAndNameRange.Value2;
                for(int i = 0; i < rownum-1; i++) 
                {
                    id = idAndName[i + 1, 1].ToString().Replace("\t", String.Empty);
                    name = idAndName[i + 1, 2].ToString().Replace("\t", String.Empty);
                    excel1IdAndName.Add(id, name);
                }
            }


            
        }

        public void fillworksheet(ArrayList res) 
        {
            //因为有标题头，都得从第二行第二列开始写
            int i = 1;
            int j = 1;
            foreach(Empolyee employee in res) 
            {
                WriteCell(i , j , employee.Id);
                WriteCell(i , j + 1, employee.Name);
                i++;
                
            }
        }
    }
}
