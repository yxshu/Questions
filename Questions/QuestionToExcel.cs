using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Text.RegularExpressions;
namespace Questions
{
    class QuestionToExcel
    {
        public QuestionToExcel() { 
        }
        public bool PutQuestionToExcel(Question question, string path)
        {
            FileStream filestream = new FileStream(path, FileMode.Append);
            return  PutQuestionToExcel(question, filestream);
        }
        public bool PutQuestionToExcel(Question question, Stream stream)
        {
            bool mark = false;
            IWorkbook workbook = new HSSFWorkbook();//创建Workbook对象  
            ISheet sheet = workbook.CreateSheet("Sheet1");//创建工作表  
            IRow headerRow = sheet.CreateRow(0);//在工作表中添加首行  
            string[] headerRowName = new string[] { "rownumber", "ID", "SN", "章", "节", "试题", "选项A", "选项B", "选项C", "选项D", "答案", "解析","备注" };
            ICellStyle style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;//设置单元格的样式：水平对齐居中
            IFont font = workbook.CreateFont();//新建一个字体样式对象
            font.Boldweight = short.MaxValue;//设置字体加粗样式
            style.SetFont(font);//使用SetFont方法将字体样式添加到单元格样式中
            for (int i = 0; i < headerRowName.Length; i++)
            {
                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(headerRowName[i]);
                cell.CellStyle = style;
            }
            int rownumber = sheet.LastRowNum;
            IRow datarow = sheet.CreateRow(rownumber + 1);
            datarow.CreateCell(0).SetCellValue(rownumber + 1);
            datarow.CreateCell(1).SetCellValue(question.Id);
            datarow.CreateCell(2).SetCellValue(question.SN);
            datarow.CreateCell(3).SetCellValue(new Regex("[_]{3,10}").Replace(question.Chapter, "_______").Trim());
            datarow.CreateCell(4).SetCellValue(question.Node.Trim());
            datarow.CreateCell(5).SetCellValue(question.Title.Trim());
            datarow.CreateCell(6).SetCellValue(question.Choosea.Trim());
            datarow.CreateCell(7).SetCellValue(question.Chooseb.Trim());
            datarow.CreateCell(8).SetCellValue(question.Choosec.Trim());
            datarow.CreateCell(9).SetCellValue(question.Choosed.Trim());
            datarow.CreateCell(10).SetCellValue(question.Answer.Trim());
            datarow.CreateCell(11).SetCellValue(question.Explain.Trim());
            datarow.CreateCell(12).SetCellValue(question.Remark.Trim());
            for (int i = 0; i < headerRow.Cells.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            using (stream)
            {
                workbook.Write(stream);
                stream.Flush();
                stream.Close();
                mark = true;
            }
            return mark;

        }
    }
}
