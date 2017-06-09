using System;
using System.IO;
//using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using NPOI.SS.UserModel;
//using NPOI.HSSF.UserModel;
using System.Reflection;

namespace Questions
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] documents = new string[] { "QuestionLibraries/equipment-hedetao.docx", "QuestionLibraries/instruction-yuxiangshu.docx", "QuestionLibraries/navigation-hedetao.docx", "QuestionLibraries/ocean-hedetao.docx", "QuestionLibraries/avoidcollision-wufei.docx", "QuestionLibraries/management-lizhite.docx", "QuestionLibraries/certificate-yuxiangshu.docx", "QuestionLibraries/english-xiangwei.docx" };
            string[] subjects = new string[] { "航海学(航海仪器)", "船舶结构与货运", "航海学(航海地文、天文)", "航海学(航海气象与海洋学)", "船舶操纵与避碰", "船舶管理", "海船船员合格证培训", "航海英语" };
            bool expstar = false;//解析开始标记
            string subject = string.Empty;
            string chapter = string.Empty;//章标题
            int chapterID = 0;//章序号
            string node = string.Empty;//节标题
            int nodeID = 0;//节序号
            int questionID = 0;//试题序号
            int questionAllID = 0;//试题的总序号
            Regex regA = new Regex("[ABCDabcd]{1}[\\.|、]", RegexOptions.IgnoreCase);//A|B|C|D
            Regex regNO = new Regex("^[0-9]+[\\.|、]", RegexOptions.IgnoreCase);//以数字开头  题干
            Regex regexpstar = new Regex("^参考答案|答案解析");//参考答案开头
            Regex regexp = new Regex("^[0-9]+[\\.|、][ABCDabcd]{1}[\\.|、|。]?");//解释
            Regex regChapter = new Regex("^第[一二三四五六七八九十]{1,3}章", RegexOptions.IgnoreCase);//章标题
            Regex regNode = new Regex("^第[一二三四五六七八九十]{1,3}节", RegexOptions.IgnoreCase);//节标题
            Regex regxhx = new Regex("[_]{3,10}", RegexOptions.IgnoreCase);//下划线
            StreamWriter writer = new StreamWriter("D://error.txt", true, System.Text.Encoding.Default, 1 * 1024);
            List<Question> list = new List<Question>();
            for (int j = 0; j < documents.Length; j++) //(string str in documents)
            {
                string str = documents[j];
                subject = subjects[j];
                string initpath = @"C:\Users\yxshu\Documents\GitHub\Questions\";
                string path = initpath + str;//C:\Users\yxshu\Documents\GitHub\Questions
                Console.WriteLine(path);
                writer.WriteLine(path);
                try
                {
                    Word.Application app = new Word.Application();
                    Word.Document doc = null;
                    object unknow = Type.Missing;
                    app.Visible = true;
                    object file = path;
                    doc = app.Documents.Open(ref file,
                        ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                        ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                        ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);
                    int paragraphsCount = doc.Paragraphs.Count;
                    for (int i = 1; i <= paragraphsCount; i++)
                    // for (int i = 1; i < 100; i++)
                    {
                        Word.Range para = doc.Paragraphs[i].Range;
                        para.Select();
                        //string text = para.Text.Trim();
                        string text = new Regex("\\r\\a").Replace(para.Text, "").Trim();
                        if (string.IsNullOrEmpty(text)) continue;
                        if (regChapter.IsMatch(text))//章标题
                        {
                            expstar = false;
                            chapter = text;
                            chapterID++;
                            nodeID = 0;
                            questionID = 0;
                            Console.WriteLine("章标题： " + chapter);
                        }
                        else if (regNode.IsMatch(text))//节标题
                        {
                            nodeID++;
                            questionID = 0;
                            expstar = false;
                            node = text;
                            Console.WriteLine("节标题： " + node);
                        }
                        else if (regexpstar.IsMatch(text))//参考答案开始
                        {
                            expstar = true;
                            Console.WriteLine("参考答案开始标记： " + text);

                        }
                        else if (regNO.IsMatch(text))//数字开头的
                        {
                            if (regxhx.IsMatch(text))//有下划线的
                            {
                                if (regA.Split(text).Length == 5)//有四个选项
                                {
                                    string[] strSplit = regA.Split(text);
                                    questionID++;
                                    questionAllID++;
                                    Question question = new Question();
                                    question.Subject = subject;
                                    question.Chapter = chapter;
                                    question.Node = node;
                                    question.AllID = questionAllID;
                                    question.Id = chapterID + "_" + nodeID + "_" + questionID;
                                    question.SN = Int32.Parse(new Regex("^[0-9]+", RegexOptions.IgnoreCase).Match(strSplit[0]).Value);
                                    question.SNID = chapterID + "_" + nodeID + "_" + question.SN;
                                    string title = regxhx.Replace(regNO.Replace(strSplit[0], ""), "_______");
                                    if (regxhx.IsMatch(title))
                                        question.Title = title;
                                    else { Console.WriteLine("题干部分无下划线。"); Console.ReadLine(); }
                                    question.Choosea = strSplit[1];
                                    question.Chooseb = strSplit[2];
                                    question.Choosec = strSplit[3];
                                    question.Choosed = strSplit[4];
                                    question.Answer = null;
                                    question.Explain = null;
                                    printQuesiton(question);
                                    list.Add(question);

                                }
                                else if (regA.Split(text).Length == 4)//有三个选项
                                {
                                    Console.WriteLine("三个选项试题： " + text);
                                    Console.ReadLine();
                                }
                                else//其它， 不知道是什么情况，有可能是判断题
                                {
                                    // writer.WriteLine("其他：数字开头，不是三个/四个选项- " + regA.Split(text).Length + "_" + text);
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("其他：数字开头，不是三个/四个选项- " + regA.Split(text).Length + "_" + text);
                                    Console.ReadLine();
                                }
                            }
                            else//无下划线
                            {
                                if (regexp.IsMatch(text) && expstar)//参考答案与解析
                                {

                                    string tou = regexp.Match(text).Value;//编号和答案
                                    string exp = text.Substring(tou.Length).Trim();//解析
                                    Regex reg = new Regex(@"同第[0-9]+题\p{P}");
                                    if (reg.IsMatch(exp))//解决关于“同第**题”的解析问题
                                    {
                                        int sameNo = Int32.Parse(new Regex("[0-9]+").Match(reg.Match(exp).Value).Value);
                                        foreach (Question q in list)
                                        {
                                            if (q.SNID == chapterID + "_" + nodeID + "_" + sameNo)
                                            {
                                                exp = reg.Replace(exp, q.Explain);
                                                break;
                                            }
                                        }
                                    }
                                    string No = chapterID + "_" + nodeID + "_" + new Regex("^[0-9]+", RegexOptions.IgnoreCase).Match(tou).Value;//试题编号（带章节）
                                    string answer = new Regex("[ABCD]{1}", RegexOptions.IgnoreCase).Match(tou).Value;//试题答案
                                    foreach (Question q in list)
                                    {
                                        if (q.SNID == No)
                                        {
                                            if (string.IsNullOrEmpty(q.Answer))
                                            {
                                                q.Answer = answer;
                                            }
                                            else if (q.Answer != answer)
                                            {
                                                Console.WriteLine("试题参考答案与解析部分答案不同。");
                                                Console.ReadLine();
                                            }
                                            q.Explain = exp;
                                            Console.WriteLine("头子：" + tou);
                                            printQuesiton(q);
                                            break;
                                        }
                                    }
                                    //Console.WriteLine("参考答案: " + text);
                                }
                                else//错误部分
                                {
                                    //writer.WriteLine("错误：数字开头无下划线 - " + text);
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("错误：数字开头无下划线 - " + text);
                                    Console.ReadLine();
                                }

                            }
                        }
                        else
                        {
                            // writer.WriteLine("错误：非章节标题，非数字开头，你是个什么鬼- " + text);
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("错误：非章节标题，非数字开头，你是个什么鬼- " + text);
                            Console.ReadLine();
                        }
                        Console.ResetColor();
                        Console.WriteLine();
                        //Thread.Sleep(1000);//每一个段落结束
                    }//一本试题结束

                    app.Documents.Close();
                    Thread.Sleep(5000);
                    Console.WriteLine("开始将试题写入到表格中……");
                    questiontoexcel(list, "d://" + str + ".xls");
                    Console.WriteLine("试题写入完成，地址：D://{0}.xls", str);
                }
                catch (Exception ex)
                {
                    writer.WriteLine(ex.Message);
                    Console.WriteLine(ex.Message);
                    Console.ReadLine();
                }
                finally
                {
                    writer.Flush();
                }
            }//所有试题结束
            writer.Close();
            Console.WriteLine("所有写入完成。");
            Console.ReadLine();
        }

        /// <summary>
        /// 打印出试题对象
        /// </summary>
        /// <param name="question">试题对象</param>
        public static void printQuesiton(Question question)
        {
            Console.WriteLine("总序号：" + question.AllID);
            Console.WriteLine("试题编号：" + question.Id);
            Console.WriteLine("原编号：" + question.SN);
            Console.WriteLine("章节+原编号:" + question.SNID);
            Console.WriteLine("科目：" + question.Subject);
            Console.WriteLine("章标题：" + question.Chapter);
            Console.WriteLine("节标题：" + question.Node);
            Console.WriteLine("试题：" + question.Title);
            Console.WriteLine("选项A:" + question.Choosea);
            Console.WriteLine("选项B:" + question.Chooseb);
            Console.WriteLine("选项C:" + question.Choosec);
            Console.WriteLine("选项D:" + question.Choosed);
            Console.WriteLine("参考答案：" + question.Answer);
            Console.WriteLine("解析：" + question.Explain);
            Console.WriteLine();
        }


        /// <summary>
        /// 将试题填充到EXCEL中
        /// </summary>
        /// <param name="list"></param>
        /// <param name="path"></param>
        public static void questiontoexcel(List<Question> list, string path)
        {
            FileStream filestream = new FileStream(path, FileMode.Append);
            Question question = new Question();
            IWorkbook workbook = new NPOI.HSSF.UserModel.HSSFWorkbook();//创建Workbook对象  
            ISheet sheet = workbook.CreateSheet("Sheet1");//创建工作表  
            IRow headerRow = sheet.CreateRow(0);//在工作表中添加首行 
            PropertyInfo[] propertyinfo = question.GetType().GetProperties();
            string[] headerRowName = new string[propertyinfo.Length];
            for (int i = 0; i < propertyinfo.Length; i++)
            {
                headerRowName[i] = propertyinfo[i].Name;
            }
            ICellStyle style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;//设置单元格的样式：水平对齐居中
            IFont font = workbook.CreateFont();//新建一个字体样式对象
            font.Boldweight = short.MaxValue;//设置字体加粗样式
            style.SetFont(font);//使用SetFont方法将字体样式添加到单元格样式中
            for (int i = 0; i < headerRowName.Length; i++)
            {
                NPOI.SS.UserModel.ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(headerRowName[i]);
                cell.CellStyle = style;
            }
            foreach (Question q in list)
            {
                int rownumber = sheet.LastRowNum;
                IRow datarow = sheet.CreateRow(rownumber + 1);
                for (int i = 0; i < q.GetType().GetProperties().Length; i++)
                {
                    ICell c = datarow.CreateCell(i);
                    c.SetCellValue("");
                    if (q.GetType().GetProperties()[i].GetValue(q) != null)
                    {
                        c.SetCellValue(q.GetType().GetProperties()[i].GetValue(q).ToString());
                    }
                }
            }
            for (int i = 0; i < headerRow.Cells.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            using (filestream)
            {
                workbook.Write(filestream);
                filestream.Flush();
                filestream.Close();
            }
        }
    }
}