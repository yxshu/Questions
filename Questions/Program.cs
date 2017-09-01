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
using System.Data.SqlClient;
using System.Text;


namespace Questions
{
    class Program
    {
        static Regex regA = new Regex("[ABCDabcd]{1}[\\.|、]", RegexOptions.IgnoreCase);//A|B|C|D
        static Regex regNO = new Regex("^[0-9]+[\\.|、]", RegexOptions.IgnoreCase);//以数字开头  题干
        static Regex regxhx = new Regex("[_]{3,10}", RegexOptions.IgnoreCase);//下划线
        static void Main(string[] args)
        {
            string[] documents = new string[] {/* "QuestionLibraries/equipment-hedetao.docx", "QuestionLibraries/ocean-hedetao.docx", "QuestionLibraries/certificate-yuxiangshu.docx", "QuestionLibraries/navigation-hedetao.docx", "QuestionLibraries/management-lizhite.docx", "QuestionLibraries/instruction-yuxiangshu.docx",*/ "QuestionLibraries/english-xiangwei.docx"/*, "QuestionLibraries/avoidcollision-wufei.docx"*/ };
            string[] subjects = new string[] { /*"航海学(航海仪器)", "航海学(航海气象与海洋学)", "海船船员培训合格证", "航海学(航海地文、天文)", "船舶管理", "船舶结构与货运", */"航海英语"/*, "船舶操纵与避碰" */};
            bool expstar = false;//解析开始标记
            string subject = string.Empty;
            string chapter = string.Empty;//章标题
            string node = string.Empty;//节标题
            string gltmark = "$关联题$";//关联题的标记，放在question.remark的开头，用于后期识别

            int numForAll = 0, questionAllID = 0, questionID = 0, chapterID = 0, nodeID = 0;//所有试题的总序号，每本试题的总序号，每章/节内的序号，章序号，节序号
            int[] numForDocument = new int[documents.Length];

            Regex regexpstar = new Regex("^参考答案|答案解析");//参考答案开头
            Regex regexp = new Regex("^[0-9]+[\\.|、][ABCDabcd]{1}[\\.|、|。]?");//解释
            Regex regPain = new Regex("^第[一二三四五六七八九十]{1,3}篇", RegexOptions.IgnoreCase);//篇标题
            Regex regChapter = new Regex("^第[一二三四五六七八九十]{1,3}章", RegexOptions.IgnoreCase);//章标题
            Regex regNode = new Regex("^第[一二三四五六七八九十]{1,3}节", RegexOptions.IgnoreCase);//节标题

            Regex regglt = new Regex(@"^passage\s*[0-9]{1,4}");//关联题，以passage开头+数字
            string initpath = @"C:\Users\yxshu\Documents\GitHub\Questions\";
            StreamWriter writer = new StreamWriter("D://error.txt", true, System.Text.Encoding.Default, 1 * 1024);
            StreamWriter log = new StreamWriter(initpath + "\\Questions\\log.txt", true, System.Text.Encoding.Default);
            List<Question> list = new List<Question>();

            DateTime[] timestar = new DateTime[documents.Length];//记录每一科的开始时间
            DateTime[] timeend = new DateTime[documents.Length];//记录每一科的结束时间

            for (int j = 0; j < documents.Length; j++) //(string str in documents)
            {
                timestar[j] = DateTime.Now;
                questionAllID = 0;
                questionID = 0;
                chapterID = 0;
                nodeID = 0;
                chapter = string.Empty;
                node = string.Empty;
                string str = documents[j];
                if (!string.IsNullOrEmpty(documents[j]))
                {
                    subject = subjects[j];
                }
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
                    Console.WriteLine("正在加载文件： {0} ", path);
                    Thread.Sleep(1000);

                    int paragraphsCount = doc.Paragraphs.Count;

                    //从这里开始检测每一段的内容
                    for (int i = 1; i <= paragraphsCount; i++)//
                    //for (int i = 1; i < 200; i++)
                    {
                        Word.Range para = doc.Paragraphs[i].Range;
                        para.Select();
                        //string text = para.Text.Trim();
                        string text = new Regex("\\r\\a").Replace(para.Text, "").Trim();
                        /***
                         * 对每一段内容进行如下检测
                         * 1、是否空行；
                         * 2、是否是章标题             第？章 开头的标识
                         * 3、是否是节标题             第？节 开头的标识
                         * 4、是否是参考答案开始标识   参考答案|答案解析  开头的标识
                         * 5、关联题检测               以passage 数字  开关的标识
                         * 5、是否是数字开头           数字.|数字、   开头的标识
                         * 6、其他情况
                         * 
                         * ****/
                        if (string.IsNullOrEmpty(text)) continue;//空行退出
                        if (regPain.IsMatch(text))//处理第？篇的问题
                        {
                            expstar = false;
                            chapter = string.Empty;
                            node = string.Empty;
                            chapterID = 0;
                            nodeID = 0;
                            questionID = 0;
                            subject = regPain.Replace(text, "").Trim();
                        }
                        else if (regChapter.IsMatch(text))//章标题
                        {
                            expstar = false;
                            chapter = text;
                            node = string.Empty;
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

                        #region  处理关联题 PASSAGE 1
                        else if (regglt.IsMatch(text.ToLower()))
                        {
                            List<int> questionrow = new List<int>();//用来装试题行的行号
                            for (int k = i + 1; ; k++)
                            {
                                doc.Paragraphs[k].Range.Select();
                                string newtext = new Regex("\\r\\a").Replace(doc.Paragraphs[k].Range.Text, "").Trim();
                                if (string.IsNullOrEmpty(newtext)) continue;//空行退出
                                if (is4ChooseQuestion(newtext))//是符合条件的选项
                                {
                                    questionrow.Add(k);
                                }
                                else if (regglt.IsMatch(newtext.ToLower()) || regexp.IsMatch(newtext))//到达下一个关联题部分，本次关联题结束
                                {
                                    if (regexp.IsMatch(newtext))
                                    {
                                        expstar = true;
                                    }
                                    Question question = new Question();
                                    string remark = string.Empty;
                                    StringBuilder sb = new StringBuilder();
                                    for (int x = i + 1; x < questionrow[0] - 1; x++)
                                    {
                                        sb.AppendLine(new Regex("\\r\\a").Replace(doc.Paragraphs[x].Range.Text, "").Trim());
                                    }
                                    TimeSpan ts = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);//给关联题加上一个时间戳
                                    remark = gltmark + "&" + Convert.ToInt64(ts.TotalMilliseconds).ToString() + "&" + sb.ToString();
                                    foreach (int n in questionrow)
                                    {
                                        questionID++;
                                        questionAllID++;
                                        question = TextToQuestionModel(subject, chapter, node, questionAllID, questionID, chapterID, nodeID, remark, new Regex("^[0-9]+", RegexOptions.IgnoreCase).Replace(doc.Paragraphs[n].Range.Text, questionID.ToString()));
                                        printQuesiton(question);
                                        list.Add(question);
                                    }
                                    i = k - 1;
                                    break;
                                }
                            }
                        }
                        #endregion

                        else if (regNO.IsMatch(text))//数字开头的
                        {
                            #region 数字开头并且有下划线的，分别检测四个选项，三个选项，其他情况
                            if (regxhx.IsMatch(text))//有下划线的
                            {
                                if (regA.Split(text).Length == 5)//有四个选项
                                {
                                    proceduQuestion(subject, chapter, node, ref questionAllID, ref questionID, chapterID, nodeID, list, text);
                                }
                                else if (regA.Split(text).Length == 4)//有三个选项
                                {
                                    text += "D、";
                                    proceduQuestion(subject, chapter, node, ref questionAllID, ref questionID, chapterID, nodeID, list, text);
                                }
                                else if (regA.Split(text).Length == 1)//判断题
                                {
                                    text += "A、 B、 C、 D、";
                                    proceduQuestion(subject, chapter, node, ref questionAllID, ref questionID, chapterID, nodeID, list, text);

                                }
                                else//其它， 不知道是什么情况，有可能是判断题
                                {

                                    consolewrite("其他：数字开头，不是三个/四个选项 ", regA.Split(text).Length + "_" + text);
                                }
                            }
                            #endregion

                            #region 数字开头没有下划线的，分别检测参考答案，错误
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
                                        if (q.SNID == No && q.Subject == subject)
                                        {
                                            if (string.IsNullOrEmpty(q.Answer))
                                            {
                                                q.Answer = answer.ToUpper();
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

                                }
                                else//错误部分
                                {
                                    //writer.WriteLine("错误：数字开头无下划线 - " + text);
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("错误：数字开头无下划线 - " + text);
                                    Console.ReadLine();
                                }

                            }
                            #endregion
                        }
                        else//错误
                        {
                            // writer.WriteLine("错误：非章节标题，非数字开头，你是个什么鬼- " + text);
                            consolewrite("错误：非章节标题，非数字开头，你是个什么鬼- ", text);

                        }

                        Console.ResetColor();
                        Console.WriteLine();
                    }//段落检测结束，全部循环完，则一本试题结束

                    app.Documents.Close();
                    Console.WriteLine("文件正在关闭。");
                    Thread.Sleep(1000);
                    Console.WriteLine("文件关闭成功，开始将试题写入到表格及数据库中……");
                    questiontoexcel(list, "d://" + str + ".xls");
                    Console.WriteLine("试题写入完成，地址：D://{0}.xls", str);
                    list.Clear();
                }
                catch (Exception ex)
                {
                    writer.WriteLine(ex.Message);
                    consolewrite(ex.Message, "");
                }
                finally
                {
                    writer.Flush();
                    timeend[j] = DateTime.Now;
                }
                numForDocument[j] = questionAllID;
                numForAll += questionAllID;
            }//所有试题结束
            writer.Close();
            printDatetime(subjects, timestar, timeend, log, numForAll, numForDocument);
            consolewrite("所有写入完成", "");

        }

        /// <summary>
        /// 打印出统计信息
        /// </summary>
        /// <param name="subjects"></param>
        /// <param name="star"></param>
        /// <param name="end"></param>
        public static void printDatetime(string[] subjects, DateTime[] star, DateTime[] end, StreamWriter writer, int allnum, int[] numForDocument)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine(string.Format("总计完成的试题数为：{0}", allnum));
            sb.AppendLine();
            sb.AppendLine("各科目的用时情况统计如下所示：");
            sb.AppendLine();
            sb.AppendLine(string.Format("总用时：{0}", (end[subjects.Length - 1] - star[0]).ToString()));
            sb.AppendLine();
            for (int i = 0; i < subjects.Length; i++)
            {
                sb.AppendLine(string.Format("其中 {0} 试题数为:{1}", subjects[i], numForDocument[i]));
                sb.AppendLine(string.Format("处理 {0} 从 {1} 开始 ——至 {2} 结束，共用时 {3} 。", subjects[i], star[i], end[i], (end[i] - star[i]).ToString()));
                sb.AppendLine();
            }
            string message = sb.ToString();
            try
            {
                writer.Write(message); writer.Flush();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                writer.Close();
            }
            Console.WriteLine(message);


        }

        /// <summary>
        /// 处理试题
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="chapter"></param>
        /// <param name="node"></param>
        /// <param name="questionAllID"></param>
        /// <param name="questionID"></param>
        /// <param name="chapterID"></param>
        /// <param name="nodeID"></param>
        /// <param name="list"></param>
        /// <param name="text"></param>
        private static void proceduQuestion(string subject, string chapter, string node, ref int questionAllID, ref int questionID, int chapterID, int nodeID, List<Question> list, string text)
        {
            questionID++;
            questionAllID++;
            Question question = TextToQuestionModel(subject, chapter, node, questionAllID, questionID, chapterID, nodeID, string.Empty, text);
            printQuesiton(question);
            list.Add(question);
        }




        /// <summary>
        /// 将一行文本填充成一个选项题的试题模型
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="chapter"></param>
        /// <param name="node"></param>
        /// <param name="questionAllID"></param>
        /// <param name="questionID"></param>
        /// <param name="chapterID"></param>
        /// <param name="nodeID"></param>
        /// <param name="remark"></param>
        /// <param name="regA"></param>
        /// <param name="regNO"></param>
        /// <param name="regxhx"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        public static Question TextToQuestionModel(string subject, string chapter, string node, int questionAllID, int questionID, int chapterID, int nodeID, string remark, string text)
        {
            string[] strSplit = regA.Split(new Regex("\\r\\a").Replace(text, "").Trim());
            if (strSplit.Length != 5)
            {
                consolewrite("试题被拆分成多个部分", strSplit.Length + "------" + text);
            }
            Question question = new Question();
            question.Subject = subject;
            question.Chapter = chapter;
            question.Node = node;
            question.AllID = questionAllID;//总编号
            question.Id = chapterID + "_" + nodeID + "_" + questionID;//章节+总编号
            question.SN = Int32.Parse(new Regex("^[0-9]+", RegexOptions.IgnoreCase).Match(strSplit[0]).Value);//原编号
            question.SNID = chapterID + "_" + nodeID + "_" + question.SN;//章节+原编号
            string title = regxhx.Replace(regNO.Replace(strSplit[0], ""), "_______");
            if (regxhx.IsMatch(title))
                question.Title = title;
            else { consolewrite("题干部分无下划线", ""); }
            question.Choosea = strSplit[1].Trim();
            question.Chooseb = strSplit[2].Trim();
            question.Choosec = strSplit[3].Trim();
            question.Choosed = strSplit[4].Trim();
            question.Answer = string.Empty;
            question.Explain = string.Empty;
            question.ImageAddress = string.Empty;
            question.Remark = remark;
            return question;
        }

        /// <summary>
        /// 处理关联试题
        /// </summary>
        /// <param name="doc">试题所在的文章</param>
        /// <param name="p">关联题开始的行号</param>
        /// <param name="k">关联题试题开始的行号</param>
        private static Question[] chuliglt(Word.Document doc, int rownum, int questionrownum, string subject, string chapter, string node, int questionAllID, int questionID, int chapterID, int nodeID, Regex regA, Regex regNO, Regex regxhx)
        {

            Question[] question = new Question[4];
            StringBuilder sb = new StringBuilder();
            string remark = string.Empty;
            for (int i = rownum; i < questionrownum; i++)
            {
                sb.AppendLine(new Regex("\\r\\a").Replace(doc.Paragraphs[i].Range.Text, "").Trim());
            }
            remark = sb.ToString();
            for (int j = questionrownum, k = 0; j < questionrownum + 4; j++, k++)
            {
                question[k] = TextToQuestionModel(subject, chapter, node, questionAllID, questionID, chapterID, nodeID, remark, doc.Paragraphs[j].Range.Text);
            }
            return question;

        }


        /// <summary>
        /// 打印出试题对象
        /// </summary>
        /// <param name="question">试题对象</param>
        public static void printQuesiton(object question)
        {
            PropertyInfo[] proper = question.GetType().GetProperties();
            for (int i = 0; i < question.GetType().GetProperties().Length; i++)
            {
                Console.WriteLine("{0}:{1}", proper[i].Name, proper[i].GetValue(question).ToString());
            }
        }


        /// <summary>
        /// 将试题填充到EXCEL并写入到数据库
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
            for (int j = 0; j < list.Count; j++)
            {
                Question q = list[j];
                printQuesiton(q);
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
                Console.WriteLine("第{0}条数据写入表格成功，总计：{1}条，剩余：{2}条", j + 1, list.Count, list.Count - j - 1);
                if (insertQuestionTODB(q, "Question", "ChooseQuestion") == 1)
                {
                    Console.WriteLine("第{0}条数据写入数据库成功，总计：{1}条，剩余：{2}条", j + 1, list.Count, list.Count - j - 1);
                }
                else
                {
                    Console.WriteLine("第{0}条数据插入数据库错误，总计：{1}条，剩余：{2}条", j + 1, list.Count, list.Count - j - 1);
                    Console.ReadLine();
                }
                Console.WriteLine();

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

        /// <summary>
        /// 将试题插入到数据库
        /// </summary>
        /// <param name="question"></param>
        /// <param name="DB"></param>
        /// <param name="table"></param>
        /// <returns>返回影响的行数</returns>
        public static int insertQuestionTODB(Question question, string DB, string table)
        {
            switch (question.Answer.ToUpper())
            {
                case "A": question.Answer = "1"; break;
                case "B": question.Answer = "2"; break;
                case "C": question.Answer = "3"; break;
                case "D": question.Answer = "4"; break;
                default: question.Answer = "0"; break;
            }
            SqlConnection conn = new SqlConnection("Data Source=localhost;Initial Catalog=" + DB + ";Integrated Security=true;");
            conn.Open();
            string sqlstr = "insert into " + table + " VALUES(" + question.AllID + ",'" + question.Id + "'," + question.SN + ",'" + question.SNID + "','" + question.Subject + "','" + question.Chapter + "','" + question.Node + "','" + question.Title + "','" + question.Choosea + "','" + question.Chooseb + "','" + question.Choosec + "','" + question.Choosed + "'," + Int32.Parse(question.Answer) + ",'" + question.Explain + "','" + question.ImageAddress + "','" + question.Remark + "')";//
            Console.WriteLine(sqlstr);
            SqlCommand command = new SqlCommand(sqlstr, conn);
            try
            {
                return command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }

        /// <summary>
        /// 检测多个正则表达式,数字开头的，有下划线，并且可以分成五个部分
        /// </summary>
        /// <param name="text"></param>
        /// <param name="reg"></param>
        /// <returns></returns>
        public static Boolean is4ChooseQuestion(string text)
        {
            Boolean istrue = true;
            if (!regNO.IsMatch(text) || !regxhx.IsMatch(text) || regA.Split(text).Length != 5)
            {
                istrue = false;
                //consolewrite("转换成试题模型不成功",text);
            }
            return istrue;
        }

        /// <summary>
        /// 从控制台打印文体，并等待输入
        /// </summary>
        /// <param name="reason">原因</param>
        /// <param name="message">信息</param>
        public static void consolewrite(string reason, string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("{0}---{1}", reason, message);
            Console.ResetColor();
            Console.ReadLine();
        }
    }
}