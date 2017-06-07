using System;
using System.IO;
using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace Questions
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] documents = new string[] { 
            "QuestionLibraries/avoidcollision-wufei.docx"  //,
            // "QuestionLibraries/certificate-yuxiangshu.docx",
            //"QuestionLibraries/english-xiangwei.docx",
            //"QuestionLibraries/equipment-hedetao.docx",
            //"QuestionLibraries/instruction-yuxiangshu.docx",
            //"QuestionLibraries/management-lizhite.docx",
            //"QuestionLibraries/navigation-hedetao.docx",
            //"QuestionLibraries/ocean-hedetao.docx"
            };
            
            bool expstar = false;//解析开始标记
            string chapter = string.Empty;//章标题
            int chapterID = 0;//章序号
            string node = string.Empty;//节标题
            int nodeID = 0;//节序号
            int questionID = 0;//试题序号
            int questionAllID = 0;//试题的总序号
            Regex regA = new Regex("[ABCDabcd]{1}[\\.|、]", RegexOptions.IgnoreCase);//A|B|C|D
            Regex regNO = new Regex("^[0-9]+[\\.|、]", RegexOptions.IgnoreCase);//以数字开头  题干
            Regex regexpstar = new Regex("^参考答案");//参考答案开头
            Regex regexp = new Regex("^[0-9]+[\\.|、][ABCDabcd]{1}[\\.|、|。]");//解释
            Regex regChapter = new Regex("^第[一二三四五六七八九十]{1,3}章", RegexOptions.IgnoreCase);//章标题
            Regex regNode = new Regex("^第[一二三四五六七八九十]{1,3}节", RegexOptions.IgnoreCase);//节标题
            Regex regxhx = new Regex("[_]{3,10}", RegexOptions.IgnoreCase);//下划线
            StreamWriter writer = new StreamWriter("D://error.txt", true, System.Text.Encoding.Default, 1 * 1024);
            List<Question> list = new List<Question>();
            foreach (string str in documents)
            {
                string path = @"d:/" + str;
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
                    for (int i = 1; i < paragraphsCount; i++)
                    {
                        Word.Range para = doc.Paragraphs[i].Range;
                        para.Select();
                        string text = para.Text.Trim();
                        if (string.IsNullOrEmpty(text)) continue;
                        if (regChapter.IsMatch(text))//章标题
                        {
                            expstar = false;
                            chapter = text;
                            chapterID++;
                            nodeID = 0;
                            questionID = 0;
                           // writer.WriteLine(text);
                            Console.WriteLine("章标题： " + chapter);
                        }
                        else if (regNode.IsMatch(text))//节标题
                        {
                            nodeID++;
                            questionID = 0;
                            expstar = false;
                            node = text;
                           // writer.WriteLine(node);
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
                                    question.Chapter = chapter;
                                    question.Node = node;
                                    question.AllID = questionAllID;
                                    question.Id = chapterID+"_"+nodeID+"_"+questionID;
                                    question.SN = Int32.Parse(new Regex("^[0-9]+", RegexOptions.IgnoreCase).Match(strSplit[0]).Value);
                                    question.Title = regNO.Replace(strSplit[0],"");
                                    question.Choosea = strSplit[1];
                                    question.Chooseb = strSplit[2];
                                    question.Choosec = strSplit[3];
                                    question.Choosed = strSplit[4];
                                    question.Answer = null;
                                    question.Explain = null;
                                    list.Add(question);
                                    Console.WriteLine("四个选项试题： " + text);

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
                                    Console.WriteLine("参考答案: " + text);
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
                        //Thread.Sleep(100);
                    }
                    //app.Documents.Close();
                }
                catch (Exception ex)
                {
                   // writer.WriteLine(ex.Message);
                    Console.WriteLine(ex.Message);
                    Console.ReadLine();
                }
                finally
                {
                    writer.Flush();
                    writer.Close();
                }
            }
            Console.WriteLine("完成");
            Console.ReadLine();
        }
    }
}