using System;
using System.IO;
using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;
using System.Threading;

namespace Questions
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] documents = new string[] { "QuestionLibraries/instruction-yuxiangshu.docx" };
            //"QuestionLibraries/avoidcollision-wufei.docx"  ,
            // "QuestionLibraries/certificate-yuxiangshu.docx",
            //"QuestionLibraries/english-xiangwei.docx",
            //"QuestionLibraries/equipment-hedetao.docx",
            //"QuestionLibraries/instruction-yuxiangshu.docx",
            //"QuestionLibraries/management-lizhite.docx",
            //"QuestionLibraries/navigation-hedetao.docx",
            //"QuestionLibraries/ocean-hedetao.docx"

            int mark = 0;//标记
            bool expstar = false;
            string chapter = string.Empty;
            string node = string.Empty;
            Regex regA = new Regex("[ABCDabcd]{1}[\\.|、]", RegexOptions.IgnoreCase);//A|B|C|D
            Regex regNO = new Regex("^[0-9]+[\\.|、]", RegexOptions.IgnoreCase);//以数字开头  题干
            Regex regexpstar = new Regex("^参考答案");
            Regex regexp = new Regex("^[0-9]+[\\.|、][ABCDabcd]{1}[\\.|、]");//解释
            Regex regChapter = new Regex("^第[一二三四五六七八九十]{1,3}章", RegexOptions.IgnoreCase);//章标题
            Regex regNode = new Regex("^第[一二三四五六七八九十]{1,3}节", RegexOptions.IgnoreCase);//节标题
            Regex regxhx = new Regex("[_]{4,10}", RegexOptions.IgnoreCase);//下划线
            StreamWriter writer = new StreamWriter("D://error.txt", true, System.Text.Encoding.Default, 1 * 1024);
            foreach (string str in documents)
            {
                string path = "d:/" + str;
                Console.WriteLine(path);
                writer.WriteLine(path);
                try
                {
                    using (FileStream stream = File.OpenRead(path))
                    {
                        XWPFDocument doc = new XWPFDocument(stream);
                        foreach (var para in doc.Paragraphs)
                        {
                            string text = para.Text.Trim();
                            if (string.IsNullOrEmpty(text)) return;
                            if (regChapter.IsMatch(text))//章标题
                            {
                                expstar = false;
                                chapter = text;
                                Console.WriteLine(mark);
                                writer.WriteLine(mark);
                                mark = 0;
                                writer.WriteLine(text);
                                Console.WriteLine("章标题： " + chapter);
                            }
                            else if (regNode.IsMatch(text))//节标题
                            {
                                expstar = false;
                                node = text;
                                Console.WriteLine(mark);
                                writer.WriteLine(mark);
                                mark = 0;
                                writer.WriteLine(node);
                                Console.WriteLine("节标题： " + node);
                            }
                            else if (regexpstar.IsMatch(text))//参考答案开始
                            {
                                expstar = true;
                                Console.WriteLine("参考答案开始： " + text);

                            }
                            else if (regNO.IsMatch(text))//数字开头的
                            {
                                if (regxhx.IsMatch(text))//有下划线的
                                {
                                    if (regA.Split(text).Length == 5)//有四个选项
                                    {
                                        mark++;
                                        Console.WriteLine("四个选项试题： " + text);

                                    }
                                    else if (regA.Split(text).Length == 4)//有三个选项
                                    {
                                        mark++;
                                        Console.WriteLine("三个选项试题： " + text);
                                    }
                                    else//其它， 不知道是什么情况，有可能是判断题
                                    {
                                        mark++;
                                        writer.WriteLine("其他： " + regA.Split(text).Length + "_" + text);
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine("其他： " + regA.Split(text).Length + "_" + text);
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
                                        writer.WriteLine("错误：数字开头无下划线 - " + text);
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine("错误：数字开头无下划线 - " + text);
                                        Console.ReadLine();
                                    }

                                }
                            }
                            else
                            {
                                writer.WriteLine("错误： " + text);
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("错误： " + text);
                                Console.ReadLine();
                            }
                            Console.ResetColor();
                            Console.WriteLine();
                            Thread.Sleep(500);
                        }
                    }
                }
                catch (Exception ex) { writer.WriteLine(ex.Message); Console.ReadLine(); }
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