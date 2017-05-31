using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Questions
{
    class PrintErrorParagraphs
    {
        private XWPFParagraph para = null;
        public PrintErrorParagraphs(XWPFParagraph p) {
            para = p;
        }
    }
}


                    //foreach (var para in doc.Paragraphs)
                    //{
                    //    string text = para.Text.Trim();
                    //    if (string.IsNullOrEmpty(text)) return;
                    //    if (regChapter.IsMatch(text))//章标题
                    //    {
                    //        expstar = false;
                    //        chapter = text;
                    //        Console.WriteLine(mark);
                    //        writer.WriteLine(mark);
                    //        mark = 0;
                    //        writer.WriteLine(text);
                    //        Console.WriteLine("章标题： " + chapter);
                    //    }
                    //    else if (regNode.IsMatch(text))//节标题
                    //    {
                    //        expstar = false;
                    //        node = text;
                    //        Console.WriteLine(mark);
                    //        writer.WriteLine(mark);
                    //        mark = 0;
                    //        writer.WriteLine(node);
                    //        Console.WriteLine("节标题： " + node);
                    //    }
                    //    else if (regexpstar.IsMatch(text))//参考答案开始
                    //    {
                    //        expstar = true;
                    //        Console.WriteLine("参考答案开始： " + text);

                    //    }
                    //    else if (regNO.IsMatch(text))//数字开头的
                    //    {
                    //        if (regxhx.IsMatch(text))//有下划线的
                    //        {
                    //            if (regA.Split(text).Length == 5)//有四个选项
                    //            {
                    //                mark++;
                    //                Console.WriteLine("四个选项试题： " + text);

                    //            }
                    //            else if (regA.Split(text).Length == 4)//有三个选项
                    //            {
                    //                mark++;
                    //                Console.WriteLine("三个选项试题： " + text);
                    //            }
                    //            else//其它， 不知道是什么情况，有可能是判断题
                    //            {
                    //                mark++;
                    //                writer.WriteLine("其他： " + regA.Split(text).Length + "_" + text);
                    //                Console.ForegroundColor = ConsoleColor.Red;
                    //                Console.WriteLine("其他： " + regA.Split(text).Length + "_" + text);
                    //            }
                    //        }
                    //        else//数字开头，无下划线
                    //        {
                    //            if (regexp.IsMatch(text) && expstar)//参考答案与解析
                    //            {
                    //                Console.WriteLine("参考答案: " + text);
                    //            }
                    //            else//错误部分
                    //            {
                    //                writer.WriteLine("错误： " + text);
                    //                Console.ForegroundColor = ConsoleColor.Red;
                    //                Console.WriteLine("错误： " + text);
                    //            }

                    //        }
                    //    }
                    //    else
                    //    {
                    //        writer.WriteLine("错误： " + text);
                    //        Console.ForegroundColor = ConsoleColor.Red;
                    //        Console.WriteLine("错误： " + text);
                    //    }
                    //    Console.ResetColor();
                    //    Console.WriteLine();
                    //    Thread.Sleep(400);
                    //}
