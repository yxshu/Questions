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
            string[] documents = new string[] {
             "QuestionLibraries/Updataed/avoidcollision-wufei.docx" }; //,
                                                                       //  "QuestionLibraries/Updataed/certificate-yuxiangshu.docx",
                                                                       //"QuestionLibraries/Updataed/english-xiangwei.docx",
                                                                       //"QuestionLibraries/Updataed/equipment-hedetao.docx",
                                                                       //"QuestionLibraries/Updataed/instruction-yuxiangshu.docx",
                                                                       //"QuestionLibraries/Updataed/management-lizhite.docx",
                                                                       //"QuestionLibraries/Updataed/navigation-hedetao.docx",
                                                                       //"QuestionLibraries/Updataed/ocean-hedetao.docx"

            int mark = 0;//标记
            bool expstar = false;
            string chapter = string.Empty;
            string node = string.Empty;
            Regex regA = new Regex("[ABCDabcd]{1}.", RegexOptions.IgnoreCase);//A|B|C|D
            Regex regNO = new Regex("^[0-9]+.", RegexOptions.IgnoreCase);//以数字开头  题干
            Regex regexpstar = new Regex("^参考答案");
            Regex regexp = new Regex("^[0-9]+.[ABCDabcd]{1}.");//解释
            Regex regChapter = new Regex("^第[一二三四五六七八九十]{1,3}章", RegexOptions.IgnoreCase);//章标题
            Regex regNode = new Regex("^第[一二三四五六七八九十]{1,3}节", RegexOptions.IgnoreCase);//节标题
            Regex regxhx = new Regex("[_]{4,10}", RegexOptions.IgnoreCase);//下划线
            StreamWriter writer = new StreamWriter("e://error.txt", true);
            foreach (string str in documents)
            {
                string path = "e:/" + str;
                Console.WriteLine(path);
                writer.WriteLine(path);
                using (FileStream stream = File.OpenRead(path))
                {
                    XWPFDocument doc = new XWPFDocument(stream);

                }
            }
            writer.Close();
            Console.WriteLine("完成");
            Console.ReadLine();
        }
    }
}