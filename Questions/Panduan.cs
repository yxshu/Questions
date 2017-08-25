using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Questions
{
    class Panduan
    {
        private int allID;
        /// <summary>
        /// 试题总的编号
        /// </summary>
        public int AllID
        {
            get { return allID; }
            set { allID = value; }
        }
        private string c_N_Id;
        /// <summary>
        /// 章节编号+试题总编号
        /// </summary>
        public string C_N_Id
        {
            get { return c_N_Id; }
            set { c_N_Id = value; }
        }
        private int sN;
        /// <summary>
        /// 试题本身的编号
        /// </summary>
        public int SN
        {
            get { return sN; }
            set { sN = value; }
        }
        private string sNID;
        /// <summary>
        /// 章节编号+试题本身的编号
        /// </summary>
        public string SNID
        {
            get { return sNID; }
            set { sNID = value; }
        }
        private string subj;
        /// <summary>
        /// 科目
        /// </summary>
        public string Subj
        {
            get { return subj; }
            set { subj = value; }
        }
        private string chapter;
        /// <summary>
        /// 章标题
        /// </summary>
        public string Chapter
        {
            get { return chapter; }
            set { chapter = value; }
        }
        private string node;
        /// <summary>
        /// 节标题
        /// </summary>
        public string Node
        {
            get { return node; }
            set { node = value; }
        }
        private string title;
        /// <summary>
        /// 题干
        /// </summary>
        public string Title
        {
            get { return title; }
            set { title = value; }
        }
        private bool answer;
        /// <summary>
        /// 参考答案
        /// </summary>
        public bool Answer
        {
            get { return answer; }
            set { answer = value; }
        }
        private string explain;
        /// <summary>
        /// 试题解析
        /// </summary>
        public string Explain
        {
            get { return explain; }
            set { explain = value; }
        }
        private string imageAddress;
        /// <summary>
        /// 图片地址
        /// </summary>
        public string ImageAddress
        {
            get { return imageAddress; }
            set { imageAddress = value; }
        }
        private string remark;
        /// <summary>
        /// 备注
        /// </summary>
        public string Remark
        {
            get { return remark; }
            set { remark = value; }
        }
    }
}
