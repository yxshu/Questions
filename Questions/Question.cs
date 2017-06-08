﻿
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Questions
{
    class Question
    {
        private int allid;
        private string id;
        private int sn;
        private string snID;
        private string subject;


        private string chapter;
        private string node;
        private string title;
        private string choosea;
        private string chooseb;
        private string choosec;
        private string choosed;
        private string answer;
        private string explain;
        public int AllID
        {
            get { return allid; }
            set { allid = value; }
        }
        /// <summary>
        /// 试题编号
        /// </summary>
        public string Id
        {
            get
            {
                return id;
            }

            set
            {
                id = value;
            }
        }
        /// <summary>
        /// 试题原编号
        /// </summary>
        public int SN
        {
            get { return sn; }
            set { sn = value; }
        }
        /// <summary>
        /// 试题原编号加章节编号
        /// </summary>
        public string SNID {
            get { return snID; }
            set { snID = value; }
        }
        /// <summary>
        /// 科目
        /// </summary>
        public string Subject
        {
            get { return subject; }
            set { subject = value; }
        }
        /// <summary>
        /// 章标题
        /// </summary>
        public string Chapter
        {
            get
            {
                return chapter;
            }

            set
            {
                chapter = value;
            }
        }
        /// <summary>
        /// 节标题
        /// </summary>
        public string Node
        {
            get
            {
                return node;
            }

            set
            {
                node = value;
            }
        }
        /// <summary>
        /// 试题题干
        /// </summary>
        public string Title
        {
            get
            {
                return title;
            }

            set
            {
                title = value;
            }
        }
        /// <summary>
        /// 选项A
        /// </summary>
        public string Choosea
        {
            get
            {
                return choosea;
            }

            set
            {
                choosea = value;
            }
        }
        /// <summary>
        /// 选项B
        /// </summary>
        public string Chooseb
        {
            get
            {
                return chooseb;
            }

            set
            {
                chooseb = value;
            }
        }
        /// <summary>
        /// 选项C
        /// </summary>
        public string Choosec
        {
            get
            {
                return choosec;
            }

            set
            {
                choosec = value;
            }
        }
        /// <summary>
        /// 选项D
        /// </summary>
        public string Choosed
        {
            get
            {
                return choosed;
            }

            set
            {
                choosed = value;
            }
        }
        /// <summary>
        /// 参考答案
        /// </summary>
        public string Answer
        {
            get
            {
                return answer;
            }

            set
            {
                answer = value;
            }
        }
        /// <summary>
        /// 解析
        /// </summary>
        public string Explain
        {
            get
            {
                return explain;
            }

            set
            {
                explain = value;
            }
        }
    }
}
