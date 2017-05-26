
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetQuestonsFromWordToExcel
{
    class Question
    {
        private int id;
        private string chapter;
        private string node;
        private string title;
        private string choosea;
        private string chooseb;
        private string choosec;
        private string choosed;
        private string answer;
        private string explain;

        public int Id
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
