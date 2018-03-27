using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExamApp
{
    class Question
    {
        public int QuestionNo { get; set; }

        public string QuestionText { get; set; }

        public List<Answer> AnswerList { get; set; }

        public int CorrectAnswerNo { get; set; }

        public string CorrectAnswerText { get; set; }

        public string Explaination { get; set; }

        public int SelectedAnswerNo { get; set; }

        public List<string> Images { get; set; }

        public bool Result { get; set; }
    }
}
