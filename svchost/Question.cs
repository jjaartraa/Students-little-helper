using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace svchost
{
    class Question
    {
        string question;
        string answer;
        double accuracy;


        public Question(string question, string answer, int accuracy)
        {
            this.question = question;
            this.answer = answer;
            this.accuracy = accuracy;
        }

        public string GetQuestion()
        {
            return question;
        }
        public string GetAnswer()
        {
            return answer;
        }
        public double GetAccuracy()
        {
            return accuracy;
        }



    }
}
