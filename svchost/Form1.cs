using System;
using System.Collections.Generic;
using System.Windows.Forms;
using static svchost.Hotkeys;

namespace svchost
{
    public partial class svchost : Form
    {
        private void AddDrag(Control Control) { Control.MouseDown += new System.Windows.Forms.MouseEventHandler(this.DragForm_MouseDown); }
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private void DragForm_MouseDown(object sender, MouseEventArgs e) //Handle for moving window by clicking anywhere
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                // Checks if Y = 0, if so maximize the form
                if (this.Location.Y == 0) { this.WindowState = FormWindowState.Maximized; }
            }
        }

        private KeyboardHook exithook = new KeyboardHook();
        private KeyboardHook changehooknext = new KeyboardHook(); //Keyboard hook to show next answer
        private KeyboardHook changehookprev = new KeyboardHook(); //Keyboard hook to show prev. answer
        private KeyboardHook answerhook = new KeyboardHook();   //Keyboard hook to show answer when key is pressed

        private List<Question> questions = new List<Question>(); //List of all questions and answers from the Excel table
        private List<Question> relevantquestions = new List<Question>(); //List of questions and answers related to question from clipboard
        string question = ""; //clipboard question
        int answerIndex = 0; //position of actual answer
        char[] delimiters = new char[] { ' ', '\r', '\n' }; //Definition of characters removed from sentences
        public svchost()
        {
            InitializeComponent();
            AddDrag(this);
            Init(); //Open excel sheet, load data to questions array

            // register the event that is fired after the key press.
            changehooknext.KeyPressed +=
                new EventHandler<KeyPressedEventArgs>(cahngehooknext_KeyPressed);
            // register X as hot key.
            changehooknext.RegisterHotKey(Hotkeys.ModifierKeys.None, Keys.X);

            changehookprev.KeyPressed +=
                new EventHandler<KeyPressedEventArgs>(cahngehookprev_KeyPressed);
            // register Y as hot key.
            changehookprev.RegisterHotKey(Hotkeys.ModifierKeys.None, Keys.Y);

            exithook.KeyPressed +=
                new EventHandler<KeyPressedEventArgs>(exithook_KeyPressed);
            // register Esc as hot key.
            exithook.RegisterHotKey(Hotkeys.ModifierKeys.None, Keys.Escape);

            answerhook.KeyPressed +=
                new EventHandler<KeyPressedEventArgs>(answerhook_KeyPressed);
            // register Alt as hot key.
            answerhook.RegisterHotKey(Hotkeys.ModifierKeys.Alt, Keys.None);

        }

        public void Init()
        {
            Excel excel = new Excel(@"source.xlsx", 1);
            
            //MessageBox.Show(excel.ReadCell(0, 0));
            int x = 0, y = 0; //X řádky Y sloupce
            while (excel.ReadCell(x, y) != "") //Load Questions
            {
                questions.Add(new Question(excel.ReadCell(x, y).ToLower(), excel.ReadCell(x, y + 1).ToLower(), 0));
                x++;
            }
            excel.Close();
        }

        private void answerhook_KeyPressed(object sender, KeyPressedEventArgs e)
        {
            if (question != Clipboard.GetText().ToLower() && Clipboard.GetText() != "")
            {
                question = Clipboard.GetText().ToLower();
                relevantquestions.Clear();
                MatchBestAnswer(Clipboard.GetText().ToLower());
                refresh();
            }

            if (this.Visible == false)
            {
                this.Visible = true;
            }
            else if(this.Visible == true)
            {
                this.Visible = false;
            }
        }

        private void exithook_KeyPressed(object sender, KeyPressedEventArgs e) //Cycle through answers X - next, Y - prev.
        {
            Application.Exit();
        }

            private void cahngehooknext_KeyPressed(object sender, KeyPressedEventArgs e) //Cycle through answers X - next, Y - prev.
        {
            if (relevantquestions.Count > 1) {
                if (answerIndex == relevantquestions.Count - 1)
                    answerIndex = 0;
                else
                    answerIndex++;
                refresh();
            }
        }
        private void cahngehookprev_KeyPressed(object sender, KeyPressedEventArgs e) //Cycle through answers X - next, Y - prev.
        {
            if (relevantquestions.Count > 1)
            {
                if (answerIndex == 0)
                    answerIndex = relevantquestions.Count - 1;
                else
                    answerIndex--;
                refresh();
            }
        }


        void MatchBestAnswer(string question)
        {
            String[] src = question.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);            //Splits clipboard sentence into words
            foreach (Question sentence in questions)    //Looks in every Excel question
            {
                int questionwordcount = sentence.GetQuestion().Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;
                int accuracy = 0;
                foreach (string word in src)            //For every clipboard word
                {
                    if (sentence.GetQuestion().Contains(word))  //Looks if Excel questions contains same words as clipboard question
                        accuracy++;                                 //Add accuracy if it does
                    else
                    { }
                }
                if (accuracy > 1)  //If both questions share more than one word, add Excel Q&A into list of relevant questions
                {
                      //Calculate accuracy based on word count (accuracy(same words) / all words) <0-1>
                    relevantquestions.Add(new Question(sentence.GetQuestion(), sentence.GetAnswer(), (accuracy / questionwordcount)));
                }
            }

            relevantquestions.Sort((x, y) => x.GetAccuracy().CompareTo(y.GetAccuracy())); //Sort Q&A by highest accuracy
            relevantquestions.Reverse();
        }

        void refresh()
        {
            if (relevantquestions.Count > 1)
            {
                label1.Text = relevantquestions[answerIndex].GetQuestion();
                label2.Text = relevantquestions[answerIndex].GetAnswer();
                label3.Text = (answerIndex + 1) + "/" + relevantquestions.Count;
            }
            else
            {
                label1.Text = "Žádné shodné otázky nenalezeny";
                label2.Text = "";
                label3.Text = "0";
            }
        }
    }
}
