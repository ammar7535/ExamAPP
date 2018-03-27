using System;
using System.Collections.Generic;
using System.Reflection;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net.Mime;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.XPath;
using iTextSharp;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Office.Interop.Word;
using Org.BouncyCastle.Apache.Bzip2;
using Org.BouncyCastle.Cms;
using Telerik.WinControls.UI;
using Telerik.WinControls;
using Telerik.WinControls.Zip;
using Telerik.WinControls.Zip.Extensions;
using Telerik.Charting;

namespace ExamApp
{
    public partial class frmMain : Form
    {
        XmlDocument docx_file_xml = new XmlDocument();
        XmlNamespaceManager mgr;
        private int test_status = 0; //0-performing test, 1-review test
        private List<Question> question_list = new List<Question>();
        private string dir_path = ""; //parent directory path
        private string word_file_path = "..\\..\\repo\\70-411.docx"; //complete word file path including file name;
        private string word_file_name = ""; //only word file name
        private string unzip_dir_path = ""; //unzip directory path
        private string zip_file_path = ""; //on zip file path
        private int serial_no = 0;
        private int current_question = 0;
        private Stopwatch sw = new Stopwatch();
        private XmlNodeList question_nodes;
        int CorrectAnswer = 0;
        string corr;
        int Result = 0;
        int TotalQuestions = 0;
        int CorectQuestion = 0;
        int Variatnt = 1;
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            LineSeries lineSeries = new LineSeries();
            //string[] lines = System.IO.File.ReadAllLines(@"..\\..\\Result\\result.txt");
            foreach (String line in File.ReadAllLines(@"..\\..\\Result\\result.txt"))
            {
                
               string[] items = line.Split(',');
               if (items[0] == "TotalQuestions")
               { 
                   TotalQuestions = Convert.ToInt32(items[1]);
               }
               if (items[0] == "CorrectAnswers")
               {
                   CorectQuestion = Convert.ToInt32(items[1]);
                   Result = (CorectQuestion * 100) / TotalQuestions;
                   lineSeries.DataPoints.Add(new CategoricalDataPoint(Result, "Test" + Variatnt));
                   Variatnt++;
               }
               
            }
            
                this.HistoryViewChart.Series.Add(lineSeries);
                HistoryViewChart.Series[0].BackColor = Color.Red;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                test_status = 0;
                if (word_file_path == "")
                {
                    OpenFileDialog ofdWord = new OpenFileDialog();
                    if (ofdWord.ShowDialog() == DialogResult.OK)
                        word_file_path = ofdWord.FileName;
                }

                dir_path = word_file_path.Substring(0, word_file_path.LastIndexOf("\\"));
                word_file_name = word_file_path.Substring(word_file_path.LastIndexOf("\\") + 1);
                zip_file_path = ConvertDocxToZip();
                unzip_dir_path = zip_file_path.Substring(0, zip_file_path.LastIndexOf("."));

                mgr = new XmlNamespaceManager(docx_file_xml.NameTable);
                ExtractWordFile();
                docx_file_xml.Load(unzip_dir_path + @"\word\document.xml");
                //ProcessWordFile();
                ProcessFile();
                DisplayQuestion();

                btnNext.Enabled = true;
                
                tmStopWatch.Enabled = true;
                sw.Start();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
            
        }

        private void tmStopWatch_Tick(object sender, EventArgs e)
        {
            lblTimeElapsed.Text = sw.Elapsed.ToString().Substring(0, 8);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            
            if (btnNext.Text == "Review")
                ShowResult();
            if (test_status == 1)
                btnNext.Text = "Next";
           
            current_question++;
            DisplayQuestion();
        }

        private void ShowResult()
        {
            string path = "..\\..\\Result\\result.txt";
            // This text is added only once to the file.
            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter swr = File.CreateText(path))
                {
                    swr.WriteLine("TestDate, " + System.DateTime.Now);
                    swr.WriteLine("TotalQuestions, " + question_list.Count);
                    swr.WriteLine("CorrectAnswers, " + CorrectAnswer);
                    swr.WriteLine("WrongAnswers, " + (question_list.Count - CorrectAnswer));
                    swr.WriteLine("TimeCompletion, " + sw.Elapsed.ToString().Substring(0, 8));
                }
            }
            else 
            { 
            // This text is always added, making the file longer over time
            // if it is not deleted.
                using (StreamWriter swr = File.AppendText(path))
                {
                    swr.WriteLine("TestDate, " + System.DateTime.Now);
                    swr.WriteLine("TotalQuestions, " + question_list.Count);
                    swr.WriteLine("CorrectAnswers, " + CorrectAnswer);
                    swr.WriteLine("WrongAnswers, " + (question_list.Count - CorrectAnswer));
                    swr.WriteLine("TimeCompletion, " + sw.Elapsed.ToString().Substring(0, 8));
                }
            }
           
            RadMessageBox.Show("Your correct answers are " + CorrectAnswer + "/" + question_list.Count);
        }

        private string ConvertDocxToZip()
        {
            if (File.Exists(word_file_path.Substring(0, word_file_path.LastIndexOf(".")) + ".zip"))
                return word_file_path.Substring(0, word_file_path.LastIndexOf(".")) + ".zip"; 

            File.Copy(word_file_path, word_file_path.Substring(0, word_file_path.LastIndexOf(".")) + ".zip");
            return word_file_path.Substring(0, word_file_path.LastIndexOf(".")) + ".zip";
        }

        private string GetPictureFromMedia(string rid)
        {
            if (rid == "")
                return "";
            XmlDocument xml_rel = new XmlDocument();
            xml_rel.Load(unzip_dir_path + @"\word\_rels\document.xml.rels");
            XmlNodeList rel_node_list = xml_rel.GetElementsByTagName("Relationship");
            string picture_path = "";

            foreach (XmlNode rel_node in rel_node_list)
            {
                if (rel_node.Attributes["Id"].Value == rid)
                {
                    picture_path = rel_node.Attributes["Target"].Value;
                    break;
                }
            }
            
            return picture_path;
        }

        private void DisplayQuestion()
        {
            if (current_question >= question_list.Count)
                return;
            if (current_question >= 1)
                btnBack.Enabled = true;
            else
                btnBack.Enabled = false;
            string question_text = question_list[current_question].QuestionText;
            pnlQuestionImages.Controls.Clear();
            lblQuestionNavigation.Text = (current_question + 1).ToString() + @"/" + question_list.Count;
            int i = 0;
            //if question contains image(s)
            if (question_list[current_question].Images.Count > 0)
            {
                foreach (string image_name in question_list[current_question].Images)
                {
                    PictureBox question_pic = new PictureBox();
                    question_pic.Image = Image.FromFile(unzip_dir_path + @"\word\" + image_name);
                    question_pic.Top = pnlQuestionArea.Top + 10;
                    question_pic.Left = pnlQuestionImages.Left + 10 + i * 200;
                    question_pic.Height = 100;
                    question_pic.Width = 200;
                    question_pic.SizeMode = PictureBoxSizeMode.StretchImage;
                    pnlQuestionImages.Controls.Add(question_pic);
                    i++;
                }
            }
            else
                pnlQuestionImages.Height = 10;
            
            txtQuestion.Text = question_text;
            DisplayAnswers();
        }

        private void DisplayAnswers()
        {
            int top = 1;
            List<Answer> answer_list = question_list[current_question].AnswerList;
            pnlAnswerArea.Controls.Clear();
            if (answer_list.Count == 0)
            {
                pnlAnswerArea.Text = "No answers are given for this question";
                return;
            }
            pnlAnswerArea.Text = "";
            
            pnlAnswerArea.Font = DefaultFont;
            
            foreach (Answer ans in answer_list)
            {

                RadRadioButton rdbAnswer = new RadRadioButton();
                rdbAnswer.Tag = ans.AnswerNo;
                rdbAnswer.Text = ans.AnswerText;
                rdbAnswer.Top = top * 30;
                rdbAnswer.Left = pnlAnswerArea.Left + 10;
                rdbAnswer.AutoSize = true;
                
                rdbAnswer.CheckStateChanged += new System.EventHandler(this.rdbAnswer_Click);
                if (test_status == 1)
                {
                    if (question_list[current_question].CorrectAnswerNo == Convert.ToInt32(rdbAnswer.Tag))
                        rdbAnswer.ForeColor = Color.Green;
                    if (question_list[current_question].SelectedAnswerNo == ans.AnswerNo)
                        if (question_list[current_question].Result)
                            rdbAnswer.ForeColor = Color.Green;
                        else
                            rdbAnswer.ForeColor = Color.Red;
                    rdbAnswer.ReadOnly = true;
                }
                pnlAnswerArea.Controls.Add(rdbAnswer);
                top++;

            }

            if (current_question == question_list.Count - 1)
            {
                btnNext.Text = "Review";
                test_status = 1;
                //btnRestartTest.Visible = true;
            }
        }

        private int GetAnswerNo(string answer)
        {
            answer = answer.Replace(" ", "") + " ";
            if (answer == "" || answer.Contains("img"))
                return 0;

            char answer_char = answer.Split(':')[1][0];
            return answer_char - 64;
        }

        private void ProcessFile()
        {
            AddWordNameSpaces(docx_file_xml.NameTable);
            XmlNodeList para_nodes = docx_file_xml.GetElementsByTagName("w:p");
            List<XmlNode> text_picture_nodes = new List<XmlNode>();
            foreach (XmlNode para in para_nodes)
            {
                XmlNodeList text_nodes = para.SelectNodes(".//w:t", mgr);
                foreach (XmlNode text_node in text_nodes)
                {
                    if (text_node.InnerText.Replace(" ", "") != "")
                        text_picture_nodes.Add(text_node);
                }

                XmlNodeList picture_nodes = para.SelectNodes(".//a:blip[@r:embed]", mgr);
                foreach (XmlNode picture_node in picture_nodes)
                    text_picture_nodes.Add(picture_node);
            }

            int i = 0;
            
            question_list = new List<Question>();
            int current_state = 0;
            string[] question_set = new string[4];
            question_set[0] = "";
            question_set[1] = "";
            question_set[2] = "";
            question_set[3] = "";
            int qno = 1;
            string question_text = "";
            string answer = "";
            string answer_text_list = "";
            string explain = "";
            List<string> picture_names = new List<string>();

            while (!text_picture_nodes[i].InnerText.ToLower().StartsWith("question"))
                i++;

            while(i < text_picture_nodes.Count)
            {
                current_state = GetState(text_picture_nodes[i].InnerText, current_state);
                
                string current_value = "";

                if (text_picture_nodes[i].OuterXml.Contains("embed"))
                {
                    string rid = text_picture_nodes[i].Attributes["r:embed"].Value;
                    string picture_path = GetPictureFromMedia(rid);
                    if (picture_path != "")
                        picture_names.Add(picture_path);
                }
                
                else
                    current_value = text_picture_nodes[i].InnerText;

                if (question_set[0].Contains("Question:") && current_value == "Question: ")
                {
                    Question question = new Question();
                    List<Answer> answer_list = GetAnswerList(question_set[1]);
                    

                    question.QuestionNo = qno++;
                    question.QuestionText = question_set[0];
                    question.AnswerList = answer_list;
                    question.CorrectAnswerNo = GetAnswerNo(question_set[2]);

                    if (question.CorrectAnswerNo > answer_list.Count)
                        question.CorrectAnswerText = question_set[2];
                    else if (question.CorrectAnswerNo < 0)
                        question.CorrectAnswerText = "Invalid Answer";
                    else
                        question.CorrectAnswerText = answer_list[question.CorrectAnswerNo - 1].AnswerText;

                    question.Explaination = question_set[3];
                    question.Images = picture_names;
                    question_list.Add(question);
                    question_set = new string[4];
                    question_set[0] = "";
                    question_set[1] = "";
                    question_set[2] = "";
                    question_set[3] = "";
                    picture_names = new List<string>();
                }
                question_set[current_state] += current_value;
                i++;
            }
        }

        private int GetState(string value, int current_state)
        {
            if (value.StartsWith("Question:"))
                return 0;
            if (value.StartsWith("A. "))
                return 1;
            if (value.StartsWith("Answer:"))
                return 2;
            if (value.StartsWith("Explanation:"))
                return 3;

            return current_state;
        }

        private List<Answer> GetAnswerList(string answers)
        {
            if (answers == "")
                return new List<Answer>();

            string[] answer_text_list = Regex.Split(answers, "[A-Z]+\\.\\s");
            int ans_no = 1;
            List<Answer> answer_list = new List<Answer>();
            foreach (string answer in answer_text_list)
            {
                if (answer != "")
                {
                    Answer ans = new Answer();
                    ans.AnswerNo = ans_no++;
                    ans.AnswerText = answer;
                    answer_list.Add(ans);
                }
            }
            return answer_list;
        }

        private void ExtractWordFile()
        {
            if (Directory.Exists(unzip_dir_path))
                return;
            Stream zip_stream = File.OpenRead(zip_file_path);

            Directory.CreateDirectory(unzip_dir_path);
            Directory.CreateDirectory(unzip_dir_path + @"\_rels");
            Directory.CreateDirectory(unzip_dir_path + @"\docProps");
            Directory.CreateDirectory(unzip_dir_path + @"\word");
            Directory.CreateDirectory(unzip_dir_path + @"\word\_rels");
            Directory.CreateDirectory(unzip_dir_path + @"\word\media");
            Directory.CreateDirectory(unzip_dir_path + @"\word\theme");
            ZipArchive zip_arch = new ZipArchive(zip_stream);

            foreach (ZipArchiveEntry compress_file in zip_arch.Entries)
            {
                FileStream fs = File.Create(unzip_dir_path + @"\" + compress_file.FullName);
                byte[] data = new byte[compress_file.Length];
                compress_file.Open().Read(data, 0, int.Parse(compress_file.Length.ToString()));
                fs.Write(data, 0, data.Length);
                fs.Close();
            }
        }

        private void AddWordNameSpaces(XmlNameTable name_table)
        {
            mgr = new XmlNamespaceManager(name_table);
            mgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            mgr.AddNamespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            mgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            mgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            mgr.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        }

        private void rdbAnswer_Click(object sender, EventArgs e)
        {
            RadRadioButton rdbAnswer = (RadRadioButton) sender;
            if (rdbAnswer.IsChecked)
            {
                question_list[current_question].SelectedAnswerNo = Convert.ToInt32(rdbAnswer.Tag.ToString());
                if (question_list[current_question].CorrectAnswerNo == question_list[current_question].SelectedAnswerNo)
                { 
                    question_list[current_question].Result = true;
                    CorrectAnswer++;
                }
                else
                    question_list[current_question].Result = false;
            }

            if (current_question == question_list.Count - 1)
                current_question = -1;

        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            current_question--;
            DisplayQuestion();
        }

    }
}