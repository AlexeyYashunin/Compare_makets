using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading.Tasks;
using System.Threading;

namespace Compare_makets {
    public partial class Form1 : Form {
        public string firstFile = "";
        public string secondFile = "";

        public TFile tFirstFile;
        public TFile tSecondFile;

        public bool bFirstFileSet=false;
        public bool bSecondFileSet=false;

        //создаем список задач
        public System.Threading.Tasks.Task[] wordTasks = null;

        public Form1() {
            InitializeComponent();

            wordTasks = new System.Threading.Tasks.Task[2];
            
            //настраиваем worker'а
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;

            //устанавливаем предел прогресс-бара
            //progressBar1.Maximum = 100;

            /*var posFirst = this.PointToScreen(label1.Location);
            posFirst = pbFirstFile.PointToClient(posFirst);

            label1.Parent = pbFirstFile;
            label1.Location = posFirst;
            //label1.ForeColor = Color.Black;
            label1.BackColor = Color.Transparent;

            var posSecond= this.PointToScreen(label2.Location);
            posSecond = pbSecondFile.PointToClient(posSecond);

            label2.Parent = pbSecondFile;
            label2.Location = posSecond;
            //label2.ForeColor = Color.Black;
            label2.BackColor = Color.Transparent;*/
        }

        private void Form1_DragDrop(object sender, DragEventArgs e) {
            int x = this.PointToClient(new System.Drawing.Point(e.X, e.Y)).X;
            int y = this.PointToClient(new System.Drawing.Point(e.X, e.Y)).Y;

            if(x >= pbFirstFile.Location.X && x <= pbFirstFile.Location.X + pbFirstFile.Width && y >= pbFirstFile.Location.Y && y <= pbFirstFile.Location.Y + pbFirstFile.Height) {
                firstFile = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];

                //pbFirstFile.Image = Properties.Resources._fill_file;
                pbFirstFile.Image = Properties.Resources._source_Certsys_filled;
            }

            if(x >= pbSecondFile.Location.X && x <= pbSecondFile.Location.X + pbSecondFile.Width && y >= pbSecondFile.Location.Y && y <= pbSecondFile.Location.Y + pbSecondFile.Height) {
                secondFile = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];

                //pbSecondFile.Image = Properties.Resources._fill_file;
                pbSecondFile.Image = Properties.Resources._source_FGIS_filled;
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e) {
            e.Effect = DragDropEffects.Move;
        }

        private void btnCompare_Click(object sender, EventArgs e) {
            if(backgroundWorker1.IsBusy){
                return;
            }

            //запускаем worker'а
            backgroundWorker1.RunWorkerAsync();
           
            /*if(firstFile=="" || secondFile=="") {
                return;
            }

            var taskOne = System.Threading.Tasks.Task.Factory.StartNew(() => {
                tFirstFile = new TFile(firstFile, "Certsys" ,file_type.WORD);
                tFirstFile.Start();
                //progressBar1.Maximum = tFirstFile.getRowsSum();
            });

            var taskTwo = System.Threading.Tasks.Task.Factory.StartNew(() => {
                tSecondFile = new TFile(secondFile, "FGIS" ,file_type.WORD);
                //progressBar1.Maximum = tSecondFile.getRowsSum();
            });

            taskOne.Wait();
            taskTwo.Wait();

            tFirstFile.Compare(tSecondFile);
            //tSecondFile.Compare(tFirstFile);*/
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e) {
            try{
                backgroundWorker1.CancelAsync();
            }catch(System.Exception ex){
            }

            try{         
                tFirstFile.Destroy();
                tSecondFile.Destroy(); 
            }catch(System.Exception ex){            	
            }                       
        }

        private void Form1_Load(object sender, EventArgs e) {
            //удаляем предыдущие выходные данные
            string curDirectory=System.IO.Directory.GetCurrentDirectory();
            string outDataDirectory = curDirectory + "\\" + "[OUT]Data";

            if (System.IO.Directory.Exists(outDataDirectory)==true) {
                try{
                    string[] files = System.IO.Directory.GetFiles(outDataDirectory, "*", System.IO.SearchOption.AllDirectories);

                    foreach(string file in files){
                        killTaskOfFile(file);
                    }

                    //удаляем директорию с файлами
                    System.IO.Directory.Delete(outDataDirectory,true);                    
                }catch(System.Exception ex){
                    MessageBox.Show("Can't delete old data files\nKill \"word.exe\" processes associated with files(...\\[OUT]Data\\)","Error");
                    
                    System.Environment.Exit(0);
                }
            }
        }

        private void killTaskOfFile(string fileName) {
            System.Diagnostics.Process tool = new System.Diagnostics.Process();

            string curDirectory = System.IO.Directory.GetCurrentDirectory();

            string noRootFileName = fileName.Substring(System.IO.Path.GetPathRoot(fileName).Length);

            tool.StartInfo.FileName = curDirectory+"\\"+"[Kill]Handle"+"\\"+"handle.exe";
            tool.StartInfo.Arguments = noRootFileName + " /accepteula";
            tool.StartInfo.UseShellExecute = false;
            tool.StartInfo.RedirectStandardOutput = true;
            tool.Start();
            tool.WaitForExit();

            string outputTool = tool.StandardOutput.ReadToEnd();

            string matchPattern = @"(?<=\s+pid:\s+)\b(\d+)\b(?=\s+)";

            foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(outputTool, matchPattern)) {
                System.Diagnostics.Process.GetProcessById(int.Parse(match.Value)).Kill();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e) {
            if (firstFile == "" || secondFile == "") {
                return;
            }
            
            var taskOne = System.Threading.Tasks.Task.Factory.StartNew(() => {
                try{
                    tFirstFile = new TFile(this, firstFile, "Certsys", file_type.WORD);
                    tFirstFile.Start();
                    //throw new IndexOutOfRangeException();
                }catch(System.Exception ex){
                    MessageBox.Show("Can't open " + System.IO.Path.GetFileName(firstFile) + " file\nKill \"word.exe\" process associated with that file","Error");
                    
                    System.Environment.Exit(0);
                }
                
                //progressBar1.Maximum = tFirstFile.getRowsSum();
            });           

            var taskTwo = System.Threading.Tasks.Task.Factory.StartNew(() => {
                try {
                    tSecondFile = new TFile(this, secondFile, "FGIS", file_type.WORD);
                } catch (System.Exception ex) {
                    MessageBox.Show("Can't open " + System.IO.Path.GetFileName(secondFile) + "file\nKill \"word.exe\" process associated with that file", "Error");

                    System.Environment.Exit(0);
                }
                
                //progressBar1.Maximum = tSecondFile.getRowsSum();
            });

            taskOne.Wait();
            taskTwo.Wait();

            try{
                tFirstFile.Compare(tSecondFile);
            }catch(System.Exception ex){
                //MessageBox.Show("Can't open output word file(s)\nKill \"word.exe\" process(es) associated with that file(s)", "Error");
                MessageBox.Show(ex.Message, "Error");

                try {
                    backgroundWorker1.CancelAsync();
                } catch(System.Exception inEx) {
                }

                try {
                    tFirstFile.Destroy();
                    tSecondFile.Destroy();
                } catch(System.Exception inEx) {
                }

                System.Environment.Exit(0);
            }

            pbFirstFile.Image = Properties.Resources._source_Certsys;
            bFirstFileSet = false;

            pbSecondFile.Image = Properties.Resources._source_FGIS;
            bSecondFileSet = false;

            backgroundWorker1.ReportProgress(0);

            //tSecondFile.Compare(tFirstFile);
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            progressBar1.Value = e.ProgressPercentage;
            labelPBar.Text = e.ProgressPercentage.ToString()+" %";
        }             

        private void Form1_MouseDoubleClick(object sender, MouseEventArgs e) {
            if(backgroundWorker1.IsBusy == true) {
                return;
            }

            /*int x = this.PointToClient(new System.Drawing.Point(e.X, e.Y)).X;
            int y = this.PointToClient(new System.Drawing.Point(e.X, e.Y)).Y;*/

            int x = e.X;
            int y = e.Y;

            if(x >= pbFirstFile.Location.X && x <= pbFirstFile.Location.X + pbFirstFile.Width && y >= pbFirstFile.Location.Y && y <= pbFirstFile.Location.Y + pbFirstFile.Height) {
                if(bFirstFileSet==true){
                    pbFirstFile.Image = Properties.Resources._source_Certsys;
                    bFirstFileSet = false;
                }

                openFileDialog1.ShowDialog();

                firstFile = openFileDialog1.FileName;

                pbFirstFile.Image = Properties.Resources._source_Certsys_filled;

                bFirstFileSet = true;
            }

            if(x >= pbSecondFile.Location.X && x <= pbSecondFile.Location.X + pbSecondFile.Width && y >= pbSecondFile.Location.Y && y <= pbSecondFile.Location.Y + pbSecondFile.Height) {
                if(bSecondFileSet == true) {
                    pbSecondFile.Image = Properties.Resources._source_FGIS;
                    bSecondFileSet = false;
                }

                openFileDialog2.ShowDialog();

                secondFile = openFileDialog2.FileName;

                pbSecondFile.Image = Properties.Resources._source_FGIS_filled;

                bSecondFileSet = true;
            }
        }
    }
}