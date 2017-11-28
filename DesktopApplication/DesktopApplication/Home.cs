using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DesktopApplication
{
    public partial class Home : Form
    {
        // There is filesname list have dowload files names 
        private List<string> FilesName = new List<string>();
        Task[] tasks = null;

        public Home()
        {
            InitializeComponent();
            DownloadData();
            PriodicallyCall();
        }

        /// <summary>
        /// this function is use for call the download data after 5 mints if there is no more dowloads are available.
        /// </summary>
        public void PriodicallyCall()
        {
            var startTimeSpan = TimeSpan.Zero;
            var periodTimeSpan = TimeSpan.FromMinutes(1);

            var timer = new System.Threading.Timer((e) =>
            {
                if ((progressBar1.Value == 100) && (progressBar2.Value == 100) && (progressBar3.Value == 100) && (progressBar4.Value == 100) && (progressBar5.Value == 100) && (progressBar6.Value == 100) && (progressBar7.Value == 100) && (progressBar8.Value == 100) && (progressBar9.Value == 100) && (progressBar10.Value == 100))
                {
                    Home NewForm = new Home();
                    NewForm.Show();
                    this.Dispose(false);
                }
            }, null, startTimeSpan, periodTimeSpan);
        }
        
        public void DownloadData()
        {
            this.button1.Enabled = false;
            AisUriProviderApi.AisUriProvider aisUriProvider = new AisUriProviderApi.AisUriProvider();
            IEnumerable<Uri> DataIEnum = aisUriProvider.Get();
            IList list = DataIEnum.ToList();
            foreach (var item in list)
            {
                FilesName.Add(item.ToString().Substring(item.ToString().LastIndexOf('/') + 1));
            }
            int counter = list.Count;

            tasks = new Task[counter];
  
            tasks[0] = Task.Run(() =>
            {
                using (WebClient wc1 = new WebClient())
                {

                    wc1.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                    wc1.DownloadFileAsync(new System.Uri(list[0].ToString()), @"D:\Files\" +
                       list[0].ToString().Substring(list[0].ToString().LastIndexOf('/') + 1));
                    wc1.DownloadProgressChanged += new DownloadProgressChangedEventHandler(Progress1Changed);

                }

            });

            tasks[1] = Task.Run(() =>
        {
            using (WebClient wc2 = new WebClient())
            {

                wc2.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                wc2.DownloadFileAsync(new System.Uri(list[1].ToString()), @"D:\Files\" +
                   list[1].ToString().Substring(list[1].ToString().LastIndexOf('/') + 1));
                wc2.DownloadProgressChanged += new DownloadProgressChangedEventHandler(Progress2Changed);

            }

        });
            tasks[2] = Task.Run(() =>
            {
                using (WebClient wc3 = new WebClient())
                {

                    wc3.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                    wc3.DownloadFileAsync(new System.Uri(list[2].ToString()), @"D:\Files\" +
                       list[2].ToString().Substring(list[2].ToString().LastIndexOf('/') + 1));
                    wc3.DownloadProgressChanged += new DownloadProgressChangedEventHandler(Progress3Changed);

                }
                //   Task.WaitAll(tasks);
            });
            tasks[3] = Task.Run(() =>
            {
                using (WebClient wc4 = new WebClient())
                {

                    wc4.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                    wc4.DownloadFileAsync(new System.Uri(list[3].ToString()), @"D:\Files\" +
                       list[3].ToString().Substring(list[3].ToString().LastIndexOf('/') + 1));

                    wc4.DownloadProgressChanged += new DownloadProgressChangedEventHandler(Progress4Changed);

                }

            });
            tasks[4] = Task.Run(() =>
            {
                using (WebClient wc5 = new WebClient())
                {

                    wc5.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                    wc5.DownloadFileAsync(new System.Uri(list[4].ToString()), @"D:\Files\" +
                       list[4].ToString().Substring(list[4].ToString().LastIndexOf('/') + 1));

                    wc5.DownloadProgressChanged += new DownloadProgressChangedEventHandler(Progress5Changed);

                }
                //  
            });
            tasks[5] = Task.Run(() =>
            {
                using (WebClient wc6 = new WebClient())
                {

                    wc6.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                    wc6.DownloadFileAsync(new System.Uri(list[5].ToString()), @"D:\Files\" +
                       list[5].ToString().Substring(list[5].ToString().LastIndexOf('/') + 1));

                    wc6.DownloadProgressChanged += new DownloadProgressChangedEventHandler(Progress6Changed);

                }

            });
            tasks[6] = Task.Run(() =>
            {
                using (WebClient wc7 = new WebClient())
                {

                    wc7.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                    wc7.DownloadFileAsync(new System.Uri(list[6].ToString()), @"D:\Files\" +
                       list[6].ToString().Substring(list[6].ToString().LastIndexOf('/') + 1));

                    wc7.DownloadProgressChanged += new DownloadProgressChangedEventHandler(Progress7Changed);

                }

            });
            tasks[7] = Task.Run(() =>
            {
                using (WebClient wc8 = new WebClient())
                {

                    wc8.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                    wc8.DownloadFileAsync(new System.Uri(list[7].ToString()), @"D:\Files\" +
                       list[7].ToString().Substring(list[7].ToString().LastIndexOf('/') + 1));

                    wc8.DownloadProgressChanged += new DownloadProgressChangedEventHandler(Progress8Changed);
                }

            });
            tasks[8] = Task.Run(() =>
            {
                using (WebClient wc9 = new WebClient())
                {

                    wc9.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                    wc9.DownloadFileAsync(new System.Uri(list[8].ToString()), @"D:\Files\" +
                       list[8].ToString().Substring(list[8].ToString().LastIndexOf('/') + 1));

                    wc9.DownloadProgressChanged += new DownloadProgressChangedEventHandler(Progress9Changed);

                }

            });
            tasks[9] = Task.Run(() =>
            {
                using (WebClient wc10 = new WebClient())
                {

                    wc10.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                    wc10.DownloadFileAsync(new System.Uri(list[9].ToString()), @"D:\Files\" +
                       list[9].ToString().Substring(list[9].ToString().LastIndexOf('/') + 1));
                    wc10.DownloadProgressChanged += new DownloadProgressChangedEventHandler(Progress10Changed);

                }

            });
            //Task.WhenAll(tasks);



        }

        #region Progress Bar and with DownlodfileCompleted.


        private void Progress10Changed(object sender, DownloadProgressChangedEventArgs e)
        {
            this.progressBar10.Invoke(new Action(() =>
            {
                progressBar10.Value = e.ProgressPercentage;
            }));
            if (this.progressBar10.Value == this.progressBar10.Maximum)
            {
                this.progressBar10.Invoke(new Action(() =>
                {
                    progressBar10.Visible = false;
                }));
                ViewFiles(9, this.progressBar10.Location);
            }
        }

        private void Progress8Changed(object sender, DownloadProgressChangedEventArgs e)
        {
            this.progressBar8.Invoke(new Action(() =>
            {
                progressBar8.Value = e.ProgressPercentage;
            }));
            if (this.progressBar8.Value == this.progressBar8.Maximum)
            {
                this.progressBar8.Invoke(new Action(() =>
                {
                    progressBar8.Visible = false;
                }));
                ViewFiles(7, this.progressBar8.Location);
            }
        }

        private void Progress9Changed(object sender, DownloadProgressChangedEventArgs e)
        {
            this.progressBar9.Invoke(new Action(() =>
            {
                progressBar9.Value = e.ProgressPercentage;
            }));
            if (this.progressBar9.Value == this.progressBar9.Maximum)
            {
                this.progressBar9.Invoke(new Action(() =>
                {
                    progressBar9.Visible = false;
                }));
                ViewFiles(8, this.progressBar9.Location);
            }
        }

        private void Progress7Changed(object sender, DownloadProgressChangedEventArgs e)
        {
            this.progressBar7.Invoke(new Action(() =>
            {
                progressBar7.Value = e.ProgressPercentage;
            }));
            if (this.progressBar7.Value == this.progressBar7.Maximum)
            {
                this.progressBar7.Invoke(new Action(() =>
                {
                    progressBar7.Visible = false;
                }));
                ViewFiles(6, this.progressBar7.Location);
            }
        }

        private void Progress6Changed(object sender, DownloadProgressChangedEventArgs e)
        {
            this.progressBar6.Invoke(new Action(() =>
            {
                progressBar6.Value = e.ProgressPercentage;
            }));
            if (this.progressBar6.Value == this.progressBar6.Maximum)
            {
                this.progressBar6.Invoke(new Action(() =>
                {
                    progressBar6.Visible = false;
                }));
                ViewFiles(5, this.progressBar6.Location);
            }
        }

        private void Progress5Changed(object sender, DownloadProgressChangedEventArgs e)
        {
            this.progressBar5.Invoke(new Action(() =>
            {
                progressBar5.Value = e.ProgressPercentage;
            }));
            if (this.progressBar5.Value == this.progressBar5.Maximum)
            {
                this.progressBar5.Invoke(new Action(() =>
                {
                    progressBar5.Visible = false;
                }));
                ViewFiles(4, this.progressBar5.Location);
            }
        }

        private void Progress4Changed(object sender, DownloadProgressChangedEventArgs e)
        {
            this.progressBar4.Invoke(new Action(() =>
            {
                progressBar4.Value = e.ProgressPercentage;
            }));
            if (this.progressBar4.Value == this.progressBar4.Maximum)
            {
                this.progressBar4.Invoke(new Action(() =>
                {
                    progressBar4.Visible = false;
                }));
                ViewFiles(3, this.progressBar4.Location);
            }
        }

        private void Progress3Changed(object sender, DownloadProgressChangedEventArgs e)
        {
            this.progressBar3.Invoke(new Action(() =>
            {
                progressBar3.Value = e.ProgressPercentage;
            }));
            if (this.progressBar3.Value == this.progressBar3.Maximum)
            {
                this.progressBar3.Invoke(new Action(() =>
                {
                    progressBar3.Visible = false;
                }));
                ViewFiles(2, this.progressBar3.Location);
            }
        }

        private void Progress2Changed(object sender, DownloadProgressChangedEventArgs e)
        {
            this.progressBar2.Invoke(new Action(() =>
            {
                progressBar2.Value = e.ProgressPercentage;
            }));
            if (this.progressBar2.Value == this.progressBar2.Maximum)
            {
                this.progressBar2.Invoke(new Action(() =>
                {
                    progressBar2.Visible = false;
                }));
                ViewFiles(1, this.progressBar2.Location);
            }
        }

        private void Progress1Changed(object sender, DownloadProgressChangedEventArgs e)
        {
            this.progressBar1.Invoke(new Action(() =>
            {
                progressBar1.Value = e.ProgressPercentage;

            }));

            if (this.progressBar1.Value == this.progressBar1.Maximum)
            {
                this.progressBar1.Invoke(new Action(() =>
                {
                    progressBar1.Visible = false;
                }));
                ViewFiles(0, this.progressBar1.Location);
            }
        }

        private void wc_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if ((progressBar1.Value == 100) && (progressBar2.Value == 100) && (progressBar3.Value == 100) && (progressBar4.Value == 100) && (progressBar5.Value == 100) && (progressBar6.Value == 100) && (progressBar7.Value == 100) && (progressBar8.Value == 100) && (progressBar9.Value == 100) && (progressBar10.Value == 100))
            {
                this.progressBar1.Invoke(new Action(() =>
                {
                    button1.Enabled = true;
                }));
            }
        } 
        #endregion

        /// <summary>
        /// we show only doc,pdf, css, txt & js files text only with a file name. 
        /// other are showing only file name.  when the download is complete then we call this function
        /// </summary>
        /// <param name="index"> </param>
        /// <param name="p"></param>

        private void ViewFiles(int index, Point p)
        {
            string FileExtsion = System.IO.Path.GetExtension(FilesName[index].ToString());
            try
            {
                if (FileExtsion == ".jpg" || FileExtsion == ".png" || FileExtsion == ".jpeg")
                {
                    try
                    {
                        var picture = new PictureBox
                        {
                            Name = "pictureBox",
                            Size = new Size(80, 80),
                            Image = Image.FromFile(@"D:\Files\" + FilesName[index].ToString()),
                            Location = new Point(p.X, p.Y),
                            SizeMode = PictureBoxSizeMode.StretchImage,
                        };
                        Invoke(new Action(() =>
                        {
                            this.Controls.Add(picture);
                        }));
                    }
                    catch (Exception)
                    {

                        // throw;
                    }
                }
                else if (FileExtsion == ".pdf")
                {
                    try
                    {
                        string response = GetTextFromPDF(FilesName[index].ToString());
                        if (response.Length > 0)
                        {
                            Label label = new Label();
                            label.Location = p;
                            label.Text = "File Name: " + FilesName[index].ToString() + " Content: " + response.Substring(0, 5);
                            label.Width = 250;
                            Invoke(new Action(() =>
                            {
                                this.Controls.Add(label);
                            }));
                        }
                        else
                        {
                            Label label = new Label();
                            label.Text = "File Name: " + FilesName[index].ToString() + " Content: File is Empty";
                            label.Location = p;
                            label.Width = 250;
                            label.ForeColor = Color.Red;
                            Invoke(new Action(() =>
                            {
                                this.Controls.Add(label);
                            }));
                        }
                    }
                    catch (Exception)
                    {

                        //  throw;
                    }
                }
                else if (FileExtsion == ".doc" || FileExtsion == ".docx")
                {
                    try
                    {
                        string response = GetTextFromWord(FilesName[index].ToString());
                        if (response.Length > 0)
                        {
                            Label label = new Label();
                            label.Location = p;
                            string text = "File Name: " + FilesName[index].ToString() + " Content: " + response.Substring(0, 5);
                            label.Text = text;
                            label.Width = 250;
                            Invoke(new Action(() =>
                            {
                                this.Controls.Add(label);
                            }));
                        }
                        else
                        {
                            Label label = new Label();
                            label.Text = "File Name: " + FilesName[index].ToString() + " Content: File is Empty";
                            label.ForeColor = Color.Red;
                            label.Location = p;
                            label.Width = 250;
                            label.ForeColor = Color.Red;
                            Invoke(new Action(() =>
                            {
                                this.Controls.Add(label);
                            }));
                        }
                    }
                    catch (Exception)
                    {

                        //  throw;
                    }
                }
                else if (FileExtsion == ".css" || FileExtsion == ".txt" || FileExtsion == ".js")
                {
                    try
                    {
                        string response = System.IO.File.ReadAllText(@"D:\Files\" + FilesName[index].ToString());
                        if (response.Trim().Length > 0)
                        {
                            Label label = new Label();
                            label.Location = p;
                            string text = "File Name: " + FilesName[index].ToString() + " Content: " + response.Substring(0, 5);
                            label.Text = text;
                            label.Width = 250;
                            Invoke(new Action(() =>
                            {
                                this.Controls.Add(label);
                            }));
                        }
                        else
                        {
                            Label label = new Label();
                            label.Text = "File Name: " + FilesName[index].ToString() + " Content: File is Empty";
                            label.Location = p;
                            label.ForeColor = Color.Red;
                            label.Width = 250;
                            Invoke(new Action(() =>
                            {
                                this.Controls.Add(label);
                            }));
                        }
                    }
                    catch (Exception)
                    {

                        // throw;
                    }
                }
                else
                {
                    try
                    {
                        Label label = new Label();
                        label.Location = p;
                        label.Text = "File Name " + FilesName[index].ToString();
                        Invoke(new Action(() =>
                        {
                            this.Controls.Add(label);
                        }));

                    }
                    catch (Exception)
                    {

                        // throw;
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Exception" + ex.Message.ToString());
            }
        }

        /// <summary>
        /// read a data from word.
        /// </summary>
        /// <param name="FilePath">code pass a path to this function</param>
        /// <returns></returns>
        private string GetTextFromWord(string FilePath)
        {
            try
            {
                StringBuilder text = new StringBuilder();
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                object miss = System.Reflection.Missing.Value;
                object path = FilePath;
                object readOnly = true;
                Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

                for (int i = 0; i < docs.Paragraphs.Count; i++)
                {
                    text.Append(" \r\n " + docs.Paragraphs[i + 1].Range.Text.ToString());
                }
                return text.ToString().Trim();
            }
            catch (Exception ex)
            {
                return "";
            }

        }
        /// <summary>
        /// read a data from pdf.
        /// </summary>
        /// <param name="FilePath">code pass a path to this function</param>
        /// <returns></returns>
        private string GetTextFromPDF(string path)
        {
            StringBuilder text = new StringBuilder();
            try
            {

                using (PdfReader reader = new PdfReader(path))
                {
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                    }
                }
                return text.ToString().Trim();
            }
            catch (Exception)
            {

                return "";
            }


        }
        /// <summary>
        /// Delete all unanccessary files
        /// </summary>
        public void DeleteAllFilesInFolder()
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(@"d:\Files");

            foreach (FileInfo file in di.GetFiles())
            {
                try
                {
                    if (file.Exists)
                        file.Delete();
                }
                catch (Exception ex)
                {

                }
            }
        }

        /// <summary>
        /// this is used for sysnc all button click.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {

            Home NewForm = new Home();
            NewForm.Show();
            this.Dispose(false);
        }

        private void Home_Load(object sender, EventArgs e)
        {

        }
        /// <summary>
        ///  this function use when user close the form while downloading
        /// </summary>
        /// <param name="e"></param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if ((progressBar1.Value == 100) && (progressBar2.Value == 100) && (progressBar3.Value == 100) && (progressBar4.Value == 100) && (progressBar5.Value == 100) && (progressBar6.Value == 100) && (progressBar7.Value == 100) && (progressBar8.Value == 100) && (progressBar9.Value == 100) && (progressBar10.Value == 100))
            {

                base.OnFormClosing(e);
            }
            else
            {
                DialogResult result1 = MessageBox.Show("Are you sure? ",
               "Alert", MessageBoxButtons.YesNo);
                if (result1.ToString().Contains("Yes"))
                {
                    base.OnFormClosing(e);
                }
                else
                {
                    e.Cancel = true;
                    base.OnFormClosing(e);
                }
            }
        }

       
    }


}