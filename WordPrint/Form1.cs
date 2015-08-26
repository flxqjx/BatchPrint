using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Aspose.Words;
using System.Net;
using System.Runtime.InteropServices;
namespace WordPrint
{

    public partial class BatchPrint : Form
    {
        public BatchPrint()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
            if (!Directory.Exists(path))//判断是否存在
            {
                Directory.CreateDirectory(path);//创建新路径
            }
            List<string> printer = GetAllPrinter();
            foreach (var item in printer)
            {
                this.cmbPrinter.Items.Add(item);
            }
            m_SynFileInfoList = new List<string>();
            m_SynFileInfoList.AddRange(new string[]
            {"Z:\\20141229\\331286\\03169ad6-ad1e-464b-a300-193e66d99d75.jpg",
            "Z:\\20141229\\331286\\B.doc",
            "Z:\\20141229\\331286\\A.pdf"
            }
            );
           // m_SynFileInfoList.AddRange(new string[]
           // {@"file://///10.10.100.52/Resource/20141229/331286/03169ad6-ad1e-464b-a300-193e66d99d75.jpg",
           // @"file://///10.10.100.52/Resource/20141229/331286/B.doc",
           // @"file://///10.10.100.52/Resource/20141229/331286/A.pdf"
           // }
           //);
            
        }
        /// <summary>
        /// 文档path
        /// </summary>
        string path = Application.StartupPath + "\\请放入要打印的文件";
        bool isStart = false;
        bool pause = false;
        Thread t = null;
        WebClient client;
        //存放下载列表
        List<string> m_SynFileInfoList;
        //检测网络状态
        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int connectionDescription, int reservedValue);

        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            int isCanResult = IsCanPrint();
            if (isCanResult == 0)
            {
                LoadWord();
            }
            else if (isCanResult == 1)
            {
                MessageBox.Show("文件夹中不能有除doc/docx或jpg/png/jpge或pdf格式的其他文件！");
            }
            else if (isCanResult == 3)
            {
                MessageBox.Show("文件夹为空！请下载文件。");

            }
        }

        /// <summary>
        /// 检查文件格式
        /// </summary>
        /// <returns></returns>
        private int IsCanPrint()
        {
            int result = 0;
            DirectoryInfo mydir = new DirectoryInfo(path);
            FileSystemInfo[] fsis = mydir.GetFileSystemInfos();
            if (fsis.Count() <= 0)
            {
                return 3;
            }
            foreach (FileSystemInfo fsi in fsis)
            {
                if (fsi is FileInfo)
                {

                    string[] fileNamearr = fsi.Name.Split('.');
                    string fileType = fileNamearr.LastOrDefault().ToLower();
                    if ((fileType != "doc" && fileType != "docx") && (fileType != "jpg" && fileType != "png" && fileType != "jpge") && (fileType != "pdf"))
                    {
                        result = 1;
                        break;
                    }
                }
            }
            return result;
            //MessageBox.Show("有不支持的文件类型！请检查PrepareFile目录里面的文件");
        }

        /// <summary>
        /// 把文件加载到ListBox
        /// </summary>
        private void LoadWord()
        {
            listBoxFile.Items.Clear();
            DirectoryInfo mydir = new DirectoryInfo(path);
            foreach (FileSystemInfo fsi in mydir.GetFileSystemInfos())
            {
                if (fsi is FileInfo)
                {
                    FileInfo fi = (FileInfo)fsi;
                    listBoxFile.Items.Add(fi.Name);
                }
            }
        }

        /// <summary>
        /// 开始 启动一个线程
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, EventArgs e)
        {

            if (!string.IsNullOrEmpty(this.cmbPrinter.Text) && listBoxFile.Items.Count > 0)
            {
                isStart = true;
                //t = new Thread(new ThreadStart(PrintMain));
                //t.Start();
                string message = DateTime.Now.ToString("MM-dd HH:mm:ss") + "打印开始……";
                listBoxLog.Items.Add(message);
                LogManage.WriteLog(LogManage.LogFile.Trace, message);
                PrintPdf();
                this.btnStart.Enabled = false;
                this.btnBreak.Enabled = true;
                this.btnPause.Enabled = true;
            }
            else if (listBoxFile.Items.Count == 0)
            {
                MessageBox.Show("文件列表为空！");
            }
            else
            {
                MessageBox.Show("请选择一个打印机！");
                this.cmbPrinter.DroppedDown = true;
            }
        }

        /// <summary>
        /// 停止方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBreak_Click(object sender, EventArgs e)
        {
            isStart = false;
            this.btnStart.Enabled = true;
            this.btnBreak.Enabled = false;
            string message = DateTime.Now.ToString("MM-dd HH:mm:ss") + "程序终止……";
            listBoxLog.Items.Add(message);
            LogManage.WriteLog(LogManage.LogFile.Trace, message);
        }

        private void WritePauseLog()
        {
            while (pause)
            {
                DateTime now = DateTime.Now;
                now = now.AddMinutes(Convert.ToInt32(txtRestTime));
                string message = now.ToString("MM-dd HH:mm:ss") + "程序暂停中：" + now.Subtract(DateTime.Now).ToString();
                this.listBoxLog.Items.Add(message);
                LogManage.WriteLog(LogManage.LogFile.Trace, message);
            }
        }
        /// <summary>
        /// 把多个PDF文件、word文件、JPG/PNG/jpge图合并成一个PDF文档
        /// </summary>
        /// <param name="fileList">需要合并文件的完整路径列表</param>
        /// <param name="outMergeFile">输出文件完整路径</param>
        public void MergePDFFile(List<string> fileList, string outMergeFile)
        {
            PdfReader reader;
            PdfImportedPage newPage;
            PdfWriter pw;
            iTextSharp.text.Document document = new iTextSharp.text.Document();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(outMergeFile, FileMode.Create));
            document.Open();
            PdfContentByte cb = writer.DirectContent;


            foreach (var itemFile in fileList)
            {
                string fileName = Path.GetFileName(itemFile);
                var ext = Path.GetExtension(itemFile).ToLower();
                if (!File.Exists(itemFile))
                {

                    LogManage.WriteLog(LogManage.LogFile.Error, string.Format("文件打印合并__{0} 文件不存在", fileName));
                    continue;
                }
                FileInfo fInfo = new FileInfo(itemFile);
                if (fInfo.Length < 1)
                {

                    LogManage.WriteLog(LogManage.LogFile.Error, string.Format("文件打印合并__文件内容为空，无法打印，{0}", fileName));
                    return;
                }



                if (".pdf".Equals(ext))
                {
                    reader = new PdfReader(itemFile);
                    int iPageNum = reader.NumberOfPages;
                    for (int j = 1; j <= iPageNum; j++)
                    {
                        document.NewPage();
                        newPage = writer.GetImportedPage(reader, j);
                        cb.AddTemplate(newPage, 0, 0);
                    }
                }
                else if (".doc".Equals(ext) || ".docx".Equals(ext))
                {
                    Aspose.Words.Document doc = new Aspose.Words.Document(itemFile);
                    string newPdf = Path.GetDirectoryName(itemFile) + "\\" + Path.GetFileNameWithoutExtension(itemFile) + ".pdf";
                    doc.Save(newPdf, SaveFormat.Pdf);
                    reader = new PdfReader(newPdf);
                    int iPageNum = reader.NumberOfPages;
                    for (int j = 1; j <= iPageNum; j++)
                    {
                        document.NewPage();
                        newPage = writer.GetImportedPage(reader, j);
                        cb.AddTemplate(newPage, 0, 0);
                    }
                }
                else if (".jpg".Equals(ext) || ".jpge".Equals(ext) || ".png".Equals(ext))
                {
                    FileStream rf = new FileStream(itemFile, FileMode.Open, FileAccess.Read);
                    int size = (int)rf.Length;
                    byte[] imext = new byte[size];
                    rf.Read(imext, 0, size);
                    rf.Close();

                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(imext);

                    //调整图片大小，使之适合A4
                    var imgHeight = img.Height;
                    var imgWidth = img.Width;
                    if (img.Height > iTextSharp.text.PageSize.A4.Height)
                    {
                        imgHeight = iTextSharp.text.PageSize.A4.Height;
                    }

                    if (img.Width > iTextSharp.text.PageSize.A4.Width)
                    {
                        imgWidth = iTextSharp.text.PageSize.A4.Width;
                    }
                    img.ScaleToFit(imgWidth, imgHeight);

                    //调整图片位置，使之居中
                    img.Alignment = iTextSharp.text.Image.ALIGN_MIDDLE;

                    document.NewPage();
                    document.Add(img);
                }
                string message = DateTime.Now.ToString("MM-dd HH:mm:ss") + "正在合拼" + itemFile + "……";
                listBoxLog.Items.Add(message);
                LogManage.WriteLog(LogManage.LogFile.Trace, message);
                listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;
            }
            document.Close();
        }
        /// <summary>
        /// 打印合并
        /// </summary>
        private void PrintPdf()
        {
            try
            {
                DirectoryInfo mydir = new DirectoryInfo(path);
                FileSystemInfo[] fsis = mydir.GetFileSystemInfos();
                List<string> pdfList = new List<string>();
                foreach (FileSystemInfo fsi in fsis)
                {

                    if (fsi is FileInfo)
                    {
                        FileInfo fi = (FileInfo)fsi;
                        pdfList.Add(fi.FullName);

                    }

                }
                var mergeFilePath = string.Format("{0}/{1}.pdf", path, DateTime.Now.ToString("yyyyMMddHHmmss") + '(' + Guid.NewGuid() + ')');
                MergePDFFile(pdfList, mergeFilePath);
                PrintHelper printHelper = new PrintHelper();
                printHelper.PrindPdf(mergeFilePath, cmbPrinter.Text);
                listBoxLog.Items.Add(DateTime.Now.ToString("MM-dd HH:mm:ss") + "打印结束。");
                LogManage.WriteLog(LogManage.LogFile.Trace, DateTime.Now.ToString("MM-dd HH:mm:ss") + "打印结束。");
                this.btnStart.Enabled = true;
                this.btnBreak.Enabled = false;
                listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                LogManage.WriteLog(LogManage.LogFile.Error, ex.Message);
            }
        }
        /// <summary>
        /// 打印主线程方法
        /// </summary>
        private void PrintMain()
        {
            try
            {
                DirectoryInfo mydir = new DirectoryInfo(path);
                FileSystemInfo[] fsis = mydir.GetFileSystemInfos();
                int i = 0;
                foreach (FileSystemInfo fsi in fsis)
                {
                    if (i != 0 && i % Convert.ToInt32(txtCount.Text) == 0)
                    {
                        listBoxLog.Items.Add(DateTime.Now.ToString("MM-dd HH:mm:ss") + "已经打印" + i + "正在休息");
                        Thread.Sleep(Convert.ToInt32(txtRestTime.Text) * 1000);
                    }
                    i++;
                    if (!isStart)
                        break;
                    if (pause)
                    {
                        pause = false;
                    }
                    if (fsi is FileInfo)
                    {
                        FileInfo fi = (FileInfo)fsi;
                        PrintHelper printHelper = new PrintHelper();
                        printHelper.Print(fi.FullName, cmbPrinter.Text);
                        string message = DateTime.Now.ToString("MM-dd HH:mm:ss") + "正在打印" + fi.Name + "……";
                        listBoxLog.Items.Add(message);
                        LogManage.WriteLog(LogManage.LogFile.Trace, message);
                    }
                    listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;

                }
                listBoxLog.Items.Add(DateTime.Now.ToString("MM-dd HH:mm:ss") + "打印结束。");
                LogManage.WriteLog(LogManage.LogFile.Trace, DateTime.Now.ToString("MM-dd HH:mm:ss") + "打印结束。");
                this.btnStart.Enabled = true;
                this.btnBreak.Enabled = false;
                listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                LogManage.WriteLog(LogManage.LogFile.Error, ex.Message);
            }

        }

        private void btnPause_Click(object sender, EventArgs e)
        {
            WritePauseLog();
            pause = true;
            this.btnBreak.Enabled = false;
            this.btnStart.Enabled = true;
            this.btnPause.Enabled = false;
        }


        /// <summary>
        /// 获取所有打印机
        /// </summary>
        /// <returns></returns>
        public List<string> GetAllPrinter()
        {
            ManagementObjectCollection queryCollection;
            string _classname = "SELECT * FROM Win32_Printer";
            Dictionary<string, ManagementObject> dict = new Dictionary<string, ManagementObject>();
            ManagementObjectSearcher query = new ManagementObjectSearcher(_classname);
            queryCollection = query.Get();
            List<string> result = new List<string>();
            foreach (ManagementObject mo in queryCollection)
            {
                string oldName = mo["Name"].ToString();
                result.Add(oldName);
            }
            return result;
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            this.listBoxLog.Items.Clear();
        }

        private void BatchPrint_Load(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 检测网络状态
        /// </summary>
        private bool isConnected()
        {
            int I = 0;
            bool state = InternetGetConnectedState(out I, 0);
            return state;
        }

        private void btnDownLoad_Click(object sender, EventArgs e)
        {
            //判断网络连接是否正常
            if (isConnected())
            {
                //设置不可用
                btnDownLoad.Enabled = false;
                //设置最大活动线程数以及可等待线程数
                ThreadPool.SetMaxThreads(3, 3);
                //判断是否还存在任务
                if (m_SynFileInfoList.Count <= 0)
                {
                    MessageBox.Show("没有文件可下载！");
                }
                else
                {
                    this.pgrShow.Visible = true;
                    foreach (string m_SynFileInfo in m_SynFileInfoList)
                    {
                        //启动下载任务
                        StartDownLoad(m_SynFileInfo);
                        this.pgrShow.Value = 0;
                    }
                }

            }
            else
            {
                MessageBox.Show("网络异常!");
            }
        }

        #region 使用WebClient下载文件

        /// <summary>
        /// HTTP下载远程文件并保存本地的函数
        /// </summary>
        void StartDownLoad(string filePath)
        {

            //再次new 避免WebClient不能I/O并发 
            WebClient client = new WebClient();
            //异步下载
            client.DownloadProgressChanged += new DownloadProgressChangedEventHandler(client_DownloadProgressChanged);
            client.DownloadFileCompleted += new AsyncCompletedEventHandler(client_DownloadFileCompleted);
            client.DownloadFileAsync(new Uri(filePath), path + filePath.Substring(filePath.LastIndexOf("\\")), filePath);

        }

        /// <summary>
        /// 下载进度条
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void client_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            this.pgrShow.Value = e.ProgressPercentage;
        }

        /// <summary>
        /// 下载完成调用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void client_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            //到此则一个文件下载完毕
            string  m_SynFileInfo = (string)e.UserState;
            m_SynFileInfoList.Remove(m_SynFileInfo);
            if (m_SynFileInfoList.Count <= 0)
            {
                //此时所有文件下载完毕
                btnDownLoad.Enabled = true;
            }
        }

        #endregion

    }
}


