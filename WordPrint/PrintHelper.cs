using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Reflection;
using System.IO;
using System.Management;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
namespace WordPrint
{
    class PrintHelper
    {


        public void PrindPdf(string mergeFilePath, string printername)
        {
            try
            {
              

                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.CreateNoWindow = false;
                proc.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                proc.StartInfo.UseShellExecute = true;
                proc.StartInfo.FileName = mergeFilePath;//打印文件路径（本地完整路径包括文件名和后缀名）
                proc.StartInfo.Verb = "print";
                string defaultPrinter = GetDefaultPrinter();
                SetDefaultPrinter(printername);
                proc.Start();
                proc.Close();
            }
            catch (Exception ex)
            {
                LogManage.WriteLog(LogManage.LogFile.Error, ex.Message);
            }
        }
        //（通过添加引用-com组件，找office的组件

        /// <summary>
        /// 打印
        /// </summary>
        /// <param name="filepath">文件路径</param>
        /// <param name="printername">指定的打印机</param>
        public void Print(string filepath, string printername)
        {
            
            if (string.IsNullOrEmpty(printername))
            {
                printername = "Microsoft XPS Document Writer";
            }
            try
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                //不现实调用程序窗口,但是对于某些应用无效
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;

                //采用操作系统自动识别的模式
                p.StartInfo.UseShellExecute = true;

                //要打印的文件路径
                p.StartInfo.FileName = filepath;

                //指定执行的动作，是打印，即print，打开是 open
                p.StartInfo.Verb = "print";

                //获取当前默认打印机

                string defaultPrinter = GetDefaultPrinter();

                //将指定的打印机设为默认打印机
                SetDefaultPrinter(printername);

                //开始打印
                p.Start();

                //等待1秒 
                p.WaitForExit(1000);

                //将默认打印机还原
                SetDefaultPrinter(defaultPrinter);
                p.Close();
            }
            catch (Exception ex)
            {

                LogManage.WriteLog(LogManage.LogFile.Error, ex.Message);
            }

        }
        [DllImport("Winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetDefaultPrinter(string printerName);

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]

        private static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int pcchBuffer);
        /// <summary>
        /// 获取默认的打印机
        /// </summary>
        /// <returns></returns>
        static string GetDefaultPrinter()
        {
            const int ERROR_FILE_NOT_FOUND = 2;

            const int ERROR_INSUFFICIENT_BUFFER = 122;

            int pcchBuffer = 0;

            if (GetDefaultPrinter(null, ref pcchBuffer))
            {
                return null;
            }

            int lastWin32Error = Marshal.GetLastWin32Error();

            if (lastWin32Error == ERROR_INSUFFICIENT_BUFFER)
            {
                StringBuilder pszBuffer = new StringBuilder(pcchBuffer);
                if (GetDefaultPrinter(pszBuffer, ref pcchBuffer))
                {
                    return pszBuffer.ToString();
                }

                lastWin32Error = Marshal.GetLastWin32Error();
            }
            if (lastWin32Error == ERROR_FILE_NOT_FOUND)
            {
                return null;
            }

            throw new Win32Exception(Marshal.GetLastWin32Error());


        }
        #region
        ///打印页面不会闪动
        public void PrintMethodOther(string filepath, string printername)
        {
            try
            {
                object wordFile = filepath;
                
                object oMissing = Missing.Value;
                //自定义object类型的布尔值
                object oTrue = true;
                object oFalse = false;

                object doNotSaveChanges = WdSaveOptions.wdDoNotSaveChanges;

                //定义WORD Application相关

                Application appWord = new Application();

                //WORD程序不可见
                appWord.Visible = false;
                //不弹出警告框
                appWord.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                //先保存默认的打印机
                string defaultPrinter = appWord.ActivePrinter;

                //打开要打印的文件
                //如果在其它函数调用中出错（doc为null),处理办法：建立临时文件夹，还要设置服务的权限属性
                Microsoft.Office.Interop.Word.Document doc = appWord.Documents.Open(
                        ref wordFile,
                        ref oMissing,
                        ref oTrue,
                        ref oFalse,
                        ref oMissing,
                        ref oMissing,
                        ref oMissing,
                        ref oMissing,
                        ref oMissing,
                        ref oMissing,
                        ref oMissing,
                        ref oMissing,
                        ref oMissing,
                        ref oMissing,
                        ref oMissing,
                        ref oMissing);

                //设置指定的打印机
                appWord.ActivePrinter = printername;
                //"Microsoft XPS Document Writer";

                //打印

                doc.PrintOut(
                    ref oTrue, //此处为true,表示后台打印
                    ref oFalse,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing,
                    ref oMissing
                    );

                //打印完关闭WORD文件
                doc.Close(ref doNotSaveChanges, ref oMissing, ref oMissing);

                //还原原来的默认打印机
                appWord.ActivePrinter = defaultPrinter;

                //退出WORD程序
                appWord.Quit(ref oMissing, ref oMissing, ref oMissing);

                doc = null;
                appWord = null;
            }
            catch (Exception ex)
            {
                //help.LogMessage(filepath + "----" + printername + "-------" + ex.Message);
            }
        }


        //// <summary>
        ///解决 word调用打印机报错问题，创建一个临时文件夹
        /// </summary>
        /// <returns></returns>
        public bool CreateFolder()
        {
            bool ifsuccesss = false;
            try
            {
                string systempath = System.Environment.GetFolderPath(Environment.SpecialFolder.System);
                string fullpath = string.Empty;
                if (FindSystemWidth() == "32")
                {
                    fullpath = systempath + "\\config\\systemprofile\\Desktop";
                }
                else
                {
                    fullpath = systempath + "\\config\\systemprofile\\Desktop";
                }
                if (!Directory.Exists(fullpath))
                {
                    Directory.CreateDirectory(fullpath);
                }
                ifsuccesss = true;
            }
            catch (Exception ex)
            {
                ifsuccesss = false;
            }
            return ifsuccesss;

        }
        /// <summary>
        /// 获取系统的位数
        /// </summary>
        /// <returns></returns>
        public string FindSystemWidth()
        {
            ConnectionOptions oConn = new ConnectionOptions();
            System.Management.ManagementScope oMs = new System.Management.ManagementScope("\\\\localhost", oConn);
            System.Management.ObjectQuery oQuery = new System.Management.ObjectQuery("select AddressWidth from Win32_Processor");
            ManagementObjectSearcher oSearcher = new ManagementObjectSearcher(oMs, oQuery);
            ManagementObjectCollection oReturnCollection = oSearcher.Get();
            string addressWidth = null;

            foreach (ManagementObject oReturn in oReturnCollection)
            {
                addressWidth = oReturn["AddressWidth"].ToString();
            }

            return addressWidth;
        }
        #endregion
    }
}