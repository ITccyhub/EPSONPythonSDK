using AxESCANOCX2Lib;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;

namespace PythonCallDll
{
    public class Class1
    {
        public string ClassName { get; set; }
        public static int _classObjectCounts = 0;

        public Class1()
        {
            ClassName = "PythonCallDll";
            _classObjectCounts = _classObjectCounts + 1;
        }
        public Class1(string name)
        {
            ClassName = name;
            _classObjectCounts = _classObjectCounts + 1;
        }
        public int AddCalc(int a, int b)
        {
            return a + b;
        }

        public void SayHello(String name)

        {

            MessageBox.Show(name);


        }
    }



    public partial class Form2 : Form
    {
        private System.ComponentModel.IContainer components = null;

        private string logFilePath = ".\\log.txt";

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        public Form2()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.axEScanControl21 = new AxESCANOCX2Lib.AxEScanControl2();
            ((System.ComponentModel.ISupportInitialize)(this.axEScanControl21)).BeginInit();
            this.Controls.Add(this.axEScanControl21);
            ((System.ComponentModel.ISupportInitialize)(this.axEScanControl21)).EndInit();
        }
        private AxESCANOCX2Lib.AxEScanControl2 axEScanControl21;

        private void LogMessage(string logFilePath, string message)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine($"{DateTime.Now}: {message}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to write log: {ex.Message}");
            }
        
    }


    public void SayHello(string deviceName, string filePath, int imageType, int docSource, int resolution, int skewCorrect, int docsize, int fileformat, string logFilePath)
    {
        LogMessage(logFilePath, "开始打印配置");

            this.logFilePath = logFilePath;
            axEScanControl21.LicenseKey = "EPSON";
        axEScanControl21.EParamDeviceInfo.DeviceName = deviceName;
        axEScanControl21.EParamSave.FilePath = filePath;

        axEScanControl21.EParamScan.DocSource = docSource;
        axEScanControl21.EParamScan.ImageType = imageType;

        // 强转 dpi 
        axEScanControl21.EParamScan.Resolution = (ushort)resolution;

        // 矫正偏差
        axEScanControl21.EParamScan.SkewCorrect = skewCorrect;

        axEScanControl21.EParamScan.DocSize = docsize; // A4
        axEScanControl21.EParamSave.FileFormat = fileformat;

        uint ret;
        ret = axEScanControl21.OpenScanner();
        if (ret != 0)
        {
            LogMessage(logFilePath, "OpenScanner failed");
            return;
        }

        ret = axEScanControl21.Execute(1);    // ET_SCAN_AND_STORE
        if (ret != 0)
        {
            LogMessage(logFilePath, "Execute failed");
            return;
        }

        ret = axEScanControl21.Execute(2);    // ET_SAVE_STORED_IMAGE
        LogMessage(logFilePath, "保存启动");
    }

    private void axEScanControl21_OnCompleted(object sender, AxESCANOCX2Lib._IEScanControl2Events_OnCompletedEvent e)
    {
        string messageText = "OnCompleted type = ";

        switch (e.execType)
        {
            case 0:
                messageText = "ET_SCAN_AND_SAVE";
                break;
            case 1:
                messageText = "ET_SCAN_AND_STORE";
                break;
            case 2:
                messageText = "ET_SAVE_STORED_IMAGE";
                break;
            case 3:
                messageText = "ET_GET_STORED_IMAGE";
                break;
            case 4:
                messageText = "ET_GET_OBR_DATA";
                break;
            case 5:
                messageText = "ET_GET_OCR_DATA";
                break;
        }

        messageText += ", error = ";

        switch (e.errorCode)
        {
            case 0:
                messageText += "ESL_SUCCESS";
                break;
            case 1:
                messageText += "ESL_CANCEL";
                break;
            case -2147483647:
                messageText += "ESL_ERR_GENERAL";
                break;
            case -2147483646:
                messageText += "ESL_ERR_NOT_INITIALIZED";
                break;
            case -2147483645:
                messageText += "ESL_ERR_FILE_MISSING";
                break;
            case -2147483644:
                messageText += "ESL_ERR_INVALID_PARAM";
                break;
            case -2147483643:
                messageText += "ESL_ERR_LOW_MEMORY";
                break;
            case -2147483642:
                messageText += "ESL_ERR_LOW_DISKSPACE";
                break;
            case -2147483641:
                messageText += "ESL_ERR_WRITE_FAIL";
                break;
            case -2147483640:
                messageText += "ESL_ERR_READ_FAIL";
                break;
            case -2147483639:
                messageText += "ESL_ERR_INVALID_OPERATION";
                break;
        }

        // Handle OBR or OCR if necessary
        if (e.execType == 4)
        {
            if (axEScanControl21.EStoredPages.count != 0)
            {
                ESCANOCX2Lib.EPage page = (ESCANOCX2Lib.EPage)axEScanControl21.EStoredPages.Item(0);
                if (page.EBarcodeBlocks.count != 0)
                {
                    ESCANOCX2Lib.ECodeData data = (ESCANOCX2Lib.ECodeData)page.EBarcodeBlocks.Item(0);
                    messageText += "\n";
                    messageText += data.DataString;
                }
            }
        }
        else if (e.execType == 5)
        {
            if (axEScanControl21.EStoredPages.count != 0)
            {
                ESCANOCX2Lib.EPage page = (ESCANOCX2Lib.EPage)axEScanControl21.EStoredPages.Item(0);
                if (page.ETextBlocks.count != 0)
                {
                    ESCANOCX2Lib.ETextData data = (ESCANOCX2Lib.ETextData)page.ETextBlocks.Item(0);
                    messageText += "\n";
                    messageText += data.DataString;
                }
            }
        }

        LogMessage(logFilePath, messageText);
    }

}


    
}
