using System;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;
//using Point = NetOffice.PowerPointApi;
using VBIDE = NetOffice.VBIDEApi;
using NOTools = NetOffice.OfficeApi.Tools;
  
namespace ClientApplication
{
    public class Form1 : System.Windows.Forms.Form
    { 
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Form1()
        {
            InitializeComponent();

            try
            {
                NetOffice.Settings.Default.PerformanceTrace["NetOffice.ExcelApi"].IntervalMS = 0;
                NetOffice.Settings.Default.PerformanceTrace["NetOffice.ExcelApi"].Enabled = true;
                NetOffice.Settings.Default.PerformanceTrace.Alert += PerformanceTrace_Alert;
                NetOffice.Settings.Default.EnableAutomaticQuit = true;
                using (Excel.Application application = new NetOffice.ExcelApi.Application())
                {
                    var book = application.Workbooks.Add();
                    book.Sheets.Add();
                    application.DisplayAlerts = false;
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
            
            //try
            //{
            //    Point.Application app = new Point.Application();
            //    dynamic application = COMDynamicObject.ConvertTo(app);
            //    application.Visible = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;
            //    Point.Presentation pres1 = application.Presentations.Add();

            //    dynamic presentations = COMDynamicObject.ConvertTo(application.Presentations);
            //    foreach (Point.Presentation item in presentations)
            //        Console.WriteLine(item.Name);

            //    var pres2 = presentations[1];
            //    dynamic pres2Dynamic = COMDynamicObject.ConvertTo(pres2);

            //    bool isEqual = pres1 == pres2Dynamic;
            //    MessageBox.Show(isEqual.ToString());

            //    NetOffice.PowerPointApi.Tools.Utils.CommonUtils utils = new Point.Tools.Utils.CommonUtils(app);
            //    string hwnd1 = utils.Application.HWND.ToString();
            //    string hwnd2 = app.HWND.ToString();
            //    Console.WriteLine(hwnd1);
            //    Console.WriteLine(hwnd2);
            //    //MessageBox.Show(hwnd1.ToString() + " " + hwnd2.ToString());
            //    app.Quit();
            //    application.Dispose();
            //}
            //catch
            //{
            //    ;
            //}
        }

        private void PerformanceTrace_Alert(PerformanceTrace sender, PerformanceTrace.PerformanceAlertEventArgs args)
        {
            Console.WriteLine(args);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {          
            try
            {
                //new MultiRegisterClient().Test();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.ToString());
            }
            finally
            {
                Close();
            }
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(292, 273);
            this.Name = "Form1";
            this.Text = "ClientApplication";
            this.Shown += new System.EventHandler(this.Form1_Shown);
            this.ResumeLayout(false);

        }

        #endregion
    }
}
