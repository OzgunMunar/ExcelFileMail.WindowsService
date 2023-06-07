using Mail.WindowsService.DataProvider;
using Mail.WindowsService.DateClasses;
using Mail.WindowsService.DTOClasses;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Timers;


namespace Mail.WindowsService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        // Initializing of the Timer.
        Timer myTimer = new Timer();
        // Initialize file path for excel file.
        string path = @"C:\test\";

        /// <summary>
        /// Works by the time Timer Starts.
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {

            myTimer.Interval = 86400000;  // It works every 24 hours
            myTimer.AutoReset = true;
            myTimer.Elapsed += MyTimer_Elapsed;
            myTimer.Start();

        }

        /// <summary>
        /// When Timer hits 24 hours, it checks if it's the first day of the month then it works or not.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MyTimer_Elapsed(object sender, ElapsedEventArgs e)
        {

            if(DateTime.Now.Day == 1)
            {

                DoExcelWork();

            }

        }

        /// <summary>
        /// When Timer stops, it is defined as null.
        /// </summary>
        protected override void OnStop()
        {

            myTimer.Stop();
            myTimer = null;

        }

        /// <summary>
        /// Processes and provides the data for the file; then sends the mail.
        /// </summary>
        private void DoExcelWork()
        {
            // Get data.
            List<Person> data = new List<Person>();
            DataProvide provider = new DataProvide();

            data = provider.ProvideData();

            // Title for the excel File
            string[] titles = { "PersonName", "Address", "Nation", "CarPlate", "StartDate", "FinishDate", "Notes"};

            // Excel app, book and sheet are initiated.
            Microsoft.Office.Interop.Excel.Application myExcelFile = new Microsoft.Office.Interop.Excel.Application();
            myExcelFile.Visible = true;

            object Missing = Type.Missing;

            Workbook workBook = myExcelFile.Workbooks.Add(Missing);

            Worksheet workSheet = (Worksheet)workBook.Sheets[1];

            int column = 2, row = 2;

            // Printing titles into the Excel Sheet.

            for (int i = 0; i < titles.Length; i++) 
            {

                // Working on appereance
                Range sheetRange = (Range)workSheet.Cells[row, column + i];
                sheetRange.Value2 = titles[i].ToString();
                sheetRange.Interior.Color = XlRgbColor.rgbLightSkyBlue;
                sheetRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                sheetRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                sheetRange.RowHeight = 50;

            }

            // After first row, we are getting to the second row to print.
            row++;

            for (int j = 0; j < data.Count; j++)
            {

                if ((j).Equals(data.Count))
                {
                    // Break the loop when data is finished.
                    break;
                }
                
                // Filling the cells.
                workSheet.Cells[row, column + j + 0].Value2 = data[j].PersonName;
                workSheet.Cells[row, column + j + 1].Value2 = data[j].Address;
                workSheet.Cells[row, column + j + 2].Value2 = data[j].Nation;
                workSheet.Cells[row, column + j + 3].Value2 = data[j].CarPlate;
                workSheet.Cells[row, column + j + 4].Value2 = data[j].StartDate.ToString("dd/MM/yyyy hh:mm");
                workSheet.Cells[row, column + j + 5].Value2 = data[j].FinishDate.ToString("dd/MM/yyyy hh:mm");
                workSheet.Cells[row, column + j + 6].Value2 = data[j].Notes;

                // Styling data cells.
                int dataRow = 2;
                for (int i = 0; i < titles.Length; i++)
                {

                    workSheet.Cells[row, dataRow + i].Font.Bold = true;
                    workSheet.Cells[row, dataRow + i].Interior.Color = XlRgbColor.rgbRed;

                    workSheet.Cells[row, dataRow + i].Select();

                    if (dataRow + i == 6 || dataRow + i == 7)
                    {
                        workSheet.Cells[row, dataRow + i].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    }

                }


                row++;
                column--;

            }

            // Total range and border properties as well as filtering.
            Microsoft.Office.Interop.Excel.Range allRange = workSheet.UsedRange;
            Microsoft.Office.Interop.Excel.Borders allBorders = allRange.Borders;
            allBorders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            allBorders.Weight = 2d;
            workSheet.Cells.AutoFilter(1, Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            workSheet.Columns.AutoFit();

            // I had to give the file a name in order to pursue.
            workBook.SaveCopyAs(path + DateTime.Now.Second.ToString() + ".xls"); 
            myExcelFile.Quit();
            BuildMail();

        }

        /// <summary>
        /// Builds the mail message that we want to send.
        /// </summary>
        public void BuildMail()
        {

            // String for body of mail.
            StringBuilder mailText = new StringBuilder();
            mailText.Append("You have received this mail from Excel Mail Service.");

            // Setting up mail properties.
            MailMessage mail = new MailMessage();

            mail.From = new MailAddress("ozgunmnr@gmail.com");
            mail.To.Add("ozgunmnr@gmail.com");
            mail.IsBodyHtml = true;
            mail.Subject = "Monthly Excel Mail Windows Service's Mail";
            mail.Body = mailText.ToString();

            Attachment attach = new Attachment(path + DateTime.Now.Second.ToString() + ".xls");
            mail.Attachments.Add(attach);

            SendMail(mail);

            mail.To.Clear();
            mail.CC.Clear();
            mail.Bcc.Clear();
            System.Threading.Thread.Sleep(5000);

        }

        /// <summary>
        /// SMTP settings and sending the mail.
        /// </summary>
        /// <param name="mailToSend"></param>
        public void SendMail(MailMessage mailToSend)
        {

            SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
            smtpClient.Credentials = new System.Net.NetworkCredential("ozgunmnr@gmail.com", "myPassword");
            smtpClient.EnableSsl = true;
            smtpClient.Send(mailToSend);

        }

    }

}
