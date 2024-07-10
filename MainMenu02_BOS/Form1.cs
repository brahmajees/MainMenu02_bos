namespace MainMenu02_BOS
{
    using OfficeOpenXml;
    using System;
    using System.Data.SqlClient;
    using System.Data;
    using System.IO;
    using System.Windows.Forms;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using OfficeOpenXml.Style;


    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void importEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileInfo fi01 = new FileInfo(@"d:\bendem\NSDL\NSDLSRC01.xlsx");
            if (fi01.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\bendem\NSDL\NSDLSRC01.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }


        private void generationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            string sourceXlsxFilePath = @"d:\\bendem\NSDL\\nsdlsrc01.xlsx";
            string targetCsvFilePath = @"d:\\bendem\NSDL\\nsdlsrc01.csv";

            ConvertXlsxToCsv(sourceXlsxFilePath, targetCsvFilePath);
            Console.WriteLine("Conversion complete.");
            string directoryPath = @"d:\bendem\nsdl\";
            var mostRecentFile = new DirectoryInfo(directoryPath)
             .GetFiles("*.txt")
             .OrderByDescending(f => f.LastWriteTime)
             .FirstOrDefault();
            string notepadPlusPlusPath = @"C:\Program Files\Notepad++\Notepad++.exe";
            if (File.Exists(notepadPlusPlusPath))
            {
                System.Diagnostics.Process.Start(notepadPlusPlusPath, mostRecentFile.FullName);
            }
        }


        internal static void ConvertXlsxToCsv(string sourceXlsxFilePath, string targetCsvFilePath)
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(sourceXlsxFilePath)))
            {
                int DATA = 0;
                var worksheet = excelPackage.Workbook.Worksheets[DATA];
                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;
                using (var streamWriter = new StreamWriter(targetCsvFilePath))
                {
                    // Write data rows
                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= columns; j++)
                        {
                            if (j > 1 && j <= 3)
                            {
                                streamWriter.Write(",");
                            }
                            var cellValue = worksheet.Cells[i, j].Value?.ToString() ?? "";
                            streamWriter.Write(cellValue);
                        }
                        streamWriter.WriteLine();
                    }
                }
            }

            {

                System.Diagnostics.Process.Start(@"d:\CAFILES\output\frnbd01.bat").WaitForExit();

                MessageBox.Show("This a option to view generated file!");
 
            }
        }


        private void notepadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("notepad++.exe");
        }

        private void calculatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("calc.exe");
        }


        private void importEntryToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FileInfo fi01 = new FileInfo(@"d:\bendem\CDSL\CDSLSRC.xlsx");
            if (fi01.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\bendem\CDSL\CDSLSRC.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void generationToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            string sourceXlsxFilePathc = @"d:\bendem\cdsl\cdslsrc.xlsx";
            string targetCsvFilePathc = @"d:\bendem\cdsl\cdslsrc.csv";

            ConvertXlsxToCsvc(sourceXlsxFilePathc, targetCsvFilePathc);
            Console.WriteLine("Conversion complete.");
            string directoryPath = @"d:\bendem\cdsl\";
            var mostRecentFile = new DirectoryInfo(directoryPath)
             .GetFiles("*.cvf")
             .OrderByDescending(f => f.LastWriteTime)
             .FirstOrDefault();
            string notepadPlusPlusPath = @"C:\Program Files\Notepad++\Notepad++.exe";
            if (File.Exists(notepadPlusPlusPath))
            {
                System.Diagnostics.Process.Start(notepadPlusPlusPath, mostRecentFile.FullName).WaitForExit();
            }
        }


        internal static void ConvertXlsxToCsvc(string sourceXlsxFilePathc, string targetCsvFilePathc)
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(sourceXlsxFilePathc)))
            {
                int DATA = 0;
                var worksheet = excelPackage.Workbook.Worksheets[DATA];
                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;

                using (var streamWriter = new StreamWriter(targetCsvFilePathc))
                {
                    // Write data rows
                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= columns; j++)
                        {
                            if (j > 1 && j <= 3)
                            {
                                streamWriter.Write(",");
                            }
                            var cellValue = worksheet.Cells[i, j].Value?.ToString() ?? "";
                            streamWriter.Write(cellValue);
                        }
                        streamWriter.WriteLine();
                    }
                }
            }

            {
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frcbd01.bat").WaitForExit();
                MessageBox.Show("This a option to view generated file!");
                
            }
        }
        //           
        private void processOfOutFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Are you sure & continue?");
            
            {
                System.Diagnostics.Process.Start(@"d:\bendem\nsdl\outfilesn\onlystep1.bat").WaitForExit(); //DTS PACKAGE FOR CONVERSION FROM RAW TO 
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\csvtoExcelBDoutNsdl01\csvtoExcelBDoutNsdl01\bin\Debug\csvtoExcelBDoutNsdl01.exe");
                System.Diagnostics.Process.Start(@"d:\bendem\nsdl\outfilesn\onlystep2.bat").WaitForExit(); //DTS PACKAGE FOR CONVERSION FROM RAW TO 
            }
            
            
        }

        
        private void processOfOutfileToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"D:\BENDEM\CDSL\OutFilesC\cnvfinal1.bat").WaitForExit();
            System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\cdslbdoutfilecsvtoexcelformat01\cdslbdoutfilecsvtoexcelformat01\bin\Debug\cdslbdoutfilecsvtoexcelformat01.exe").WaitForExit();
        }

        private void consolidatedFileToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DialogResult dresult = MessageBox.Show("Are you sure & continue", "Alert", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dresult == DialogResult.OK)
            {
                SqlConnection con = new SqlConnection("Data Source=VCCIPL-TECH\\VENTURESQLEXP;Initial Catalog=VCCIPL;Integrated Security=True;");
                con.Open();
                SqlCommand sql_cmnd2 = new SqlCommand("FINALCONSOLIDATENSDLCDSLBENDEMDATA", con);

                sql_cmnd2.CommandType = CommandType.StoredProcedure;
                sql_cmnd2.ExecuteNonQuery();
                con.Close();
            }
            MessageBox.Show("PROCESS COMPLETED View the File!!!");
            System.Diagnostics.Process.Start(@"d:\bendem\consfiles\exportexcel.bat").WaitForExit();

        }
        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("excel.exe");
        }

        private void wordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("winword.exe");
        }

        private void explorerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe");
        }

        private void vCCIPLWebSiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.vccipl.com/");
        }

        private void vcciplClientLoginToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://client.vccipl.com/");
        }

        private void bendemToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void eVotingSummaryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"d:\evote\readme.txt").WaitForExit();
            SqlConnection con = new SqlConnection("Data Source=VCCIPL-TECH\\VENTURESQLEXP;Initial Catalog=EVOTING;Integrated Security=True;");
            con.Open();
            SqlCommand sql_cmnd1 = new SqlCommand("sp_evotingprocess01", con);
            sql_cmnd1.CommandType = CommandType.StoredProcedure;
            sql_cmnd1.ExecuteNonQuery();

            System.Diagnostics.Process.Start(@"d:\EVOTE\EVOTEPRO01.BAT").WaitForExit();
            MessageBox.Show("PROCESS COMPLETED View the File!!!");
            string directoryPath = @"d:\EVOTE\PRESENTVOTE\";
            var mostRecentFile = new DirectoryInfo(directoryPath)
             .GetFiles("*.txt")
             .OrderByDescending(f => f.LastWriteTime)
             .FirstOrDefault();
            string notepadPlusPlusPath = @"C:\Program Files\Notepad++\Notepad++.exe";
            if (File.Exists(notepadPlusPlusPath))
            {
                System.Diagnostics.Process.Start(notepadPlusPlusPath, mostRecentFile.FullName).WaitForExit();
            }
            MessageBox.Show("!!!PROCESS COMPLETED!!!");
            string textFilePath = mostRecentFile.FullName;
            string excelFilePath = (@"D:\evote\presentvote\Vote_Summary.xls");
            string[] lines = File.ReadAllLines(textFilePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var excel = new ExcelPackage();
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("Sheet1");
                for (int i = 0; i < lines.Length; i++)
                {
                    string[] columns = lines[i].Split(',');
                    for (int j = 0; j < columns.Length; j++)
                    {
                        worksheet.Cells[i + 1, j + 1].Value = columns[j];
                    }
                    FileInfo excelFile = new FileInfo(excelFilePath);
                    excel.SaveAs(excelFile);
                }
                Console.WriteLine("text file has been successfully convert to Excel File");
            }
            this.Close();
        }

        private void summaryAILFirstCallToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void fIlesCopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\filesCopyfromSrcToDest\filesCopyfromSrcToDest\bin\Debug\filesCopyfromSrcToDest.exe").WaitForExit();
            
        }

        private void analysisISINToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"D:\pbi_Files\isin_master.pbix").WaitForExit();
        }

        private void closeToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void closeToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void refundGenBankWiseToolStripMenuItem_Click(object sender, EventArgs e)
        {
         //   System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\SplitExcelToCsvByColumn01\SplitExcelToCsvByColumn01\bin\Debug\SplitExcelToCsvByColumn01.exe").WaitForExit();
         //   MessageBox.Show("Generation is over pls check it once!!!! at test  out Folder ");
        }

        private void summaryAILFirstCallToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            
        }

        private void convFromBDOutputToExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("csvtoexcelwithheadersforbendemnsdl01");
            MessageBox.Show("[D:BENDEM_NSDL_OutFilesN_IN200537_BENDEM_20231206_105201_20231206105334_9999_VCCILNSD.OUT]");
            MessageBox.Show("Input File must be NSDL BENDEM OUT FILE (DELETED last three lines");

            System.Threading.Thread.Sleep(8000);
            System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\csvtoexcelwithheadersforbendemnsdl01\csvtoexcelwithheadersforbendemnsdl01\bin\Debug\csvtoexcelwithheadersforbendemnsdl01.exe").WaitForExit();
            MessageBox.Show("[D:_BENDEM_NSDL_OutFilesN]");
            MessageBox.Show("Process Over please Check in the Folder ");
            
        }

        private void cSVFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\SplitExcelToCsvByColumn01\SplitExcelToCsvByColumn01\bin\Debug\SplitExcelToCsvByColumn01.exe").WaitForExit();
            MessageBox.Show("Generation is over pls check it once!!!! at test  CSVFiles Folder ");
        }

        private void xLSFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\SplitExcelFileintoIndExcelFilesonColumn02\SplitExcelFileintoIndExcelFilesonColumn02\bin\Debug\SplitExcelFileintoIndExcelFilesonColumn02.exe").WaitForExit();
            MessageBox.Show("Generation is over [D:][VRights][FirstCall][refund][excelfiles]pls check it once!!!!");
        }

        private void nCAFilestoExcelFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\texttoexcel01\texttoexcel01\bin\Debug\texttoexcel01.exe").WaitForExit();
        }

        private void sCAHDFilesToXLSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\texttoexcel02\texttoexcel02\bin\Debug\texttoexcel02.exe").WaitForExit();
        
        }

        private void nCAHDCSVtoXLSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\texttoExcelNgrepfile01\texttoExcelNgrepfile01\bin\Debug\texttoExcelNgrepfile01.exe").WaitForExit();
            MessageBox.Show("Your File is Processed !!!");
        }

        private void sCAHDCSVtoXLSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\texttoexcel02\texttoexcel02\bin\Debug\texttoexcel02.exe").WaitForExit();
            MessageBox.Show("Your File is Processed !!!");
        }

        private void stampDutyCalculatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://nsdl.co.in/stampduty_calculator.php");
        }

        private void ASBABANKLISTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"d:\brdata\ASBA_BANKS_LIST.xlsx");
        }

        private void adroitInfotechToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"d:\vrights\firstcall\readme.txt").WaitForExit();
            SqlConnection con = new SqlConnection("Data Source=VCCIPL-TECH\\VENTURESQLEXP;Initial Catalog=IPORIGHTSBONUS;Integrated Security=True;");
            con.Open();
            SqlCommand sql_cmnd1 = new SqlCommand("vbidProcess01", con);
            sql_cmnd1.CommandType = CommandType.StoredProcedure;
            sql_cmnd1.ExecuteNonQuery();
            System.Diagnostics.Process.Start(@"D:\VRights\FirstCall\FcallMoneypro01.bat").WaitForExit();
            MessageBox.Show("SUMMARY OF BID FIRST CALL PROCESS COMPLETED View the File!!!");
            string directoryPath = @"d:\VRIGHTS\FIRSTCALL\";
            var mostRecentFile = new DirectoryInfo(directoryPath)
             .GetFiles("*.txt")
             .OrderByDescending(f => f.LastWriteTime)
             .FirstOrDefault();
            string notepadPlusPlusPath = @"C:\Program Files\Notepad++\Notepad++.exe";
            if (File.Exists(notepadPlusPlusPath))
            {
                System.Diagnostics.Process.Start(notepadPlusPlusPath, mostRecentFile.FullName).WaitForExit();
            }
            MessageBox.Show("!!!PROCESS COMPLETED!!!");
            string textFilePath = mostRecentFile.FullName;
            string excelFilePath = (@"D:\VRights\FirstCall\FirstCall_Summary.xls");
            string[] lines = File.ReadAllLines(textFilePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var excel = new ExcelPackage();
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("Sheet1");
                for (int i = 0; i < lines.Length; i++)
                {
                    string[] columns = lines[i].Split(',');
                    for (int j = 0; j < columns.Length; j++)
                    {
                        worksheet.Cells[i + 1, j + 1].Value = columns[j];
                    }
                    FileInfo excelFile = new FileInfo(excelFilePath);
                    excel.SaveAs(excelFile);
                }
                Console.WriteLine("text file has been successfully convert to Excel File");
            }
        }

        private void adroitInfotechLtdToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"W:\ARHAM\MERGER\RIGHTS24\1STCALL\REFUND\BID_14062024_Final.xlsx");
        }
    }
}
