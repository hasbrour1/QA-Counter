using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

/*
 ************************************************************
 * @author: Robert Hasbrouck                                *
 * @Version: 1.1                                            *
 *                                                          *
 * After version 0.4 changed from sequential sort to        *
 * binary sort to decrease run time                         *
 ************************************************************ 
 */

namespace QA_Counter
{
    public partial class Form1 : Form
    {
        //File locations for each month
        String[] month = new String[12];

        //excel workbooks
        public static Excel._Workbook book = null;
        public static Excel.Application app = null;
        public static Excel._Worksheet sheet = null;
        public static Excel._Workbook srcBook = null;
        public static Excel.Application srcApp = null;
        public static Excel._Worksheet srcSheet = null;

        List<Analyte> waterAnalytes = new List<Analyte>();
        List<Analyte> drinkingWaterAnalytes = new List<Analyte>();
        List<Analyte> airAnalytes = new List<Analyte>();
        List<Analyte> solidAnalytes = new List<Analyte>();
        List<Analyte> strangeAnalytes = new List<Analyte>();

        int row = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[0] = openFileDialog1.FileName;
                this.label1.Text = "Opening " + month[0];
                this.label1.Text = month[0];

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[1] = openFileDialog1.FileName;
                this.label2.Text = "Opening " + month[1];
                this.label2.Text = month[1];

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[2] = openFileDialog1.FileName;
                this.label3.Text = "Opening " + month[2];
                this.label3.Text = month[2];

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[3] = openFileDialog1.FileName;
                this.label4.Text = "Opening " + month[3];
                this.label4.Text = month[3];

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[4] = openFileDialog1.FileName;
                this.label5.Text = "Opening " + month[4];
                this.label5.Text = month[4];

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[5] = openFileDialog1.FileName;
                this.label6.Text = "Opening " + month[5];
                this.label6.Text = month[5];

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[6] = openFileDialog1.FileName;
                this.label7.Text = "Opening " + month[6];
                this.label7.Text = month[6];

            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[7] = openFileDialog1.FileName;
                this.label8.Text = "Opening " + month[7];
                this.label8.Text = month[7];

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[8] = openFileDialog1.FileName;
                this.label9.Text = "Opening " + month[8];
                this.label9.Text = month[8];

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[9] = openFileDialog1.FileName;
                this.label10.Text = "Opening " + month[9];
                this.label10.Text = month[9];

            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[10] = openFileDialog1.FileName;
                this.label11.Text = "Opening " + month[10];
                this.label11.Text = month[10];

            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                month[11] = openFileDialog1.FileName;
                this.label12.Text = "Opening " + month[11];
                this.label12.Text = month[11];
                
            }
        }

        /*
         * button13 is Finish button
         * when clicked it will launch counting thread
         * 
         */
        private void button13_Click(object sender, EventArgs e)
        {
            //Start Thread to process data
            this.label13.Text = "Compiling...";
            Thread t = new Thread(new ThreadStart(startThread));
            t.Start();   
        }

        /*
         * startThread method creates the workbooks
         * and begins the counting process from the 
         * addMonth() method then saves the workbook
         * in the directory of january's file
         * 
         */
        public void startThread()
        {
                //Template setup
                app = new Excel.Application();
                app.Visible = false;
                book = (Excel._Workbook)(app.Workbooks.Add());
                sheet = (Excel._Worksheet)book.Sheets[1];

                //source setup
                srcApp = new Excel.Application();
                srcApp.Visible = false;

                //Headers
                sheet.Cells[2, 1] = "Destination Lab";
                sheet.Cells[2, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                sheet.Cells[2, 2] = "Matrix";
                sheet.Cells[2, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                sheet.Cells[2, 3] = "Alternate_Matrix";
                sheet.Cells[2, 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                sheet.Cells[2, 4] = "State";
                sheet.Cells[2, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                sheet.Cells[2, 5] = "Analyte";
                sheet.Cells[2, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                sheet.Cells[1, 6] = "Analyte_Count";
                sheet.Cells[1, 6].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                sheet.Cells[2, 18] = "Year Total";
                sheet.Cells[2, 18].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                sheet.Cells[2, 6] = "January";
                sheet.Cells[2, 7] = "February";
                sheet.Cells[2, 8] = "March";
                sheet.Cells[2, 9] = "April";
                sheet.Cells[2, 10] = "May";
                sheet.Cells[2, 11] = "June";
                sheet.Cells[2, 12] = "July";
                sheet.Cells[2, 13] = "August";
                sheet.Cells[2, 14] = "September";
                sheet.Cells[2, 15] = "October";
                sheet.Cells[2, 16] = "November";
                sheet.Cells[2, 17] = "December";

                for (int i = 0; i < month.Length; i++)
                {
                    addMonth(month[i], i);
                }

                //write counts
                writeAnalytes(airAnalytes);
                writeAnalytes(waterAnalytes);
                writeAnalytes(drinkingWaterAnalytes);
                writeAnalytes(solidAnalytes);
                writeAnalytes(strangeAnalytes);

                /***************************************
                 * This code will add boarders
                 * Took out because it takes too long
                 *************************************** 
                 *
                for(int i = 0; i < sheet.UsedRange.Rows.Count; i++){
                    for(int j = 0; j < sheet.UsedRange.Columns.Count; j++){
                        sheet.Cells[i + 1, j + 1].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black); 
                    }
                }
                 */

                
                for (int i = 0; i < sheet.UsedRange.Columns.Count; i++)
                {
                    sheet.Cells[1, i + 1].EntireColumn.ColumnWidth = 13;
                }

                sheet.Cells[1, 3].EntireColumn.ColumnWidth = 14.5;
                sheet.Cells[1, 5].EntireColumn.ColumnWidth = 30;


                //save workbook
                //find out how to save
                SetText("Complete");
                string directoryPath = Path.GetDirectoryName(@month[0]);
                object m_objOpt = System.Reflection.Missing.Value;
                book.SaveAs(directoryPath + "\\YearCounted.xls", m_objOpt,
                        m_objOpt, m_objOpt, false, m_objOpt, Excel.XlSaveAsAccessMode.xlNoChange,
                        m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                book.Close(true);
                app.Workbooks.Close();
                app.Quit();
                srcBook.Close(true); 
                srcApp.Workbooks.Close();
                srcApp.Quit();

                //re initalize row so if ran again 
                //it starts from row 0 again
                row = 0;

                //open file location after saved
                if (File.Exists(directoryPath + "\\YearCounted.xls"))
                {
                    Process.Start("explorer.exe", "/select, " + directoryPath + "\\YearCounted.xls");
                }
        }

        public void addMonth(String src, int month)
        {
            //If the template file is found, run program
            if (File.Exists(src))
            {
                //get src
                srcBook = (Excel._Workbook)(srcApp.Workbooks.Add(src));
                srcSheet = (Excel._Worksheet)srcBook.Sheets[1];

                //Merge Both files here
                switch (month){
                    case 0://Jan
                        SetText("Counting January...");
                        countMonth(month);
                        break;
                    case 1://Feb
                        SetText("Counting February...");                     
                        countMonth(month);
                        break;
                    case 2://Mar
                        SetText("Counting March...");                      
                        countMonth(month);
                        break;
                    case 3://April 
                        SetText("Counting April...");                      
                        countMonth(month);
                        break;
                    case 4://May
                        SetText("Counting May...");
                        countMonth(month);
                        break;
                    case 5://June
                        SetText("Counting June...");
                        countMonth(month);
                        break;
                    case 6://July
                        SetText("Counting July...");
                        countMonth(month);
                        break;
                    case 7: //Aug
                        SetText("Counting August...");
                        countMonth(month);
                        break;
                    case 8: //Sept
                        SetText("Counting September...");
                        countMonth(month);
                        break;
                    case 9://Oct
                        SetText("Counting October...");                        
                        countMonth(month);
                        break;
                    case 10: //Nov
                        SetText("Counting November...");
                        countMonth(month);
                        break;
                    case 11://Dec
                        SetText("Counting December...");                   
                        countMonth(month);
                        break;
                    default: 
                        break;
                }
            }
            else
            {
                SetText("File not selected. Skipping...");
            }
        }

        /*
         * countMonth sorts the months analytes
         * into the respective Analyte List via 
         * its detailed matrix
         * 
         */
        public void countMonth(int month)
        {

            //Sort analite to correct matrix array
            //and count correct ammount
            Analyte anly;
            int foundIndex;
            for (int i = 2; i < srcSheet.UsedRange.Rows.Count; i++)
            {
                //ignore all results not from New York
                if (srcSheet.Cells[i, 4].Value2 == "NY")
                {
                    //anly is instance of Analyte class 
                    //which will populate itself and add itself
                    //to the appropriate List<Analyte> which are 
                    //a list of Analyte class instances
                    //If analyte is already populated in list
                    //it just adds the count to the analyte already
                    //in the array
                    anly = new Analyte(srcSheet.Cells[i, 5].Value2, srcSheet.Cells[i, 2].Value2, srcSheet.Cells[i, 3].Value2, Convert.ToInt32(srcSheet.Cells[i, 6].Value2), month);
                    String str = srcSheet.Cells[i, 3].Value2;
                    switch (str)
                    {
                        case "Air":
                            foundIndex = getIndex(anly, airAnalytes);
                            airAnalytes = insertAnly(foundIndex, airAnalytes, anly, i, month);
                            break;
                        case "Drinking Water":
                            foundIndex = getIndex(anly, drinkingWaterAnalytes);
                            drinkingWaterAnalytes = insertAnly(foundIndex, drinkingWaterAnalytes, anly, i, month);
                            break;
                        case "Solid":
                            foundIndex = getIndex(anly, solidAnalytes);
                            solidAnalytes = insertAnly(foundIndex, solidAnalytes, anly, i, month);
                            break;
                        case "Water":
                            foundIndex = getIndex(anly, waterAnalytes);
                            waterAnalytes =insertAnly(foundIndex, waterAnalytes, anly, i, month);
                            break;
                        case "Ground Water":
                            anly.setDetMatrix("Water");
                            foundIndex = getIndex(anly, waterAnalytes);
                            waterAnalytes = insertAnly(foundIndex, waterAnalytes, anly, i, month);
                            break;
                        default:
                            foundIndex = getIndex(anly, strangeAnalytes);
                            strangeAnalytes = insertAnly(foundIndex, strangeAnalytes, anly, i, month);
                            break;
                    }
                }
            }
        }

        public List<Analyte> insertAnly(int foundIndex, List<Analyte> analytes, Analyte anly, int i, int month)
        {
            if (foundIndex <= -1)
            {
                analytes.Add(anly);
                analytes = analytes.OrderBy(si => si.name).ToList();
                return analytes;
            }
            else
            {
                analytes[foundIndex].addCount(Convert.ToInt32(srcSheet.Cells[i, 6].Value2), month);
                analytes = analytes.OrderBy(si => si.name).ToList();
                return analytes;
            }
        }

        public int getIndex(Analyte anly, List<Analyte> analytes)
        {
            return analytes.BinarySearch(anly, new MyObjectIdComparer());
        }

        public void writeAnalytes(List<Analyte> analytes)
        {
            //write to file
            int size = analytes.Count + row;
            int initRow = row;

            for (int i = row; i < size; i++)
            {
                sheet.Cells[i + 3, 1] = "Newburgh";
                sheet.Cells[i + 3, 2] = analytes[(i - initRow)].getMatrix();
                sheet.Cells[i + 3, 3] = analytes[(i - initRow)].getDetMatrix();
                sheet.Cells[i + 3, 4] = "NY";
                sheet.Cells[i + 3, 5] = analytes[(i - initRow)].getName();
                sheet.Cells[i + 3, 6] = analytes[(i - initRow)].getCount(0);
                sheet.Cells[i + 3, 7] = analytes[(i - initRow)].getCount(1);
                sheet.Cells[i + 3, 8] = analytes[(i - initRow)].getCount(2);
                sheet.Cells[i + 3, 9] = analytes[(i - initRow)].getCount(3);
                sheet.Cells[i + 3, 10] = analytes[(i - initRow)].getCount(4);
                sheet.Cells[i + 3, 11] = analytes[(i - initRow)].getCount(5);
                sheet.Cells[i + 3, 12] = analytes[(i - initRow)].getCount(6);
                sheet.Cells[i + 3, 13] = analytes[(i - initRow)].getCount(7);
                sheet.Cells[i + 3, 14] = analytes[(i - initRow)].getCount(8);
                sheet.Cells[i + 3, 15] = analytes[(i - initRow)].getCount(9);
                sheet.Cells[i + 3, 16] = analytes[(i - initRow)].getCount(10);
                sheet.Cells[i + 3, 17] = analytes[(i - initRow)].getCount(11);

                //write total
                int total = 0;
                for(int j = 0; j < 12; j++){
                    total += analytes[(i - initRow)].getCount(j);
                }

                sheet.Cells[i + 3, 18] = total;

                row++;
                SetText("Compiling forms... ");
            }
        }

        delegate void SetTextCallback(string text);

        private void SetText(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.label13.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.label13.Text = text;
            }
        }

        public class MyObjectIdComparer : Comparer<Analyte>
        {
            public override int Compare(Analyte x, Analyte y)
            {
                // argument checking etc removed for brevity

                return x.name.CompareTo(y.name);
            }
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            String year;

            year = textBox1.Text;

            month[0] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Jan 2014 Count.xls";
            if (File.Exists(month[0]))
            {
                this.label1.Text = month[0];
            }
            else
            {
                this.label1.Text = "File not found.";
                month[0] = "";
            }
            month[1] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Feb 2014 Count.xls";
            if (File.Exists(month[1]))
            {
                this.label2.Text = month[1];
            }
            else
            {
                this.label2.Text = "File not found.";
                month[1] = "";
            }
            month[2] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Mar 2014 Count.xls";
            if (File.Exists(month[2]))
            {
                this.label3.Text = month[2];
            }
            else
            {
                this.label3.Text = "File not found.";
                month[2] = "";
            }
            month[3] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Apr 2014 Count.xls";
            if (File.Exists(month[3]))
            {
                this.label4.Text = month[3];
            }
            else
            {
                this.label4.Text = "File not found.";
                month[3] = "";
            }
            month[4] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\May 2014 Count.xls";
            if (File.Exists(month[4]))
            {
                this.label5.Text = month[4];
            }
            else
            {
                this.label5.Text = "File not found.";
                month[4] = "";
            }
            month[5] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Jun 2014 Count.xls";
            if (File.Exists(month[5]))
            {
                this.label6.Text = month[5];
            }
            else
            {
                this.label6.Text = "File not found.";
                month[5] = "";
            }
            month[6] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Jul 2014 Count.xls";
            if (File.Exists(month[6]))
            {
                this.label7.Text = month[6];
            }
            else
            {
                this.label7.Text = "File not found.";
                month[6] = "";
            }
            month[7] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Aug 2014 Count.xls";
            if (File.Exists(month[7]))
            {
                this.label8.Text = month[7];
            }
            else
            {
                this.label8.Text = "File not found.";
                month[7] = "";
            }
            month[8] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Sep 2014 Count.xls";
            if (File.Exists(month[8]))
            {
                this.label9.Text = month[8];
            }
            else
            {
                this.label9.Text = "File not found.";
                month[8] = "";
            }
            month[9] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Oct 2014 Count.xls";
            if (File.Exists(month[9]))
            {
                this.label10.Text = month[9];
            }
            else
            {
                this.label10.Text = "File not found.";
                month[9] = "";
            }
            month[10] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Nov 2014 Count.xls";
            if (File.Exists(month[10]))
            {
                this.label11.Text = month[10];
            }
            else
            {
                this.label11.Text = "File not found.";
                month[10] = "";
            }
            month[11] = @"\\newbsvr8\sys\RENEE\NYS COUNTS\" + year + "\\Dec 2014 Count.xls";
            if (File.Exists(month[11]))
            {
                this.label12.Text = month[11];
            }
            else
            {
                this.label12.Text = "File not found.";
                month[11] = "";
            }
        }
    }
}
