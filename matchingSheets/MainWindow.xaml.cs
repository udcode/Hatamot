using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using _excel = Microsoft.Office.Interop.Excel;
namespace matchingSheets
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        _Application excel = new _excel.Application();
        Workbook hulon, azrieli;
        Worksheet hSheet, aSheet;
        BackgroundWorker bw = new BackgroundWorker();
        OpenFileDialog fd = new OpenFileDialog();
        List<string> results;
        public MainWindow()
        {
            InitializeComponent();
            bw.DoWork += processSheets;
            bw.RunWorkerCompleted += completedProcessing;
            bool result = (bool)fd.ShowDialog();
            if (result)
            {
                hulon = excel.Workbooks.Open(fd.FileName);
                hSheet = hulon.Worksheets[1];
                result = (bool)fd.ShowDialog();
                if (result)
                {
                    azrieli = excel.Workbooks.Open(fd.FileName);
                    aSheet = azrieli.Worksheets[1];
                    loading.Visibility = Visibility.Visible;
                    bw.RunWorkerAsync();


                }

            }
        }

        private void completedProcessing(object sender, RunWorkerCompletedEventArgs e)
        {
            loading.Visibility = Visibility.Collapsed;
            resultListBox.DataContext = results;
        }

        private void processSheets(object sender, DoWorkEventArgs e)
        {

            results = findUnMatchingReciets();
        }

        List<string> findUnMatchingReciets()
        {
            List<string> unMatcings = new List<string>();
            List<string> notFounds = new List<string>();
            List<string> worngSum = new List<string>();
            List<string> maamChozList = new List<string>();
            maamChozList.Add("---------------------------------- מעמ חוז לא זהה----------------------------------");
            worngSum.Add("----------------------------------החשבוניות הבאות סכומם לא זהה----------------------------------");
            notFounds.Add("----------------------------------החשבוניות הבאות לא נמצאו בקניון חולון----------------------------------");
            int hColumn = 12, hsumCol = 15, asumCol = 16, aColumn = 11;
            _excel.Range hlast = hSheet.Cells.SpecialCells(_excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            _excel.Range hrange = hSheet.get_Range("L1", hlast);
            _excel.Range alast = aSheet.Cells.SpecialCells(_excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            _excel.Range arange = aSheet.get_Range("L1", alast);
            int hlastRow = hlast.Row;
            for (int hrow = 7; hrow < hlastRow; ++hrow)
            {

                if (hSheet.Cells[hrow, hColumn].Value2 != null)
                {
                    double hreciet = hSheet.Cells[hrow, hColumn].Value2;
                    // MessageBox.Show(reciet.ToString());
                    double check = hSheet.Cells[hrow, hColumn - 1].Value2;
                    string nbm = hSheet.Cells[hrow, 7].Value2;
                    string maamChoz = hSheet.Cells[hrow, 8].Value2;
                    bool isMaamChoz = maamChoz != null && maamChoz.Contains("מעמ חוז");
                    bool isCheck = check > 0;
                    isCheck = nbm != null && nbm.Contains("נבמ");
                    if (isCheck && hreciet > 0)
                    {
                      
                        int alastRow = alast.Row;
                        bool foundReciet = false;
                        bool recietMatch = false;
                        bool checkFound = false;

                        for (int arow = 7; arow < hlastRow; ++arow)
                        {
                            if (aSheet.Cells[arow, asumCol].Value2 != null && hSheet.Cells[hrow, hsumCol].Value2 != null)
                            {
                                if (aSheet.Cells[arow, asumCol].Value2 is string || hSheet.Cells[hrow, hsumCol].Value2 is string) continue;
                                double acheck = aSheet.Cells[arow, asumCol].Value2;
                                double hcheck = hSheet.Cells[hrow, hsumCol].Value2;
                                if (acheck == hcheck)
                                {
                                    foundReciet = true;
                                    checkFound = true;
                                }

                            }

                            if (aSheet.Cells[arow, aColumn].Value2 != null)
                            {
                                double areciet = aSheet.Cells[arow, aColumn].Value2;
                                if (areciet == hreciet)
                                {
                                    foundReciet = true;
                                    //MessageBox.Show($" holon reciet = {hreciet} azrieli reciet = {areciet} ");
                                    if (aSheet.Cells[arow, asumCol].Value2 != null && hSheet.Cells[hrow, hsumCol].Value2 != null)
                                    {
                                        double acheck = aSheet.Cells[arow, asumCol].Value2;
                                        double hcheck = hSheet.Cells[hrow, hsumCol].Value2;
                                        if (acheck == hcheck)
                                        {

                                            //MessageBox.Show($"acheck ={acheck}  hcheck = {hcheck}");
                                        }
                                        else
                                        {
                                            unMatcings.Add($"ho no in reciet number: {hreciet} checks dosn't match!! azrieli check = {acheck}  holon check = {hcheck}");
                                            // MessageBox.Show($"ho no in reciet number: {hreciet} checks dosn't match!! azrieli check = {acheck}  holon check = {hcheck}");
                                        }
                                    }
                                }
                            }
                        }
                        if (!foundReciet)
                        {
                            notFounds.Add($"holon reciet - {hreciet} ");
                            //MessageBox.Show($"holon reciet - {hreciet} was not found in azrieli");
                        }
                        else if (!recietMatch)
                        {

                        }
                        if (!checkFound)
                        {
                            unMatcings.Add($"לא הופקד {check} צ'ק מספר");
                        }
                    }
                    else if(isMaamChoz && (hSheet.Cells[hrow, hsumCol].Value2 != null || hSheet.Cells[hrow, hsumCol+1].Value2 != null))
                    {

                        double hsum = hSheet.Cells[hrow, hsumCol].Value2 != null ? hSheet.Cells[hrow, hsumCol].Value2: hSheet.Cells[hrow, hsumCol + 1].Value2;
                        bool foundMaamChoz = false;
                        for (int arow = 7; arow < hlastRow; ++arow)
                        {
                            if (aSheet.Cells[arow, asumCol].Value2 is string || hSheet.Cells[hrow, hsumCol].Value2 is string) continue;
                            if (aSheet.Cells[arow, asumCol].Value2 != null || aSheet.Cells[arow, asumCol-1].Value2 != null)
                            {
                                double aSum = aSheet.Cells[arow, asumCol].Value2 != null ? aSheet.Cells[arow, asumCol].Value2 : aSheet.Cells[arow, asumCol - 1].Value2;

                                if(Math.Abs(aSum-hsum) < 10)
                                {
                                    foundMaamChoz = true;
                                }

                            }
                        }
                        if (!foundMaamChoz)
                        {
                            maamChozList.Add(hsum.ToString());
                        }

                        
                    }
                }
            }
            asumCol = 15;
            hsumCol = 16;
            for (int arow = 7; arow < hlastRow; ++arow)
            {
                bool foundReciet = false;
                if (aSheet.Cells[arow, aColumn].Value2 != null && aSheet.Cells[arow, asumCol].Value2 != null)
                {
                    double areciet = aSheet.Cells[arow, aColumn].Value2;
                    double aSum = aSheet.Cells[arow, asumCol].Value2;
                    double hsum = 0;
                    
                    for (int hrow = 7; hrow < hlastRow; ++hrow)
                    {
                        string nbm = hSheet.Cells[hrow, 7].Value2;
                        bool isNotCheck = nbm != null && nbm.Contains("חס7");
                        if (!isNotCheck) continue;
                        if ( hSheet.Cells[hrow, hColumn].Value2 != null && hSheet.Cells[hrow, hsumCol].Value2 != null)
                        {
                            double hreciet = hSheet.Cells[hrow, hColumn].Value2;
                           
                            if (areciet == hreciet)
                            {
                                hsum += hSheet.Cells[hrow, hsumCol].Value2;
                                foundReciet = true;

                            }
                        }
                    }
                    if(aSum != hsum)
                    {
                        worngSum.Add($" חשבונית מספר {areciet}");
                    }
                    if (!foundReciet && areciet>0)
                    {
                        notFounds.Add($"azrieli reciet - {areciet} ");
                        //MessageBox.Show($"holon reciet - {hreciet} was not found in azrieli");
                    }
                }
            }
            hulon.Close();
            azrieli.Close();
            unMatcings.AddRange(worngSum);
            unMatcings.AddRange(notFounds);
            unMatcings.AddRange(maamChozList);
            return unMatcings;
        }

    }

}
