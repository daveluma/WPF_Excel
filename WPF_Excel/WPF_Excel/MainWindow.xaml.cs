using System;
using System.Collections.Generic;
using System.IO;
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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ClosedXML.Excel;

namespace WPF_textFile
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string name;
        string lname;
        string fileName = @"C:\Users\Public\csharp-Excel.xls";
        string text;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            name = FirstNameBox.Text;
            lname = LastNameBox.Text;

            //new Excel Application Object
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            //check if Excel is installed on device
            if (xlApp == null)
            {
                Console.WriteLine("Excel is not installed");
            }
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            object misValue = System.Reflection.Missing.Value;
            Excel.Workbook XLWorkbook = xlApp.Workbooks.Add(misValue);

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //place information into the cells
            xlWorkSheet.Cells[1, 1] = "First-Name";
            xlWorkSheet.Cells[1, 2] = "Last-Name";
            xlWorkSheet.Cells[2, 1] = name;
            xlWorkSheet.Cells[2, 2] = lname;

            //export the excel file
            xlWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
