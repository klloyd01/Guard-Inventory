using System;
using System.Collections.Generic;
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
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace Guard_Inventory
{
    /// <summary>
    /// Interaction logic for Trans_Log.xaml
    /// </summary>
    public partial class Trans_Log : Page
    {
        public Trans_Log()
        {
            InitializeComponent();

            
            

            //Open up door5Inventory
            Excel.Application transHistoryApp = new Excel.Application();
            transHistoryApp.Visible = false;
            Excel.Workbook transHistoryWorkbook = (Excel.Workbook)(transHistoryApp.Workbooks.Open(@"E:\Data\transHistory"));
            Excel.Worksheet transHistoryWorksheet = (Excel.Worksheet)transHistoryWorkbook.ActiveSheet;

            //Find first available row in transHistory
            string foundEmptyInv = "no";
            Int32 emptyInvRow = 1;
            while (foundEmptyInv != "yes")
            {
                Excel.Range emptyInv = transHistoryWorksheet.get_Range("A" + emptyInvRow);
                //MessageBox.Show(emptyInv.Text);
                if (emptyInv.Text == "")
                {
                    foundEmptyInv = "yes";
                    //MessageBox.Show("Found an empty row");
                }
                if (foundEmptyInv != "yes")
                {
                    //MessageBox.Show("Row used, looking at next one");
                    emptyInvRow = emptyInvRow + 1;
                }

            }

            //Ranges with data from excel to populate text blocks
            //First row
            Excel.Range tID1 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 1));
            Excel.Range d1 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 1));
            Excel.Range i1 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 1));
            Excel.Range q1 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 1));
            Excel.Range pp1 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 1));
            Excel.Range c1 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 1));
            Excel.Range uID1 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 1));
            textBlocktransID1.Text = tID1.Text;
            date1.Text = d1.Text;
            item1.Text = i1.Text;
            quantity1.Text = q1.Text;
            purchasePrice1.Text = pp1.Text;
            container1.Text = c1.Text;
            userID1.Text = uID1.Text;

            //Second Row
            Excel.Range tID2 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 2));
            Excel.Range d2 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 2));
            Excel.Range i2 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 2));
            Excel.Range q2 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 2));
            Excel.Range pp2 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 2));
            Excel.Range c2 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 2));
            Excel.Range uID2 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 2));
            textBlocktransID2.Text = tID2.Text;
            date2.Text = d2.Text;
            item2.Text = i2.Text;
            quantity2.Text = q2.Text;
            purchasePrice2.Text = pp2.Text;
            container2.Text = c2.Text;
            userID2.Text = uID2.Text;

            //Third Row
            Excel.Range tID3 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 3));
            Excel.Range d3 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 3));
            Excel.Range i3 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 3));
            Excel.Range q3 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 3));
            Excel.Range pp3 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 3));
            Excel.Range c3 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 3));
            Excel.Range uID3 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 3));
            textBlocktransID3.Text = tID3.Text;
            date3.Text = d3.Text;
            item3.Text = i3.Text;
            quantity3.Text = q3.Text;
            purchasePrice3.Text = pp3.Text;
            container3.Text = c3.Text;
            userID3.Text = uID3.Text;

            //Fourth Row
            Excel.Range tID4 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 4));
            Excel.Range d4 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 4));
            Excel.Range i4 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 4));
            Excel.Range q4 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 4));
            Excel.Range pp4 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 4));
            Excel.Range c4 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 4));
            Excel.Range uID4 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 4));
            textBlocktransID4.Text = tID4.Text;
            date4.Text = d4.Text;
            item4.Text = i4.Text;
            quantity4.Text = q4.Text;
            purchasePrice4.Text = pp4.Text;
            container4.Text = c4.Text;
            userID4.Text = uID4.Text;

            //Fifth Row
            Excel.Range tID5 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 5));
            Excel.Range d5 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 5));
            Excel.Range i5 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 5));
            Excel.Range q5 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 5));
            Excel.Range pp5 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 5));
            Excel.Range c5 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 5));
            Excel.Range uID5 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 5));
            textBlocktransID5.Text = tID5.Text;
            date5.Text = d5.Text;
            item5.Text = i5.Text;
            quantity5.Text = q5.Text;
            purchasePrice5.Text = pp5.Text;
            container5.Text = c5.Text;
            userID5.Text = uID5.Text;

            //Sixth Row
            Excel.Range tID6 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 6));
            Excel.Range d6 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 6));
            Excel.Range i6 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 6));
            Excel.Range q6 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 6));
            Excel.Range pp6 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 6));
            Excel.Range c6 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 6));
            Excel.Range uID6 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 6));
            textBlocktransID6.Text = tID6.Text;
            date6.Text = d6.Text;
            item6.Text = i6.Text;
            quantity6.Text = q6.Text;
            purchasePrice6.Text = pp6.Text;
            container6.Text = c6.Text;
            userID6.Text = uID6.Text;

            //Seventh Row
            Excel.Range tID7 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 7));
            Excel.Range d7 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 7));
            Excel.Range i7 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 7));
            Excel.Range q7 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 7));
            Excel.Range pp7 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 7));
            Excel.Range c7 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 7));
            Excel.Range uID7 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 7));
            textBlocktransID7.Text = tID7.Text;
            date7.Text = d7.Text;
            item7.Text = i7.Text;
            quantity7.Text = q7.Text;
            purchasePrice7.Text = pp7.Text;
            container7.Text = c7.Text;
            userID7.Text = uID7.Text;

            //Eighth Row
            Excel.Range tID8 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 8));
            Excel.Range d8 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 8));
            Excel.Range i8 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 8));
            Excel.Range q8 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 8));
            Excel.Range pp8 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 8));
            Excel.Range c8 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 8));
            Excel.Range uID8 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 8));
            textBlocktransID8.Text = tID8.Text;
            date8.Text = d8.Text;
            item8.Text = i8.Text;
            quantity8.Text = q8.Text;
            purchasePrice8.Text = pp8.Text;
            container8.Text = c8.Text;
            userID8.Text = uID8.Text;

            //Nineth Row
            Excel.Range tID9 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 9));
            Excel.Range d9 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 9));
            Excel.Range i9 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 9));
            Excel.Range q9 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 9));
            Excel.Range pp9 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 9));
            Excel.Range c9 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 9));
            Excel.Range uID9 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 9));
            textBlocktransID9.Text = tID9.Text;
            date9.Text = d9.Text;
            item9.Text = i9.Text;
            quantity9.Text = q9.Text;
            purchasePrice9.Text = pp9.Text;
            container9.Text = c9.Text;
            userID9.Text = uID9.Text;

            //Tenth Row
            Excel.Range tID10 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 10));
            Excel.Range d10 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 10));
            Excel.Range i10 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 10));
            Excel.Range q10 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 10));
            Excel.Range pp10 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 10));
            Excel.Range c10 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 10));
            Excel.Range uID10 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 10));
            textBlocktransID10.Text = tID10.Text;
            date10.Text = d10.Text;
            item10.Text = i10.Text;
            quantity10.Text = q10.Text;
            purchasePrice10.Text = pp10.Text;
            container10.Text = c10.Text;
            userID10.Text = uID10.Text;

            //Eleventh Row
            Excel.Range tID11 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 11));
            Excel.Range d11 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 11));
            Excel.Range i11 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 11));
            Excel.Range q11 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 11));
            Excel.Range pp11 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 11));
            Excel.Range c11 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 11));
            Excel.Range uID11 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 11));
            textBlocktransID11.Text = tID11.Text;
            date11.Text = d11.Text;
            item11.Text = i11.Text;
            quantity11.Text = q11.Text;
            purchasePrice11.Text = pp11.Text;
            container11.Text = c11.Text;
            userID11.Text = uID11.Text;

            //Twelfth Row
            Excel.Range tID12 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 12));
            Excel.Range d12 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 12));
            Excel.Range i12 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 12));
            Excel.Range q12 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 12));
            Excel.Range pp12 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 12));
            Excel.Range c12 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 12));
            Excel.Range uID12 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 12));
            textBlocktransID12.Text = tID12.Text;
            date12.Text = d12.Text;
            item12.Text = i12.Text;
            quantity12.Text = q12.Text;
            purchasePrice12.Text = pp12.Text;
            container12.Text = c12.Text;
            userID12.Text = uID12.Text;

            //Thirteenth Row
            Excel.Range tID13 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 13));
            Excel.Range d13 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 13));
            Excel.Range i13 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 13));
            Excel.Range q13 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 13));
            Excel.Range pp13 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 13));
            Excel.Range c13 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 13));
            Excel.Range uID13 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 13));
            textBlocktransID13.Text = tID13.Text;
            date13.Text = d13.Text;
            item13.Text = i13.Text;
            quantity13.Text = q13.Text;
            purchasePrice13.Text = pp13.Text;
            container13.Text = c13.Text;
            userID13.Text = uID13.Text;

            //Fourteenth Row
            Excel.Range tID14 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 14));
            Excel.Range d14 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 14));
            Excel.Range i14 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 14));
            Excel.Range q14 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 14));
            Excel.Range pp14 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 14));
            Excel.Range c14 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 14));
            Excel.Range uID14 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 14));
            textBlocktransID14.Text = tID14.Text;
            date14.Text = d14.Text;
            item14.Text = i14.Text;
            quantity14.Text = q14.Text;
            purchasePrice14.Text = pp14.Text;
            container14.Text = c14.Text;
            userID14.Text = uID14.Text;

            //Fifteenth Row
            Excel.Range tID15 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 15));
            Excel.Range d15 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 15));
            Excel.Range i15 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 15));
            Excel.Range q15 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 15));
            Excel.Range pp15 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 15));
            Excel.Range c15 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 15));
            Excel.Range uID15 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 15));
            textBlocktransID15.Text = tID15.Text;
            date15.Text = d15.Text;
            item15.Text = i15.Text;
            quantity15.Text = q15.Text;
            purchasePrice15.Text = pp15.Text;
            container15.Text = c15.Text;
            userID15.Text = uID15.Text;

            //Sixteenth Row
            Excel.Range tID16 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 16));
            Excel.Range d16 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 16));
            Excel.Range i16 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 16));
            Excel.Range q16 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 16));
            Excel.Range pp16 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 16));
            Excel.Range c16 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 16));
            Excel.Range uID16 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 16));
            textBlocktransID16.Text = tID16.Text;
            date16.Text = d16.Text;
            item16.Text = i16.Text;
            quantity16.Text = q16.Text;
            purchasePrice16.Text = pp16.Text;
            container16.Text = c16.Text;
            userID16.Text = uID16.Text;

            //Seventeenth Row
            Excel.Range tID17 = transHistoryWorksheet.get_Range("A" + (emptyInvRow - 17));
            Excel.Range d17 = transHistoryWorksheet.get_Range("B" + (emptyInvRow - 17));
            Excel.Range i17 = transHistoryWorksheet.get_Range("C" + (emptyInvRow - 17));
            Excel.Range q17 = transHistoryWorksheet.get_Range("D" + (emptyInvRow - 17));
            Excel.Range pp17 = transHistoryWorksheet.get_Range("E" + (emptyInvRow - 17));
            Excel.Range c17 = transHistoryWorksheet.get_Range("F" + (emptyInvRow - 17));
            Excel.Range uID17 = transHistoryWorksheet.get_Range("G" + (emptyInvRow - 17));
            textBlocktransID17.Text = tID17.Text;
            date17.Text = d17.Text;
            item17.Text = i17.Text;
            quantity17.Text = q17.Text;
            purchasePrice17.Text = pp17.Text;
            container17.Text = c17.Text;
            userID17.Text = uID17.Text;

            //Close the excel doc
            transHistoryWorkbook.Close();
            transHistoryApp.Quit();
        }
    }
}
