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
    /// Interaction logic for Add_Remove.xaml
    /// </summary>
    /// 


    public static class MyGlobals
    {
        public static int x = 2;
        public static int addRemove = 0;
        public static int y = 1;
        public static string curItem = "---";
        public static string purchasePrice = "---";

        
    }

    public partial class Add_Remove : Page
    {

        public Add_Remove()
        {
            InitializeComponent();

            //Open up the excel doc
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = false;
            Excel.Workbook workbook = (Excel.Workbook)(excelapp.Workbooks.Open(@"E:\Data\door5Inventory"));
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            MyGlobals.x = 2;

            //Create range to get item names from
            Excel.Range range1 = worksheet.get_Range("B" + MyGlobals.x);
            String item1 = range1.Text;
            itemPicker.Items.Add(item1);
            // Populate itemPicker.Listbox
            while (range1.Text != null && MyGlobals.x <= 100)
            {
                MyGlobals.x = MyGlobals.x + 1;
                range1 = worksheet.get_Range("B" + MyGlobals.x);
                item1 = range1.Text;
                itemPicker.Items.Add(item1);
            }

            //Close excel workbook
            workbook.Close();
            excelapp.Quit();

            refreshAlert.Visibility = Visibility.Hidden;
        }

        // Refresh button that updates info on page based on item you have selected
        private void refresh_Click(object sender, RoutedEventArgs e)
        {
            //Repopulate itemPicker
            //Open up the excel doc
            Excel.Application excelapp5 = new Excel.Application();
            excelapp5.Visible = false;
            Excel.Workbook workbook5 = (Excel.Workbook)(excelapp5.Workbooks.Open(@"E:\Data\door5Inventory"));
            Excel.Worksheet worksheet5 = (Excel.Worksheet)workbook5.ActiveSheet;
            MyGlobals.x = 2;
            //Create range to get item names from
            Excel.Range range1 = worksheet5.get_Range("B" + MyGlobals.x);
            String item1 = range1.Text;
            itemPicker.Items.Add(item1);
            // Populate itemPicker.Listbox
            while (range1.Text != null && MyGlobals.x <= 100)
            {
                MyGlobals.x = MyGlobals.x + 1;
                range1 = worksheet5.get_Range("B" + MyGlobals.x);
                item1 = range1.Text;
                itemPicker.Items.Add(item1);
            }

            // Hide the refresh alert
            refreshAlert.Visibility = Visibility.Hidden;

            // Reset variables and text boxes
            MyGlobals.y = 1;
            MyGlobals.curItem = (string)itemPicker.SelectedItem;
            var excelFoundItem = false;
            string searchItem;
            Excel.Range range2;
            textBoxPurchasePrice.Text = "---";
            refreshAlert.Visibility = Visibility.Hidden;
            if (MyGlobals.curItem == "")
            {
                excelFoundItem = true;
            }

            //Match itemPicker selection with cell in excel doc
            while (excelFoundItem != true && MyGlobals.y < 100)
            {
                MyGlobals.y = MyGlobals.y + 1;
                range2 = worksheet5.get_Range("B" + MyGlobals.y);
                searchItem = (string)range2.Text;
                //MessageBox.Show(searchItem);
                //Once found, stop search and update textBoxCurrentQt
                if (searchItem == MyGlobals.curItem)
                {
                    excelFoundItem = true;
                    Excel.Range foundItemLocation = worksheet5.get_Range("D" + MyGlobals.y);
                    string foundItemQt = foundItemLocation.Text;
                    textBoxCurrentQt.Text = foundItemQt;
                    
                }
                
            }

            //Set Add/Remove to 0
            MyGlobals.addRemove = 0;
            string addRemoveString = Convert.ToString(MyGlobals.addRemove);
            textBoxAddRemove.Text = addRemoveString;

            //Display Container
            Excel.Range itemContainer = worksheet5.get_Range("C" + MyGlobals.y);
            string itemCont = itemContainer.Text;
            textBoxContainer.Text = itemCont;

            //Display Condition
            Excel.Range itemCondition = worksheet5.get_Range("E" + MyGlobals.y);
            string itemCond = itemCondition.Text;
            textBoxCondition.Text = itemCond;

            //Display Date
            string purchaseDate = DateTime.Now.ToString();
            textBlockDate.Text = purchaseDate;


            //Check quantity to see if it is low or zero, return appropriate alert
            Int32 lowQtCheck = Convert.ToInt32(textBoxCurrentQt.Text);
            if (lowQtCheck < 4 && lowQtCheck != 0)
            {
                MessageBox.Show("Running low!");
            }
            if (lowQtCheck == 0)
            {
                MessageBox.Show("All out!");
            }
            //Checks to see if quantity is negative and returns alert
            if (lowQtCheck < 0)
            {
                MessageBox.Show("There is currentley negative " + MyGlobals.curItem + ". Please check the spreadsheet");
            }
            
            // Close workbook
            workbook5.Close();
            excelapp5.Quit();


        }

        private void updateButton_Click(object sender, RoutedEventArgs e)
        {
            //Reopen excel doc
            Excel.Application excelapp3 = new Excel.Application();
            excelapp3.Visible = false;
            Excel.Workbook workbook3 = (Excel.Workbook)(excelapp3.Workbooks.Open(@"E:\Data\door5Inventory"));
            Excel.Worksheet worksheet3 = (Excel.Worksheet)workbook3.ActiveSheet;
            //Define a few more variables

            string addRemoveString = textBoxAddRemove.Text;;
            Int32 addRemoveQt = Convert.ToInt32(addRemoveString);
            string currentQtString;
            Int32 currentQt = 0;
            Int32 newQt;
            
            //Find first available row in Door5Inventory
            string foundEmptyInv = "no";
            Int32 emptyInvRow = 1;
            while (foundEmptyInv != "yes")
            {
                Excel.Range emptyInv = worksheet3.get_Range("A" + emptyInvRow);
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
            //Add/Remove
            //MessageBox.Show(Convert.ToString(MyGlobals.addRemove));
            addRemoveString = textBoxAddRemove.Text;
            MyGlobals.addRemove = Convert.ToInt32(addRemoveString);
            //MessageBox.Show(Convert.ToString(MyGlobals.addRemove));
            addRemoveString = Convert.ToString(MyGlobals.addRemove);
            addRemoveQt = Convert.ToInt32(addRemoveString);
            Excel.Range currentQtRange = worksheet3.get_Range("D" + MyGlobals.y);
            currentQtString = currentQtRange.Text;
            if (newItemCheckBox.IsChecked == false)
            {
                currentQt = Convert.ToInt32(currentQtString);
            }
            else
            {
                currentQt = 0;
            }
            newQt = addRemoveQt + currentQt;
            //Checks to see if the current transaction will make the item now running low
            if (newQt < 4 && newQt != 0)
            {
                MessageBox.Show("This item is now running low");
            }
            //Checks to see if the current transaction will completely use up the item
            if (newQt == 0)
            {
                MessageBox.Show("This item is completely depleted");
            }
            //Checks to see if newQt is negative
            if (newQt < 0)
            {
                MessageBox.Show("This will result in a negative current quantity. Please examine spreadsheet");
            }
            if (newItemCheckBox.IsChecked == true)
            {
                worksheet3.Cells[emptyInvRow, 4] = newQt;
            }
            if (newItemCheckBox.IsChecked == false)
            { 
                worksheet3.Cells[(MyGlobals.y), 4] = newQt;
            }

            //Item ID (if new item
            if (newItemCheckBox.IsChecked == true)
            {
                string newItemID = Convert.ToString(emptyInvRow - 1);
                worksheet3.Cells[emptyInvRow, 1] = newItemID;
            }

            //Item name (if new item)
            if (newItemCheckBox.IsChecked == true)
            {
                MyGlobals.curItem = itemPicker.Text;
                worksheet3.Cells[emptyInvRow, 2] = MyGlobals.curItem;
            }

            //Purchase Price
            if (textBoxPurchasePrice.Text != "---")
            {
                MyGlobals.purchasePrice = textBoxPurchasePrice.Text;
                if (newItemCheckBox.IsChecked == true)
                {
                    worksheet3.Cells[emptyInvRow, 8] = MyGlobals.purchasePrice;
                }
                if (newItemCheckBox.IsChecked == false)
                {
                    worksheet3.Cells[(MyGlobals.y), 8] = MyGlobals.purchasePrice;
                }
            }

            //Container
            if (textBoxContainer.Text != "---")
            {
                string container = textBoxContainer.Text;
                if (newItemCheckBox.IsChecked == true)
                {
                    worksheet3.Cells[emptyInvRow, 3] = container;
                }
                if (newItemCheckBox.IsChecked == false)
                {
                    worksheet3.Cells[(MyGlobals.y), 3] = container;
                }
            }

            //Condition
            if (addRemoveQt > 0)
            {
                string condition = textBoxCondition.Text;
                if (newItemCheckBox.IsChecked == true)
                {
                    worksheet3.Cells[emptyInvRow, 5] = "NEW";
                }
                if (newItemCheckBox.IsChecked == false)
                {
                    worksheet3.Cells[(MyGlobals.y), 5] = condition;
                }
                
            }
          

            //Purchase Date
            if (MyGlobals.purchasePrice != "---")
            {
                string purchaseDate1 = textBlockDate.Text;
                if (newItemCheckBox.IsChecked == true)
                {
                    worksheet3.Cells[emptyInvRow, 7] = purchaseDate1;
                }
                if (newItemCheckBox.IsChecked == false)
                {
                    worksheet3.Cells[(MyGlobals.y), 7] = purchaseDate1;
                }   
            }

            

            //Close workbook
            workbook3.Save();
            workbook3.Close();
            excelapp3.Quit();
            

            //Create new transaction in transactionHistory.xlsx
            //Open up transactionHistory.xlsx
            Excel.Application excelapp4 = new Excel.Application();
            excelapp4.Visible = false;
            Excel.Workbook workbook4 = (Excel.Workbook)(excelapp4.Workbooks.Open(@"E:\Data\transHistory"));
            Excel.Worksheet worksheet4 = (Excel.Worksheet)workbook4.ActiveSheet;
            //MessageBox.Show("Successfully opened up transactionHistory.xlsx");
            //Find first unused transaction history spot
            string foundEmptyTrans = "no";
            Int32 emptyTransRow = 1;
            while (foundEmptyTrans != "yes")
            {
                Excel.Range emptyTrans = worksheet4.get_Range("A" + emptyTransRow);
                //MessageBox.Show(emptyTrans.Text);
                if (emptyTrans.Text == "")
                {
                    foundEmptyTrans = "yes";
                    //MessageBox.Show("Found an empty row");
                }
                if (foundEmptyTrans != "yes")
                {
                    //MessageBox.Show("Row used, looking at next one");
                    emptyTransRow = emptyTransRow + 1;
                }

            }
            //Post Transaction id
            
            
            string transactionId = Convert.ToString(emptyTransRow - 1);
            worksheet4.Cells[emptyTransRow, 1] = transactionId;



            //Post Date
            string purchaseDate = DateTime.Now.ToString();
            textBlockDate.Text = purchaseDate;
            worksheet4.Cells[(emptyTransRow), 2] = purchaseDate;

            //Post Item
            worksheet4.Cells[emptyTransRow, 3] = MyGlobals.curItem;

            //Post Quantity
            string transQt = Convert.ToString(addRemoveQt);
            worksheet4.Cells[emptyTransRow, 4] = transQt;

            //Post Purchase Price
            if (textBoxPurchasePrice.Text != "---")
            {
                MyGlobals.purchasePrice = textBoxPurchasePrice.Text;
                worksheet4.Cells[(emptyTransRow), 5] = MyGlobals.purchasePrice;
            }

            //Post Container
            worksheet4.Cells[emptyTransRow, 6] = textBoxContainer.Text;


            //Post UserID
            worksheet4.Cells[emptyTransRow, 7] = textBoxUserID.Text;

            //Close workbook
            workbook4.Save();
            workbook4.Close();
            excelapp4.Quit();


            //Show refreshAlert image so user knows to refresh the page
            refreshAlert.Visibility = Visibility.Visible;
        }
    }
}