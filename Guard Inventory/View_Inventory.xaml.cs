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
    /// Interaction logic for View_Inventory.xaml
    /// </summary>
    public partial class View_Inventory : Page
    {
        public View_Inventory()
        {
            InitializeComponent();

            //Open up door5Inventory
            Excel.Application door5app = new Excel.Application();
            door5app.Visible = false;
            Excel.Workbook door5Workbook = (Excel.Workbook)(door5app.Workbooks.Open(@"E:\Data\door5Inventory"));
            Excel.Worksheet door5Worksheet = (Excel.Worksheet)door5Workbook.ActiveSheet;

            //Populate item pickers
            ///Item picker 1
            MyGlobals.x = 2;

            //Create range to get item names from
            Excel.Range range1 = door5Worksheet.get_Range("B" + MyGlobals.x);
            String item1 = range1.Text;
            listBoxItem1.Items.Add(item1);
            // Populate itemPicker.Listbox
            while (range1.Text != null && MyGlobals.x <= 100)
            {
                MyGlobals.x = MyGlobals.x + 1;
                range1 = door5Worksheet.get_Range("B" + MyGlobals.x);
                item1 = range1.Text;
                listBoxItem1.Items.Add(item1);
            }
            ///Item picker 2
            MyGlobals.x = 2;

            //Create range to get item names from
            Excel.Range range2 = door5Worksheet.get_Range("B" + MyGlobals.x);
            String item2 = range2.Text;
            listBoxItem2.Items.Add(item2);
            // Populate itemPicker.Listbox
            while (range2.Text != null && MyGlobals.x <= 100)
            {
                MyGlobals.x = MyGlobals.x + 1;
                range2 = door5Worksheet.get_Range("B" + MyGlobals.x);
                item2 = range2.Text;
                listBoxItem2.Items.Add(item2);
            }
            ///Item picker 3
            MyGlobals.x = 2;

            //Create range to get item names from
            Excel.Range range3 = door5Worksheet.get_Range("B" + MyGlobals.x);
            String item3 = range3.Text;
            listBoxItem3.Items.Add(item3);
            // Populate itemPicker.Listbox
            while (range3.Text != null && MyGlobals.x <= 100)
            {
                MyGlobals.x = MyGlobals.x + 1;
                range3 = door5Worksheet.get_Range("B" + MyGlobals.x);
                item3 = range3.Text;
                listBoxItem3.Items.Add(item3);
            }
            ///Item picker 4
            MyGlobals.x = 2;

            //Create range to get item names from
            Excel.Range range4 = door5Worksheet.get_Range("B" + MyGlobals.x);
            String item4 = range4.Text;
            listBoxItem4.Items.Add(item4);
            // Populate itemPicker.Listbox
            while (range4.Text != null && MyGlobals.x <= 100)
            {
                MyGlobals.x = MyGlobals.x + 1;
                range4 = door5Worksheet.get_Range("B" + MyGlobals.x);
                item4 = range4.Text;
                listBoxItem4.Items.Add(item4);
            }
            ///Item picker 5
            MyGlobals.x = 2;

            //Create range to get item names from
            Excel.Range range5 = door5Worksheet.get_Range("B" + MyGlobals.x);
            String item5 = range5.Text;
            listBoxItem5.Items.Add(item5);
            // Populate itemPicker.Listbox
            while (range5.Text != null && MyGlobals.x <= 100)
            {
                MyGlobals.x = MyGlobals.x + 1;
                range5 = door5Worksheet.get_Range("B" + MyGlobals.x);
                item5 = range5.Text;
                listBoxItem5.Items.Add(item5);
            }

            //Hide the text blocks
            textBlockItemID1.Visibility = Visibility.Hidden;
            textBlockContainer1.Visibility = Visibility.Hidden;
            textBlockQuantity1.Visibility = Visibility.Hidden;
            textBlockCondition1.Visibility = Visibility.Hidden;
            textBlockItemID2.Visibility = Visibility.Hidden;
            textBlockContainer2.Visibility = Visibility.Hidden;
            textBlockQuantity2.Visibility = Visibility.Hidden;
            textBlockCondition2.Visibility = Visibility.Hidden;
            textBlockItemID3.Visibility = Visibility.Hidden;
            textBlockContainer3.Visibility = Visibility.Hidden;
            textBlockQuantity3.Visibility = Visibility.Hidden;
            textBlockCondition3.Visibility = Visibility.Hidden;
            textBlockItemID4.Visibility = Visibility.Hidden;
            textBlockContainer4.Visibility = Visibility.Hidden;
            textBlockQuantity4.Visibility = Visibility.Hidden;
            textBlockCondition4.Visibility = Visibility.Hidden;
            textBlockItemID5.Visibility = Visibility.Hidden;
            textBlockContainer5.Visibility = Visibility.Hidden;
            textBlockQuantity5.Visibility = Visibility.Hidden;
            textBlockCondition5.Visibility = Visibility.Hidden;
            

            //Close excel
            door5Workbook.Close();
            door5app.Quit();

        }

        //Add item 1
        private void add1_Click(object sender, RoutedEventArgs e)
        {
            //Open up door5Inventory
            Excel.Application door5app = new Excel.Application();
            door5app.Visible = false;
            Excel.Workbook door5Workbook = (Excel.Workbook)(door5app.Workbooks.Open(@"E:\Data\door5Inventory"));
            Excel.Worksheet door5Worksheet = (Excel.Worksheet)door5Workbook.ActiveSheet;

            // Reset variables and text boxes
            MyGlobals.y = 1;
            MyGlobals.curItem = (string)listBoxItem1.SelectedItem;
            var excelFoundItem = false;
            string searchItem;
            Excel.Range range2;
            if (MyGlobals.curItem == "")
            {
                excelFoundItem = true;
            }

            //Match itemPicker selection with cell in excel doc
            while (excelFoundItem != true && MyGlobals.y < 100)
            {
                MyGlobals.y = MyGlobals.y + 1;
                range2 = door5Worksheet.get_Range("B" + MyGlobals.y);
                searchItem = (string)range2.Text;
                //MessageBox.Show(searchItem);
                //Once found, stop search and update textBoxCurrentQt
                if (searchItem == MyGlobals.curItem)
                {
                    excelFoundItem = true;
                    Excel.Range foundItemID1 = door5Worksheet.get_Range("A" + MyGlobals.y);
                    string itemID1 = foundItemID1.Text;
                    textBlockItemID1.Text = itemID1;
                    textBlockItemID1.Visibility = Visibility.Visible;
                    Excel.Range foundContainer1 = door5Worksheet.get_Range("C" + MyGlobals.y);
                    string container1 = foundContainer1.Text;
                    textBlockContainer1.Text = container1;
                    textBlockContainer1.Visibility = Visibility.Visible;
                    Excel.Range foundQuantity1 = door5Worksheet.get_Range("D" + MyGlobals.y);
                    string quantity1 = foundQuantity1.Text;
                    textBlockQuantity1.Text = quantity1;
                    textBlockQuantity1.Visibility = Visibility.Visible;
                    Excel.Range foundCondition1 = door5Worksheet.get_Range("E" + MyGlobals.y);
                    string condition1 = foundCondition1.Text;
                    textBlockCondition1.Text = condition1;
                    textBlockCondition1.Visibility = Visibility.Visible;

                }
            }

            //Close excel
            door5Workbook.Close();
            door5app.Quit();

        }

        //Add item 2
        private void add2_Click(object sender, RoutedEventArgs e)
        {
            //Open up door5Inventory
            Excel.Application door5app = new Excel.Application();
            door5app.Visible = false;
            Excel.Workbook door5Workbook = (Excel.Workbook)(door5app.Workbooks.Open(@"E:\Data\door5Inventory"));
            Excel.Worksheet door5Worksheet = (Excel.Worksheet)door5Workbook.ActiveSheet;

            // Reset variables and text boxes
            MyGlobals.y = 1;
            MyGlobals.curItem = (string)listBoxItem2.SelectedItem;
            var excelFoundItem = false;
            string searchItem;
            Excel.Range range2;
            if (MyGlobals.curItem == "")
            {
                excelFoundItem = true;
            }

            //Match itemPicker selection with cell in excel doc
            while (excelFoundItem != true && MyGlobals.y < 100)
            {
                MyGlobals.y = MyGlobals.y + 1;
                range2 = door5Worksheet.get_Range("B" + MyGlobals.y);
                searchItem = (string)range2.Text;
                //MessageBox.Show(searchItem);
                //Once found, stop search and update textBoxCurrentQt
                if (searchItem == MyGlobals.curItem)
                {
                    excelFoundItem = true;
                    Excel.Range foundItemID2 = door5Worksheet.get_Range("A" + MyGlobals.y);
                    string itemID2 = foundItemID2.Text;
                    textBlockItemID2.Text = itemID2;
                    textBlockItemID2.Visibility = Visibility.Visible;
                    Excel.Range foundContainer2 = door5Worksheet.get_Range("C" + MyGlobals.y);
                    string container2 = foundContainer2.Text;
                    textBlockContainer2.Text = container2;
                    textBlockContainer2.Visibility = Visibility.Visible;
                    Excel.Range foundQuantity2 = door5Worksheet.get_Range("D" + MyGlobals.y);
                    string quantity2 = foundQuantity2.Text;
                    textBlockQuantity2.Text = quantity2;
                    textBlockQuantity2.Visibility = Visibility.Visible;
                    Excel.Range foundCondition2 = door5Worksheet.get_Range("E" + MyGlobals.y);
                    string condition2 = foundCondition2.Text;
                    textBlockCondition2.Text = condition2;
                    textBlockCondition2.Visibility = Visibility.Visible;

                }
            }

            //Close excel
            door5Workbook.Close();
            door5app.Quit();
        }

        //Add item 3
        private void add3_Click(object sender, RoutedEventArgs e)
        {
            //Open up door5Inventory
            Excel.Application door5app = new Excel.Application();
            door5app.Visible = false;
            Excel.Workbook door5Workbook = (Excel.Workbook)(door5app.Workbooks.Open(@"E:\Data\door5Inventory"));
            Excel.Worksheet door5Worksheet = (Excel.Worksheet)door5Workbook.ActiveSheet;

            // Reset variables and text boxes
            MyGlobals.y = 1;
            MyGlobals.curItem = (string)listBoxItem3.SelectedItem;
            var excelFoundItem = false;
            string searchItem;
            Excel.Range range3;
            if (MyGlobals.curItem == "")
            {
                excelFoundItem = true;
            }

            //Match itemPicker selection with cell in excel doc
            while (excelFoundItem != true && MyGlobals.y < 100)
            {
                MyGlobals.y = MyGlobals.y + 1;
                range3 = door5Worksheet.get_Range("B" + MyGlobals.y);
                searchItem = (string)range3.Text;
                //MessageBox.Show(searchItem);
                //Once found, stop search and update textBoxCurrentQt
                if (searchItem == MyGlobals.curItem)
                {
                    excelFoundItem = true;
                    Excel.Range foundItemID3 = door5Worksheet.get_Range("A" + MyGlobals.y);
                    string itemID3 = foundItemID3.Text;
                    textBlockItemID3.Text = itemID3;
                    textBlockItemID3.Visibility = Visibility.Visible;
                    Excel.Range foundContainer3 = door5Worksheet.get_Range("C" + MyGlobals.y);
                    string container3 = foundContainer3.Text;
                    textBlockContainer3.Text = container3;
                    textBlockContainer3.Visibility = Visibility.Visible;
                    Excel.Range foundQuantity3 = door5Worksheet.get_Range("D" + MyGlobals.y);
                    string quantity3 = foundQuantity3.Text;
                    textBlockQuantity3.Text = quantity3;
                    textBlockQuantity3.Visibility = Visibility.Visible;
                    Excel.Range foundCondition3 = door5Worksheet.get_Range("E" + MyGlobals.y);
                    string condition3 = foundCondition3.Text;
                    textBlockCondition3.Text = condition3;
                    textBlockCondition3.Visibility = Visibility.Visible;

                }
            }

            //Close excel
            door5Workbook.Close();
            door5app.Quit();
        }

        //Add item 4
        private void add4_Click(object sender, RoutedEventArgs e)
        {
            //Open up door5Inventory
            Excel.Application door5app = new Excel.Application();
            door5app.Visible = false;
            Excel.Workbook door5Workbook = (Excel.Workbook)(door5app.Workbooks.Open(@"E:\Data\door5Inventory"));
            Excel.Worksheet door5Worksheet = (Excel.Worksheet)door5Workbook.ActiveSheet;

            // Reset variables and text boxes
            MyGlobals.y = 1;
            MyGlobals.curItem = (string)listBoxItem4.SelectedItem;
            var excelFoundItem = false;
            string searchItem;
            Excel.Range range4;
            if (MyGlobals.curItem == "")
            {
                excelFoundItem = true;
            }

            //Match itemPicker selection with cell in excel doc
            while (excelFoundItem != true && MyGlobals.y < 100)
            {
                MyGlobals.y = MyGlobals.y + 1;
                range4 = door5Worksheet.get_Range("B" + MyGlobals.y);
                searchItem = (string)range4.Text;
                //MessageBox.Show(searchItem);
                //Once found, stop search and update textBoxCurrentQt
                if (searchItem == MyGlobals.curItem)
                {
                    excelFoundItem = true;
                    Excel.Range foundItemID4 = door5Worksheet.get_Range("A" + MyGlobals.y);
                    string itemID4 = foundItemID4.Text;
                    textBlockItemID4.Text = itemID4;
                    textBlockItemID4.Visibility = Visibility.Visible;
                    Excel.Range foundContainer4 = door5Worksheet.get_Range("C" + MyGlobals.y);
                    string container4 = foundContainer4.Text;
                    textBlockContainer4.Text = container4;
                    textBlockContainer4.Visibility = Visibility.Visible;
                    Excel.Range foundQuantity4 = door5Worksheet.get_Range("D" + MyGlobals.y);
                    string quantity4 = foundQuantity4.Text;
                    textBlockQuantity4.Text = quantity4;
                    textBlockQuantity4.Visibility = Visibility.Visible;
                    Excel.Range foundCondition4 = door5Worksheet.get_Range("E" + MyGlobals.y);
                    string condition4 = foundCondition4.Text;
                    textBlockCondition4.Text = condition4;
                    textBlockCondition4.Visibility = Visibility.Visible;

                }
            }

            //Close excel
            door5Workbook.Close();
            door5app.Quit();
        }

        //Add item 5
        private void add5_Click(object sender, RoutedEventArgs e)
        {
            //Open up door5Inventory
            Excel.Application door5app = new Excel.Application();
            door5app.Visible = false;
            Excel.Workbook door5Workbook = (Excel.Workbook)(door5app.Workbooks.Open(@"E:\Data\door5Inventory"));
            Excel.Worksheet door5Worksheet = (Excel.Worksheet)door5Workbook.ActiveSheet;

            // Reset variables and text boxes
            MyGlobals.y = 1;
            MyGlobals.curItem = (string)listBoxItem5.SelectedItem;
            var excelFoundItem = false;
            string searchItem;
            Excel.Range range5;
            if (MyGlobals.curItem == "")
            {
                excelFoundItem = true;
            }

            //Match itemPicker selection with cell in excel doc
            while (excelFoundItem != true && MyGlobals.y < 100)
            {
                MyGlobals.y = MyGlobals.y + 1;
                range5 = door5Worksheet.get_Range("B" + MyGlobals.y);
                searchItem = (string)range5.Text;
                //MessageBox.Show(searchItem);
                //Once found, stop search and update textBoxCurrentQt
                if (searchItem == MyGlobals.curItem)
                {
                    excelFoundItem = true;
                    Excel.Range foundItemID5 = door5Worksheet.get_Range("A" + MyGlobals.y);
                    string itemID5 = foundItemID5.Text;
                    textBlockItemID5.Text = itemID5;
                    textBlockItemID5.Visibility = Visibility.Visible;
                    Excel.Range foundContainer5 = door5Worksheet.get_Range("C" + MyGlobals.y);
                    string container5 = foundContainer5.Text;
                    textBlockContainer5.Text = container5;
                    textBlockContainer5.Visibility = Visibility.Visible;
                    Excel.Range foundQuantity5 = door5Worksheet.get_Range("D" + MyGlobals.y);
                    string quantity5 = foundQuantity5.Text;
                    textBlockQuantity5.Text = quantity5;
                    textBlockQuantity5.Visibility = Visibility.Visible;
                    Excel.Range foundCondition5 = door5Worksheet.get_Range("E" + MyGlobals.y);
                    string condition5 = foundCondition5.Text;
                    textBlockCondition5.Text = condition5;
                    textBlockCondition5.Visibility = Visibility.Visible;

                }
            }

            //Close excel
            door5Workbook.Close();
            door5app.Quit();
        }
        //X buttons
        private void x1_Click(object sender, RoutedEventArgs e)
        {
            //Hide the text blocks
            textBlockItemID1.Visibility = Visibility.Hidden;
            textBlockContainer1.Visibility = Visibility.Hidden;
            textBlockQuantity1.Visibility = Visibility.Hidden;
            textBlockCondition1.Visibility = Visibility.Hidden;

            //Hide the calculated item, if it is even up
            textBlockCalculatedItem1.Visibility = Visibility.Hidden;
        }

        private void x2_Click(object sender, RoutedEventArgs e)
        {
            //Hide the text blocks
            textBlockItemID2.Visibility = Visibility.Hidden;
            textBlockContainer2.Visibility = Visibility.Hidden;
            textBlockQuantity2.Visibility = Visibility.Hidden;
            textBlockCondition2.Visibility = Visibility.Hidden;

            //Hide the calculated item, if it is even up
            textBlockCalculatedItem2.Visibility = Visibility.Hidden;
        }

        private void x3_Click(object sender, RoutedEventArgs e)
        {
            //Hide the text blocks
            textBlockItemID3.Visibility = Visibility.Hidden;
            textBlockContainer3.Visibility = Visibility.Hidden;
            textBlockQuantity3.Visibility = Visibility.Hidden;
            textBlockCondition3.Visibility = Visibility.Hidden;

            //Hide the calculated item, if it is even up
            textBlockCalculatedItem3.Visibility = Visibility.Hidden;
        }

        private void x4_Click(object sender, RoutedEventArgs e)
        {
            //Hide the text blocks
            textBlockItemID4.Visibility = Visibility.Hidden;
            textBlockContainer4.Visibility = Visibility.Hidden;
            textBlockQuantity4.Visibility = Visibility.Hidden;
            textBlockCondition4.Visibility = Visibility.Hidden;

            //Hide the calculated item, if it is even up
            textBlockCalculatedItem4.Visibility = Visibility.Hidden;
        }

        private void x5_Click(object sender, RoutedEventArgs e)
        {
            //Hide the text blocks
            textBlockItemID5.Visibility = Visibility.Hidden;
            textBlockContainer5.Visibility = Visibility.Hidden;
            textBlockQuantity5.Visibility = Visibility.Hidden;
            textBlockCondition5.Visibility = Visibility.Hidden;

            //Hide the calculated item, if it is even up
            textBlockCalculatedItem5.Visibility = Visibility.Hidden;
        }

        //Set table to show 5 items with lowest current quantity
        private void runningLow_Click(object sender, RoutedEventArgs e)
        {
            //Open up door5Inventory
            Excel.Application door5app = new Excel.Application();
            door5app.Visible = false;
            Excel.Workbook door5Workbook = (Excel.Workbook)(door5app.Workbooks.Open(@"E:\Data\door5Inventory"));
            Excel.Worksheet door5Worksheet = (Excel.Worksheet)door5Workbook.ActiveSheet;

            //Find first empty door5Inventory row
            string foundEmptyInv = "no";
            Int32 emptyInvRow = 1;
            while (foundEmptyInv != "yes")
            {
                Excel.Range emptyInv = door5Worksheet.get_Range("A" + emptyInvRow);
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




            //Search through the quantities, searching for first 0, then 1, then 2, then 3.  If
            //a quantity of current search quantity is found, all data from that row populates first
            //available View Inventory row

            //Set up counting variables (also using emptyInvRow)
            string eIRString = Convert.ToString(emptyInvRow);
            //MessageBox.Show(eIRString);
            int populatedRows = 0;
            int searchQuantity = 0;
            int zeroSearchRow = 2;
            int oneSearchRow = 2;
            int twoSearchRow = 2;
            int threeSearchRow = 2;
            int fourSearchRow = 2;
            //While populatedRows is less than 5 this is what will happen
            while (populatedRows < 5)
            {
                //Start searching for quantities that have 0
                while (searchQuantity == 0 && zeroSearchRow < emptyInvRow)
                {
                    //Get value, check if it is zero
                    Excel.Range queryQuant = door5Worksheet.get_Range("D" + zeroSearchRow);
                    string queryQuantString = queryQuant.Text;
                    int queryQuantInt = Convert.ToInt32(queryQuantString);
                    //MessageBox.Show(queryQuantString);
                    //If the value is zero, post data to textBlocks, add 1 to populatedRows
                    if (queryQuantString == "0")
                    {
                        //Use ranges to find other values
                        Excel.Range foundItemID = door5Worksheet.get_Range("A" + zeroSearchRow);
                        Excel.Range foundItem = door5Worksheet.get_Range("B" + zeroSearchRow);
                        Excel.Range foundCont = door5Worksheet.get_Range("C" + zeroSearchRow);
                        Excel.Range foundQaunt = door5Worksheet.get_Range("D" + zeroSearchRow);
                        Excel.Range foundCondition = door5Worksheet.get_Range("E" + zeroSearchRow);

                        //Set text blocks based on how many populated rows there are
                        if (populatedRows == 4)
                        {


                            textBlockItemID5.Text = foundItemID.Text;
                            textBlockItemID5.Visibility = Visibility.Visible;
                            textBlockCalculatedItem5.Text = foundItem.Text;
                            textBlockCalculatedItem5.Visibility = Visibility.Visible;
                            textBlockContainer5.Text = foundCont.Text;
                            textBlockContainer5.Visibility = Visibility.Visible;
                            textBlockQuantity5.Text = foundQaunt.Text;
                            textBlockQuantity5.Visibility = Visibility.Visible;
                            textBlockCondition5.Text = foundCondition.Text;
                            textBlockCondition5.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 5;
                        }


                        if (populatedRows == 3)
                        {


                            textBlockItemID4.Text = foundItemID.Text;
                            textBlockItemID4.Visibility = Visibility.Visible;
                            textBlockCalculatedItem4.Text = foundItem.Text;
                            textBlockCalculatedItem4.Visibility = Visibility.Visible;
                            textBlockContainer4.Text = foundCont.Text;
                            textBlockContainer4.Visibility = Visibility.Visible;
                            textBlockQuantity4.Text = foundQaunt.Text;
                            textBlockQuantity4.Visibility = Visibility.Visible;
                            textBlockCondition4.Text = foundCondition.Text;
                            textBlockCondition4.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 4;
                        }


                        if (populatedRows == 2)
                        {


                            textBlockItemID3.Text = foundItemID.Text;
                            textBlockItemID3.Visibility = Visibility.Visible;
                            textBlockCalculatedItem3.Text = foundItem.Text;
                            textBlockCalculatedItem3.Visibility = Visibility.Visible;
                            textBlockContainer3.Text = foundCont.Text;
                            textBlockContainer3.Visibility = Visibility.Visible;
                            textBlockQuantity3.Text = foundQaunt.Text;
                            textBlockQuantity3.Visibility = Visibility.Visible;
                            textBlockCondition3.Text = foundCondition.Text;
                            textBlockCondition3.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 3;
                        }


                        if (populatedRows == 1)
                        {


                            textBlockItemID2.Text = foundItemID.Text;
                            textBlockItemID2.Visibility = Visibility.Visible;
                            textBlockCalculatedItem2.Text = foundItem.Text;
                            textBlockCalculatedItem2.Visibility = Visibility.Visible;
                            textBlockContainer2.Text = foundCont.Text;
                            textBlockContainer2.Visibility = Visibility.Visible;
                            textBlockQuantity2.Text = foundQaunt.Text;
                            textBlockQuantity2.Visibility = Visibility.Visible;
                            textBlockCondition2.Text = foundCondition.Text;
                            textBlockCondition2.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 2;
                        }


                        if (populatedRows == 0)
                        {


                            textBlockItemID1.Text = foundItemID.Text;
                            textBlockItemID1.Visibility = Visibility.Visible;
                            textBlockCalculatedItem1.Text = foundItem.Text;
                            textBlockCalculatedItem1.Visibility = Visibility.Visible;
                            textBlockContainer1.Text = foundCont.Text;
                            textBlockContainer1.Visibility = Visibility.Visible;
                            textBlockQuantity1.Text = foundQaunt.Text;
                            textBlockQuantity1.Visibility = Visibility.Visible;
                            textBlockCondition1.Text = foundCondition.Text;
                            textBlockCondition1.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 1;
                        }
                        
                        
                        
                        
                    }

                    //Add one to zeroSearchRow because no matter what you will need to look at the next row...
                    //MessageBox.Show("zeroSearchRow =" + Convert.ToString(zeroSearchRow));
                    zeroSearchRow = zeroSearchRow + 1;
                    //MessageBox.Show("zeroSearchRow =" + Convert.ToString(zeroSearchRow));


                    //... unless you are out of rows
                    if (zeroSearchRow == emptyInvRow)
                    {
                        //MessageBox.Show("zeroSearchRow =" + Convert.ToString(zeroSearchRow) + " emptyInvRow =" + Convert.ToString(emptyInvRow));
                        //At which point you set the searchQuantity equal to itself plus one, so you start looking for the next quantity
                        searchQuantity = 1;
                        //MessageBox.Show(Convert.ToString(searchQuantity));
                    }

                }

                //Start searching for quantities that have 1
                while (searchQuantity == 1)
                {
                    //MessageBox.Show("Yep, its working");
                    //Get value, check if it is zero
                    //Also set zeroSearchRow to oneSearchRow so that it starts from the beginning looking for 1
                    
                    Excel.Range queryQuant1 = door5Worksheet.get_Range("D" + oneSearchRow);
                    string queryQuantString = queryQuant1.Text;
                    int queryQuantInt = Convert.ToInt32(queryQuantString);
                    //MessageBox.Show(queryQuantString);
                    //If the value is one, post data to textBlocks, add 1 to populatedRows
                    if (queryQuantString == "1")
                    {
                        //Use ranges to find other values
                        Excel.Range foundItemID = door5Worksheet.get_Range("A" + oneSearchRow);
                        Excel.Range foundItem = door5Worksheet.get_Range("B" + oneSearchRow);
                        Excel.Range foundCont = door5Worksheet.get_Range("C" + oneSearchRow);
                        Excel.Range foundQaunt = door5Worksheet.get_Range("D" + oneSearchRow);
                        Excel.Range foundCondition = door5Worksheet.get_Range("E" + oneSearchRow);

                        //Set text blocks based on how many populated rows there are
                        if (populatedRows == 4)
                        {


                            textBlockItemID5.Text = foundItemID.Text;
                            textBlockItemID5.Visibility = Visibility.Visible;
                            textBlockCalculatedItem5.Text = foundItem.Text;
                            textBlockCalculatedItem5.Visibility = Visibility.Visible;
                            textBlockContainer5.Text = foundCont.Text;
                            textBlockContainer5.Visibility = Visibility.Visible;
                            textBlockQuantity5.Text = foundQaunt.Text;
                            textBlockQuantity5.Visibility = Visibility.Visible;
                            textBlockCondition5.Text = foundCondition.Text;
                            textBlockCondition5.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 5;
                        }


                        if (populatedRows == 3)
                        {


                            textBlockItemID4.Text = foundItemID.Text;
                            textBlockItemID4.Visibility = Visibility.Visible;
                            textBlockCalculatedItem4.Text = foundItem.Text;
                            textBlockCalculatedItem4.Visibility = Visibility.Visible;
                            textBlockContainer4.Text = foundCont.Text;
                            textBlockContainer4.Visibility = Visibility.Visible;
                            textBlockQuantity4.Text = foundQaunt.Text;
                            textBlockQuantity4.Visibility = Visibility.Visible;
                            textBlockCondition4.Text = foundCondition.Text;
                            textBlockCondition4.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 4;
                        }


                        if (populatedRows == 2)
                        {


                            textBlockItemID3.Text = foundItemID.Text;
                            textBlockItemID3.Visibility = Visibility.Visible;
                            textBlockCalculatedItem3.Text = foundItem.Text;
                            textBlockCalculatedItem3.Visibility = Visibility.Visible;
                            textBlockContainer3.Text = foundCont.Text;
                            textBlockContainer3.Visibility = Visibility.Visible;
                            textBlockQuantity3.Text = foundQaunt.Text;
                            textBlockQuantity3.Visibility = Visibility.Visible;
                            textBlockCondition3.Text = foundCondition.Text;
                            textBlockCondition3.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 3;
                        }


                        if (populatedRows == 1)
                        {


                            textBlockItemID2.Text = foundItemID.Text;
                            textBlockItemID2.Visibility = Visibility.Visible;
                            textBlockCalculatedItem2.Text = foundItem.Text;
                            textBlockCalculatedItem2.Visibility = Visibility.Visible;
                            textBlockContainer2.Text = foundCont.Text;
                            textBlockContainer2.Visibility = Visibility.Visible;
                            textBlockQuantity2.Text = foundQaunt.Text;
                            textBlockQuantity2.Visibility = Visibility.Visible;
                            textBlockCondition2.Text = foundCondition.Text;
                            textBlockCondition2.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 2;
                        }


                        if (populatedRows == 0)
                        {


                            textBlockItemID1.Text = foundItemID.Text;
                            textBlockItemID1.Visibility = Visibility.Visible;
                            textBlockCalculatedItem1.Text = foundItem.Text;
                            textBlockCalculatedItem1.Visibility = Visibility.Visible;
                            textBlockContainer1.Text = foundCont.Text;
                            textBlockContainer1.Visibility = Visibility.Visible;
                            textBlockQuantity1.Text = foundQaunt.Text;
                            textBlockQuantity1.Visibility = Visibility.Visible;
                            textBlockCondition1.Text = foundCondition.Text;
                            textBlockCondition1.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 1;
                        }




                    }

                    //Add one to oneSearchRow because no matter what you will need to look at the next row...
                    oneSearchRow = oneSearchRow + 1;

                    //... unless you are out of rows
                    if (oneSearchRow == emptyInvRow)
                    {
                        //At which point you set the searchQuantity equal to itself plus one, so you start looking for the next quantity
                        searchQuantity = 2;
                    }

                }

                //Start searching for quantities that have 2
                while (searchQuantity == 2)
                {
                    //Get value, check if it is zero
                    //Also reset oneSearchRow to twoSearchRow so that it starts from the beginning
                    
                    Excel.Range queryQuant = door5Worksheet.get_Range("D" + twoSearchRow);
                    string queryQuantString = queryQuant.Text;
                    int queryQuantInt = Convert.ToInt32(queryQuantString);
                    //If the value is one, post data to textBlocks, add 1 to populatedRows
                    if (queryQuantString == "2")
                    {
                        //Use ranges to find other values
                        Excel.Range foundItemID = door5Worksheet.get_Range("A" + twoSearchRow);
                        Excel.Range foundItem = door5Worksheet.get_Range("B" + twoSearchRow);
                        Excel.Range foundCont = door5Worksheet.get_Range("C" + twoSearchRow);
                        Excel.Range foundQaunt = door5Worksheet.get_Range("D" + twoSearchRow);
                        Excel.Range foundCondition = door5Worksheet.get_Range("E" + twoSearchRow);

                        //Set text blocks based on how many populated rows there are
                        if (populatedRows == 4)
                        {


                            textBlockItemID5.Text = foundItemID.Text;
                            textBlockItemID5.Visibility = Visibility.Visible;
                            textBlockCalculatedItem5.Text = foundItem.Text;
                            textBlockCalculatedItem5.Visibility = Visibility.Visible;
                            textBlockContainer5.Text = foundCont.Text;
                            textBlockContainer5.Visibility = Visibility.Visible;
                            textBlockQuantity5.Text = foundQaunt.Text;
                            textBlockQuantity5.Visibility = Visibility.Visible;
                            textBlockCondition5.Text = foundCondition.Text;
                            textBlockCondition5.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 5;
                        }


                        if (populatedRows == 3)
                        {


                            textBlockItemID4.Text = foundItemID.Text;
                            textBlockItemID4.Visibility = Visibility.Visible;
                            textBlockCalculatedItem4.Text = foundItem.Text;
                            textBlockCalculatedItem4.Visibility = Visibility.Visible;
                            textBlockContainer4.Text = foundCont.Text;
                            textBlockContainer4.Visibility = Visibility.Visible;
                            textBlockQuantity4.Text = foundQaunt.Text;
                            textBlockQuantity4.Visibility = Visibility.Visible;
                            textBlockCondition4.Text = foundCondition.Text;
                            textBlockCondition4.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 4;
                        }


                        if (populatedRows == 2)
                        {


                            textBlockItemID3.Text = foundItemID.Text;
                            textBlockItemID3.Visibility = Visibility.Visible;
                            textBlockCalculatedItem3.Text = foundItem.Text;
                            textBlockCalculatedItem3.Visibility = Visibility.Visible;
                            textBlockContainer3.Text = foundCont.Text;
                            textBlockContainer3.Visibility = Visibility.Visible;
                            textBlockQuantity3.Text = foundQaunt.Text;
                            textBlockQuantity3.Visibility = Visibility.Visible;
                            textBlockCondition3.Text = foundCondition.Text;
                            textBlockCondition3.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 3;
                        }


                        if (populatedRows == 1)
                        {


                            textBlockItemID2.Text = foundItemID.Text;
                            textBlockItemID2.Visibility = Visibility.Visible;
                            textBlockCalculatedItem2.Text = foundItem.Text;
                            textBlockCalculatedItem2.Visibility = Visibility.Visible;
                            textBlockContainer2.Text = foundCont.Text;
                            textBlockContainer2.Visibility = Visibility.Visible;
                            textBlockQuantity2.Text = foundQaunt.Text;
                            textBlockQuantity2.Visibility = Visibility.Visible;
                            textBlockCondition2.Text = foundCondition.Text;
                            textBlockCondition2.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 2;
                        }


                        if (populatedRows == 0)
                        {


                            textBlockItemID1.Text = foundItemID.Text;
                            textBlockItemID1.Visibility = Visibility.Visible;
                            textBlockCalculatedItem1.Text = foundItem.Text;
                            textBlockCalculatedItem1.Visibility = Visibility.Visible;
                            textBlockContainer1.Text = foundCont.Text;
                            textBlockContainer1.Visibility = Visibility.Visible;
                            textBlockQuantity1.Text = foundQaunt.Text;
                            textBlockQuantity1.Visibility = Visibility.Visible;
                            textBlockCondition1.Text = foundCondition.Text;
                            textBlockCondition1.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 1;
                        }




                    }

                    //Add one to twoSearchRow because no matter what you will need to look at the next row...
                    twoSearchRow = twoSearchRow + 1;

                    //... unless you are out of rows
                    if (twoSearchRow == emptyInvRow)
                    {
                        //At which point you set the searchQuantity equal to itself plus one, so you start looking for the next quantity
                        searchQuantity = 3;
                    }

                }

                //Start searching for quantities that have 3
                while (searchQuantity == 3)
                {
                    //Get value, check if it is zero
                    //Also reset threeSearchRow so that it starts from the beginning
                    
                    Excel.Range queryQuant = door5Worksheet.get_Range("D" + threeSearchRow);
                    string queryQuantString = queryQuant.Text;
                    int queryQuantInt = Convert.ToInt32(queryQuantString);
                    //If the value is one, post data to textBlocks, add 1 to populatedRows
                    if (queryQuantString == "3")
                    {
                        //Use ranges to find other values
                        Excel.Range foundItemID = door5Worksheet.get_Range("A" + threeSearchRow);
                        Excel.Range foundItem = door5Worksheet.get_Range("B" + threeSearchRow);
                        Excel.Range foundCont = door5Worksheet.get_Range("C" + threeSearchRow);
                        Excel.Range foundQaunt = door5Worksheet.get_Range("D" + threeSearchRow);
                        Excel.Range foundCondition = door5Worksheet.get_Range("E" + threeSearchRow);

                        //Set text blocks based on how many populated rows there are
                        if (populatedRows == 4)
                        {


                            textBlockItemID5.Text = foundItemID.Text;
                            textBlockItemID5.Visibility = Visibility.Visible;
                            textBlockCalculatedItem5.Text = foundItem.Text;
                            textBlockCalculatedItem5.Visibility = Visibility.Visible;
                            textBlockContainer5.Text = foundCont.Text;
                            textBlockContainer5.Visibility = Visibility.Visible;
                            textBlockQuantity5.Text = foundQaunt.Text;
                            textBlockQuantity5.Visibility = Visibility.Visible;
                            textBlockCondition5.Text = foundCondition.Text;
                            textBlockCondition5.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 5;
                        }


                        if (populatedRows == 3)
                        {


                            textBlockItemID4.Text = foundItemID.Text;
                            textBlockItemID4.Visibility = Visibility.Visible;
                            textBlockCalculatedItem4.Text = foundItem.Text;
                            textBlockCalculatedItem4.Visibility = Visibility.Visible;
                            textBlockContainer4.Text = foundCont.Text;
                            textBlockContainer4.Visibility = Visibility.Visible;
                            textBlockQuantity4.Text = foundQaunt.Text;
                            textBlockQuantity4.Visibility = Visibility.Visible;
                            textBlockCondition4.Text = foundCondition.Text;
                            textBlockCondition4.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 4;
                        }


                        if (populatedRows == 2)
                        {


                            textBlockItemID3.Text = foundItemID.Text;
                            textBlockItemID3.Visibility = Visibility.Visible;
                            textBlockCalculatedItem3.Text = foundItem.Text;
                            textBlockCalculatedItem3.Visibility = Visibility.Visible;
                            textBlockContainer3.Text = foundCont.Text;
                            textBlockContainer3.Visibility = Visibility.Visible;
                            textBlockQuantity3.Text = foundQaunt.Text;
                            textBlockQuantity3.Visibility = Visibility.Visible;
                            textBlockCondition3.Text = foundCondition.Text;
                            textBlockCondition3.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 3;
                        }


                        if (populatedRows == 1)
                        {


                            textBlockItemID2.Text = foundItemID.Text;
                            textBlockItemID2.Visibility = Visibility.Visible;
                            textBlockCalculatedItem2.Text = foundItem.Text;
                            textBlockCalculatedItem2.Visibility = Visibility.Visible;
                            textBlockContainer2.Text = foundCont.Text;
                            textBlockContainer2.Visibility = Visibility.Visible;
                            textBlockQuantity2.Text = foundQaunt.Text;
                            textBlockQuantity2.Visibility = Visibility.Visible;
                            textBlockCondition2.Text = foundCondition.Text;
                            textBlockCondition2.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 2;
                        }


                        if (populatedRows == 0)
                        {


                            textBlockItemID1.Text = foundItemID.Text;
                            textBlockItemID1.Visibility = Visibility.Visible;
                            textBlockCalculatedItem1.Text = foundItem.Text;
                            textBlockCalculatedItem1.Visibility = Visibility.Visible;
                            textBlockContainer1.Text = foundCont.Text;
                            textBlockContainer1.Visibility = Visibility.Visible;
                            textBlockQuantity1.Text = foundQaunt.Text;
                            textBlockQuantity1.Visibility = Visibility.Visible;
                            textBlockCondition1.Text = foundCondition.Text;
                            textBlockCondition1.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 1;
                        }




                    }

                    //Add one to threeSearchRow because no matter what you will need to look at the next row...
                    threeSearchRow = threeSearchRow + 1;

                    //... unless you are out of rows
                    if (threeSearchRow == emptyInvRow)
                    {
                        //At which point you set the searchQuantity equal to itself plus one, so you start looking for the next quantity
                        searchQuantity = 4;
                    }

                }

                //Start searching for quantities that have 4
                while (searchQuantity == 4)
                {
                    //Get value, check if it is four
                    //Also reset fourSearchRow so that it starts from the beginning
                    
                    Excel.Range queryQuant = door5Worksheet.get_Range("D" + fourSearchRow);
                    string queryQuantString = queryQuant.Text;
                    int queryQuantInt = Convert.ToInt32(queryQuantString);
                    //If the value is one, post data to textBlocks, add 1 to populatedRows
                    if (queryQuantString == "4")
                    {
                        //Use ranges to find other values
                        Excel.Range foundItemID = door5Worksheet.get_Range("A" + fourSearchRow);
                        Excel.Range foundItem = door5Worksheet.get_Range("B" + fourSearchRow);
                        Excel.Range foundCont = door5Worksheet.get_Range("C" + fourSearchRow);
                        Excel.Range foundQaunt = door5Worksheet.get_Range("D" + fourSearchRow);
                        Excel.Range foundCondition = door5Worksheet.get_Range("E" + fourSearchRow);

                        //Set text blocks based on how many populated rows there are
                        if (populatedRows == 4)
                        {


                            textBlockItemID5.Text = foundItemID.Text;
                            textBlockItemID5.Visibility = Visibility.Visible;
                            textBlockCalculatedItem5.Text = foundItem.Text;
                            textBlockCalculatedItem5.Visibility = Visibility.Visible;
                            textBlockContainer5.Text = foundCont.Text;
                            textBlockContainer5.Visibility = Visibility.Visible;
                            textBlockQuantity5.Text = foundQaunt.Text;
                            textBlockQuantity5.Visibility = Visibility.Visible;
                            textBlockCondition5.Text = foundCondition.Text;
                            textBlockCondition5.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 5;
                        }


                        if (populatedRows == 3)
                        {


                            textBlockItemID4.Text = foundItemID.Text;
                            textBlockItemID4.Visibility = Visibility.Visible;
                            textBlockCalculatedItem4.Text = foundItem.Text;
                            textBlockCalculatedItem4.Visibility = Visibility.Visible;
                            textBlockContainer4.Text = foundCont.Text;
                            textBlockContainer4.Visibility = Visibility.Visible;
                            textBlockQuantity4.Text = foundQaunt.Text;
                            textBlockQuantity4.Visibility = Visibility.Visible;
                            textBlockCondition4.Text = foundCondition.Text;
                            textBlockCondition4.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 4;
                        }


                        if (populatedRows == 2)
                        {


                            textBlockItemID3.Text = foundItemID.Text;
                            textBlockItemID3.Visibility = Visibility.Visible;
                            textBlockCalculatedItem3.Text = foundItem.Text;
                            textBlockCalculatedItem3.Visibility = Visibility.Visible;
                            textBlockContainer3.Text = foundCont.Text;
                            textBlockContainer3.Visibility = Visibility.Visible;
                            textBlockQuantity3.Text = foundQaunt.Text;
                            textBlockQuantity3.Visibility = Visibility.Visible;
                            textBlockCondition3.Text = foundCondition.Text;
                            textBlockCondition3.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 3;
                        }


                        if (populatedRows == 1)
                        {


                            textBlockItemID2.Text = foundItemID.Text;
                            textBlockItemID2.Visibility = Visibility.Visible;
                            textBlockCalculatedItem2.Text = foundItem.Text;
                            textBlockCalculatedItem2.Visibility = Visibility.Visible;
                            textBlockContainer2.Text = foundCont.Text;
                            textBlockContainer2.Visibility = Visibility.Visible;
                            textBlockQuantity2.Text = foundQaunt.Text;
                            textBlockQuantity2.Visibility = Visibility.Visible;
                            textBlockCondition2.Text = foundCondition.Text;
                            textBlockCondition2.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 2;
                        }


                        if (populatedRows == 0)
                        {


                            textBlockItemID1.Text = foundItemID.Text;
                            textBlockItemID1.Visibility = Visibility.Visible;
                            textBlockCalculatedItem1.Text = foundItem.Text;
                            textBlockCalculatedItem1.Visibility = Visibility.Visible;
                            textBlockContainer1.Text = foundCont.Text;
                            textBlockContainer1.Visibility = Visibility.Visible;
                            textBlockQuantity1.Text = foundQaunt.Text;
                            textBlockQuantity1.Visibility = Visibility.Visible;
                            textBlockCondition1.Text = foundCondition.Text;
                            textBlockCondition1.Visibility = Visibility.Visible;

                            //Add 1 to populatedRows becuase thats what we just did
                            populatedRows = 1;
                        }




                    }

                    //Add one to fourSearchRow because no matter what you will need to look at the next row...
                    fourSearchRow = fourSearchRow + 1;

                    //... unless you are out of rows
                    if (fourSearchRow == emptyInvRow)
                    {
                        //Since we are making the decision (for the test purposes, anyways) that quantities that are "running low" are 4 or less,
                        //and we have just finished checking the spreadsheet for quantities of 4, we need to get the function to stop.
                        //We do that by declaring populatedRows = "five", which will end the over-arching while statement
                        populatedRows = 5;
                        searchQuantity = 8;
                        //goodluck
                    }

                }
                //make sure that it stops
                populatedRows = 5;
            }


            //Close excel
            //door5Workbook.Save();
            door5Workbook.Close();
            door5app.Quit();
        }
    }
}
