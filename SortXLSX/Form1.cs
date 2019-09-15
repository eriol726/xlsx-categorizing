using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader.Core;
using GemBox.Spreadsheet;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

using Data = Google.Apis.Sheets.v4.Data;
using System.Configuration;

namespace SortXLSX
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            InitializeComponent();
        }
        /*
         * Download the xml file from SEB containing all the data from one month
         * place it in the /bin/debug folder and rename it to seb_month
         * change the variable 'monthIndex' to the current month 
         * */

        DataSet monthSet, yearSet;
        public static DataTable tempDataTable;
        RootObject categoryList = new RootObject();
        Category c = new Category();
        int categorySum = 0;
        int categoryIndex = 0;
        int monthIndex = 7;
        string spreadsheetId = ConfigurationManager.AppSettings["spreadsheetId"];
        string[] monthStrings = {"Januari", "Februari", "Mars", "April", "Maj", "Juni", "Juli", "Augusti", "September", "Oktober", "November", "December", "Summering" };
        public SpreadsheetsResource.ValuesResource.GetRequest request1;

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
            {
                //ofd.ShowDialog() == DialogResult.OK
                if (true)
                {
                    FileStream fsYear = File.Open("Ekonomi_2019_downloaded.xlsx", FileMode.Open, FileAccess.Read);
                    FileStream fsMonth = File.Open("seb_"+monthStrings[monthIndex]+".xlsx", FileMode.Open, FileAccess.Read);
                    IExcelDataReader readerYear, readerMonth;
                    
                    if (ofd.FilterIndex == 1)
                    {
                        readerYear = ExcelReaderFactory.CreateOpenXmlReader(fsYear);
                        readerMonth = ExcelReaderFactory.CreateOpenXmlReader(fsMonth);
                    }
                    else
                    {
                        readerYear = ExcelReaderFactory.CreateReader(fsYear);
                        readerMonth = ExcelReaderFactory.CreateReader(fsMonth);
                    }


                    monthSet = readerMonth.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    //yearSet = readerYear.AsDataSet(new ExcelDataSetConfiguration()
                    //{
                    //    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    //    {
                    //        UseHeaderRow = true
                    //    }
                    //});

                    yearSet = new DataSet();
                    DataTable monthTable = new DataTable("Januari");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("Februari");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("Mars");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("April");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("Maj");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("Juni");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("Juli");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("Augusti");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("September");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("Oktober");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("November");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("December");
                    yearSet.Tables.Add(monthTable);
                    monthTable = new DataTable("Summering");
                    yearSet.Tables.Add(monthTable);


                    // remove unnecessary columns
                    DataTable dt = monthSet.Tables[0];  

                    dt.Columns.Remove(dt.Columns[5]);
                    dt.Columns.Remove(dt.Columns[2]);
                    dt.Columns.Remove(dt.Columns[1]);
                    monthSet.Tables[0].AcceptChanges();

                    foreach (DataRow row in monthSet.Tables[0].Rows)
                    {
                        string year = row[0].ToString();

                        // remove rows with a an empty description
                        if (row[1] == DBNull.Value)
                        {
                            row.Delete();
                        }
                        //radera alla rader som inte börjar med ett datum
                        else if(!year.ToString().Contains("2019") )
                        {
                            row.Delete();
                        }
                        else if(row[1].ToString().Contains("ISK"))
                        {
                            row.Delete();
                        }
                        else if (row[1].ToString().Contains("PREMIE FÖRS"))
                        {
                            row.Delete();
                        }
                        else if (Convert.ToInt32(row[2]) > 0)
                        {
                            row.Delete();
                        }
                    }

                    //important to reorder table when we delete
                    monthSet.Tables[0].AcceptChanges();

                    categoryList.items = new List<Item>
                          {
                             new Item { title = "Hyra", sum = 0, annualSum = 0 },
                             new Item { title = "Bredband/Mobil", sum = 0, annualSum = 0 },
                             new Item { title = "Basutgifter", sum = 0, annualSum = 0 },
                             new Item { title = "Mat", sum = 0, annualSum = 0 },
                             new Item { title = "Sprit", sum = 0, annualSum = 0 },
                             new Item { title = "Swish", sum = 0, annualSum = 0 },
                             new Item { title = "Resor", sum = 0, annualSum = 0 },
                             new Item { title = "Övrigt", sum = 0, annualSum = 0 }
                          };

                    // sumerize induvidual caregory
                    monthSet.Tables[0].AcceptChanges();
                    tempDataTable = monthSet.Tables[0].Clone();

                    DataRow titleRow = tempDataTable.NewRow();
                    DataRow emptyRow = tempDataTable.NewRow();

                    var hyra = new List<string>() { "HYRES" };
                    categorization(titleRow, emptyRow, hyra);

                    var bredband = new List<string>() { "TELEN", "COMV" };
                    categorization(titleRow, emptyRow, bredband);

                    var basutgifter = new List<string>() { "DFS", "SPOTIFY", "AEA", "SVEGOT", "WORLDCLASS", "ENKLA VARDAG", "CSN", "EON ", "E ON" };
                    categorization(titleRow, emptyRow, basutgifter);

                    var mat = new List<string>() { "COOP", "NETT", "ICA", "WILLYS", "LIDL", "SELECTA", "PIZZ", "PIZ", "VISUALISERIN", "ENOTEKET", "PRESSBYRÅN", "RESQ CLUB", "GRILLEN", "ESPRESSO", "KROG", "CAFE", "CIRCLE K", "PREEM", "MAX", "KOND", "HEMKÖP", "FALAFEL", "PIGEONSTREET" };
                    categorization(titleRow, emptyRow, mat);

                    var sprit = new List<string>() { "SYSTEM", "SALIGA", "ARBIS", "CROMWELL", "LERO", "VÄRDENS", "KARHUSET", "BROOKLYN", "TRÄDGÅR", "LION", "WATTS", "S 12", "O LEARYS", "BAR", "SOFO", "SKANETRAFIKE", "HUSET UNDER", "SODERKELLARE", "BROADWAY", "RESTAURANG K" };
                    categorization(titleRow, emptyRow, sprit);

                    var swish = new List<string>() { "56940769587" };
                    categorization(titleRow, emptyRow, swish);

                    var resor = new List<string>() { "ÖSTG", "SNELLTAGET", "SJ INTERNETB", "TAXI", "SJ MOB" };
                    categorization(titleRow, emptyRow, resor);

                    // we dont need to insert rows from category 'others' because they are already in source table 
                    titleRow = tempDataTable.NewRow();
                    titleRow[0] = c.Other.title;
                    tempDataTable.Rows.InsertAt(titleRow, tempDataTable.Rows.Count);
                    
                    // sumerize category others
                    foreach (DataRow row in monthSet.Tables[0].Rows)
                    {
                        categoryList.items[7].sum += Convert.ToInt32(row[2]);
                    }

                    monthSet.Tables[0].AcceptChanges();

                    int rowIndex = 0;
                    // insert categorized tabel from tempDataTable to monthSet table
                    foreach (DataRow row in tempDataTable.Rows)
                    {
                        DataRow newRow = monthSet.Tables[0].NewRow();
                        newRow[0] = row[0];
                        newRow[1] = row[1];
                        newRow[2] = row[2];
                        monthSet.Tables[0].Rows.InsertAt(newRow, rowIndex);

                        rowIndex++;
                    }

                    //Addning 3 new columns to the right
                    monthSet.Tables[0].Columns[0].ColumnName = "Rubrik";
                    monthSet.Tables[0].Columns.Add("", typeof(System.Int32));
                    monthSet.Tables[0].Columns.Add("Kategorier", typeof(System.String));
                    monthSet.Tables[0].Columns.Add("Summa", typeof(System.Int32));

                    //Assaigning title and sums in the summarize columns
                    monthSet.Tables[0].Rows[0][4] = categoryList.items[0].title;
                    monthSet.Tables[0].Rows[0][5] = categoryList.items[0].sum;
                    monthSet.Tables[0].Rows[1][4] = categoryList.items[1].title;
                    monthSet.Tables[0].Rows[1][5] = categoryList.items[1].sum;
                    monthSet.Tables[0].Rows[2][4] = categoryList.items[2].title;
                    monthSet.Tables[0].Rows[2][5] = categoryList.items[2].sum;
                    monthSet.Tables[0].Rows[3][4] = categoryList.items[3].title;
                    monthSet.Tables[0].Rows[3][5] = categoryList.items[3].sum;
                    monthSet.Tables[0].Rows[4][4] = categoryList.items[4].title;
                    monthSet.Tables[0].Rows[4][5] = categoryList.items[4].sum;
                    monthSet.Tables[0].Rows[5][4] = categoryList.items[5].title;
                    monthSet.Tables[0].Rows[5][5] = categoryList.items[5].sum;
                    monthSet.Tables[0].Rows[6][4] = categoryList.items[6].title;
                    monthSet.Tables[0].Rows[6][5] = categoryList.items[6].sum;
                    monthSet.Tables[0].Rows[7][4] = categoryList.items[7].title;
                    monthSet.Tables[0].Rows[7][5] = categoryList.items[7].sum;
                    
                    // sumerize all categories for the entire month 
                    int categorySums = 0;
                    foreach (DataRow row in monthSet.Tables[0].Rows)
                    {
                        if(row[5] != DBNull.Value)
                        {
                            categorySums += Convert.ToInt32(row[5]);
                        }
                    }
                    monthSet.Tables[0].Rows[8][4] = "Summa: ";
                    monthSet.Tables[0].Rows[8][5] = categorySums;

                    // read data from google sheets and assign it to 
                    readGoogleSheet(monthSet);

                    // replace month table in year file
                    yearSet.Tables[monthIndex].Clear();
                    yearSet.Tables[monthIndex].Columns.RemoveAt(0);
                    yearSet.Tables[monthIndex].Columns.RemoveAt(0);
                    yearSet.Tables[monthIndex].Columns.RemoveAt(0);
                    yearSet.Tables[monthIndex].Columns.RemoveAt(0);
                    yearSet.Tables[monthIndex].Columns.RemoveAt(0);
                    yearSet.Tables[monthIndex].Columns.RemoveAt(0);

                    // coppy data from month table to year data set
                    DataTableReader reader = new DataTableReader(monthSet.Tables[0]);
                    yearSet.Tables[monthIndex].Load(reader);

                    sumerizeMonth();

                    writeToGoogleSheet(yearSet.Tables[monthIndex]);
                    monthIndex = 12;
                    writeToGoogleSheet(yearSet.Tables[monthIndex]);

                    //tempDataTable = monthSet.Tables[0];

                    foreach (DataTable dataTable in yearSet.Tables)
                    {
                        cboSheet.Items.Add(dataTable.TableName);
                    }

                    readerMonth.Close();
                    readerYear.Close();
                }
            }
        }

        private void CboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView.DataSource = yearSet.Tables[cboSheet.SelectedIndex];
        }

        // Find a category name in month table, add it to tempDataTable, sumerize the amounts, delete the row from the month table 
        private void search_description(string title)
        {
            
            foreach (DataRow row in monthSet.Tables[0].Rows)
            {
                string description = row[1].ToString();

                if (description.Contains(title) )
                {
                    tempDataTable.Rows.Add(row.ItemArray);
                    categorySum += Convert.ToInt32(row[2]);
                    row.Delete();
                }
            }

            monthSet.Tables[0].AcceptChanges();
            categoryList.items[categoryIndex].sum = categorySum;
        }

        private void categorization(DataRow titleRow, DataRow emptyRow, List<string> Category)
        {
            // adding header
            titleRow = tempDataTable.NewRow();
            titleRow[0] = categoryList.items[categoryIndex].title;
            tempDataTable.Rows.InsertAt(titleRow, tempDataTable.Rows.Count);

            foreach (String categoryStr in Category)
            {
                search_description(categoryStr);
            }

            // adding space
            emptyRow = tempDataTable.NewRow();
            tempDataTable.Rows.InsertAt(emptyRow, tempDataTable.Rows.Count);
            // reset the categorySum for the new category
            categorySum = 0;

            categoryIndex++;
        }

        private void saveDStoFile(DataSet yearDs)
        {
            var workbook = new ExcelFile();

            ExcelWorksheet worksheet;

            for (int i = 0; i < yearSet.Tables.Count; i++)
            {
                worksheet = workbook.Worksheets.Add("sheet" + i);

                worksheet.InsertDataTable(yearSet.Tables[i],
                new InsertDataTableOptions()
                {
                    ColumnHeaders = true,
                    StartRow = 0
                });
            }
        }

        private void readGoogleSheet(DataSet ds)
        {
            // Define request parameters.
            List<string> ranges = new List<string>();

            foreach (var monthString in monthStrings)
            {
                ranges.Add(monthString);
            }

            // choose input format between 0: FORMATED_VALUE, 1: UNFORMATED_VALUE and 2: FORMULA
            SpreadsheetsResource.ValuesResource.BatchGetRequest.ValueRenderOptionEnum valueRenderOption = (SpreadsheetsResource.ValuesResource.BatchGetRequest.ValueRenderOptionEnum)1;
            SpreadsheetsResource.ValuesResource.BatchGetRequest.DateTimeRenderOptionEnum dateTimeRenderOption = (SpreadsheetsResource.ValuesResource.BatchGetRequest.DateTimeRenderOptionEnum)0;
            SpreadsheetsResource.ValuesResource.BatchGetRequest request = Program.SpreadsheetService.Spreadsheets.Values.BatchGet(spreadsheetId);
            request.Ranges = ranges;
            request.ValueRenderOption = valueRenderOption;
            request.DateTimeRenderOption = dateTimeRenderOption;

            Data.BatchGetValuesResponse response = request.Execute();

            for (int month = 0; month < response.ValueRanges.Count; month++)
            {
                for (int i = 0; i < 6; i++)
                {
                        var columnSpec = new DataColumn
                        {
                            DataType = typeof(string),
                            ColumnName = "Database Name" + month + " " + i
                        };
                        yearSet.Tables[month].Columns.Add(columnSpec);
                }

                if (response.ValueRanges[month].Values != null && response.ValueRanges[month].Values.Count > 0)
                {
                    for (int row = 0; row < response.ValueRanges[month].Values.Count; row++)
                    {
                        DataRow newRow = yearSet.Tables[month].NewRow();

                        yearSet.Tables[month].Rows.Add(newRow);
                        for (int col = 0; col < response.ValueRanges[month].Values[row].Count; col++)
                        {
                            
                           yearSet.Tables[month].Rows[row][col] = response.ValueRanges[month].Values[row][col];
                        }

                    }
                }
            }
        }

        private void writeToGoogleSheet(DataTable dt) {

            List<IList<Object>> googleDataRows = new List<IList<Object>>();
            var googleDataRow = new List<object>();
            foreach(DataRow row in dt.Rows)
            {
                for(int i = 0; i < 6; i++)
                {
                    googleDataRow.Add(row[i]);
                }

                googleDataRows.Add(googleDataRow);
                googleDataRow = new List<object>();
            }

            // The new values to apply to the spreadsheet.
            List<Data.ValueRange> data = new List<Data.ValueRange>();  // TODO: Update placeholder value.

            Data.BatchUpdateValuesRequest requestBody = new Data.BatchUpdateValuesRequest();
            requestBody.ValueInputOption = "RAW";

            ValueRange vr = new ValueRange();
            vr.Values = googleDataRows;
            vr.Range = monthStrings[monthIndex];
            data.Add(vr);
    
            requestBody.Data = data;

            SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = Program.SpreadsheetService.Spreadsheets.Values.BatchUpdate(requestBody, spreadsheetId);

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Data.BatchUpdateValuesResponse response = request.Execute();

            Console.Read();
        }

        private void sumerizeMonth()
        {
            yearSet.Tables[12].Rows.Clear();

            int completedMonths = 0;
            int avgMonthTotal = 0;

            for (int i = 0; i < 20; i++)
            {
                DataRow newRow = yearSet.Tables[12].NewRow();
                yearSet.Tables[12].Rows.Add(newRow);
            }

            for (int i = 0; i < yearSet.Tables.Count-1; i++)
            {
                if (yearSet.Tables[i].Rows.Count > 0)
                {
                    DataRow newRow = yearSet.Tables[12].NewRow();

                    yearSet.Tables[12].Rows[i][0] = monthStrings[i];
                    yearSet.Tables[12].Rows[i][1] = yearSet.Tables[i].Rows[8][5];
                    avgMonthTotal += Convert.ToInt32(yearSet.Tables[12].Rows[i][1]);

                    // adding other avrage for category other
                    categoryList.items[0].annualSum += Convert.ToDouble(yearSet.Tables[i].Rows[0][5]);
                    categoryList.items[1].annualSum += Convert.ToDouble(yearSet.Tables[i].Rows[1][5]);
                    categoryList.items[2].annualSum += Convert.ToDouble(yearSet.Tables[i].Rows[2][5]);
                    categoryList.items[3].annualSum += Convert.ToDouble(yearSet.Tables[i].Rows[3][5]);
                    categoryList.items[4].annualSum += Convert.ToDouble(yearSet.Tables[i].Rows[4][5]);
                    categoryList.items[5].annualSum += Convert.ToDouble(yearSet.Tables[i].Rows[5][5]);
                    categoryList.items[6].annualSum += Convert.ToDouble(yearSet.Tables[i].Rows[6][5]);
                    categoryList.items[7].annualSum += Convert.ToDouble(yearSet.Tables[i].Rows[7][5]);

                    completedMonths++;
                }
            }

            for (int i = 0; i < categoryList.items.Count; i++)
            {
                Console.WriteLine("i: " + i);
                yearSet.Tables[12].Rows[i][3] = categoryList.items[i].title;
                yearSet.Tables[12].Rows[i][4] = (int)(categoryList.items[i].annualSum / completedMonths);
            }

            yearSet.Tables[12].Rows[9][3] = "Genomsnitt månad";
            yearSet.Tables[12].Rows[9][4] = avgMonthTotal / completedMonths;



        }
    }
}

public class Category
{
    public Item Hyra = new Item { title = "Hyra", sum = 0.0 };
    public Item Bredband = new Item { title = "Bredband/Modil", sum = 0 };
    public Item Manadsutgifter = new Item { title = "Månadsutgifter", sum = 0 };
    public Item Swish = new Item { title = "Swish", sum = 0 };
    public Item Mat = new Item { title = "Mat", sum = 0 };
    public Item Sprit = new Item { title = "Sprit", sum = 0 };
    public Item Resor = new Item { title = "Resor", sum = 0 };
    public Item Other = new Item { title = "Övrigt", sum = 0 };
    public Item Snus = new Item { title = "Snus", sum = 0 };


}

public class Item
{
    public string title { get; set; }
    public double sum { get; set; }
    public double annualSum { get; set; }
}

public class RootObject
{
    public List<Item> items { get; set; }
}