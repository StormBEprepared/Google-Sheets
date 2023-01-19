using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Drawing.Drawing2D;
using System.Data.SqlClient;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Dynamic;
using System.Timers;
using Data = Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;

namespace GgoogleSheetsApi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {// Credential json file downloaded from google cloud-> IAM admin -> Service accounts-> Keys must be moved to app directory, such as
         // ..\GgoogleSheetsApi\GgoogleSheetsApi\bin\Debug\net6.0-windows
            GoogleCredential credential;
            using (var stream = new FileStream("ordinal-door-blablabla.json", FileMode.Open, FileAccess.ReadWrite)) 
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }
            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "WHATEVER";
        static readonly string SpreadsheetID = "YOURSHEETIDHERE";//Taken from Google link , should be something like: 1vIJyz346365BPSW_jzKtQRasgaV3YAF54OwPphOMwhv1Q
        static readonly string NameOfSheet = "YOURSHEETNAMEHERE";//Name of the actual google sheet 

        static SheetsService service;
        public void ReadEntries(string sheet, string Cells, DataGridView dataGridView, int NoOfColumns4Table)
        {
            var range = $"{sheet}!{Cells}";//A2:F
            var request = service.Spreadsheets.Values.Get(SpreadsheetID, range);
            var response = request.Execute();
            var values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    DataGridViewRow row1 = (DataGridViewRow)dataGridView.Rows[0].Clone();
                    for (int i = 0; i < NoOfColumns4Table; i++)
                    {
                        row1.Cells[i].Value = row[i];
                    }

                    dataGridView.Rows.Add(row1);
                }
            }
            else
            {
                //Console.WriteLine("nO DATA");
            }
        }
        public static void CreateEntry(string sheet, string Cells)
        {
            var range = $"{sheet}!{Cells}";//A:F for next row
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "Hello!", "This", "was", "inserted", "via", "c#." };//i think the cells range must correspond with the number of objects to be inserted
            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }
        public static void UpdateEntry(string sheet, string Cells)
        {
            var range = $"{sheet}!{Cells}"; // c15 for specific cell, update row by making a cells range ex: a2:f
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "updated" };// this is for single cell update, add more if you change more cells.
            valueRange.Values = new List<IList<object>> { objectList };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }




        public static void DeleteEntry(string sheet, string Cells)
        {
            var range = $"{sheet}!{Cells}";//a15:f
            var requestBody = new ClearValuesRequest();
            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetID, range);
            var deleteResponse = deleteRequest.Execute();
        }




        public static void UpdateEntryWmessage(string sheet, string Cells, string message)//UPDATE MEANS OVERWRITES DATA THAT ALREADY EXISTS IN RANGE
        {
            var range = $"{sheet}!{Cells}"; // c15 for specific cell, update row by making a cells range ex: a2:f
            var valueRange = new ValueRange();

            var objectList = new List<object>();
            objectList.AddRange(message.Split(' ').ToList());
            valueRange.Values = new List<IList<object>> { objectList };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }
        public static void CreateEntryWmessage(string sheet, string Cells, string message)//CREATE MEANS IT DOES NOT OVERWRITE DATA FROM RANGE BUT GOES TO NEX AVAILABLE ROW.
        {
            var range = $"{sheet}!{Cells}";//A:F for next row
            var valueRange = new ValueRange();

            var objectList = new List<object>();//I think the cells range must correspond with the number of objects to be inserted - > No, you can give range A1:A2 and a long string. It will write all string given and extend the row over the column limit given in range.
            objectList.AddRange(message.Split(' ').ToList());
            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);//.Append is for adding another new row in the next available (empty on all columns) row.
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }



        
       


        private void button1_Click(object sender, EventArgs e)
        {
            //Reads range and outputs in datagridview. The last argument given is the number of columns for table, in this case 4.
            //If the RANGE given exceeds the number of columns, like in our case it does as we have columns A,B,C,D,E = 5, then it reads the
            //table up until column 4.
            //If the RANGE given is smaller than the number of columns, for ex ReadEntries(NameOfSheet,"A1:E6",dataGridView1,8); then an error will be printed.
            //Please be accurate when choosing range and no of columns until further update.
            ReadEntries(NameOfSheet,"A1:E6",dataGridView1,4);


        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Function with standard string inside. Made for testing purposes.
            //CreateEntry(NameOfSheet, "A13:H");

            //Adds value 'update' written and changeable in main function in a certain cell. Works only with single cells. Made for testing purposes.
            //UpdateEntry(NameOfSheet, "F:");




            //Deletes value in single cell EX:("A3") or in row EX:("A3:F") or 'square' range such as (C8:G11).

            //DeleteEntry(NameOfSheet, "C8:G11");





            //Adds string values splitted by space. if more words than columns it can still add data but reading it might be an issue for null cells
            //Does not add value in range, just enlarges the table by adding extra columns if the string is that big that the
            //number of words is bigger than the existing no of columns before input. This results in endless row, not in data added in a 'sqare' range.
            //If range is already occupied by data then it will not overwrite but go to next row.

            //CreateEntryWmessage(NameOfSheet,"A5:F13","data data data data");






            //Adds string values splitted by space. if more words than columns it can still add data but reading it might be an issue for null cells
            //Does not add value in range, just enlarges the table by adding extra columns if the string is that big that the
            //number of words is bigger than the existing no of columns before input. This results in endless row, not in data added in a 'sqare' range.
            //If range is already occupied by data then IT WILL OVERWRITE.

            //UpdateEntryWmessage(NameOfSheet, "B5", "Id know,");



            //The functions to create and update are splitting the given string in words and each word will be inserted into a cell, on the same row, in order.
            //Both functions can be used to input/update single cells, however, the way that they are built atm are unsafe in case of space misplaced in string as it can mess up the whole table.
            //An idea to avoid it is to create a string trimmer before sending the request to the API
            //OR
            //create a separate function that does not add a list of string elements but A SINGLE element.

        }
















        //MOVING THE EDGELESS WINDOWS FORMS BY CLICK AND DRAG   
        bool mouseDown;
        private Point offset;
        private void MouseDown_Event(object sender, MouseEventArgs e)
        {
            offset.X = e.X;
            offset.Y = e.Y;
            mouseDown = true;
        }
        private void MouseMove_Event(object sender, MouseEventArgs e)
        {
            if (mouseDown == true)
            {
                Point CurrentScreenPos = PointToScreen(e.Location);
                Location = new Point(CurrentScreenPos.X - offset.X, CurrentScreenPos.Y - offset.Y);
            }
        }
        private void MouseUp_Event(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }
        //MOVING THE EDGELESS WINDOWS FORMS BY CLICK AND DRAG  


    }
}