using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections.Specialized;
using System.Linq.Expressions;
using Twilio.Rest.Api.V2010.Account;
using Twilio;

namespace ReminderAppV1
{
    class Program
    {
        static DataTable dataTableObj = new DataTable("ExcelDataTable"), SelectedValues = new DataTable("SelectedValues");
        static void Main(string[] args)
        {
            GetDetails();
        }
        static DataTable ReadSheet()
        {
            string file_path = Properties.Settings.Default.filepath;
            String excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file_path + ";Extended Properties=Excel 12.0;Persist Security Info=False";
            OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
            OleDbDataAdapter adp = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", excelConnectionString);
            adp.Fill(dataTableObj);
            return dataTableObj;
        }

        public static void GetDetails()
        {
            ReadSheet();
            DateTime now = DateTime.Now.AddDays(+30);
            string asString = now.ToShortDateString(), Expression = "ExpiryDate ='" + asString + "'"; 
            DataRow[] Result = dataTableObj.Select(Expression);
            SelectedValues.Columns.Add("Customer"); SelectedValues.Columns.Add("Details"); SelectedValues.Columns.Add("ExpiryDate"); SelectedValues.Columns.Add("Property");
            foreach (var item in Result)             {
                SelectedValues.Rows.Add(item.ItemArray);
            }
            SendReminder();
        }
        static void SendReminder()
        {
             string accountSid = Properties.Settings.Default.accountSid, authToken = Properties.Settings.Default.authToken;
            TwilioClient.Init(accountSid, authToken);
            foreach (DataRow row in SelectedValues.Rows)
            {
                var date = row["ExpiryDate"].ToString();
                date = date.Remove(date.Length - 8);
                var message = MessageResource.Create(
                    body: "Hi " + row["Customer"].ToString() + ", Please note that: " + row["Property"].ToString() + " is due for a Gas Safety certificate which is set to expire on "+ date,
                    from: new Twilio.Types.PhoneNumber("+12016543514"),to: new Twilio.Types.PhoneNumber("+" +row["Details"].ToString())
                  );                
            }                  
        }
    }
}
