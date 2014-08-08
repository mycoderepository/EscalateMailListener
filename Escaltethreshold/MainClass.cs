#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Diagnostics;
using Microsoft.Exchange.WebServices.Data;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.Globalization;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.IO;
using System.Configuration;
using Oracle.DataAccess.Client;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Configuration;
using Twilio;
using MySql.Data.MySqlClient;
using com.IBL.Utility;
using System.Data.SqlClient;

#endregion


namespace Escaltethreshold
{


    public struct mDataReader
    {
        public MySqlDataReader mysqlrdr;
        public OracleDataReader orardr;
        public SqlDataReader sqlrdr;


    }


    class MainClass
    {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
        string omysqlserver = ConfigurationManager.AppSettings["mysqlserver"];
        string oserveruname = ConfigurationManager.AppSettings["serveruname"];
        string oserverpwd = ConfigurationManager.AppSettings["serverpwd"];
        string oserverport = ConfigurationManager.AppSettings["serverport"];
        
        #region find outlook items
        public void FindItems()
        {
            ItemView view = new ItemView(10);
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
            view.PropertySet = new PropertySet(
                BasePropertySet.IdOnly,
                ItemSchema.Subject,
                ItemSchema.DateTimeReceived);

            FindItemsResults<Item> findResults = service.FindItems(
                WellKnownFolderName.Inbox,
                new SearchFilter.SearchFilterCollection(
                    LogicalOperator.Or,
                    new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Threshold"),
                    new SearchFilter.ContainsSubstring(ItemSchema.Body, "Nigeria")),
                view);


            //return findResults
            //Console.WriteLine("Total number of items found: " + findResults.TotalCount.ToString());

            foreach (Item item in findResults)
            {
                // Do something with the item.
            }
        }
        #endregion

        #region Email reply method
        public void ReplyToMessage(EmailMessage messageToReplyTo, string reply, string cc)
        {
            messageToReplyTo.Reply(reply, true /* replyAll */);
            // Or
            ResponseMessage responseMessage = messageToReplyTo.CreateReply(true);
            responseMessage.BodyPrefix = reply;
            responseMessage.CcRecipients.Add(cc);
            responseMessage.SendAndSaveCopy();
        }
        #endregion

        #region Email forwarder
        public void ForwardMessage(EmailMessage messageToForward, string forward, string ccrec)
        {
            messageToForward.Forward(forward);
            // Or
            ResponseMessage responseMessage = messageToForward.CreateForward();
            responseMessage.BodyPrefix = forward;
            responseMessage.CcRecipients.Add(ccrec);
            responseMessage.SendAndSaveCopy();
        }
        #endregion

        #region SMS forwarder
        public void sendtextmessage(string xTo, string xmsg)
        {
            try
            {
                if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                {

                    string AccountSid = Devsecurity.StringDecrypt(ConfigurationManager.AppSettings["Authid"]);
                    string AuthToken = Devsecurity.StringDecrypt(ConfigurationManager.AppSettings["secret"]);

                    var twilio = new TwilioRestClient(AccountSid, AuthToken);

                    var message = twilio.SendMessage("+17314724935", xTo, xmsg);
                    Trace.WriteLine(">>>>>>>>>>>>>>>>>> Message sent successfully to - " + xTo + "");
                }
                else
                {
                    MessageBox.Show("There is a Network Issue", "", MessageBoxButtons.OK);

                }
            }
            catch (Exception ex)
            {
                
                 Trace.WriteLine(">>>>>>>>>>>>>>>>>> Message not sent to - "+xTo+" with error  "+ex+"");
            }

        }
        #endregion

        #region creating Appointment.
        public int createAppointment(string xsubject, string xbody, DateTime xsentdate, string iemail)
        {


            try
            {


                Outlook.Application outlookApp = new Outlook.Application(); // creates new outlook app
                Outlook.AppointmentItem oAppointment = (Outlook.AppointmentItem)outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem); // creates a new appointment

                oAppointment.Subject = xsubject; // set the subject
                oAppointment.Body = xbody; // set the body
                oAppointment.Location = " Office"; // set the location
                oAppointment.Start = xsentdate; // Set the start date 
                oAppointment.End = xsentdate.AddHours(3); // End date 
                oAppointment.ReminderSet = true; // Set the reminder
                oAppointment.ReminderMinutesBeforeStart = 15; // reminder time
                oAppointment.Importance = Outlook.OlImportance.olImportanceHigh; // appointment importance
                oAppointment.BusyStatus = Outlook.OlBusyStatus.olBusy;
                oAppointment.Save();

                Outlook.MailItem mailItem = oAppointment.ForwardAsVcal();

                // email address to send to 
                mailItem.To = iemail; // ConfigurationManager.AppSettings["xemail"]; 
                mailItem.Send();



                return 1;

            }
            catch (Exception e)
            {
                Trace.WriteLine(e.ToString());
                return 0;

            }



        }
        #endregion

        #region creates outlook app instance.
        public Outlook.Application runapplicationoutlook()
        {



            Outlook.Application application = null;

            // Check whether there is an Outlook process running.
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {

                try
                {
                    // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                    application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                }
                catch (COMException ce)
                {
                    Type type = Type.GetTypeFromProgID("Outlook.Application");
                    application = (Outlook.Application)System.Activator.CreateInstance(type);
                    throw ;
                }



            }
            else
            {

                // If not, create a new instance of Outlook and log on to the default profile.
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("", "", Type.Missing, Type.Missing);
                nameSpace = null;
            }

            // Return the Outlook Application object.
            return application;





        }
        #endregion

        #region insert update delete class
        public int insupddelClass(string osql)
        {
            try
            {
                var xconn = Properties.Settings.Default.ConnectionString;    //ConfigurationSettings.AppSettings["conOracle"];
                OracleConnection conn = new OracleConnection((xconn));
               


                string isql = osql;

                OracleCommand cmd = new OracleCommand(isql, conn);
                conn.Open();
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                Trace.WriteLine("Information Saved in Database >>>>>>");
                return 1;


                conn.Close();
                conn.Dispose();


            }
            catch (Exception ex)
            {
                //string elog = Convert.ToString(ex);
                //this.writelog(elog);
                Trace.WriteLine("Error Message", ex.ToString() + "\n");
                return 0;

            }
        }
        #endregion

        #region NewMail event handler.
        private static void outLookApp_NewMailEx(string EntryIDCollection)
        {
            MessageBox.Show("You've got a new mail whose EntryIDCollection is \n" + EntryIDCollection,
                    "NOTE", MessageBoxButtons.OK);
        }
        #endregion

        #region query database
        public mDataReader Populate1(string xsql, int i)
        {
            mDataReader mDataReader = new mDataReader();
            if (i == 1)   // Sql connection
            {
                SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["conSQL1"].ConnectionString);
                SqlCommand sqlCommand = new SqlCommand();
                string str = xsql;
                sqlCommand.Connection = sqlConnection;
                sqlCommand.CommandType = CommandType.Text;
                sqlCommand.CommandText = str;
                sqlConnection.Open();
                mDataReader.sqlrdr = sqlCommand.ExecuteReader();
                mDataReader.mysqlrdr = (MySqlDataReader)null;

                return mDataReader;
            }
            else if (i == 2)  // Oracle Connecetions
            {
                //OracleConnection oracleConnection = new OracleConnection(Devsecurity.StringDecrypt(ConfigurationManager.ConnectionStrings["conOracle"].ConnectionString));
                //OracleCommand oracleCommand = new OracleCommand();
                //try
                //{
                //    string str = xsql;
                //    oracleCommand.Connection = oracleConnection;
                //    oracleCommand.CommandType = CommandType.Text;
                //    oracleCommand.CommandText = str;
                //    oracleConnection.Open();
                //    mDataReader.orclrdr = oracleCommand.ExecuteReader();
                //    mDataReader.sqlrdr = (SqlDataReader)null;
                //}
                //catch (Exception ex)
                //{
                //    this.writelog(Convert.ToString((object)ex));
                //}
                return mDataReader;
            }
            else if (i == 3)   // MySql Connection
            {

                try
                {

                    string oenusername = (oserveruname);
                    string oenpassword = (oserverpwd);
                    string xconn = "Server=" + omysqlserver + ";Port=" + oserverport + ";Database=isng;Uid=" + oenusername + ";Pwd=" + oenpassword + "";


                    MySqlConnection conn = new MySqlConnection((xconn));
                    MySqlCommand cmd = new MySqlCommand();

                    string osql = xsql;


                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = osql;
                    conn.Open();


                    mDataReader.mysqlrdr = cmd.ExecuteReader();
                    mDataReader.sqlrdr = null;



                    return mDataReader;
                    conn.Close();
                    conn.Dispose();
                }
                catch (Exception ex)
                {

                    Trace.WriteLine(">>>>>>>> Error : "+((object)ex));
                }

            }
            return mDataReader;
        }
        #endregion

        #region new insert class
        public int iClass(string osql, int i)
        {


            if (i == 1)
            {

                try
                {

                    string xconn = ConfigurationManager.ConnectionStrings["conSQL1"].ConnectionString;
                    SqlConnection conn = new SqlConnection((xconn));


                    string isql = osql;
                    SqlCommand cmd = new SqlCommand(isql, conn);
                    conn.Open();
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                    return 1;

                    conn.Close();
                    conn.Dispose();


                }
                catch (Exception ex)
                {
                    string elog = Convert.ToString(ex);
                    Trace.WriteLine(elog);
                    return 0;

                }

            }
            else if (i == 2)
            {
             
                try
                {
                    string oenusername = (oserveruname);
                    string oenpassword =(oserverpwd);
                    string xconn = "Server=" + omysqlserver + ";Port=" + oserverport + ";Database=isng;Uid=" + oenusername + ";Pwd=" + oenpassword + "";

                  
                    MySqlConnection conn = new MySqlConnection((xconn));


                    string isql = osql;
            

                    MySqlCommand cmd = new MySqlCommand(isql, conn);
                    conn.Open();
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                    return 1;


                    conn.Close();
                    conn.Dispose();

                }
                catch (Exception ex)
                {
                    string elog = Convert.ToString(ex);
                    Trace.WriteLine(elog);
                    return 0;

                }

            }
            return 0;
          

        }

            #endregion



    }
}
