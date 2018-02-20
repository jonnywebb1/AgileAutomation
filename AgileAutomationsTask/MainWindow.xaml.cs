using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using System.Data.OleDb;
using System.Data;
using Path = System.IO.Path;
using mshtml;
using System.Reflection;
using System.Threading;

namespace AgileAutomationsTask
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //GLOBALS
        const string url = "https://agileautomations.co.uk/home/inputform";
        List<ContactDetails> contactDetails;
        WebBrowser wb;
        int Tries;
        bool loaded;
        int tries;
        //bool docLoaded;

        public MainWindow()
        {
            InitializeComponent();

            wb = new WebBrowser();

            wb.LoadCompleted += new LoadCompletedEventHandler(bws_LoadCompleted);

            wb.Navigate(new Uri(url));

            contactDetails = new List<ContactDetails>();
        }

        private void Button_SelectExcel_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialogue = new Microsoft.Win32.OpenFileDialog();
            dialogue.DefaultExt = ".xlsx";
            bool? result = dialogue.ShowDialog();

            if (result != true) return;

            string workbookPath = dialogue.FileName;

            var conn = OpenDbConnection(workbookPath);

            OleDbCommand cmd = new OleDbCommand(); ;
            OleDbDataAdapter dbAdapter;
            DataSet contactInfo = new DataSet();

            cmd.Connection = conn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM [ContactInfo$]"; //sheet name

            dbAdapter = new OleDbDataAdapter(cmd);
            dbAdapter.Fill(contactInfo, "Employee");

            foreach (var c in contactInfo.Tables[0].AsEnumerable())
            {
                tries = 0;
                loaded = false;

                PostContactDetails(new ContactDetails 
                { 
                    ContactName = Convert.ToString(c["Name"] != DBNull.Value ? c["Name"] : ""),
                    ContactEmail = Convert.ToString(c["Email"] != DBNull.Value ? c["Email"] : ""),
                    ContactSubject = Convert.ToString(c["Subject"] != DBNull.Value ? c["Subject"] : ""),
                    Message = Convert.ToString(c["Message"] != DBNull.Value ? c["Message"] : "")
                });

                HTMLDocument doc = wb.Document as HTMLDocument;

                //TODO: Get Form Submission response & get the reference number & append to row in Excel

                c["Reference"] = doc.getElementById("ReferenceNo").innerText ?? "";
                Console.WriteLine("Ref:" + c["Reference"]);
            }
        }

        /// <summary>
        /// UNUSED - can be used if there is an event for the reloading of the form after submission.
        /// </summary>
        /// <returns></returns>
        private string TryGetReference() 
        { 
            Thread.Sleep(500); 
            if(!loaded && tries < 6)
            {
                Tries++;
                TryGetReference();
            }
            else
	        {
                HTMLDocument doc = wb.Document as HTMLDocument;
                return doc.getElementById("ReferenceNo").innerText ?? "";
	        }
            return "";
        }

        void bws_LoadCompleted(object sender, NavigationEventArgs e)
        {
            loaded = true;
        }

        /// <summary>
        /// Updates the value of the element 'name' in the num form 'formid'
        /// </summary>
        /// <param name="formId"></param>
        /// <param name="name"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        private bool UpdateTextInput(int formId, string name, string text)
        {
            bool successful = false;
            IHTMLFormElement form = GetForm(formId);
            if (form != null)
            {
                var element = form.item(name: name);
                if (element != null)
                {
                    var textinput = element as HTMLInputElement;
                    textinput.value = text;
                    successful = true;
                }
            }

            return successful;
        }

        /// <summary>
        /// Posts contact details form with contact details
        /// </summary>
        /// <param name="cd"></param>
        private void PostContactDetails(ContactDetails cd)
        {
            var form = GetForm(0);
            UpdateTextInput(0, "ContactName", cd.ContactName);
            UpdateTextInput(0, "ContactEmail", cd.ContactEmail);
            UpdateTextInput(0, "ContactSubject", cd.ContactSubject);
            UpdateTextInput(0, "ContactName", cd.Message);
            
            form.submit();

            //TODO:: Find event or method for form submission page reload completion
            //Thread.Sleep(1000);  //wait 1 sec for page to reload 
        }
        
        /// <summary>
        /// Gets form number from web page
        /// </summary>
        /// <param name="formNo"></param>
        /// <returns></returns>
        private IHTMLFormElement GetForm(int formNo)
        {
            IHTMLDocument2 doc = (IHTMLDocument2)wb.Document;
            IHTMLElementCollection forms = doc.forms;
            var ix = 0;
            foreach (IHTMLFormElement f in forms)
                if (ix++ == formNo)
                    return f;

            return null;
        }
        
        /// <summary>
        /// connects to workbook specified
        /// </summary>
        /// <param name="workbookPath"></param>
        /// <returns></returns>
        private OleDbConnection OpenDbConnection(string workbookPath)
        {
            OleDbConnection conn = null;

            try
            {
                if (Path.GetExtension(workbookPath) == ".xls")
                    conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + workbookPath +
                        "; Extended Properties= \"Excel 8.0;HDR=Yes;IMEX=2\"");
                else if (Path.GetExtension(workbookPath) == ".xlsx")
                    conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" +
                        workbookPath + "; Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';");

                conn.Open();
            }
            catch (Exception ex)
            {
                //TODO: HANDLE THE EXCEPTION
                throw;
            }
            return conn;
        }

        /// <summary>
        /// UNUSED - First approach => get IEnumberable of objects to submit.
        /// </summary>
        /// <param name="dbConn"></param>
        /// <returns></returns>
        private List<ContactDetails> GetContactDetailsFromExcel(OleDbConnection dbConn)
        {
            OleDbCommand cmd = new OleDbCommand(); ;
            OleDbDataAdapter dbAdapter;
            DataSet contactInfo = new DataSet();

            cmd.Connection = dbConn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM [ContactInfo$]"; //sheet name

            dbAdapter = new OleDbDataAdapter(cmd);
            dbAdapter.Fill(contactInfo, "Employee");

            var contactList = contactInfo.Tables[0].AsEnumerable().Select(c => new ContactDetails
            {
                ContactName = Convert.ToString(c["Name"] != DBNull.Value ? c["Name"] : ""),
                ContactEmail = Convert.ToString(c["Email"] != DBNull.Value ? c["Email"] : ""),
                ContactSubject = Convert.ToString(c["Subject"] != DBNull.Value ? c["Subject"] : ""),
                Message = Convert.ToString(c["Message"] != DBNull.Value ? c["Message"] : "")
            }).ToList();

            return contactList;
        }

        /// <summary>
        /// UNUSED - First approach - loop through IENumerable & Submit
        /// Issue was getting the reference number & appending back to excel.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SubmitContactDetails_Click(object sender, RoutedEventArgs e)
        {
            var form = GetForm(0);
            foreach (var item in contactDetails)
            {
                PostContactDetails(item);
            }
        }
    }
}
