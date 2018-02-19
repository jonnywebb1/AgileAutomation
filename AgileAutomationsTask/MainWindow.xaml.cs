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
using Microsoft.Office.Interop;
using System.Data.OleDb;
using System.IO;
using System.Data;
using Path = System.IO.Path;
using mshtml;
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

        public MainWindow()
        {
            InitializeComponent();

            wb = new WebBrowser();
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
            contactDetails = GetContactDetailsFromExcel(conn);

            btn_SubmitContactDetails.Visibility = System.Windows.Visibility.Visible;
        }

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
                //TODO HANDLE THE EXCEPTION
                throw;
            }
            return conn;
        }

        private List<ContactDetails> GetContactDetailsFromExcel(OleDbConnection dbConn) 
        { 
            OleDbCommand cmd = new OleDbCommand (); ;
            OleDbDataAdapter dbAdapter;
            DataSet contactInfo = new DataSet ();

            cmd.Connection  = dbConn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM [ContactInfo$]"; //sheet name
            
            dbAdapter = new OleDbDataAdapter(cmd);
            dbAdapter.Fill(contactInfo, "Employee");

            var contactList = contactInfo.Tables[0].AsEnumerable().Select(c => new ContactDetails
            {
                Name = Convert.ToString(c["Name"] != DBNull.Value ? c["Name"] : ""),
                Email = Convert.ToString(c["Email"] != DBNull.Value ? c["Email"] : ""),
                Subject = Convert.ToString(c["Subject"] != DBNull.Value ? c["Subject"] : ""),
                Message = Convert.ToString(c["Message"] != DBNull.Value ? c["Message"] : "")
            }).ToList();

            return contactList;
        }

        private string SubmitContactDetails(HTMLDocument htmlDoc, ContactDetails contactDetails) 
        {
            HTMLFormElement form = htmlDoc.getElementById("sky-form3") as HTMLFormElement;

            var contactName = htmlDoc.getElementById("ContactName") as IHTMLInputTextElement;
            contactName.value = contactDetails.Name;

            var email = htmlDoc.getElementById("ContactEmail") as IHTMLInputElement;
            email.value = contactDetails.Email;
            
            var subject = htmlDoc.getElementById("ContactSubject") as IHTMLInputElement;
            subject.value = contactDetails.Subject;

            var message = htmlDoc.getElementById("Message") as IHTMLTextAreaElement;
            message.value = contactDetails.Message;
            form.submit();

            Thread.Sleep(500);

            var label = htmlDoc.getElementById("ReferenceNo") as IHTMLLabelElement;

            return label.htmlFor;
        }

        private void btn_SubmitContactDetails_Click(object sender, RoutedEventArgs e)
        {
            mshtml.HTMLDocument htmldoc = wb.Document as HTMLDocument;

            foreach (var item in contactDetails)
            {
                item.Reference = SubmitContactDetails(htmldoc, item);
            }
        }
    }
}
