using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Net.NetworkInformation;
using System.Configuration;
using System.Text.RegularExpressions;



namespace VPN_App
{
    public partial class Form1 : Form
    {
        //collect settings from app.exe.config

        static NameValueCollection configsettings = ConfigurationManager.AppSettings;

        string passwordaddendum = configsettings["passwordaddendum"];
        string userdomain = configsettings["userdomain"];
        string titlename = configsettings["titlename"];
        string companyname = configsettings["companyname"];
        string companylogo = configsettings["companylogo"];
        string accountinfo = configsettings["accountinfo"];
        string custommessage1 = configsettings["custommessage1"];

        private BackgroundWorker backgroundWorker;

        public Form1()
        {
            InitializeComponent();
            //initialize background worker and add handlers
            InitializeBackgroundWorker();
            //set form to use the submit button when the enter key is hit
            this.AcceptButton = button1;
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            //apply label config settings if they exist
            if (!String.IsNullOrEmpty(companyname))
            {
                label1.Text = "Please Enter Your " + companyname + " Username:";
                label2.Text = "Please Enter Your " + companyname + " Password:";
            }
            else
            {
                //default settings for label
                label1.Text = "Please Enter Your Username:";
                label2.Text = "Please Enter Your Password:";
            }
            //apply window title settings from app.exe.config
            if (!String.IsNullOrEmpty(titlename))
            {
                this.Text = titlename;
            }
            else
            {
                //default title settings
                this.Text = "VPN Connection";
            }
            //apply custom icon to window from app.exe.config
            if (!String.IsNullOrEmpty(companylogo))
            {
                try
                {
                    Icon ico = new Icon(companylogo);
                    this.Icon = ico;
                }
                catch
                {
                    MessageBox.Show("Unable to find Icon specified in configuration file");
                }
                
            }
            

            //get path to rasphone.pbk
            String f = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Network\Connections\Pbk\rasphone.pbk";
            //check if path is a valid path
            if (File.Exists(f))
            {
                //creat collection array of strings
                List<string> lines = new List<string>();
                //read each line of pbk file and add to collection
                using (StreamReader r = new StreamReader(f))
                {
                    string line;
                    while ((line = r.ReadLine()) != null)
                    {
                        
                        lines.Add(line);
                    }
                }
                //find the name of the connection denoted inside of brackets and add as items to listbox
                foreach (string s in lines)
                {
                    if (s.StartsWith("["))
                    {
                        char[] MyChar = { ']' };
                        string NewString = s.TrimEnd(MyChar);
                        char[] MyChar2 = { '[' };
                        string NewString2 = NewString.TrimStart(MyChar2);
                        listBox1.Items.Add(NewString2);
                    }
                    
                }

                string text = File.ReadAllText(f);
                string newText = AlterRasFile(text);
                File.WriteAllText(f, newText);
            }
            else
            {
                //return popup if unable to find rasphone.pbk
                MessageBox.Show("VPN is not installed. \n Please Contact your System Administrator");
                return;
            }
            //sort listbox
            listBox1.Sorted = true;
            try
            {
                //select the first item by default
                listBox1.SelectedIndex = 0;
            } catch
            {
                Console.WriteLine("failed to index combobox");
            }
        }

        public void ChangeRasCredentials(string credentialLine)
        {
            credentialLine = "UseRasCredentials=0";
        }


        //this region handles the asynchronous aspect of the program
        #region
        private void InitializeBackgroundWorker()
        {
            this.backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(backgroundWorker_DoWork);
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backGroundWorker_RunWorkerCompleted);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(backGroundWorker_ProgressChanged);
        }
        //connect the the VPN through the background worker
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            e.Result = RunVPN((ConnectionObj)e.Argument);
        }
        //alert user what the program did based upon exitcode
        private void backGroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            if(e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }

            int exitcode = (int)e.Result;
            switch (exitcode)
            {
                case 0:
                    listBoxOut.Items.Clear();
                    MessageBox.Show("You are now connected to the " + listBox1.Text + " VPN System");
                    Application.Exit();
                    break;
                case 602:
                    MessageBox.Show("VPN is already connected");
                    Application.Exit();
                    break;
                case 691:
                    MessageBox.Show("Your Username or Password is Incorrect. Please try again.");
                    listBoxOut.Items.Clear();
                    break;
                default:
                    MessageBox.Show("System threw error " + exitcode.ToString() + "\n Please contact your system administrator");
                    listBoxOut.Items.Clear();
                    break;
            }
        }

        private void backGroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }
        #endregion




        public static int RunVPN(ConnectionObj connectionObj)
        {
            //create process object information
            //Console.WriteLine(connectionstring); //For Debugging
            ProcessStartInfo rasdialinfo = new ProcessStartInfo("\"rasdial.exe\"", connectionObj.ConnectionString);
            Process process = new Process();
            process.StartInfo = rasdialinfo;
            //keep command shell from popping up
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            try
            {
                //start process
                process.Start();
            }
            catch
            {
                MessageBox.Show("Failed to Start connection.\r\n Please contact your system administrator");
                return 0;
            }
            //add custom message to listbox if applicable
            
            //variables for performance measuring to be output into console window for debugging
            long peakPagedMem = 0,
            peakWorkingSet = 0,
            peakVirtualMem = 0;
            //display connecting status
            
            int i = 0;
            do
            {
                //loop for displaying program status in console for debugging while waiting for process to exit
                //runs every second
                if (!process.HasExited)
                {
                    
                    
                    
                    // Refresh the current process property values.
                    process.Refresh();


                    Console.WriteLine();

                    // Display current process statistics.

                    Console.WriteLine("{0} -", process.ToString());
                    Console.WriteLine("-------------------------------------");

                    Console.WriteLine("  physical memory usage: {0}",
                        process.WorkingSet64);
                    Console.WriteLine("  base priority: {0}",
                        process.BasePriority);
                    Console.WriteLine("  priority class: {0}",
                        process.PriorityClass);
                    Console.WriteLine("  user processor time: {0}",
                        process.UserProcessorTime);
                    Console.WriteLine("  privileged processor time: {0}",
                        process.PrivilegedProcessorTime);
                    Console.WriteLine("  total processor time: {0}",
                        process.TotalProcessorTime);
                    Console.WriteLine("  PagedSystemMemorySize64: {0}",
                        process.PagedSystemMemorySize64);
                    Console.WriteLine("  PagedMemorySize64: {0}",
                       process.PagedMemorySize64);

                    // Update the values for the overall peak memory statistics.
                    peakPagedMem = process.PeakPagedMemorySize64;
                    peakVirtualMem = process.PeakVirtualMemorySize64;
                    peakWorkingSet = process.PeakWorkingSet64;

                    //show whether or not process is running or not responding
                    if (process.Responding)
                    {
                        Console.WriteLine("Status = Running");
                    }
                    else
                    {
                        Console.WriteLine("Status = Not Responding");
                    }
                }
            }
            while (!process.WaitForExit(1000));

            //display exit code in console for debugging purposes
            Console.WriteLine();
            Console.WriteLine("Process exit code: {0}",
                process.ExitCode);

            // Display peak memory statistics for the process.
            Console.WriteLine("Peak physical memory usage of the process: {0}",
                peakWorkingSet);
            Console.WriteLine("Peak paged memory usage of the process: {0}",
                peakPagedMem);
            Console.WriteLine("Peak virtual memory usage of the process: {0}",
                peakVirtualMem);
            //handling for several exit codes
            //most likely exit codes are given display messages.
            //list of all exit codes: https://support.microsoft.com/en-us/help/824864/list-of-error-codes-for-dial-up-connections-or-vpn-connections

            return process.ExitCode;
        }

        //code that runs when the button is clicked
        private void button1_Click(object sender, EventArgs e)
        {

            //create ping object
            Ping PingSender = new Ping();
            PingOptions MyPingOptions = new PingOptions();
            //build ping packet
            MyPingOptions.DontFragment = true;
            string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
            byte[] buffer = Encoding.ASCII.GetBytes(data);
            int timeout = 120;
            //ping IP 8.8.8.8 to test internet connectivity
            int p = 0;
            int err = 0;
            PingReply reply;
            do
            {
                try
                {
                    reply = PingSender.Send("8.8.8.8", timeout, buffer, MyPingOptions);
                }
                catch
                {
                    Console.WriteLine("No Network Connection Available");
                    MessageBox.Show("No Network Connection Available");
                    break;
                }
                if (reply.Status == IPStatus.Success)
                {
                    Console.WriteLine("Ping success");
                    break;
                }
                else
                {
                    Console.WriteLine("Ping Failed");
                    err++;
                }

                p++;
            }
            while (p < 4);
            if (err == 4)
            {
                MessageBox.Show("You may not be connected to the internet. \r\n If connection fails, please double check your internet connection");
            }

            string username = textBoxUserName.Text;
            string password = textBoxPassword.Text;
            string connection = listBox1.Text;

            //validate that a VPN connection is selected
            if (String.IsNullOrEmpty(connection))
            {
                MessageBox.Show("Please Select a Connection");
                return;
            }
            //create local variables for user information
            string UserName = username;
            string PlainTextPw = password;
            //check for any possible password addendum from the config file and apply if necessary
            if (!String.IsNullOrEmpty(passwordaddendum))
            {
                int pushcount = PlainTextPw.Length - passwordaddendum.Length;
                try
                {
                    PlainTextPw.Substring(pushcount, passwordaddendum.Length);
                }
                catch
                {
                    MessageBox.Show("Please enter the phrase \"" + passwordaddendum + "\" after your password, or enter your 6 digit Okta Code");
                    return;
                }
                if (PlainTextPw.Substring(pushcount, passwordaddendum.Length) != passwordaddendum && !matchNumbers(PlainTextPw.Substring((PlainTextPw.Length - 6), 6)))
                {
                    MessageBox.Show("Please enter the phrase \"" + passwordaddendum + "\" after your password, or enter your 6 digit Okta Code");
                    return;
                }
            }
            //check for a user domain from config file and applies if necessary
            if (!String.IsNullOrEmpty(userdomain))
            {
                int usercount = UserName.Length - userdomain.Length;
                try
                {
                    UserName.Substring(usercount, userdomain.Length);
                }
                catch
                {
                    MessageBox.Show("Please use your " + accountinfo + userdomain + " account to log in");
                    return;
                }
                if (!UserName.Substring(usercount, userdomain.Length).Equals(userdomain, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("Please use your " + accountinfo + userdomain + " account to log in");
                    return;
                }
            }
            //build connection string to connect to VPN
            string connectionstring = "\"" + connection + "\"" + " " + username + " " + password;

            ConnectionObj connectionObj = new ConnectionObj();
            connectionObj.ConnectionString = connectionstring;
            connectionObj.CustomMessage = custommessage1;
            try
            {
                if (!String.IsNullOrEmpty(custommessage1))
                {
                    listBoxOut.Items.Add(custommessage1);
                }

                listBoxOut.Items.Add("Connecting...");

                backgroundWorker.RunWorkerAsync(connectionObj);

            }
            catch
            {
                MessageBox.Show("Failed to Connect to VPN. Please check your credentials");
                return;
            }

        }
        //method to check if string matches exactly 6 digits
        private static bool matchNumbers(string text)
        {
            Regex rx = new Regex("[0-9][0-9][0-9][0-9][0-9][0-9]");
            bool isMatch = rx.IsMatch(text);
            if (isMatch)
            {
                return true;
            }
            return false;
        }

        private static string AlterRasFile(string text)
        {
            string output = Regex.Replace(text, "UseRasCredentials=.*", "UseRasCredentials=0");
            return output;
        }
    }
    //class for the background worker to pass between the run and completed methods
    public class ConnectionObj
    {
        public string ConnectionString { get; set; }
        public string CustomMessage { get; set; }
    }
}

