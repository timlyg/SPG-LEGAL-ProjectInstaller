using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

namespace SPG_LEGAL_ProjectInstaller
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // Temporary section
            // Get all current running processes into Combo Box1
            comboBox1.Items.Clear();
            Process[] MyProcess = Process.GetProcesses();
            for (int i = 0; i < MyProcess.Length; i++)
                comboBox1.Items.Add(MyProcess[i].ProcessName);
            // comboBox1.Items.Add(MyProcess[i].ProcessName + "-" + MyProcess[i].Id);
            comboBox1.Sorted = true;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Ready";
        }

        // VMWare Installation
        private void button1_Click(object sender, EventArgs e)
        {
            string adminname = "SPGadmin";
            string adminpass = "";
            toolStripStatusLabel1.Text = "Analyzing Installed packages...";

            // Verify checked installation & if they had already been installed
            if (getVMWare.Checked == true || getCitrix.Checked == true || getRelativity.Checked == true)
            {
                Process myProcess = new Process();

                myProcess.StartInfo.UseShellExecute = false;
                myProcess.StartInfo.FileName = "wmic.exe";
                myProcess.StartInfo.Arguments = "product list brief";
                // myProcess.StartInfo.CreateNoWindow = true;
                myProcess.StartInfo.RedirectStandardOutput = true;
                myProcess.Start();

                string standard_output;
                while ((standard_output = myProcess.StandardOutput.ReadLine()) != null)
                {
                    if (standard_output.Contains("VMware Horizon Client"))
                    {
                        //MessageBox.Show("VMWare had already been installed");
                        getVMWare.Checked = false;
                    }
                    if (standard_output.Contains("Citrix Receiver Inside"))
                    {
                        //MessageBox.Show("Citrix had already been installed");
                        getCitrix.Checked = false;
                    }
                    if (standard_output.Contains("Relativity Web Client"))
                    {
                        //MessageBox.Show("Relativity had already been installed");
                        getRelativity.Checked = false;
                    }
                }

                myProcess.WaitForExit();

            }
            // Handle VMWare installation
            if (getVMWare.Checked == true)
            {
                string fileName = "VMware-Horizon-View-Client-x86_64-3.5.2-3150477.exe";
                string remoteUri = "http://nycphantom.com/upload/uploaded/" + fileName;
                // Create a new WebClient instance.
                WebClient myWebClient = new WebClient();
                toolStripStatusLabel1.Text = "Downloading File " + fileName + " from " + remoteUri + ".......";
                // Download the Web resource and save it into the current filesystem folder.
                myWebClient.DownloadFile(remoteUri, @"C:\Users\Public\Downloads\" + fileName);
                toolStripStatusLabel1.Text = "Successfully Downloaded File " + fileName + " from " + remoteUri;
                toolStripStatusLabel1.Text = "Downloaded file saved in the following file system folder: " + "C:\\Users\\Public\\Downloads\\";

                // Install VMWare in silent mode
                var psi = new ProcessStartInfo();

                psi.UserName = adminname;
                psi.Password = new NetworkCredential("", adminpass).SecurePassword;
                psi.UseShellExecute = false;
                psi.WorkingDirectory = @"C:\Users\Public\Downloads\";
                psi.FileName = @"C:\Users\Public\Downloads\" + fileName;
                psi.Arguments = "/s /v\" /qn REBOOT=Reallysuppress\"";
                Process.Start(psi);
            }


            // Handle Citrix installation
            if (getCitrix.Checked == true)
            {
                string fileName = "CitrixReceiver-4.5.exe";
                string remoteUri = "http://nycphantom.com/upload/uploaded/" + fileName;
                // Create a new WebClient instance.
                WebClient myWebClient = new WebClient();
                toolStripStatusLabel1.Text = "Downloading File " + fileName + " from " + remoteUri + ".......";
                // Download the Web resource and save it into the current filesystem folder.
                myWebClient.DownloadFile(remoteUri, @"C:\Users\Public\Downloads\" + fileName);
                toolStripStatusLabel1.Text = "Successfully Downloaded File " + fileName + " from " + remoteUri;
                toolStripStatusLabel1.Text = "Downloaded file saved in the following file system folder: " + "C:\\Users\\Public\\Downloads\\";

                // Install Citrix 4.5 in silent mode
                var psi = new ProcessStartInfo();
                // Citrix installation doesn't require admin elevation
                //psi.UserName = "administrator";
                //psi.Password = new NetworkCredential("", "password").SecurePassword;
                //psi.UseShellExecute = false;
                psi.WorkingDirectory = @"C:\Users\Public\Downloads\";
                psi.FileName = @"C:\Users\Public\Downloads\" + fileName;
                psi.Arguments = "/silent";
                Process.Start(psi);
            }

            // Handle Relativity installation
            if (getRelativity.Checked == true)
            {
                string fileName = "ViewerInstallationKit9.4.496.3.exe";
                string remoteUri = "http://nycphantom.com/upload/uploaded/" + fileName;
                // Create a new WebClient instance.
                WebClient myWebClient = new WebClient();
                toolStripStatusLabel1.Text = "Downloading File " + fileName + " from " + remoteUri + ".......";
                // Download the Web resource and save it into the current filesystem folder.
                myWebClient.DownloadFile(remoteUri, @"C:\Users\Public\Downloads\" + fileName);
                toolStripStatusLabel1.Text = "Successfully Downloaded File " + fileName + " from " + remoteUri;
                toolStripStatusLabel1.Text = "Downloaded file saved in the following file system folder: " + "C:\\Users\\Public\\Downloads\\";

                // Install Relativity in silent mode
                var psi = new ProcessStartInfo();
                //psi.UserName = adminname;
                //psi.Password = new NetworkCredential("", adminpass).SecurePassword;
                //psi.UseShellExecute = false;
                psi.UseShellExecute = true;
                psi.WorkingDirectory = @"C:\Windows\System32";
                //psi.WorkingDirectory = @"C:\Users\Public\Downloads\";
                psi.FileName = @"C:\Users\Public\Downloads\" + fileName;
                psi.Arguments = "/silent";
                psi.Verb = "runas";
                Process.Start(psi);
            }




            // Handle Citrix WebPlugin installation
            if (getCitrixWebPlugin.Checked == true)
            {
                string fileName = "CitrixOnlinePluginWeb.exe";
                string remoteUri = "http://nycphantom.com/upload/uploaded/" + fileName;
                // Create a new WebClient instance.
                WebClient myWebClient = new WebClient();
                toolStripStatusLabel1.Text = "Downloading File " + fileName + " from " + remoteUri + ".......";
                // Download the Web resource and save it into the current filesystem folder.
                myWebClient.DownloadFile(remoteUri, @"C:\Users\Public\Downloads\" + fileName);
                toolStripStatusLabel1.Text = "Successfully Downloaded File " + fileName + " from " + remoteUri;
                toolStripStatusLabel1.Text = "Downloaded file saved in the following file system folder: " + "C:\\Users\\Public\\Downloads\\";

                // Install Relativity in silent mode
                var psi = new ProcessStartInfo();
                //psi.UserName = adminname;
                //psi.Password = new NetworkCredential("", adminpass).SecurePassword;
                //psi.UseShellExecute = false;
                psi.UseShellExecute = true;
                psi.WorkingDirectory = @"C:\Windows\System32";
                //psi.WorkingDirectory = @"C:\Users\Public\Downloads\";
                psi.FileName = @"C:\Users\Public\Downloads\" + fileName;
                psi.Arguments = "/silent";
                psi.Verb = "runas";
                Process.Start(psi);
            }




            // Handle Clear_Outook
            if (getClearOutlook.Checked == true)
            {
                string fileName = "clear_outlook.bat";
                string remoteUri = "http://nycphantom.com/upload/uploaded/" + fileName;
                // Create a new WebClient instance.
                WebClient myWebClient = new WebClient();
                toolStripStatusLabel1.Text = "Downloading File " + fileName + " from " + remoteUri + ".......";
                // Download the Web resource and save it into the current filesystem folder.
                myWebClient.DownloadFile(remoteUri, @"C:\Users\Public\Downloads\" + fileName);
                toolStripStatusLabel1.Text = "Successfully Downloaded File " + fileName + " from " + remoteUri;
                toolStripStatusLabel1.Text = "Downloaded file saved in the following file system folder: " + "C:\\Users\\Public\\Downloads\\";

                // Run Clear_Outlook batch file
                var psi = new ProcessStartInfo();
                // Citrix installation doesn't require admin elevation
                //psi.UserName = "administrator";
                //psi.Password = new NetworkCredential("", "password").SecurePassword;
                //psi.UseShellExecute = false;
                psi.WorkingDirectory = @"C:\Users\Public\Downloads\";
                psi.FileName = @"C:\Users\Public\Downloads\" + fileName;
                Process.Start(psi);
            }




            // Handle TeamViewer with Admin rights installation
            if (getTeamViewer.Checked == true)
            {
                string fileName = "TeamViewer_Setup.exe";
                string remoteUri = "http://nycphantom.com/upload/uploaded/" + fileName;
                // Create a new WebClient instance.
                WebClient myWebClient = new WebClient();
                toolStripStatusLabel1.Text = "Downloading File " + fileName + " from " + remoteUri + ".......";
                // Download the Web resource and save it into the current filesystem folder.
                myWebClient.DownloadFile(remoteUri, @"C:\Users\Public\Downloads\" + fileName);
                toolStripStatusLabel1.Text = "Downloaded file saved in the following file system folder: " + "C:\\Users\\Public\\Downloads\\";

                // Install VMWare in silent mode
                var psi = new ProcessStartInfo();

                psi.UserName = adminname;
                psi.Password = new NetworkCredential("", adminpass).SecurePassword;
                psi.UseShellExecute = false;
                psi.WorkingDirectory = @"C:\Users\Public\Downloads\";
                psi.FileName = @"C:\Users\Public\Downloads\" + fileName;
                Process.Start(psi);
                // "directory name invalid" error is referring to the SPGadmin user not able to access LAUser folder, this sort of thing. If this is an issue, work with file in Users/Public then.
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process[] processlist = Process.GetProcesses();

            foreach (Process process in processlist)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {
                    Console.WriteLine("Process: {0} ID: {1} Window title: {2}", process.ProcessName, process.Id, process.MainWindowTitle);
                }
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            string[] desktopFiles =     // Desktop Folder files to remain
            {
                "desktop.ini",
                "VMware Horizon Client.lnk",
                "Google Chrome.lnk",
                "GoToAssist Customer.lnk",
                "Outlook 2016.lnk",
                "Relativity.url",
                "SPG-LEGAL-ProjectInstaller.exe",
                "Recycle Bin"
            };
            string[] documentsFiles =     // My Docs Folder files to remain
            {
                "SPG-LEGAL-ProjectInstaller.exe",
                "Custom Office Templates",
                "OneNote Notebooks",
                "Sound recordings"
            };
            string[] downloadsFiles =     // Downloads Folder files to remain
            {
                "CitrixOnlinePluginWeb.exe",
                "ViewerInstallationKit.exe",
                "VMware-Horizon-View-Client-x86_64-3.5.2-3150477.exe",
                "Setup.X86.en-us_O365BusinessRetail_0ac07f2e-0865-40af-bf1b-7077409aecd4_TX_PR_b_32_.exe",
                "SPG-LEGAL-ProjectInstaller.exe",
                "CutePDF Writer",
                "Citrix Receiver",
                "Citrix Receiver4.7",
                "Citrix Receiver 4.7",
                "CitrixReceiver4.5",
                "CitrixReceiver 4.5",
                "GoToAssist Installer 3.1.0.1251",
                "GoToAssist 3_5 1488",
                "My Company_Default Group_WIN64BIT",
                "Reboot Restore Rx 20160810 v2.1"
            };

            string userProfile = Environment.GetEnvironmentVariable("USERPROFILE");

            if (!userProfile.Contains("LAuser") && !userProfile.Contains("NYuser"))
            {
                MessageBox.Show(userProfile + ": User Profile Cleaning not allowed");
                return;
            }




            string desktopPath = Path.Combine(userProfile, "Desktop");
            string documentsPath = Path.Combine(userProfile, "Documents");
            string downloadsPath = Path.Combine(userProfile, "Downloads");

            // Start with Desktop cleaning
            string[] filePaths = Directory.GetFiles(desktopPath);
            foreach (string filePath in filePaths)
            {
                if (desktopFiles.Contains(Path.GetFileName(filePath)))
                {
                   // keep these files
                }
                else
                {
                    File.Delete(filePath);
                }
            }
            filePaths = Directory.GetDirectories(desktopPath);
            foreach (string filePath in filePaths)
            {
                if (desktopFiles.Contains(Path.GetFileName(filePath)))
                {
                    // keep these files
                }
                else
                {
                    Directory.Delete(filePath,true);
                }
            }


            // Start with Documents cleaning
            filePaths = Directory.GetFiles(documentsPath);
            foreach (string filePath in filePaths)
            {
                if (documentsFiles.Contains(Path.GetFileName(filePath)))
                {
                    // keep these files
                }
                else
                {
                    File.Delete(filePath);
                }
            }
            filePaths = Directory.GetDirectories(documentsPath);
            foreach (string filePath in filePaths)
            {
                if (documentsFiles.Contains(Path.GetFileName(filePath)))
                {
                    // keep these files
                }
                else
                {
                    Directory.Delete(filePath, true);
                }
            }


            // Start with Downloads cleaning
            filePaths = Directory.GetFiles(downloadsPath);
            foreach (string filePath in filePaths)
            {
                if (downloadsFiles.Contains(Path.GetFileName(filePath)))
                {
                    // keep these files
                }
                else
                {
                    File.Delete(filePath);
                }
            }
            filePaths = Directory.GetDirectories(downloadsPath);
            foreach (string filePath in filePaths)
            {
                if (downloadsFiles.Contains(Path.GetFileName(filePath)))
                {
                    // keep these files
                }
                else
                {
                    Directory.Delete(filePath, true);
                }
            }


        }
    }
}
