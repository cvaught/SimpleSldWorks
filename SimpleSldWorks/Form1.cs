using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SldWorks;
using SwConst;
using System.IO;
using System.Diagnostics;


namespace SimpleSldWorks
{
    public partial class Form1 : Form
    {
        const string progID = "SldWorks.Application";
        SldWorks.SldWorks swApp;
        Boolean alreadyOpen = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.connectToSolidWorks())
            {
                MessageBox.Show("Hello SolidWorks User Group!");
                // get the path to the example part
                //String path = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Example.SLDPRT");

                this.releaseSolidWorks();
            }
            else
            {
                MessageBox.Show("Error starting SolidWorks.");
            }
            
        }

        private Boolean connectToSolidWorks()
        {
            if (swApp == null)
            {
                alreadyOpen = false;
                try
                {
                    swApp = System.Runtime.InteropServices.Marshal.GetActiveObject(progID) as SldWorks.SldWorks;
                    swApp.Visible = true;
                    alreadyOpen = true;
                }
                catch
                {

                }

                if (!alreadyOpen)
                {
                    this.killAnyInstances();
                    Type swAppType = System.Type.GetTypeFromProgID(progID);
                    swApp = System.Activator.CreateInstance(swAppType) as SldWorks.SldWorks;
                    swApp.Visible = true;
                }
                
            }

            if (swApp != null)
            {
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swLockRecentDocumentsList, true);
            }

            return true;
            
        }

        private void releaseSolidWorks()
        {
            if (swApp != null)
            {
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swLockRecentDocumentsList, false);
            }

            if (!alreadyOpen)
            {
                try
                {
                    swApp.CloseAllDocuments(true);
                    swApp.ExitApp();
                }
                catch
                {

                }
                this.killAnyInstances();

            }
            GC.Collect();
            swApp = null;
        }

        public void killAnyInstances()
        {
            try
            {
                foreach (Process proc in Process.GetProcessesByName("SLDWORKS"))
                {
                    proc.Kill();
                }
            }
            catch
            {

            }
            swApp = null;
        }
    }
}
