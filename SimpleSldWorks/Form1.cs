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
                // get the path to the example part
                String path = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Example.SLDPRT");

                if (File.Exists(path))
                {
                    // open the example part in solidworks
                    int errors = 0;
                    ModelDoc2 swModel = swApp.OpenDoc2(path, (int)swDocumentTypes_e.swDocPART, true, false, true, ref errors);
                    if (swModel != null)
                    {
                        PartDoc partDoc = (PartDoc)swModel;

                        // get the solid bodies in the part
                        var bodies = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, false);

                        ModelDocExtension swModExt = (ModelDocExtension)swModel.Extension;

                        // iterate through each solid body and delete any hidden solidbodies
                        foreach (Body2 body in bodies)
                        {
                            if (!body.Visible)
                            {
                                swModExt.SelectByID2(body.Name, "SOLIDBODY", 0, 0, 0, false, 0, null, 0);
                                swModel.FeatureManager.InsertDeleteBody();
                                swModel.ClearSelection();
                            }
                        }

                        // export the result
                        String desktopPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                        int warnings = 0;

                        swModExt.SaveAs(Path.Combine(desktopPath, "Result.step"), 0, 0, null, ref errors, ref warnings);

                        swApp.CloseAllDocuments(true);
                    }
                }

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
            }
            GC.Collect();
            swApp = null;
        }
    }
}
