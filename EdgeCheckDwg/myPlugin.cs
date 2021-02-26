using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;
using System.Configuration;

using Autodesk.AutoCAD.Windows;
using Autodesk.Windows;

[assembly: ExtensionApplication(typeof(EdgeCheckDwg.MyPlugin))]

namespace EdgeCheckDwg
{

    // This class is instantiated by AutoCAD once and kept alive for the 
    // duration of the session. If you don't do any one time initialization 
    // then you should remove this class.
    public class MyPlugin : IExtensionApplication
    {
        public const string THISPLUGIN_GUID = "1438AD09-EABB-4D39-988A-74E027F7FCB3";

        public static bool m_FirstStart = true;
        static bool m_isClosing = false;            // indica che l'applicazione è in chiusura

        void IExtensionApplication.Initialize()
        {
            Application.Idle += Application_Idle;
        }

        void IExtensionApplication.Terminate()
        {
        }

        private void Application_Idle(object sender, EventArgs e)
        {
            if (Autodesk.Windows.ComponentManager.Ribbon != null)
            {
                // rimuovi evento, in modo che questa sia la prima e ultima volta che viene eseguito
                Autodesk.AutoCAD.ApplicationServices.Application.Idle -= Application_Idle;

                generaRibbon();
            }
        }

        public void generaRibbon()
        {
            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon != null)
            {
                RibbonTab rtab = ribbon.FindTab("Edge");
                if (rtab != null)
                {
                    ribbon.Tabs.Remove(rtab);
                }
                rtab = new RibbonTab();
                rtab.Title = "Edge";
                rtab.Id = "Testing";
                //Add the Tab
                ribbon.Tabs.Add(rtab);
                addContent(rtab);
            }
        }

        static void addContent(RibbonTab rtab)
        {
            rtab.Panels.Add(AddOnePanel());
        }

        static RibbonPanel AddOnePanel()
        {
            RibbonButton rb;
            RibbonPanelSource rps = new RibbonPanelSource();
            rps.Title = "Genera Emissione";
            RibbonPanel rp = new RibbonPanel();
            rp.Source = rps;

            //Create a Command Item that the Dialog Launcher can use,
            // for this test it is just a place holder.
            //RibbonButton rci = new RibbonButton();
            //rci.Name = "generaEmi";

            //assign the Command Item to the DialgLauncher which auto-enables
            // the little button at the lower right of a Panel
            //rps.DialogLauncher = rci;

            rb = new RibbonButton();
            rb.ShowText = true;
            rb.CommandParameter = "generaEmi ";
            rb.CommandHandler = new SimpleButtonCmdHandler();
            rb.Image = GenericFunction.ToImageSource(Properties.Resources.logo);
            //rb.Size = ;

            rps.Items.Add(rb);
            return rp;
        }
    }

}