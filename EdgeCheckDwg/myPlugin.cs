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

[assembly: ExtensionApplication(typeof(EdgeAutocadPlugins.MyPlugin))]

namespace EdgeAutocadPlugins
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

            rb = new RibbonButton();

            rb.Text = "Show";
            rb.Size = RibbonItemSize.Large;
            rb.ShowText = true;
            rb.Orientation = System.Windows.Controls.Orientation.Vertical;

            rb.CommandParameter = "_generaEmi ";
            rb.CommandHandler = new SimpleButtonCmdHandler();

            rb.Image = GenericFunction.Convert(Properties.Resources.icon_messages_app_27x20_1x);
            rb.LargeImage = GenericFunction.Convert(Properties.Resources.icon_messages_app_27x20_1x);


            rps.Items.Add(rb);

            RibbonButton rb1;
            rb1 = new RibbonButton();

            rb1.Text = "Show";
            rb1.Size = RibbonItemSize.Large;
            rb1.ShowText = true;
            rb1.Orientation = System.Windows.Controls.Orientation.Vertical;

            rb1.CommandParameter = "_checkPE ";
            rb1.CommandHandler = new SimpleButtonCmdHandler();

            rb1.Image = GenericFunction.Convert(Properties.Resources.Icon_32);
            rb1.LargeImage = GenericFunction.Convert(Properties.Resources.Icon_32);


            rps.Items.Add(rb1);


            return rp;
        }
    }
}