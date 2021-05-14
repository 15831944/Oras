using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace EdgeAutocadPlugins
{
    public static class GenericFunction
    {
        public static ImageSource ToImageSource(this Icon icon)
        {
            ImageSource imageSource = Imaging.CreateBitmapSourceFromHIcon(
                icon.Handle,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());

            return imageSource;
        }

        public static string choosePath()
        {
            string path = null;


            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = @"X:\Commesse\Focchi\200000 40L";

                fbd.ShowNewFolderButton = true;

                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    path = fbd.SelectedPath;
                }
            }

            // ! Selezione Cartella tramite interop di excel
            //Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Core.FileDialog fileDialog = app.get_FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFolderPicker);
            //fileDialog.InitialFileName = @"X:\Commesse\Focchi\200000 40L\02.Shop_Drawings\01.Da inviare";
            //int nres = fileDialog.Show();
            //if (nres == -1) //ok
            //{
            //    Microsoft.Office.Core.FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;

            //    string[] selectedFolders = selectedItems.Cast<string>().ToArray();

            //    if (selectedFolders.Length > 0)
            //    {
            //        path = selectedFolders[0];
            //    }
            //}

            return path;
        }

        internal static string[] getRecentiRev(string[] vs)
        {
            List<string> ls = new List<string>();

            List<string> tmp = new List<string>();

            List<string> list = new List<string>(vs);

            foreach (string file in vs)
            {
                string nomeFile = Path.GetFileName(file);
                string codice = nomeFile.Split('_')[0];

                if (tmp.Contains(codice)) continue;

                tmp.Add(codice);
                
                string f = vs.Where(n => n.Contains(codice)).OrderByDescending(x=>x).FirstOrDefault();

                ls.Add(f);
            }

            return ls.ToArray();
        }

        internal static string findParentesi(string s)
        {
            string r = "";

            if(s.Contains("(") && s.Contains(")"))
            {
                int wO = s.IndexOf("(");
                int wC = s.IndexOf(")");

                r = s.Substring(wO, wC-wO);

                r = r.Replace("(", "");
                r = r.Replace(")", "");

                r = r.Trim();
            }

            return r;
        }

        public static BitmapImage Convert(Bitmap src)
        {
            MemoryStream ms = new MemoryStream();
            ((System.Drawing.Bitmap)src).Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
            BitmapImage image = new BitmapImage();
            image.BeginInit();
            ms.Seek(0, SeekOrigin.Begin);
            image.StreamSource = ms;
            image.EndInit();
            return image;
        }

    }
}
