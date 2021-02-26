using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace EdgeCheckDwg
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
                fbd.SelectedPath = @"X:\Commesse\Focchi\200000 40L\02.Shop_Drawings\01.Da inviare";

                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    path = fbd.SelectedPath;
                }
            }
            return path;
        }
    }
}
