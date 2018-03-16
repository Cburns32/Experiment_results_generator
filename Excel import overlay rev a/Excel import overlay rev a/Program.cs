using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Excel_import_overlay_rev_a
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

            //         // ListFiles(); //(ref folderProjects, ref fileNames, ref fileCount);
        }

        /*
        static void ListFiles()//String folderProject, )
        {
            string folderProject = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\INBRE_WORK";
            String template = folderProject + "\\template.xlsx";
            int fileCount = 0;

            DirectoryInfo di = new DirectoryInfo(folderProject);
            foreach (FileInfo fi in di.GetFiles())
            {
                if (fi.Name != ".")
                {
                    string fileName = fi.Name;
                    string fullFileName = fileName.Substring(0, fileName.Length - 4);
                    fileCount += 1;
                    string fullprintout = " Full file name: " + fullFileName + " |  File name: " + fileName + " | File count: " + fileCount + Environment.NewLine;
                }
            }
        }
            //MessageBox.Show("filecount: " fileCount); //figure out how to pass filecount to outside function
       */
    }

}