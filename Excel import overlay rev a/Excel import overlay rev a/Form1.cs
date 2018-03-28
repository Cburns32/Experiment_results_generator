using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ClosedXML.Excel;

namespace Excel_import_overlay_rev_a
{
    public partial class Form1 : Form
    {
        private FolderBrowserDialog folderBrowserDialog1;
        private DialogResult impulseResult = new DialogResult();
        private bool FilesAvailable = false;
        private string folderProject;
        string[] stringSeparator = { "  ", "   ", "\t", "\t\t", " \t", "  \t" };
        private string[] impulseFolders = { "01_IL12_strength_4cycle", "02_OXP_dose", "03_OXP_dose_treatment_cycle+1", "04_Treatment_Cycle" };
        private const int numberGenes = 48;
        private string[] wsNamesOriginal, wsNamesParameters, wsNamesInputs1, wsNamesInputs2; 
        private bool button1WasClicked = false;

        public Form1()
        {
            InitializeComponent();
            System.Windows.Forms.MessageBox.Show("Select the main project folder:", "Message");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1WasClicked = true;
            // if (button1WasClicked == true && impulseFilesAvailable == false || parameterFilesAvailable == false)
            //      MessageBox.Show("Error: Select folders for each input");

            if (button1WasClicked == true && FilesAvailable == true)// && parameterFilesAvailable == true)
            {
                System.Windows.Forms.MessageBox.Show("Ready to do stuff with passed files", "Message");

                //worksheet names in the Excel template
                wsNamesOriginal = new string[4];
                wsNamesParameters = new string[4];
                wsNamesInputs1 = new string[4];
                wsNamesInputs2 = new string[4];

                for (int wwss = 0; wwss < 4; wwss++)
                {
                    string sequenceString = "data" + (wwss + 1);
                    wsNamesOriginal[wwss] = sequenceString + "_Original_Data";
                    wsNamesParameters[wwss] = sequenceString + "_Parameters";
                    wsNamesInputs1[wwss] = sequenceString + "_Inputs-1";
                    wsNamesInputs2[wwss] = sequenceString + "_Inputs-2";
                }

                //excel template
                string template = folderProject + "\\template.xlsx";
                //directory of solution files - under "Controlxx\Impulsesx_0x_...\
                string folderSolutions = folderProject + "\\Results\\Solutions";
                //directory for generated Excel files
                string folderExcels = folderProject + "\\Results\\Excels";

                for (int control = 1; control <= 30; control++)
                {
                    string tempy0;
                    if (control < 10)
                    {
                        tempy0 = "0" + control;
                    }
                    else
                        tempy0 = control.ToString();

                    // "\\results\\Solution" Control Folders 
                    string controlFolder = folderProject + "\\Results\\Solutions\\Control" + tempy0;

                    for (int pp = 0; pp < 4; pp++)
                    {
                        // Control variable impulse set folder
                        string impulseFolder = controlFolder + "\\impulses" + (pp + 1) + "_" + impulseFolders[pp];

                        var wb = new XLWorkbook(template);
                        var wsheets = wb.Worksheets;
                        var files = Directory.GetFiles(impulseFolder);

                        foreach (var file in files)
                        {
                            string filename = Path.GetFileName(file);
                            int nameIndex = Convert.ToInt32(filename.Substring(6, 1)) - 1;
                          //  MessageBox.Show(nameIndex.ToString());  //Used for tracking names during debug

                            var lines = File.ReadAllLines(file);
                            //Console.WriteLine(lines.Length);
                            //fill out original data sheet
                            writeSolution(lines, wsheets, wsNamesOriginal[nameIndex]);
                            //fill out parameter sheet
                            string[] dataParamSheet = getSubArray(lines, 13, 8);
                            writeSolution(dataParamSheet, wsheets, wsNamesParameters[nameIndex]);
                            //fill out inputs-1
                            string[] dataInputs1 = getSubArray(lines, 25, 720);
                            writeSolution(dataInputs1, wsheets, wsNamesInputs1[nameIndex]);
                            //fill out inputs-2
                            string[] dataInputs2 = getSubArray(lines, 748, 720);
                            writeSolution(dataInputs2, wsheets, wsNamesInputs2[nameIndex]);
                        }
                        //save the workbook
                        wb.SaveAs(folderExcels + "\\ParSet" + tempy0 + "_" + impulseFolders[pp] + ".xlsx");
                    }
                }
            }
        }

          //          new projectFolder
        private void openFolderButton_Click(object sender, EventArgs e)
        {
            using (var impulseFbd = new FolderBrowserDialog())
            {
                impulseResult = impulseFbd.ShowDialog();
                if (impulseResult == DialogResult.OK && !string.IsNullOrWhiteSpace(impulseFbd.SelectedPath))
                {
                    folderProject = impulseFbd.SelectedPath;
                    string[] files = Directory.GetFiles(folderProject);

                    System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
                    System.Windows.Forms.MessageBox.Show("Folder Path: " + impulseFbd.SelectedPath.ToString(), "Message");


                    System.Windows.Forms.MessageBox.Show("FolderProject:", folderProject.ToString());


                    System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(impulseFbd.SelectedPath);
                    System.IO.FileSystemInfo[] impulseFiles = di.GetFileSystemInfos();

                   /* foreach (var impulse in impulseFiles)
                        System.Windows.Forms.MessageBox.Show("impulsefiles[]: " + impulse.ToString(), "Message");
                        */
                    impulseCheckedListBox1.Items.AddRange(impulseFiles);

                }
            }
        }
        public string[] getSubArray(string[] data, int index, int length)
        {
            string[] result = new string[length];
            Array.Copy(data, index, result, 0, length);
            return result;
        }

        public void writeSolution(string[] data, IXLWorksheets ws, string sheetName)
        {
            var originalSolution = ws.Worksheet(sheetName);
            importData(data, originalSolution);
        }

        public void importData(string[] data, IXLWorksheet wkst)
        {
            int line_num = data.Length;
            for (int i = 0; i < line_num; i++)
            {
                var cells = data[i].Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);
                int cell_num = cells.Length;
                for (int k = 0; k < cell_num; k++)
                {
                    wkst.Cell(i + 1, k + 1).Value = cells[k];
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }

}