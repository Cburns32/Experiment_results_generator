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
        /*private OpenFileDialog openFileDialog1;
          private string openFileName, folderName;
          private RichTextBox richTextBox1;
          private bool folderSelected = false;
          private MainMenu mainMenu1;
          private MenuItem fileMenuItem, openMenuItem;
          private MenuItem folderMenuItem, closeMenuItem;
        */

        //for importing files from folder to text in checkedBoxList
        private DialogResult result = new DialogResult();
        private bool parameterFilesAvailable = false;
        private bool impulseFilesAvailable = false;
        private string folderProject;
        // Variables migrated from console application
        string[] stringSeparator = { "  ", "   ", "\t", "\t\t", " \t", "  \t" };
        //string[] wsNames = { "control", "10", "0.1", "100", "0.01", "1000", "0.001" };
        private string[] impulseFolders = { "01_IL12_strength_4cycle", "02_OXP_dose", "03_OXP_dose_treatment_cycle+1", "04_Treatment_Cycle" };
        private const int numberGenes = 48;
        private string[] wsNamesOriginal, wsNamesParameters, wsNamesInputs1, wsNamesInputs2;

        public Form1()
        {
            InitializeComponent();
        }

        // global variables   
        private bool button1WasClicked = false;

        //      private string folderProject1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\INBRE_WORK";



        private void button1_Click(object sender, EventArgs e)
        {
            button1WasClicked = true;
           // if (button1WasClicked == true && impulseFilesAvailable == false || parameterFilesAvailable == false)
          //      MessageBox.Show("Error: Select folders for each input");

            if (button1WasClicked == true && impulseFilesAvailable == true && parameterFilesAvailable == true)
            {
                System.Windows.Forms.MessageBox.Show("Ready to do stuff with passed files", "Message");
                for (int pp = 0; pp < 4; pp++)
                {
 
                    //string folderProject = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\F7Res_TreatmentStrtegy";
                   //
                    //tring folderProject = impulseFbd.SelectedPath;
                    string template = folderProject + "\\template.xlsx";
                    var wb = new XLWorkbook(template);
                    var wsheets = wb.Worksheets;
                    var files = Directory.GetFiles(folderProject);

                    //worksheet names in the Excel template
                    wsNamesOriginal = new string[4];
                    wsNamesParameters = new string[4];
                    wsNamesInputs1 = new string[4];
                    wsNamesInputs2 = new string[4];
                    //excel template

                    for (int wwss = 0; wwss < 4; wwss++)
                    {
                        string sequenceString = "data" + (wwss + 1);
                        wsNamesOriginal[wwss] = sequenceString + "_Original_Data";
                        wsNamesParameters[wwss] = sequenceString + "_Parameters";
                        wsNamesInputs1[wwss] = sequenceString + "_Inputs-1";
                        wsNamesInputs2[wwss] = sequenceString + "_Inputs-2";
                    }

                    //excel template
                    //directory of solution files - under "Controlxx\Impulsesx_0x_...\
                    string folderSolutions = folderProject + "\\Results\\Solutions";
                    //directory for generated Excel files
                    string folderExcels = folderProject + "\\Results\\Excels";

                    for (int control = 1; control <= 30; control++)
                    {
                        string tempy0;
                        if (control < 10)
                            tempy0 = "0" + control;
                        else
                            tempy0 = control.ToString();

                        //solution control folder
                        string controlFolder = folderProject + "\\Results\\Solutions\\Control" + tempy0;
                        string impulseFolder = controlFolder + "\\Impulses" + (pp + 1) + "_" + impulseFolders[pp];

                        foreach (var file in files)
                        {
                            string filename = Path.GetFileName(file);
                            int nameIndex = Convert.ToInt32(filename.Substring(6, 1)) - 1;

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
 
    private void Form1_Load(object sender, EventArgs e)
    {

    }
    // Idea to make eventhandlers for drag and drop to sort files into impulses/solutions sections
    private void fileLabel1_MouseDown(object sender, MouseEventArgs e)
    {
        //Get the data to pass in the DnD operation
        string strText = ((Label)sender).Text;

        //Start the DnD operation
        DoDragDrop(strText, DragDropEffects.Copy);
    }
    private void panel1_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent("Text"))
            e.Effect = DragDropEffects.Copy;
        else
            e.Effect = DragDropEffects.None;
    }
    /*      private void panel1_DragDrop(object sender, DragEventArgs e)
    {
        //Get the text
        string strText = e.Data.GetData("Text", true) as string;

        //Get the position relative to the client where the drop occurred
        Point pt = panel1.PointToClient(new Point(e.X, e.Y));

        //Create a new label to store in the panel
        Label lbl = new Label();
        lbl.Text = strText;
        lbl.Location = pt;

        //Add to the panel
        panel1.Controls.Add(lbl);
    }*/
    private void button3_Click(object sender, EventArgs e)
    {

    }

    private void fileLabel1_MouseDown(object sender, EventArgs e)
    {
        //Get the data to pass in the DnD operation
        string strText = ((Label)sender).Text;

        //Start the DnD operation
        DoDragDrop(strText, DragDropEffects.Copy);
    }
    //Linklable implemented just for convenience
    /*  private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        if (button1WasClicked == true)
        {
            System.Diagnostics.Process.Start(folderProject);
            button1WasClicked = false;
        }
    }*/

    private void openParametersFolderButton_Click(object sender, EventArgs e)
    {
        using (var parameterFbd = new FolderBrowserDialog())
        {
            result = parameterFbd.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(parameterFbd.SelectedPath))
            {
                string[] files = Directory.GetFiles(parameterFbd.SelectedPath);

                System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
                System.Windows.Forms.MessageBox.Show("Folder Path: " + parameterFbd.SelectedPath.ToString(), "Message");

                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(parameterFbd.SelectedPath);
                System.IO.FileSystemInfo[] impulseFiles = di.GetFileSystemInfos();
                parameterCheckedListBox2.Items.AddRange(impulseFiles);
                parameterFilesAvailable = true;

            }
        }

    }

    //          new projectFolder
    private void openFolderButton_Click(object sender, EventArgs e)
    {
        using (var impulseFbd = new FolderBrowserDialog())
        {
            result = impulseFbd.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(impulseFbd.SelectedPath))
            {
                string[] files = Directory.GetFiles(impulseFbd.SelectedPath);

                System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
                System.Windows.Forms.MessageBox.Show("Folder Path: " + impulseFbd.SelectedPath.ToString(), "Message");
                folderProject = impulseFbd.SelectedPath.ToString();
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(impulseFbd.SelectedPath);
                System.IO.FileSystemInfo[] impulseFiles = di.GetFileSystemInfos();
                impulseCheckedListBox1.Items.AddRange(impulseFiles);
                impulseFilesAvailable = true;

            }
        }
    }
    /*
    openFileDialog1 = new OpenFileDialog();
    // If a file is not opened, then set the initial directory to the
    // FolderBrowserDialog.SelectedPath value.
    if (!fileOpened)
    {
        openFileDialog1.InitialDirectory = folderBrowserDialog1.SelectedPath;
        openFileDialog1.FileName = null;
    }

    // Display the openFile dialog.
    DialogResult result = openFileDialog1.ShowDialog();

    // OK button was pressed.
    if (result == DialogResult.OK)
    {
        openFileName = openFileDialog1.FileName;
        MessageBox.Show(openFileName);
        string FolderName = new DirectoryInfo(System.IO.Path.GetDirectoryName(openFileName)).Name;

        //DirectoryInfo openFolderName = new DirectoryInfo(openFileDialog1.FileName);
        MessageBox.Show(FolderName);

        try
        {
            // Output the requested file in richTextBox1.
            Stream s = openFileDialog1.OpenFile();
            richTextBox1.LoadFile(s, RichTextBoxStreamType.RichText);
            s.Close();

            fileOpened = true;

        }
        catch (Exception exp)
        {
            MessageBox.Show("An error occurred while attempting to load the file. The error is:"
                            + System.Environment.NewLine + exp.ToString() + System.Environment.NewLine);
            fileOpened = false;
        }
        Invalidate();

        closeMenuItem.Enabled = fileOpened;
    }

    // Cancel button was pressed.
    else if (result == DialogResult.Cancel)
    {
        return;
    }*/
    public string[] getSubArray(string[] data, int index, int length)
    {
        string[] result = new string[length];
        Array.Copy(data, index, result, 0, length);
        return result;
    }
    //      static void Main(string[] args)
    //    {
    //     new Program();
    //}
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
}
}

