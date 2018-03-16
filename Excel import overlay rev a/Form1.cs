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
            if (button1WasClicked == true && impulseFilesAvailable == false || parameterFilesAvailable == false)
                MessageBox.Show("Error: Select folders for each input");

            else if (button1WasClicked == true && impulseFilesAvailable == true && parameterFilesAvailable == true)
            {
                System.Windows.Forms.MessageBox.Show("Ready to do stuff with passed files", "Message");

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

        //           openImpulseFolderButton_Click
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

                    System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(impulseFbd.SelectedPath);
                    System.IO.FileSystemInfo[] impulseFiles = di.GetFileSystemInfos();
                    impulseCheckedListBox1.Items.AddRange(impulseFiles);
                    impulseFilesAvailable = true;

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
        }
    }
}

