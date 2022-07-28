using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;

namespace findandreplacenarr
{
    public partial class Form1 : Form
    {
        List <Panel> listPanel = new List <Panel> ();
        int panelIndex;
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        //Find and Replace Method
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        //Fire Separation Distance Find and Replace Function
        private void FSDFindAndReplace(string Findfsd, string FindfsdRating, string FindfsdOpening, int input, Microsoft.Office.Interop.Word.Application wordApp, string FindfsdOccupancy, string buildingType)
        {
            int IntFSDInput = input;
            string noccupancy = FindfsdOccupancy;
            string FSDRating;
            string FSDOpening;
            if (IntFSDInput > 0 && IntFSDInput < 3)
            {
                FSDOpening = "Not Permitted";
                this.FindAndReplace(wordApp, FindfsdOpening, FSDOpening);
                this.FindAndReplace(wordApp, Findfsd, IntFSDInput);
                if (noccupancy == "F-1" || noccupancy == "M" || noccupancy == "S-1")
                {
                    FSDRating = "2";
                    this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                }
                else
                {
                    FSDRating = "1";
                    this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                }

            }
            else if (IntFSDInput >= 3 && IntFSDInput < 5)
            {
                FSDOpening = "15%";
                this.FindAndReplace(wordApp, FindfsdOpening, FSDOpening);
                this.FindAndReplace(wordApp, Findfsd, IntFSDInput);
                if (noccupancy == "F-1" || noccupancy == "M" || noccupancy == "S-1")
                {
                    FSDRating = "2";
                    this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                }
                else
                {
                    FSDRating = "1";
                    this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                }
            }
            else if (IntFSDInput >= 5 && IntFSDInput < 10)
            {
                FSDOpening = "25%";
                this.FindAndReplace(wordApp, FindfsdOpening, FSDOpening);
                this.FindAndReplace(wordApp, Findfsd, IntFSDInput);
                if (noccupancy == "F-1" || noccupancy == "M" || noccupancy == "S-1")
                {
                    if (buildingType == "Type IA")
                    {
                        FSDRating = "2";
                        this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                    }
                    else
                    {
                        FSDRating = "1";
                        this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                    }
                }
                else
                {
                    FSDRating = "1";
                    this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                }
            }
            else if (IntFSDInput >= 10 && IntFSDInput < 15)
            {
                FSDOpening = "45%";
                this.FindAndReplace(wordApp, FindfsdOpening, FSDOpening);
                this.FindAndReplace(wordApp, Findfsd, IntFSDInput);
                if (buildingType == "Type IIB" || buildingType == "Type VB")
                {

                    FSDRating = "0";
                    this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                }
                else
                {
                    FSDRating = "1";
                    this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                }
            }
            else if (IntFSDInput >= 15 && IntFSDInput < 20)
            {
                FSDOpening = "75%";
                this.FindAndReplace(wordApp, FindfsdOpening, FSDOpening);
                this.FindAndReplace(wordApp, Findfsd, IntFSDInput);
                if (buildingType == "Type IIB" || buildingType == "VB")
                {

                    FSDRating = "0";
                    this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                }
                else
                {
                    FSDRating = "1";
                    this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
                }
            }
            else if(IntFSDInput >= 20)
            {
                FSDOpening = "No Limit";
                this.FindAndReplace(wordApp, FindfsdOpening, FSDOpening);
                this.FindAndReplace(wordApp, Findfsd, IntFSDInput);
                FSDRating = "0";
                this.FindAndReplace(wordApp, FindfsdRating, FSDRating);
            }

        }


        //Create DOC
        private void CreateWordDocument(object filename, object SaveAs)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            //object missing = Missing.Value;
            Microsoft.Office.Interop.Word.Document myWordDoc = null;


            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = true;
                wordApp.Visible = true;

                myWordDoc = wordApp.Documents.Open(ref filename, ref readOnly);

                myWordDoc.Activate();

                //find and replace
                this.FindAndReplace(wordApp, "PNAME", ProjectNameInput.Text);
                this.FindAndReplace(wordApp, "PADDRESS", ProjectAddressInput.Text);
                this.FindAndReplace(wordApp, "PCITY", ProjectCityInput.Text);
                this.FindAndReplace(wordApp, "PSTATE", ProjectStateInput.Text);
                this.FindAndReplace(wordApp, "PZIPCODE", ProjectZipcodeInput.Text);
                this.FindAndReplace(wordApp, "ARCH", AccountNameInput.Text);
                this.FindAndReplace(wordApp, "ARCHADD", AccountAddressInput.Text);
                this.FindAndReplace(wordApp, "ARCHZIP", AccountZipcodeInput.Text);

                string buildingType = "Type";
                bool VOA = false;
                bool VOB = false;
                bool VOM = false;
                bool VOR1 = false;
                bool VOR2 = false;
                bool VOI1 = false;
                bool VOI3 = false;
                bool VOS = false;

                if (occupancyClassificationInput.CheckedItems.Contains("A"))
                {
                    VOA = true;
                }
                if (occupancyClassificationInput.CheckedItems.Contains("B"))
                {
                    VOB = true;
                }
                if (occupancyClassificationInput.CheckedItems.Contains("M"))
                {
                    VOM = true;
                }
                if (occupancyClassificationInput.CheckedItems.Contains("R-1"))
                {
                    VOR1 = true;
                }
                if (occupancyClassificationInput.CheckedItems.Contains("R-2"))
                {
                    VOR2 = true;
                }
                if (occupancyClassificationInput.CheckedItems.Contains("I-1"))
                {
                    VOI1 = true;
                }
                if (occupancyClassificationInput.CheckedItems.Contains("I-3"))
                {
                    VOI3 = true;
                }
                if (occupancyClassificationInput.CheckedItems.Contains("S"))
                {
                    VOS = true;
                }


                int intBuildingHeight = int.Parse(BuildingHeightInput.Text);

                if (intBuildingHeight <= 75)
                {
                    //Set, find and replace BUILDTYPE
                    buildingType = "Type IIB";
                    this.FindAndReplace(wordApp, "BUILDTYPE", buildingType);

                    //Delete columns that aren't IIB
                    Microsoft.Office.Interop.Word.Table table2 = myWordDoc.Tables[2];
                    for (int i = 0; i < 3; i++)
                    {

                        table2.Columns[2].Delete();
                    }

                    //Delete the other table
                    Microsoft.Office.Interop.Word.Table table2AndAHalf = myWordDoc.Tables[3];
                    table2AndAHalf.Delete();


                }
                else if (intBuildingHeight > 75 && intBuildingHeight <= 85)
                {
                    //Set, find and replace BUILDTYPE
                    buildingType = "Type IIA";
                    this.FindAndReplace(wordApp, "BUILDTYPE", buildingType);

                    //Delete columns that aren't IIA
                    Microsoft.Office.Interop.Word.Table table2 = myWordDoc.Tables[2];
                    table2.Columns[5].Delete();
                    for (int i = 0; i < 2; i++)
                    {
                        table2.Columns[2].Delete();
                    }

                    //Delete the other table
                    Microsoft.Office.Interop.Word.Table table2AndAHalf = myWordDoc.Tables[3];
                    table2AndAHalf.Delete();

                }
                else if (intBuildingHeight > 85 && intBuildingHeight <= 180)
                {

                    if (isSprinklered.Checked)
                    {
                        //Set, find and replace BUILDTYPE
                        buildingType = "Type IIA";
                        this.FindAndReplace(wordApp, "BUILDTYPE", buildingType);

                        //Delete columns that aren't IIA
                        Microsoft.Office.Interop.Word.Table table2 = myWordDoc.Tables[2];
                        table2.Columns[5].Delete();
                        for (int i = 0; i < 2; i++)
                        {
                            table2.Columns[2].Delete();
                        }

                        //Change Primary Column
                        Microsoft.Office.Interop.Word.Cell primColChange = table2.Cell(3, 2);
                        primColChange.Range.Text = "2";

                        //Delete the other table
                        Microsoft.Office.Interop.Word.Table table2AndAHalf = myWordDoc.Tables[3];
                        table2AndAHalf.Delete();

                    }
                    else
                    {
                        //Set, find and replace BUILDTYPE
                        buildingType = "Type IB";
                        this.FindAndReplace(wordApp, "BUILDTYPE", buildingType);

                        //Delete columns that aren't IB
                        Microsoft.Office.Interop.Word.Table table2 = myWordDoc.Tables[2];
                        table2.Columns[5].Delete();
                        table2.Columns[4].Delete();
                        table2.Columns[2].Delete();

                        //Delete the other table
                        Microsoft.Office.Interop.Word.Table table2AndAHalf = myWordDoc.Tables[3];
                        table2AndAHalf.Delete();

                    }

                }
                else if (intBuildingHeight > 180)
                {
                    
                    if (intBuildingHeight >= 420)
                    {
                        //Set, find and replace BUILDTYPE
                        buildingType = "Type IA Reduced";
                        this.FindAndReplace(wordApp, "BUILDTYPE", buildingType);

                        //Delete columns that aren't IB
                        Microsoft.Office.Interop.Word.Table table2 = myWordDoc.Tables[2];
                        table2.Columns[5].Delete();
                        table2.Columns[4].Delete();
                        table2.Columns[2].Delete();

                        //Change Primary Column
                        Microsoft.Office.Interop.Word.Cell primColChange = table2.Cell(3, 2);
                        primColChange.Range.Text = "3";

                        //Change Building Element Header
                        Microsoft.Office.Interop.Word.Cell buildElemChange = table2.Cell(1, 2);
                        buildElemChange.Range.Text = "Type IA Reduced";

                        //Delete the other table
                        Microsoft.Office.Interop.Word.Table table2AndAHalf = myWordDoc.Tables[3];
                        table2AndAHalf.Delete();

                    }
                    else
                    {
                        //Set, find and replace BUILDTYPE
                        buildingType = "Type IA";
                        this.FindAndReplace(wordApp, "BUILDTYPE", buildingType);

                        //Delete columns that aren't IA
                        Microsoft.Office.Interop.Word.Table table2 = myWordDoc.Tables[2];
                        for (int i = 0; i < 3; i++)
                        {
                            table2.Columns[3].Delete();
                        }

                        //Delete the other table
                        Microsoft.Office.Interop.Word.Table table2AndAHalf = myWordDoc.Tables[3];
                        table2AndAHalf.Delete();

                    }
                }

                //Fire Separation Distance North
                FSDFindAndReplace("NFSD", "NFSDRating", "NFSDOpening", int.Parse(NFSDInput.Text), wordApp, NFSDOccupancy.GetItemText(NFSDOccupancy.SelectedItem), buildingType);
                //Fire Separation Distance South
                FSDFindAndReplace("SFSD", "SFSDRating", "SFSDOpening", int.Parse(SFSDInput.Text), wordApp, SFSDOccupancy.GetItemText(SFSDOccupancy.SelectedItem), buildingType);
                //Fire Separation Distance East
                FSDFindAndReplace("EFSD", "EFSDRating", "EFSDOpening", int.Parse(EFSDInput.Text), wordApp, EFSDOccupancy.GetItemText(EFSDOccupancy.SelectedItem), buildingType);
                //Fire Separation Distance West
                FSDFindAndReplace("WFSD", "WFSDRating", "WFSDOpening", int.Parse(WFSDInput.Text), wordApp, WFSDOccupancy.GetItemText(WFSDOccupancy.SelectedItem), buildingType);


            }

        }


        private void button1_Click(object sender, EventArgs e)
        {
            CreateWordDocument(@"C:\Users\alexa\Documents\SLS\auto narr\findandreplacenarr\Narrative Template.docx", @"C:\Users\alexa\Downloads\DDMMYY_SLS XXXX_Project Name_FPLS Narrative_7th Edit. Code_Template 2020 output.docx");
        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void label1_Click_2(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_2(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void PreviousButton_Click(object sender, EventArgs e)
        {
            if (panelIndex > 0)
            {
                listPanel[--panelIndex].BringToFront();
            }
        }

        private void NextButton_Click(object sender, EventArgs e)
        {
            if (panelIndex < listPanel.Count - 1)
            {
                listPanel[++panelIndex].BringToFront();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listPanel.Add(panel1);
            listPanel.Add(panel2);
            listPanel.Add(panel3);
            listPanel[panelIndex].BringToFront();
        }
    }
}
