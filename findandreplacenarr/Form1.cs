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

            //Building Occupancy Input
            

            //Fire Separation Distance
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
                
                int intBuildingHeight = int.Parse(BuildingHeightInput.Text);

                bool A1 = false;
                bool A2 = false;
                bool A3 = false;
                bool B = false;
                bool M = false;
                bool R1 = false;
                bool R2 = false;
                bool I1 = false;
                bool I3 = false;
                bool S1 = false;
                bool S2 = false;
                A1 = intBuildingOccupancy.GetItemChecked(0);
                A2 = intBuildingOccupancy.GetItemChecked(1);
                A3 = intBuildingOccupancy.GetItemChecked(2);
                B = intBuildingOccupancy.GetItemChecked(3);
                M = intBuildingOccupancy.GetItemChecked(4);
                R1 = intBuildingOccupancy.GetItemChecked(5);
                R2 = intBuildingOccupancy.GetItemChecked(6);
                I1 = intBuildingOccupancy.GetItemChecked(7);
                I3 = intBuildingOccupancy.GetItemChecked(8);
                S1 = intBuildingOccupancy.GetItemChecked(9);
                S2 = intBuildingOccupancy.GetItemChecked(10);

                //PARAGRAGPH REPLACE PER OCCUPANCY
                if ((A1 == true || A2 == true || A3 == true) && R1 == true)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "For Assembly occupancies, FFPC, NFPA 101 Section 12.3.2 and Hotel occupancies, FFPC, NFPA 101 Section 28.3.2 states that rooms containing high-pressure boilers, large transformers, or other service equipment subject to explosion shall not be located directly under or abutting required exits." +
                        "Hotel units must be separated from adjacent hotel units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 28.3.7.  The hotel unit separation in FBC Section 708 is 1-hour fire partition.");
                }
                else if ((A1 == false && A2 == false && A3 == false) && R1 == true)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "For Hotel occupancies, FFPC, NFPA 101 Section 28.3.2 states that rooms containing high-pressure boilers, large transformers, or other service equipment subject to explosion shall not be located directly under or abutting required exits." +
                        "Hotel units must be separated from adjacent hotel units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 28.3.7.  The hotel unit separation in FBC Section 708 is 1-hour fire partition.");
                }
                if ((A1 == true || A2 == true || A3 == true) && R1 == false)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "For Assembly occupancies, FFPC, NFPA 101 Section 12.3.2 states that rooms containing high-pressure boilers, large transformers, or other service equipment subject to explosion shall not be located directly under or abutting required exits. ");
                }

                if (R2 == true)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "Dwelling units must be separated from adjacent dwelling units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 30.3.7.  The dwelling unit separation in Section FBC Section 708 is 1-hour fire partition.");
                }

                if (I1 == true || I3 == true)
                {
                    this.FindAndReplace(wordApp, "I1I3PARTITION", "Every story shall be divided into not less than two smoke compartments (FBC §420.4, FFPC, NFPA 101 FBC §32.3.3.7).  Each smoke compartment shall have an area not exceeding 22,500 square feet and the maximum travel distance from any point to reach a door in the smoke barrier shall not exceed 200 feet. Smoke barriers shall be constructed in accordance with FFPC, NFPA 101 §8.5 and shall have a minimum 1 - hour fire resistance rating(FFPC, NFPA 101 §32.3.3.7.8).Smoke barrier doors shall be at least 1 ¼ in.thick, solid - bonded wood - core doors, or shall be fire rated for at least 20 minutes(FFPC, NFPA 101 §32.3.3.7.13).At least 15 net square feet per resident shall be provided within the aggregate area of corridors, lounge or dining areas, and other low hazard areas on each side of the smoke barrier(FBC §420.4.1, FFPC, NFPA 101 §32.3.3.7.11), and not less than 6 net square feet for other occupants.");
                }
                else
                {
                    this.FindAndReplace(wordApp, "I1I3PARTITION", "DELETE");
                }

                // Table 1 Editing per Occupancy Input
                Microsoft.Office.Interop.Word.Table table1 = myWordDoc.Tables[1];
                if (A1 == false)
                {
                    table1.Rows[1].Delete();
                }
                if (A2 == false)
                {
                    table1.Rows[2].Delete();
                }
                if (A3 == false)
                {
                    table1.Rows[3].Delete();
                }
                if (B == false)
                {
                    table1.Rows[4].Delete();
                }
                if (M == false)
                {
                    table1.Rows[5].Delete();
                }
                if (R1 == false)
                {
                    table1.Rows[6].Delete();
                }
                if (R2 == false)
                {
                    table1.Rows[7].Delete();
                }
                if (I1 == false)
                {
                    table1.Rows[8].Delete();
                }
                if (I3 == false)
                {
                    table1.Rows[9].Delete();
                }
                if (S1 == false)
                {
                    table1.Rows[10].Delete();
                }
                if (S2 == false)
                {
                    table1.Rows[11].Delete();
                }

                //Table 8 Fire rating of spaces
                
                Microsoft.Office.Interop.Word.Table table8 = myWordDoc.Tables[8];
                if (R1 == false && R2 == false)
                {
                    table8.Rows[8].Delete();
                    table8.Rows[11].Delete();
                }

                // Table 10 Editing per Occupancy Input
                Microsoft.Office.Interop.Word.Table table10 = myWordDoc.Tables[10];
                if (A1 == false && A2 == false && A3 == false)
                {
                    table10.Rows[1].Delete();
                }
               
                if (B == false)
                {
                    table10.Rows[2].Delete();
                }
                if (M == false)
                {
                    table10.Rows[7].Delete();
                }
                if (R1 == false)
                {
                    table10.Rows[6].Delete();
                }
                if (R2 == false)
                {
                    table10.Rows[5].Delete();
                }
                if (I1 == false)
                {
                    table10.Rows[3].Delete();
                }
                if (I3 == false)
                {
                    table10.Rows[4].Delete();
                }
                if (S1 == false || S2 == false)
                {
                    table10.Rows[8].Delete();
                }

                //Table 11 editing
                Microsoft.Office.Interop.Word.Table table11 = myWordDoc.Tables[11];
                if (A1 == false && A2 == false && A3 == false)
                {
                    table11.Rows[1].Delete();
                }

                if (B == false)
                {
                    table11.Rows[2].Delete();
                }
                if (M == false)
                {
                    table11.Rows[5].Delete();
                }
                if (R1 == false)
                {
                    table11.Rows[6].Delete();
                }
                if (R2 == false)
                {
                    table11.Rows[7].Delete();
                }
                if (I1 == false)
                {
                    table11.Rows[3].Delete();
                }
                if (I3 == false)
                {
                    table11.Rows[4].Delete();
                }
                if (S1 == false)
                {
                    table11.Rows[8].Delete();
                }
                if (S2 == false)
                {
                    table11.Rows[9].Delete();
                }

                //EXIT ACCESS SECTION
                
                if (R1 == true || I1 == true)
                {
                    this.FindAndReplace(wordApp, "R1I1UnitExit", "For Hotel Group R-1 occupancies and Res B/C, the FFPC requires two exit access doors from the unit when the guest room or guest suite is over 2,000 sq.ft.  The exit access doors must be located remotely from each other (FFPC, NFPA 101 Section 28.2.5.7).  If limits shown in Table 12 are exceeded, then additional exits must be provided.");
                }

                //OCCUPANT EVAC OR ADDITIONAL STAIR

                if (intBuildingHeight >= 420 && R2 == false)
                {
                    this.FindAndReplace(wordApp, "OEESection", "For buildings greater than 420 ft. in building height, other than R-2 buildings, one additional stairway must be provided in addition to above exit stairs per FBC Section 403.5.2.  There is an alternate provision to the additional stair, which states that an occupant evacuation elevator can be provided in lieu of the stair.  The occupant evacuation elevator, separate from fire service access elevator, must comply with FBC Section 3008 and FFPC, NFPA 101 Section 7.14." + "NOTE: If the building is divided into R - 1 lower floors and R - 2 upper floors where it exceeds 420 ft. in height, then the exception can be applied, and the extra stair is not required since the upper occupancy is R - 2.");
                }

                //LOOPED CORRIDOR FOR R1/R2

                if ((R2 == true || R1 == true) && loopedcorridor == true)
                {
                    this.FindAndReplace(wordApp, "Loopedcorridor", "For R - 2 and R - 1 occupancies, the distance between exits is not applicable to common nonlooped exit access corridors in a building that has corridor doors from the guestroom or guest suite or dwelling unit, which are arranged so that the exits are located in opposite directions from such doors (FBC Section 1007.1.1 Exception 3).The exit discharge must also meet the remoteness requirement.");
                }


                //VERTICAL OPENING SECTION

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

                        this.FindAndReplace(wordApp, "IA REDUCED P1", "FBC Section 403.2.1.1 allows Type IA construction if building height is under 420 ft. to be reduced to Type IB Construction except the required fire resistance rating of columns supporting floors cannot be reduced." + "Based on the type of construction, Type IA Reduced, Tables 504.3 and 506.2 permit unlimited building height and area.  The reduced fire resistance rating of the building elements does not change the building height and building area limitations for the same building without such reductions.");

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
            CreateWordDocument(@"C:\Users\Owner\Desktop\SLS\findreplaceFPLS\Narrative Template.docx", @"C:\Users\alexa\Downloads\DDMMYY_SLS XXXX_Project Name_FPLS Narrative_7th Edit. Code_Template 2020 output.docx");
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

        private void intBuildingOccupancy_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
