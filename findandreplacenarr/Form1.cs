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
using DocumentFormat.OpenXml;

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

        //Find Text and Replace with Image Method
        private void FindTextAndReplaceImage(Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document myWordDoc, string textToFind, string imgLocation)
        {
            // Find text and replace with image
            Microsoft.Office.Interop.Word.Find fnd = wordApp.ActiveWindow.Selection.Find;
            fnd.ClearFormatting();
            fnd.Replacement.ClearFormatting();
            fnd.Forward = true;
            fnd.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;

            string imagePath = imgLocation;
            var keyword = textToFind;
            var sel = wordApp.Selection;
            sel.Find.Text = string.Format("[{0}]", keyword);
            wordApp.Selection.Find.Execute(keyword);

            Microsoft.Office.Interop.Word.Range range = wordApp.Selection.Range;
            if (range.Text.Contains(keyword))
            {
                //gets desired range here it gets last character to make superscript in range 
                Microsoft.Office.Interop.Word.Range temprange = myWordDoc.Range(range.End - 7, range.End);//keyword is of 4 charecter range.End - 4
                temprange.Select();
                Microsoft.Office.Interop.Word.Selection currentSelection = wordApp.Selection;
                //currentSelection.Font.Superscript = 1;

                sel.Find.Execute(Replace: WdReplace.wdReplaceOne);
                sel.Range.Select();
                var imagePath1 = Path.GetFullPath(string.Format(imagePath, keyword));
                sel.InlineShapes.AddPicture(FileName: imagePath1, LinkToFile: false, SaveWithDocument: true);
            }
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

                string buildingTypeHeight;
                bool VOA = false;
                bool VOB = false;
                bool VOM = false;
                bool VOR1 = false;
                bool VOR2 = false;
                bool VOI1 = false;
                bool VOI3 = false;
                bool VOS = false;
                int numVOLevels = VOLevelsInput.Text.Count(c => char.IsDigit(c) && c != ',');

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
                A1 = BuildingOccupancyListBox.GetItemChecked(0);
                A2 = BuildingOccupancyListBox.GetItemChecked(1);
                A3 = BuildingOccupancyListBox.GetItemChecked(2);
                B = BuildingOccupancyListBox.GetItemChecked(3);
                M = BuildingOccupancyListBox.GetItemChecked(4);
                R1 = BuildingOccupancyListBox.GetItemChecked(5);
                R2 = BuildingOccupancyListBox.GetItemChecked(6);
                I1 = BuildingOccupancyListBox.GetItemChecked(7);
                I3 = BuildingOccupancyListBox.GetItemChecked(8);
                S1 = BuildingOccupancyListBox.GetItemChecked(9);
                S2 = BuildingOccupancyListBox.GetItemChecked(10);

                //PARAGRAGPH REPLACE PER OCCUPANCY
                if ((A1 == true || A2 == true || A3 == true) && R1 == true && R2 == false)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "For Assembly occupancies, FFPC, NFPA 101 Section 12.3.2 and Hotel occupancies, FFPC, NFPA 101 Section 28.3.2 states that rooms containing high-pressure boilers, large transformers, or other service equipment subject to explosion shall not be located directly under or abutting required exits." +
                        "Hotel units must be separated from adjacent hotel units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 28.3.7.  The hotel unit separation in FBC Section 708 is 1-hour fire partition.");
                }
                else if ((A1 == false && A2 == false && A3 == false) && R1 == true && R2 == false)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "For Hotel occupancies, FFPC, NFPA 101 Section 28.3.2 states that rooms containing high-pressure boilers, large transformers, or other service equipment subject to explosion shall not be located directly under or abutting required exits." +
                        "Hotel units must be separated from adjacent hotel units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 28.3.7.  The hotel unit separation in FBC Section 708 is 1-hour fire partition.");
                }
                else if ((A1 == true || A2 == true || A3 == true) && R1 == false && R2 == false)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "For Assembly occupancies, FFPC, NFPA 101 Section 12.3.2 states that rooms containing high-pressure boilers, large transformers, or other service equipment subject to explosion shall not be located directly under or abutting required exits.");
                }
                else if ((A1 == false && A2 == false && A3 == false) && R1 == false && R2 == true)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "Dwelling units must be separated from adjacent dwelling units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 30.3.7.  The dwelling unit separation in Section FBC Section 708 is 1-hour fire partition.");
                }
                else if((A1 == true || A2 == true || A3 == true) && R1 == false && R2 == true)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "For Assembly occupancies, FFPC, NFPA 101 Section 12.3.2 states that rooms containing high-pressure boilers, large transformers, or other service equipment subject to explosion shall not be located directly under or abutting required exits."
                        + "Dwelling units must be separated from adjacent dwelling units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 30.3.7.  The dwelling unit separation in Section FBC Section 708 is 1-hour fire partition.");
                }
                else if((A1 == true && A2 == true && A3 == true) && R1 == true && R2 == true)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "For Assembly occupancies, FFPC, NFPA 101 Section 12.3.2 and Hotel occupancies, FFPC, NFPA 101 Section 28.3.2 states that rooms containing high-pressure boilers, large transformers, or other service equipment subject to explosion shall not be located directly under or abutting required exits." +
                        "Hotel units must be separated from adjacent hotel units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 28.3.7.  The hotel unit separation in FBC Section 708 is 1-hour fire partition." + "Dwelling units must be separated from adjacent dwelling units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 30.3.7.  The dwelling unit separation in Section FBC Section 708 is 1-hour fire partition.");
                }
                else if((A1 == false && A2 == false && A3 == false) && R1 == true && R2 == true)
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "For Hotel occupancies, FFPC, NFPA 101 Section 28.3.2 states that rooms containing high-pressure boilers, large transformers, or other service equipment subject to explosion shall not be located directly under or abutting required exits." +
                        "Hotel units must be separated from adjacent hotel units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 28.3.7.  The hotel unit separation in FBC Section 708 is 1-hour fire partition." + "Dwelling units must be separated from adjacent dwelling units by ½-hr fire barriers in accordance with FFPC, NFPA 101 Section 30.3.7.  The dwelling unit separation in Section FBC Section 708 is 1-hour fire partition.");
                }
                else
                {
                    this.FindAndReplace(wordApp, "FR ASSEMBLY HOTEL P1", "DELETE");
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
                //-----------------------------------------------------------------How do these work?--------------------------------------------------------------------
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

                //TABLE 13 EDITING
                Microsoft.Office.Interop.Word.Table table13 = myWordDoc.Tables[13];
                if (A1 == false && A2 == false && A3 == false)
                {
                    table13.Rows[1 - 6].Delete();
                    table13.Rows[10 - 12].Delete();
                    table13.Rows[16].Delete();
                    table13.Rows[18].Delete();
                    table13.Rows[20].Delete();
                }
                if (B == false)
                {
                    table13.Rows[15].Delete();
                }
                if (M == false)
                {
                    table13.Rows[13].Delete();
                }
                if (R1 == false && R2 == false)
                {
                    table13.Rows[19].Delete();
                }
                if (I1 == false && I3 == false)
                {
                    table13.Rows[5].Delete();
                    table13.Rows[7 - 10].Delete();
                }
                if (S2 == false)
                {
                    table13.Rows[17].Delete();
                }

                //TABLE 14 EDITING
                Microsoft.Office.Interop.Word.Table table14 = myWordDoc.Tables[14];
                if (A1 == false && A2 == false && A3 == false)
                {
                    table14.Rows[1].Delete();
                }
                if (B == false)
                {
                    table14.Rows[2].Delete();
                }
                if (M == false)
                {
                    table14.Rows[3].Delete();
                }
                if (R2 == false)
                {
                    table14.Rows[5].Delete();
                }
                if (R1 == false)
                {
                    table14.Rows[6].Delete();
                }
                if (I1 == false)
                {
                    table14.Rows[7].Delete(); ;
                }
                if (I3 == false)
                {
                    table14.Rows[8].Delete();
                }

                //TABLE 15 AND 16 EDITING

                Microsoft.Office.Interop.Word.Table table15 = myWordDoc.Tables[15];
                Microsoft.Office.Interop.Word.Table table16 = myWordDoc.Tables[16];
                if (isEmergencyVoiceSystem.Checked && intBuildingHeight >= 75 && isSprinklered.Checked)
                {

                    //Delete Table 15 ------------------------------
                    table15.Delete();
                    if (I1 == true || I3 == true)
                    {
                        table16.Rows[2].Delete();
                        //what do we do for healthcare table16.Rows[3].Delete();
                        table16.Rows[4].Delete();
                    }
                    else
                    {
                        table16.Rows[1].Delete();
                        table16.Rows[2].Delete();
                        table16.Rows[3].Delete();
                    }
                }
                else
                {
                    //Delete Table 16 -------------------------------
                    table16.Delete();
                    if (I1 == true || I3 == true)
                    {
                        table15.Rows[2].Delete();
                        //what do we do for healthcare table15.Rows[3].Delete();
                        table15.Rows[4].Delete();
                    }
                    else
                    {
                        table15.Rows[1].Delete();
                        table15.Rows[2].Delete();
                        table15.Rows[3].Delete();
                    }
                }

                //MEANS OF ESCAPE SECTION
                if ((R1 == true || R2 == true) && isSprinklered.Checked) //required for R2 with only one exit =========================
                {
                    this.FindAndReplace(wordApp, "MeansEscape", "Secondary means of escape windows are not required in dwelling units  [hotel units] when the building is protected by an automatic sprinkler system per FFPC, NFPA 101.  Emergency escape/rescue windows are required by FBC Section 1030 for only for R-2 occupancies in buildings that have only one exit.  The rescue windows are required even if the building is protected by an automatic sprinkler system.");
                }

                //LOW LEVEL EXIT SIGNAGE
                if (R1 == true)
                {
                    this.FindAndReplace(wordApp, "LOWEXIT", "NOTE:   FBC Section 1013.2 requires floor-level exit signs in all R-1 (Hotel) Occupancies.  The bottom of the sign shall not be less than 10 inches and no more than 12 inches above the floor.  The sign shall be flush mounted to the door or wall.  The edge of the sign shall be within 4 inches of the door frame on the latch side.");
                }

                //LUMINOUS EGRESS MARKINGS - NEED TO ADD CODE FOR R2 BUILDINGS WITH NON R2 ACCESSORY SPACES ABOVE THE 75FT
                if (R2 == false && intBuildingHeight >= 75)
                {
                    this.FindAndReplace(wordApp, "LUMINOUSMARKSECTION", "As a high-rise building, FBC §403.5.5 states that approved luminous egress path markings delineating the exit path must be provided in Group A, B, E, I, M and R-1 occupancies in accordance with FBC §1025.  Markings within the exit enclosures are required to be provided on steps, landings, handrails, perimeter demarcation lines, and discharge doors from the exit enclosure.  Materials should comply with either UL 1994 or ASTM E2072. ");
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
                if ((R2 == true || R1 == true) && isNONLooped.Checked)
                {
                    this.FindAndReplace(wordApp, "NONLoopedcorridor", "For R - 2 and R - 1 occupancies, the distance between exits is not applicable to common nonlooped exit access corridors in a building that has corridor doors from the guestroom or guest suite or dwelling unit, which are arranged so that the exits are located in opposite directions from such doors (FBC Section 1007.1.1 Exception 3).The exit discharge must also meet the remoteness requirement.");
                }

                //STREET FLOOR REQ
                if (B == true || R1 == true || I1 == true)
                {
                    this.FindAndReplace(wordApp, "StreetFloorREQ", "For Business (FFPC, NFPA 101 Section 38.2.3.3), Hotel (FFPC, NFPA 101 Section 28.2.3.2) and Res B/C,(check others), the code requires that street floor exits must accommodate the occupant load of street floor plus stair discharging onto street floor.");
                }

                //DOOR LOCK I1
                if (I1 == true)
                {
                    this.FindAndReplace(wordApp, "DoorLockI1", "Door locking arrangements shall be permitted where clinical needs of residents require specialized security measures or where residents pose a security threat provided the staff can always readily unlock doors and the building is protected with an approved automatic sprinkler system (FFPC §32.3.2.2.2(6)). Doors in the means of egress permitted to be locked must have provisions for the rapid removal of occupants by means of remote-control locks from within the locked building, keying of all locks to keys always carried by staff, or other reliable means. Only one locking device shall always be permitted (FFPC §32.3.2.2.2 (7)(8)).");
                }
                else
                {
                    this.FindAndReplace(wordApp, "DoorLockI1", "DELETE");
                }

                //PANIC HARDWARE A OR E ----- Need to add E to Program

                if (A1 == true || A2 == true || A3 == true) //Or E is true)
                {
                    this.FindAndReplace(wordApp, "PanicHardwareREQ", "Panic hardware (or fire exit hardware for fire doors) must be installed in all doors serving rooms or spaces with an occupant load of 50 persons or more in a Group A or E occupancy per FBC Section 1010.1.10.  The FFPC, Section 12.2.2.2.3, has a similar requirement for assembly occupancies where the occupancy load is 100 or more. Therefore, the FBC has the more stringent requirement and must be implemented.  Panic hardware must be installed in electrical rooms as stated in other section of this report.");
                }

                //-------------------------------------------Building Occupancy-------------------------------------------------------------
                string buildingTypeA2Floor;
                string buildingTypeA3Floor;
                string buildingTypeBFloor;
                string buildingTypeMFloor;
                string buildingTypeR1Floor;
                string buildingTypeR2Floor;
                string buildingTypeS1Floor;
                string buildingTypeS2Floor;

                string buildingTypeA2SQF;
                string buildingTypeA3SQF;
                string buildingTypeBSQF;
                string buildingTypeMSQF;
                string buildingTypeR1SQF;
                string buildingTypeR2SQF;
                string buildingTypeS1SQF;
                string buildingTypeS2SQF;

                if (A2 == true || A3 == true)
                {
                    //---------A2--------------------------//
                    if (int.Parse(A2HighestFloor.Text) <= 3)
                    {
                        buildingTypeA2Floor = "Type IIB";
                    }
                    else if (int.Parse(A2HighestFloor.Text) == 4)
                    {
                        buildingTypeA2Floor = "Type IIA";
                    }
                    else if ((int.Parse(A2HighestFloor.Text) >= 5 && int.Parse(A2HighestFloor.Text) <= 12))
                    {
                        buildingTypeA2Floor = "Type IB";
                    }
                    else if (int.Parse(A2HighestFloor.Text) > 12)
                    {
                        buildingTypeA2Floor = "Type IA";
                    }
                    //area value
                    if(int.Parse(A2Area.Text) <= 28500)
                    {
                        buildingTypeA2SQF = "Type IIB";
                    }
                    else if(int.Parse(A2Area.Text) > 28500 && int.Parse(A2Area.Text) <= 46500)
                    {
                        buildingTypeA2SQF = "Type IIA";
                    }
                    else if(int.Parse(A2Area.Text) > 46500)
                    {
                        buildingTypeA2SQF = "Type IB";
                    }

                    //-------------A3--------------------//
                    if (int.Parse(A3HighestFloor.Text) <= 3)
                    {
                        buildingTypeA3Floor = "Type IIB";
                    }
                    else if (int.Parse(A3HighestFloor.Text) == 4)
                    {
                        buildingTypeA3Floor = "Type IIA";
                    }
                    else if ((int.Parse(A3HighestFloor.Text) >= 5 && int.Parse(A3HighestFloor.Text) <= 12))
                    {
                        buildingTypeA3Floor = "Type IB";
                    }
                    else if (int.Parse(A3HighestFloor.Text) > 12)
                    {
                        buildingTypeA3Floor = "Type IA";
                    }
                    //Area Value
                    if (int.Parse(A3Area.Text) <= 28500)
                    {
                        buildingTypeA3SQF = "Type IIB";
                    }
                    else if (int.Parse(A3Area.Text) > 28500 && int.Parse(A3Area.Text) <= 46500)
                    {
                        buildingTypeA3SQF = "Type IIA";
                    }
                    else if (int.Parse(A3Area.Text) > 46500)
                    {
                        buildingTypeA3SQF = "Type IB";
                    }


                }

                if (B == true)
                {
                    //What about when Bfloor is below 4?
                    if (int.Parse(BHighestFloor.Text) == 4)
                    {
                        buildingTypeBFloor = "Type IIB";
                    }
                    else if (int.Parse(BHighestFloor.Text) > 4 && int.Parse(BHighestFloor.Text) <= 6)
                    {
                        buildingTypeBFloor = "Type IIA";
                    }
                    else if (int.Parse(BHighestFloor.Text) > 6 && int.Parse(BHighestFloor.Text) <= 12)
                    {
                        buildingTypeBFloor = "Type IB";
                    }
                    else if (int.Parse(BHighestFloor.Text) > 12)
                    {
                        buildingTypeBFloor = "Type IA";
                    }
                    //Area Value
                    if (int.Parse(BArea.Text) <= 69000)
                    {
                        buildingTypeBSQF = "Type IIB";
                    }
                    else if (int.Parse(BArea.Text) > 69000 && int.Parse(BArea.Text) <= 112500)
                    {
                        buildingTypeBSQF = "Type IIA";
                    }
                    else if (int.Parse(BArea.Text) > 112500)
                    {
                        buildingTypeBSQF = "Type IB";
                    }
                }

                if (M == true)
                {
                    //What if it is less than 3?
                    if (int.Parse(MHighestFloor.Text) == 3)
                    {
                        buildingTypeMFloor = "Type IIB";
                    }
                    else if (int.Parse(MHighestFloor.Text) > 3 && int.Parse(MHighestFloor.Text) <= 5)
                    {
                        buildingTypeMFloor = "Type IIA";
                    }
                    else if (int.Parse(MHighestFloor.Text) > 5 && int.Parse(MHighestFloor.Text) <= 12)
                    {
                        buildingTypeMFloor = "Type IB";
                    }
                    else if (int.Parse(MHighestFloor.Text) > 12)
                    {
                        buildingTypeMFloor = "Type IA";
                    }
                    //Area Value
                    if (int.Parse(MArea.Text) <= 37500)
                    {
                        buildingTypeMSQF = "Type IIB";
                    }
                    else if (int.Parse(MArea.Text) > 37500 && int.Parse(MArea.Text) <= 64500)
                    {
                        buildingTypeMSQF = "Type IIA";
                    }
                    else if (int.Parse(MArea.Text) > 64500)
                    {
                        buildingTypeMSQF = "Type IB";
                    }
                }

                if (R1 == true)
                {
                    //What if R1 highest floor is below 5?
                    if (int.Parse(R1HighestFloor.Text) == 5)
                    {
                        buildingTypeR1Floor = "Type IIB";
                    }
                    else if (int.Parse(R1HighestFloor.Text) > 5 && int.Parse(R1HighestFloor.Text) <= 12)
                    {
                        buildingTypeR1Floor = "Type IB";
                    }
                    else if (int.Parse(R1HighestFloor.Text) > 12)
                    {
                        buildingTypeR1Floor = "Type IA";
                    }
                    //Area Value
                    if (int.Parse(R1Area.Text) <= 48000)
                    {
                        buildingTypeR1SQF = "Type IIB";
                    }
                    else if (int.Parse(R1Area.Text) > 48000 && int.Parse(R1Area.Text) <= 72000)
                    {
                        buildingTypeR1SQF = "Type IIA";
                    }
                    else if (int.Parse(R1Area.Text) > 72000)
                    {
                        buildingTypeR1SQF = "Type IB";
                    }
                }

                if (R2 == true)
                {
                    //What about when R2Highest is below 5?
                    if (int.Parse(R2HighestFloor.Text) == 5)
                    {
                        buildingTypeR2Floor = "Type IIB";
                    }
                    else if (int.Parse(R2HighestFloor.Text) > 5 && int.Parse(R2HighestFloor.Text) <= 12)
                    {
                        buildingTypeR2Floor = "Type IB";
                    }
                    else if (int.Parse(R2HighestFloor.Text) > 12)
                    {
                        buildingTypeR2Floor = "Type IA";
                    }
                    //Area Value
                    if (int.Parse(R2Area.Text) <= 48000)
                    {
                        buildingTypeR2SQF = "Type IIB";
                    }
                    else if (int.Parse(R2Area.Text) > 48000 && int.Parse(R2Area.Text) <= 72000)
                    {
                        buildingTypeR2SQF = "Type IIA";
                    }
                    else if (int.Parse(R2Area.Text) > 72000)
                    {
                        buildingTypeR2SQF = "Type IB";
                    }
                }

                if (S1 == true)
                {
                    //What about when S1Highest is below 3?
                    if (int.Parse(S1HighestFloor.Text) == 3)
                    {
                        buildingTypeS1Floor = "Type IIB";
                    }
                    else if (int.Parse(S1HighestFloor.Text) > 3 && int.Parse(S1HighestFloor.Text) <= 5)
                    {
                        buildingTypeS1Floor = "Type IIA";
                    }
                    else if (int.Parse(S1HighestFloor.Text) > 5 && int.Parse(S1HighestFloor.Text) <= 12)
                    {
                        buildingTypeS1Floor = "Type IB";
                    }
                    else if (int.Parse(S1HighestFloor.Text) > 12)
                    {
                        buildingTypeS1Floor = "Type IA";
                    }
                    //Area Value
                    if (int.Parse(S1Area.Text) <= 52500)
                    {
                        buildingTypeS1SQF = "Type IIB";
                    }
                    else if (int.Parse(S1Area.Text) > 52500 && int.Parse(S1Area.Text) <= 78000)
                    {
                        buildingTypeS1SQF = "Type IIA";
                    }
                    else if (int.Parse(S1Area.Text) > 78000 && int.Parse(S1Area.Text) <= 144000)
                    {
                        buildingTypeS1SQF = "Type IB";
                    }
                    else if (int.Parse(S1Area.Text) > 144000)
                    {
                        buildingTypeS1SQF = "Type IA";
                    }
                }

                if (S2 == true)
                {
                    //What about when S2Highest is below 4?
                    if (int.Parse(S2HighestFloor.Text) == 4)
                    {
                        buildingTypeS2Floor = "Type IIB";
                    }
                    else if (int.Parse(S2HighestFloor.Text) > 4 && int.Parse(S2HighestFloor.Text) <= 6)
                    {
                        buildingTypeS2Floor = "Type IIA";
                    }
                    else if (int.Parse(S2HighestFloor.Text) > 6 && int.Parse(S2HighestFloor.Text) <= 12)
                    {
                        buildingTypeS2Floor = "Type IB";
                    }
                    else if (int.Parse(S2HighestFloor.Text) > 12)
                    {
                        buildingTypeS2Floor = "Type IA";
                    }
                    //Area Value
                    if (int.Parse(S2Area.Text) <= 78000)
                    {
                        buildingTypeS2SQF = "Type IIB";
                    }
                    else if (int.Parse(S2Area.Text) > 78000 && int.Parse(S2Area.Text) <= 117000)
                    {
                        buildingTypeS2SQF = "Type IIA";
                    }
                    else if (int.Parse(S2Area.Text) > 117000 && int.Parse(S2Area.Text) <= 237000)
                    {
                        buildingTypeS2SQF = "Type IB";
                    }
                    else if (int.Parse(S2Area.Text) > 237000)
                    {
                        buildingTypeS2SQF = "Type IA";
                    }
                }



                if (intBuildingHeight <= 75)
                {
                    //Set, find and replace BUILDTYPE
                    buildingTypeHeight = "Type IIB";

                    /*
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
                    */

                }
                else if (intBuildingHeight > 75 && intBuildingHeight <= 85)
                {
                    //Set, find and replace BUILDTYPE
                    buildingTypeHeight = "Type IIA";

                    /*
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
                    */

                }
                else if (intBuildingHeight > 85 && intBuildingHeight <= 180)
                {

                    if (isSprinklered.Checked)
                    {
                        //Set, find and replace BUILDTYPE
                        buildingTypeHeight = "Type IIA";

                        /*
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
                        */

                    }
                    else
                    {
                        //Set, find and replace BUILDTYPE
                        buildingTypeHeight = "Type IB";

                        /*
                        this.FindAndReplace(wordApp, "BUILDTYPE", buildingType);

                        //Delete columns that aren't IB
                        Microsoft.Office.Interop.Word.Table table2 = myWordDoc.Tables[2];
                        table2.Columns[5].Delete();
                        table2.Columns[4].Delete();
                        table2.Columns[2].Delete();

                        //Delete the other table
                        Microsoft.Office.Interop.Word.Table table2AndAHalf = myWordDoc.Tables[3];
                        table2AndAHalf.Delete();
                        */

                    }

                }
                else if (intBuildingHeight > 180)
                {

                    //if (intBuildingHeight >= 420)
                    //{
                    //Set, find and replace BUILDTYPE
                    //buildingTypeHeight = "Type IA Reduced";
                    /*
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
                    */

                    //}
                    buildingTypeHeight = "Type IA";
                }

                //Fire Separation Distance North
                FSDFindAndReplace("NFSD", "NFSDRating", "NFSDOpening", int.Parse(NFSDInput.Text), wordApp, NFSDOccupancy.GetItemText(NFSDOccupancy.SelectedItem), buildingType);
                //Fire Separation Distance South
                FSDFindAndReplace("SFSD", "SFSDRating", "SFSDOpening", int.Parse(SFSDInput.Text), wordApp, SFSDOccupancy.GetItemText(SFSDOccupancy.SelectedItem), buildingType);
                //Fire Separation Distance East
                FSDFindAndReplace("EFSD", "EFSDRating", "EFSDOpening", int.Parse(EFSDInput.Text), wordApp, EFSDOccupancy.GetItemText(EFSDOccupancy.SelectedItem), buildingType);
                //Fire Separation Distance West
                FSDFindAndReplace("WFSD", "WFSDRating", "WFSDOpening", int.Parse(WFSDInput.Text), wordApp, WFSDOccupancy.GetItemText(WFSDOccupancy.SelectedItem), buildingType);

                //Insert uploaded pictures
                FindTextAndReplaceImage(wordApp, myWordDoc, "NEWRPIC", NFSDImage.ImageLocation);
                FindTextAndReplaceImage(wordApp, myWordDoc, "SEWRPIC", SFSDImage.ImageLocation);
                FindTextAndReplaceImage(wordApp, myWordDoc, "EEWRPIC", EFSDImage.ImageLocation);
                FindTextAndReplaceImage(wordApp, myWordDoc, "WEWRPIC", WFSDImage.ImageLocation);


                //Start vertical opening section of narrative
                if (VOA = true)
                {
                    if(numVOLevels == 2)
                    {
                        //Assembly spaces shall be permitted to have unprotected vertical openings between any two adjacent floors, provided that such openings are separated from unprotected vertical openings serving other floors by a barrier complying with 8.6.5;
                    }
                    else if(numVOLevels > 2 && numVOLevels <= 4)
                    {
                        /*
                         * Assembly spaces are allowed per NFPA 101 Section 12.3 to implement Convenience Stairways that comply with all of the following:
                            (1) The convenience stair openings shall not serve as a required means of egress.
	                        (2) THe building shall be protected throughout by an approved, supervised automatic sprinkler system in accordance with Section 9.7.
	                        (3) The convenience stair openings shall be protected in accordance with the method detailed for the protection of vertical openings in NFPA 13(Sprinkler-Draft Curtain protection method).
	                        (4) In new construction, the area of the floor opening shall not exceed twice the horizontal projected area of the stairway.
	                        (5) For new construction, such openings shall not connect more than four contiguous stories.
                         */
                    }
                }
                else if(VOR2 == true)
                {
                    if(numVOLevels == 2)
                    {
                        //INCLUDE R-2 SOMEWHERE HERE AND USE 8.6.9.1 CONVENIENT OPENING SECTION
                    }
                    else if(numVOLevels > 2 && numVOLevels <= 3)
                    {
                        //INCLUDE R-2 SOMEWHERE HERE AND USE 8.6.6 COMMUNICATING SPACE SECTION
                    }
                }
                else if(VOR1 == true)
                {
                    if(numVOLevels == 2)
                    {
                        //INCLUDE R-2 SOMEWHERE HERE AND USE 8.6.9.1 CONVENIENT OPENING SECTION
                    }
                }
                else if(VOB == true)
                {
                    if(numVOLevels <= 2)
                    {
                        //INCLUDE BUSINESS SOMEWHERE HERE AND USE 8.6.9.1 CONVENIENT OPENING SECTION
                    }
                    else if(numVOLevels > 2)
                    {
                        /*
                         * Business spaces are allowed per NFPA 101 Section 38.3 to implement Convenience Stairways that comply with all of the following:
	                        (1) The convenience stair openings shall not serve as a required means of egress.
	                        (2) THe building shall be protected throughout by an approved, supervised automatic sprinkler system in accordance with Section 9.7.
	                        (3) The convenience stair openings shall be protected in accordance with the method detailed for the protection of vertical openings in NFPA 13(Sprinkler-Draft Curtain protection method).
	                        (4) In new construction, the area of the floor opening shall not exceed twice the horizontal projected area of the stairway.
	                        (5) For new construction, such openings shall not have a limit of floors in business occupancy.
                         */
                    }
                }
                

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
            listPanel.Add(panel4);
            listPanel[panelIndex].BringToFront();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String imageLocation = "";
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "jpeg files(*.jpg)|*.jpg| PNG files(*.png)|*.png| All files(*.*)|*.*";

                if(dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    imageLocation = dialog.FileName;

                    SFSDImage.ImageLocation = imageLocation;

                }


            }
            catch (Exception)
            {
                MessageBox.Show("An Error Occured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void uploadNFSD_Click(object sender, EventArgs e)
        {
            String imageLocation = "";
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "jpeg files(*.jpg)|*.jpg| PNG files(*.png)|*.png| All files(*.*)|*.*";

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    imageLocation = dialog.FileName;

                    NFSDImage.ImageLocation = imageLocation;

                }


            }
            catch (Exception)
            {
                MessageBox.Show("An Error Occured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void uploadESFD_Click(object sender, EventArgs e)
        {
            String imageLocation = "";
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "jpeg files(*.jpg)|*.jpg| PNG files(*.png)|*.png| All files(*.*)|*.*";

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    imageLocation = dialog.FileName;

                    EFSDImage.ImageLocation = imageLocation;

                }


            }
            catch (Exception)
            {
                MessageBox.Show("An Error Occured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void uploadWFSD_Click(object sender, EventArgs e)
        {
            String imageLocation = "";
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "jpeg files(*.jpg)|*.jpg| PNG files(*.png)|*.png| All files(*.*)|*.*";

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    imageLocation = dialog.FileName;

                    WFSDImage.ImageLocation = imageLocation;

                }


            }
            catch (Exception)
            {
                MessageBox.Show("An Error Occured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
