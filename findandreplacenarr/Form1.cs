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
            //test1
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

                //TABLE 13 EDITING

                Microsoft.Office.Interop.Word.Table table13 = myWordDoc.Tables[13];
                if (A1 == false && A2 == false && A3 == false)
                {
                    table13.Rows[1-6].Delete();
                    table13.Rows[10-12].Delete();
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
                    table13.Rows[7-10].Delete();
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
                    table14.Rows[7].Delete();              ;
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

                //FCC Section Miami Dade and City of Miami

                //FDPTINPUT is the text box
                //if (FDPTName == "City of Miami" || FDPTName == "Miami Dade")
                //{
                //    this.FindAndReplace(wordApp, "MDandCOMFCC", "Miami Dade County and City of Miami Fire Department requires a door opening into the lobby and additional door opening to the outside to provide direct access without entering the lobby.   The fire command center shall be located on the address side/main entrance of the building and shall be within proximity to the fire service access elevators and stairs that have a standpipe available for fire operations.");
                //}

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



                //Building Type Logic

                int A3FloorIn = int.Parse(A3FloorInput.Text);
                int A3AreaIn = int.Parse(A3AreaInput.Text); 

                if (intBuildingHeight <= 75)
                {
                    //Set, find and replace BUILDTYPE
                    buildingType = "Type IIB";
                    this.FindAndReplace(wordApp, "BUILDTYPE", buildingType);

                    /*
                    //occupancy specific floors and SQF Logic --------------------- assuming sprinklered checked -------------------------------------
                    if (A2 == true || A3 == true)
                    {
                        if (A2FloorIn <= 3 || A3FloorIn <= 3)
                        {
                            buildingTypeAFloor = "Type IIB";
                        }
                        else if (A2FloorIn == 4 || A3FloorIn == 4)
                        {
                            buildingTypeAFloor = "Type IIA";
                        }
                        else if ((A2FloorIn >=5 && A2FloorIn <= 12) || (A3FloorIn >=5 && A3FloorIn <= 12))
                        {
                            buildingTypeAFloor = "Type IB";
                        }
                        else if (A2FloorIn > 12 || A3FloorIn > 12)
                        {
                            buildingTypeAFloor = "Type IA";
                        }
                        if (A2AreaIn <= 28500 || A3AreaIn <= 28500)
                        {
                            buildingTypeASQF = "Type IIB";
                        }
                        else if ((A2AreaIn > 28500 && A2AreaIn <= 46500) || (A3AreaIn > 28500 && A3AreaIn <= 46500))
                        {
                            buildingTypeASQF = "Type IIA";
                        }
                        else if (A2AreaIn > 46500 || A3AreaIn > 46500)
                        {
                            buildingTypeASQF = "Type IB";
                        }
                    }

                    if (B == true)
                    {
                        if (BFloor == 4)
                        {
                            buildingTypeBFloor = "Type IIB";
                        }
                        else if (BFloor > 4 && BFloor <= 6)
                        {
                            buildingTypeBFloor = "Type IIA";
                        }
                        else if (BFloor > 6 && BFloor <= 12)
                        {
                            buildingTypeBFloor = "Type IB";
                        }
                        else if (BFloor > 12)
                        {
                            buildingTypeBFloor = "Type IA";
                        }
                        if (BSQF <= 69000)
                        {
                            buildingTypeBSQF = "Type IIB";
                        }
                        else if (BSQF > 69000 && BSQF <= 112500)
                        {
                            buildingTypeBSQF = "Type IIA";
                        }
                        else if (BSQF > 112500)
                        {
                            buildingTypeBSQF = "Type IB";
                        }
                    }

                    if (M == true)
                    {
                        if (MFloor == 3)
                        {
                            buildingTypeMFloor = "Type IIB";
                        }
                        else if (MFloor > 3 && MFloor <= 5)
                        {
                            buildingTypeMFloor = "Type IIA";
                        }
                        else if (MFloor > 5 && MFloor <= 12)
                        {
                            buildingTypeMFloor = "Type IB";
                        }
                        else if (MFloor > 12)
                        {
                            buildingTypeMFloor = "Type IA";
                        }
                        if (MSQF <= 37500)
                        {
                            buildingTypeMSQF = "Type IIB";
                        }
                        else if (MSQF > 37500 && MSQF <= 64500)
                        {
                            buildingTypeMSQF = "Type IIA";
                        }
                        else if (MSQF > 64500)
                        {
                            buildingTypeMSQF = "Type IB";
                        }
                    }

                    if (R1 == true)
                    {
                        if (R1Floor == 5)
                        {
                            buildingTypeR1Floor = "Type IIB";
                        }
                        //else if (R1Floor == 5)
                        //{
                        //    buildingTypeR1Floor = "Type IIA";
                        //}
                        else if (R1Floor > 5 && R1Floor <= 12)
                        {
                            buildingTypeR1Floor = "Type IB";
                        }
                        else if (R1Floor > 12)
                        {
                            buildingTypeR1Floor = "Type IA";
                        }
                        if (R1SQF <= 48000)
                        {
                            buildingTypeR1SQF = "Type IIB";
                        }
                        else if (R1SQF > 48000 && R1SQF <= 72000)
                        {
                            buildingTypeR1SQF = "Type IIA";
                        }
                        else if (R1SQF > 72000)
                        {
                            buildingTypeR1SQF = "Type IB";
                        }
                    }
                    if (R2 == true)
                    {
                        if (R2Floor == 5)
                        {
                            buildingTypeR2Floor = "Type IIB";
                        }
                        //else if (R2Floor == 5)
                        //{
                        //    buildingTypeR2Floor = "Type IIA";
                        //}
                        else if (R2Floor > 5 && R2Floor <= 12)
                        {
                            buildingTypeR2Floor = "Type IB";
                        }
                        else if (R2Floor > 12)
                        {
                            buildingTypeR2Floor = "Type IA";
                        }
                        if (R2SQF <= 48000)
                        {
                            buildingTypeR2SQF = "Type IIB";
                        }
                        else if (R2SQF > 48000 && R2SQF <= 72000)
                        {
                            buildingTypeR2SQF = "Type IIA";
                        }
                        else if (R2SQF > 72000)
                        {
                            buildingTypeR2SQF = "Type IB";
                        }
                    }

                    if (S1 == true)
                    {
                        if (S1Floor == 3)
                        {
                            buildingTypeS1Floor = "Type IIB";
                        }
                        else if (S1Floor > 3 && S1Floor <= 5)
                        {
                            buildingTypeS1Floor = "Type IIA";
                        }
                        else if (S1Floor > 5 && S1Floor <= 12)
                        {
                            buildingTypeS1Floor = "Type IB";
                        }
                        else if (S1Floor > 12)
                        {
                            buildingTypeS1Floor = "Type IA";
                        }
                        if (S1SQF <= 52500)
                        {
                            buildingTypeS1SQF = "Type IIB";
                        }
                        else if (S1SQF > 52500 && S1SQF <= 78000)
                        {
                            buildingTypeS1SQF = "Type IIA";
                        }
                        else if (S1SQF > 78000 && S1SQF <= 144000)
                        {
                            buildingTypeS1SQF = "Type IB";
                        }
                        else if (S1SQF > 144000)
                        {
                            buildingTypeS1SQF = "Type IA";
                        }
                    }

                    if (S2 == true)
                    {
                        if (S2Floor == 4)
                        {
                            buildingTypeS2Floor = "Type IIB";
                        }
                        else if (S2Floor > 4 && S2Floor <= 6)
                        {
                            buildingTypeS2Floor = "Type IIA";
                        }
                        else if (S2Floor > 6 && S2Floor <= 12)
                        {
                            buildingTypeS2Floor = "Type IB";
                        }
                        else if (S2Floor > 12)
                        {
                            buildingTypeS2Floor = "Type IA";
                        }
                        if (S2SQF <= 78000)
                        {
                            buildingTypeS2SQF = "Type IIB";
                        }
                        else if (S2SQF > 78000 && S2SQF <= 117000)
                        {
                            buildingTypeS2SQF = "Type IIA";
                        }
                        else if (S2SQF > 117000 && S2SQF <= 237000)
                        {
                            buildingTypeS2SQF = "Type IB";
                        }
                        else if (S2SQF > 237000)
                        {
                            buildingTypeS2SQF = "Type IA";
                        }
                    }
                    if (buildingTypeHeight == "IA" || buildingType__Floor == "IA" || buildingType__SQF == "IA")
                        {
                            buildingType = "IA";
                        }
                    else if buildingTypeHeight == "IB" || buildingType__Floor == "IB" || buildingType__SQF == "IB")
                        {
                            buildingType = "IB";
                        }
                    else if buildingTypeHeight == "IIA" || buildingType__Floor == "IIA" || buildingType__SQF == "IIA")
                        {
                            buildingType = "IIA";
                        }
                    else if buildingTypeHeight == "IIB" || buildingType__Floor == "IIB" || buildingType__SQF == "IIB")
                        {
                            buildingType = "IIB";
                        }

                    //COMPARE ALL RESULTS AND PICK MOST STRINGENT
                    */

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

        private void isLooped_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
