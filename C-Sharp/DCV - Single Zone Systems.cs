/*
This C# script automates the application of Demand-Controlled Ventilation (DCV) to single-zone HVAC systems in a DesignBuilder/EnergyPlus model, using the ASHRAE Standard 62.1 
Ventilation Rate Procedure (VRP). Instead of manually building minimum outdoor air schedules for each air loop, the script:
1) Extracts ventilation parameters from the IDF — outdoor air rates (Rp, Ra), occupancy density, and the zone's occupancy and fan schedules.
2) Calculates an hourly DCV schedule per the VRP, scaling outdoor airflow to actual occupancy at each hour.
3) Writes the new schedule into the IDF and assigns it as the Minimum Outdoor Air Schedule for each air loop's outdoor air controller.
4) Reuses schedules across air loops with identical parameters, avoiding duplicates.
5) Validates inputs (air loop names, single-zone systems, compatible OA sizing method) and warns the user if something doesn't match.
*/
using DB.Extensibility.Contracts;
using EpNet;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Text.RegularExpressions;
using System.Windows.Forms;


namespace DB.Extensibility.Scripts
{
    public class DCVScheduleScript : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(ApiEnvironment.EnergyPlusInputIdfPath, ApiEnvironment.EnergyPlusInputIddPath);

            //
            // Instructions:

            //    1) Only modify the "airLoops" variable. Enter the exact names of the air loops where DCV should be applied. 
            //       Names must match exactly as shown in DesignBuilder (including spaces, special characters, and capitalization).
            //       Each air loop name must be enclosed in quotation marks (" ").

            //    2) If you want to apply DCV to multiple air loops, separate them with commas inside the list.
            //    Example:
            //    List<string> airLoops = new List<string> { "Air Loop 1", "Air Loop 2" };
             
            //    3) This script is intended for single-zone systems only.
            //


            List<string> airLoops = new List<string> { "Air Loop - RTU MovementActivity", "Air Loop - RTU PlayArea" }; // Add the air loops to apply DCV here :D
            DCVScheduleAssignor assignor = new DCVScheduleAssignor(airLoops, idf);
        }
        public class DCVScheduleAssignor
        {
            public List<string> AirLoopsList { get; private set; }

            public List<SingleZoneSystem> singleZoneSystems = new List<SingleZoneSystem>();
            private IdfReader Idf;

            public DCVScheduleAssignor(List<string> airLoopsList, IdfReader idf)
            {
                this.AirLoopsList = airLoopsList;
                this.Idf = idf;
                CreateDCVSchedule();
            }
            private void CreateDCVSchedule()
            {
                Dictionary<string, IdfObject> airTerminalSingleDuctList = Idf["AirTerminal:SingleDuct:ConstantVolume:NoReheat"].ToDictionary(x => x["Air Inlet Node Name"].Value, x => x);
                Dictionary<string, IdfObject> designSpecificationOutdoorAirList = Idf["DesignSpecification:OutdoorAir"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> peopleList = Idf["People"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> availabilityManagerNightCycleList = Idf["AvailabilityManager:NightCycle"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> controllerOutdoorAirList = Idf["Controller:OutdoorAir"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> AirLoopHVACList = Idf["AirLoopHVAC"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> ZoneSplitterList = Idf["AirLoopHVAC:ZoneSplitter"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> schedules = Idf["Schedule:Compact"].ToDictionary(x => x["Name"].Value, x => x);

                foreach (string airloop in AirLoopsList)
                {
                    VerifyAirLoopName(airloop, AirLoopHVACList);
                    IsSingleZone(airloop, ZoneSplitterList);
                    SingleZoneSystem system = new SingleZoneSystem(airloop, airTerminalSingleDuctList, designSpecificationOutdoorAirList, peopleList, availabilityManagerNightCycleList);
                    string minimumOASchedule = VerifyIfDCVScheduleExists(system);
                    if (minimumOASchedule == null)
                    {
                        Schedule detailedOccupancySchedule = BreakDownScheduleFromIDFObject(schedules[system.OccupancySchedule]);
                        Schedule detailedFanSchedule = BreakDownScheduleFromIDFObject(schedules[system.FanSchedule]);
                        string dcvSchedule1 = @"Schedule:Compact,
{0} DCV Schedule,
Any Number,
Through: 31 Dec,
For:Weekdays,
";
                        string dcvSchedule2 = @"
For: Saturday,
";
                        string dcvSchedule3 = @"
For: Sunday,
";
                        double outdoorAirflowPerAreaDesign = system.Ra + (system.Rp * (1 / system.OccupancyDensity));

                        for (int i = 0; i < 24; i++)
                        {
                            double weekdayValue = (system.Ra + (system.Rp * (1 / system.OccupancyDensity) * detailedOccupancySchedule.WeekdaysValues[i])) * detailedFanSchedule.WeekdaysValues[i] / outdoorAirflowPerAreaDesign;
                            double saturdayValue = (system.Ra + (system.Rp * (1 / system.OccupancyDensity) * detailedOccupancySchedule.SaturdayValues[i])) * detailedFanSchedule.SaturdayValues[i] / outdoorAirflowPerAreaDesign;
                            double sundayValue = (system.Ra + (system.Rp * (1 / system.OccupancyDensity) * detailedOccupancySchedule.SundayValues[i])) * detailedFanSchedule.SundayValues[i] / outdoorAirflowPerAreaDesign;
                            dcvSchedule1 += "Until: " + (i + 1).ToString() + ":00, " + weekdayValue.ToString() + ",\n";
                            dcvSchedule2 += "Until: " + (i + 1).ToString() + ":00, " + saturdayValue.ToString() + ",\n";
                            string delimiter = i < 23 ? "," : ";";
                            dcvSchedule3 += "Until: " + (i + 1).ToString() + ":00, " + sundayValue.ToString() + delimiter + "\n";
                        }
                        string DCVSchedule = String.Format(dcvSchedule1 + dcvSchedule2 + dcvSchedule3, airloop);
                        this.Idf.Load(DCVSchedule);
                        system.MinimimumOAScheduleName = airloop + " DCV Schedule";
                        ModifyFieldMinimumOutdoorAirSchedule(controllerOutdoorAirList, system.MinimimumOAScheduleName, airloop);
                        this.singleZoneSystems.Add(system);
                    }
                    else
                    {
                        system.MinimimumOAScheduleName = minimumOASchedule;
                        ModifyFieldMinimumOutdoorAirSchedule(controllerOutdoorAirList, system.MinimimumOAScheduleName, airloop);
                    }
                }
                this.Idf.Save();
            }

            private string VerifyIfDCVScheduleExists(SingleZoneSystem airloop)
            {
                foreach (SingleZoneSystem airSystem in this.singleZoneSystems)
                {
                    if (airloop.OccupancySchedule.Equals(airSystem.OccupancySchedule))
                    {
                        if (airloop.FanSchedule.Equals(airSystem.FanSchedule))
                        {
                            if (airloop.Ra.Equals(airSystem.Ra))
                            {
                                if (airloop.Rp.Equals(airSystem.Rp))
                                {
                                    if (airloop.OccupancyDensity.Equals(airSystem.OccupancyDensity))
                                    {
                                        return airSystem.MinimimumOAScheduleName;
                                    }
                                }
                            }
                        }
                    }
                }
                return null;
            }

            private Schedule BreakDownScheduleFromIDFObject(IdfObject schedule)
            {
                Schedule detailedSchedule = new Schedule();
                int startHour = 0;
                int endHour = 0;
                bool weekdaysStatus = false;
                bool saturdayStatus = false;
                bool sundayStatus = false;


                for (int i = 0; i < schedule.Count; i++)
                {
                    if (schedule[i].Value.Contains("For"))
                    {
                        weekdaysStatus = false;
                        saturdayStatus = false;
                        sundayStatus = false;
                        startHour = 0;
                        if (schedule[i].Value.Contains("Weekdays"))
                        {
                            weekdaysStatus = true;
                        }
                        if (schedule[i].Value.Contains("Saturday"))
                        {
                            saturdayStatus = true;
                        }
                        if (schedule[i].Value.Contains("Sunday"))
                        {
                            sundayStatus = true;
                        }
                        if (schedule[i].Value.Contains("Weekends"))
                        {
                            saturdayStatus = true;
                            sundayStatus = true;
                        }
                        if (schedule[i].Value.Contains("AllOtherDays"))
                        {
                            if (detailedSchedule.SaturdayValues.Count == 0)
                            {
                                saturdayStatus = true;
                            }
                            if (detailedSchedule.SundayValues.Count == 0)
                            {
                                sundayStatus = true;
                            }
                        }
                    }
                    else if (schedule[i].Value.Contains("Until"))
                    {
                        endHour = ExtractTime(schedule[i].Value);
                        double fractionValue = Double.Parse(schedule[i + 1].Value);
                        for (int j = startHour; j < endHour; j++)
                        {
                            if (weekdaysStatus)
                            {
                                detailedSchedule.AddWeekdayValue(fractionValue);
                            }
                            if (saturdayStatus)
                            {
                                detailedSchedule.AddSaturdayValue(fractionValue);
                            }
                            if (sundayStatus)
                            {
                                detailedSchedule.AddSundayValue(fractionValue);
                            }
                        }
                        startHour = endHour;
                        i++;
                    }
                }
                return detailedSchedule;
            }
            private int ExtractTime(string timeString)
            {
                var match = Regex.Match(timeString, @"Until:\s*(\d{1,2})");

                if (match.Success)
                {
                    return int.Parse(match.Groups[1].Value);
                }
                return -1;
            }
            private void ModifyFieldMinimumOutdoorAirSchedule(Dictionary<string, IdfObject> controllerOutdoorAirList, string nameSchedule, string nameAirLoop)
            {
                IdfObject controllerOudoorAirObject = controllerOutdoorAirList[nameAirLoop + " AHU Outdoor Air Controller"];
                controllerOudoorAirObject["Minimum Outdoor Air Schedule Name"].Value = nameSchedule;
            }
            private void VerifyAirLoopName(string airLoopName, Dictionary<string, IdfObject> airLoopHVACList)
            {
                if (!airLoopHVACList.ContainsKey(airLoopName))
                {
                    MessageBox.Show(string.Format("\"{0}\" was not found. Please check the name.", airLoopName));
                }
            }
            private void IsSingleZone(string airLoopName, Dictionary<string, IdfObject> zoneSplitterList)
            {
                IdfObject zoneSplitter = zoneSplitterList[airLoopName + " Zone Splitter"];
                if (zoneSplitter.Count > 3)
                {
                    MessageBox.Show(string.Format("\"{0}\" is not single-zone. This script supports single-zone systems only.", airLoopName));
                }
            }
        }
        public class SingleZoneSystem
        {
            public string AirLoopName { get; private set; }
            public string ZoneName { get; private set; }
            public double Rp { get; private set; }
            public double Ra { get; private set; }
            public double OccupancyDensity { get; private set; }
            public string OccupancySchedule { get; private set; }
            public string FanSchedule { get; private set; }
            public string MinimimumOAScheduleName { get; set; }

            private Dictionary<string, IdfObject> AirTerminalSingleDuctList;
            private Dictionary<string, IdfObject> DesignSpecificationOutdoorAirList;
            private Dictionary<string, IdfObject> PeopleList;
            private Dictionary<string, IdfObject> AvailabilityManagerNightCycleList;
            public SingleZoneSystem(string airLoopName, Dictionary<string, IdfObject> airTerminalSingleDuctList, Dictionary<string, IdfObject> designSpecificationOutdoorAirList, Dictionary<string, IdfObject> peopleList, Dictionary<string, IdfObject> availabilityManagerNightCycleList)
            {
                this.AirLoopName = airLoopName;
                AirTerminalSingleDuctList = airTerminalSingleDuctList;
                this.AirTerminalSingleDuctList = airTerminalSingleDuctList;
                ExtractZoneName();
                DesignSpecificationOutdoorAirList = designSpecificationOutdoorAirList;
                PeopleList = peopleList;
                AvailabilityManagerNightCycleList = availabilityManagerNightCycleList;
                ExtractOutdoorAirValues();
                ExtractOccupancyData();
                ExtractFanSchedule();
                VerifyVentilationInformation();
            }
            private void ExtractZoneName()
            {
                this.ZoneName = AirTerminalSingleDuctList[this.AirLoopName + " Zone Splitter Outlet Node 1"]["Name"].Value.Replace(" Single Duct CAV No Reheat", "").Trim();
            }
            private void ExtractOutdoorAirValues()
            {
                if (DesignSpecificationOutdoorAirList[this.ZoneName + " Design Specification Outdoor Air Object"]["Outdoor Air Method"].Value == "Sum")
                {
                    this.Ra = DesignSpecificationOutdoorAirList[this.ZoneName + " Design Specification Outdoor Air Object"]["Outdoor Air Flow per Zone Floor Area"].Number;
                    this.Rp = DesignSpecificationOutdoorAirList[this.ZoneName + " Design Specification Outdoor Air Object"]["Outdoor Air Flow per Person"].Number;
                }
                else
                {
                    this.Ra = 0;
                    this.Rp = 0;
                    MessageBox.Show(string.Format("The Outdoor Air Method for zone \"{0}\" is not valid. It should be set to \"Sum\" or \"Flow per Person\".", this.ZoneName));
                }
            }
            private void ExtractOccupancyData()
            {
                if (PeopleList["People " + this.ZoneName]["Number of People Calculation Method"].Value == "Area/Person")
                {
                    this.OccupancyDensity = PeopleList["People " + this.ZoneName]["Floor Area per Person"].Number;
                    this.OccupancySchedule = PeopleList["People " + this.ZoneName]["Number of People Schedule Name"].Value;
                }
                else
                {
                    this.OccupancyDensity = 0;
                    this.OccupancySchedule = PeopleList["People " + this.ZoneName]["Number of People Schedule Name"].Value;
                    MessageBox.Show(string.Format("The zone \"{0}\" must use \"Floor Area per Person\" as the occupancy method for this script to run.", this.ZoneName));
                }
            }

            private void ExtractFanSchedule()
            {
                this.FanSchedule = AvailabilityManagerNightCycleList[this.AirLoopName + " AHU Night Cycle Operation"]["Fan Schedule Name"].Value;
            }
            private void VerifyVentilationInformation()
            {
                if (this.Rp == 0)
                {
                    MessageBox.Show(string.Format("Zone \"{0}\" has an Outdoor Air Rate per Person (Rp) of zero. DCV is not applicable to this zone.", this.ZoneName));
                    if (this.Ra == 0)
                    {
                        MessageBox.Show(string.Format("Zone \"{0}\" cannot have both Outdoor Air Rate per Person and Outdoor Air Rate per Area set to zero.", this.ZoneName));
                    }
                }
                if (this.OccupancyDensity == 0)
                {
                    MessageBox.Show(string.Format("Zone \"{0}\" has an Occupancy Density of zero. DCV is not applicable to this zone.", this.ZoneName));
                }
            }
        }
        public class Schedule
        {
            public List<double> WeekdaysValues;
            public List<double> SaturdayValues;
            public List<double> SundayValues;
            public Schedule()
            {
                this.WeekdaysValues = new List<double>();
                this.SaturdayValues = new List<double>();
                this.SundayValues = new List<double>();
            }
            public void AddWeekdayValue(double fractionValue)
            {
                WeekdaysValues.Add(fractionValue);
            }
            public void AddSaturdayValue(double fractionValue)
            {
                SaturdayValues.Add(fractionValue);
            }
            public void AddSundayValue(double fractionValue)
            {
                SundayValues.Add(fractionValue);
            }
        }
    }
}
