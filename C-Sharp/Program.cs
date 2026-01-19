using System;
using System.Runtime;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DB.Extensibility.Contracts;
using EpNet;
using EpNet.Results;
using System.Text.RegularExpressions;


namespace DB.Extensibility.Scripts
{
    public class NECBSystem3Generator : ScriptBase, IScript
    {
        public class GlobalVariables
        {
            public const string ClimateZone = "7";   // Climate zone (NECB): e.g., "5", "6", "7", "8"
            public const string ProjectName = "Shoppers";  // Project name (same as in DesignBuilder)
            public const string ProjectPath = @"C:\Antonio\Energy Modelling\DesignBuilder\EpNet\Script - NECB System 3\Test Files"; // Folder path where the report CSV will be saved
            public static readonly string ResultsSummaryPath = Path.Combine(ProjectPath, ProjectName + " - AirLoops Report.csv");
        }
        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(ApiEnvironment.EnergyPlusInputIdfPath, ApiEnvironment.EnergyPlusInputIddPath);
            HVACSystemsExtractor systemsFinder = new HVACSystemsExtractor();
            List<HVACSystem> hvacSystemList = systemsFinder.GetHVACSystemFullInfo(idf);
            NECBSystem3 newNECBSystem3 = new NECBSystem3(hvacSystemList);
            newNECBSystem3.ModifyAirLoops(idf);
            newNECBSystem3.CreateHotWaterPlant(idf);
            idf.Save();
        }
        public override void AfterEnergySimulation()
        {
            IdfReader idf = new IdfReader(ApiEnvironment.EnergyPlusInputIdfPath, ApiEnvironment.EnergyPlusInputIddPath);
            HVACSystemsExtractor systemsFinder = new HVACSystemsExtractor();
            List<HVACSystem> hvacSystemList = systemsFinder.GetHVACSystemFullInfo(idf);
            CSVResultsGenerator resultsGenerator = new CSVResultsGenerator();
            resultsGenerator.GenerateResultsCSV(hvacSystemList);
        }
        public class HVACSystemsExtractor
        {
            private bool IsCSVResultsFilePresent = false;

            public string[] ImportCSVResults()
            {
                if (File.Exists(GlobalVariables.ResultsSummaryPath))
                {
                    string[] lines = File.ReadAllLines(GlobalVariables.ResultsSummaryPath);
                    IsCSVResultsFilePresent = true;
                    return lines;
                }
                return Array.Empty<string>();
            }
            public string GetAirLoopName(IdfObject singleDuctCAVNoReheatObject)
            {
                string patron = @"\s*Zone Splitter Outlet Node\s*\d*";
                string airLoopName = Regex.Replace(singleDuctCAVNoReheatObject["Air Inlet Node Name"].Value, patron, "").Trim();
                return airLoopName;
            }
            public string GetZoneName(IdfObject singleDuctCAVNoReheatObject)
            {
                string zoneName = singleDuctCAVNoReheatObject["Air Outlet Node Name"].Value.Replace(" Single Duct CAV No Reheat Supply Outlet", "").Trim();
                return zoneName;
            }

            public Dictionary<string, List<string>> GetZonesPerAirLoop(IdfReader idfFile)
            {
                Dictionary<string, List<string>> zonesPerAirLoop = new Dictionary<string, List<string>>();
                List<IdfObject> singleDuctCAVNoReheatObjectsList =
                    idfFile["AirTerminal:SingleDuct:ConstantVolume:NoReheat"]
                      .OrderBy(obj => obj["Name"].Value, StringComparer.OrdinalIgnoreCase)
                      .ToList();

                foreach (IdfObject objetoIDF in singleDuctCAVNoReheatObjectsList)
                {
                    string airLoopName = GetAirLoopName(objetoIDF);
                    string zoneName = GetZoneName(objetoIDF);

                    if (!zonesPerAirLoop.ContainsKey(airLoopName))
                    {
                        zonesPerAirLoop[airLoopName] = new List<string>();
                    }
                    zonesPerAirLoop[airLoopName].Add(zoneName);
                }
                return zonesPerAirLoop;
            }
            public void ExtractDataFromCSVResultsFile(int index, string[] allLinesCSVFile, out string coolingEfficiencyMetric, out double coolingCapacity, out bool economizerRequired, out bool ervRequired)
            {
                if (!this.IsCSVResultsFilePresent)
                {
                    coolingEfficiencyMetric = "SEER";
                    coolingCapacity = 0;
                    economizerRequired = false;
                    ervRequired = false;
                }
                else
                {
                    string line = allLinesCSVFile[index];
                    string[] columns = line.Split(',');
                    coolingEfficiencyMetric = columns[6];
                    coolingCapacity = Double.Parse(columns[1]);
                    economizerRequired = columns[4].Equals("TRUE", StringComparison.OrdinalIgnoreCase);
                    ervRequired = columns[5].Equals("TRUE", StringComparison.OrdinalIgnoreCase);
                }
            }
            public List<HVACSystem> GetHVACSystemFullInfo(IdfReader idfFile)
            {
                string[] linesCVSResultsFile = ImportCSVResults();
                List<HVACSystem> airSystemsList = new List<HVACSystem>();
                Dictionary<string, List<string>> zonesPerAirLoop = GetZonesPerAirLoop(idfFile);

                int i = 0;
                foreach (string airloop in zonesPerAirLoop.Keys)
                {
                    string airLoopName = airloop;
                    List<string> zonesName = zonesPerAirLoop[airloop];
                    string coolingEfficiencyMetric;
                    double coolingCapacity;
                    bool economizerRequired, ervRequired;
                    ExtractDataFromCSVResultsFile(i + 1, linesCVSResultsFile, out coolingEfficiencyMetric, out coolingCapacity, out economizerRequired, out ervRequired); //i+1: avoid taking table headers in CSV file
                    airSystemsList.Add(new HVACSystem(airLoopName, zonesName, coolingEfficiencyMetric, coolingCapacity, economizerRequired, ervRequired));
                    i++;
                }
                return airSystemsList;
            }
        }
        public class HVACSystem
        {
            public string AirLoopName { get; set; }
            public List<string> ZoneNames { get; set; }
            public string CoolingEfficiencyMetric { get; set; }
            public double CoolingCapacity { get; set; }
            public bool EconomizerRequired { get; set; }
            public bool ERVRequired { get; set; }
            public string FanSchedule { get; set; }

            public HVACSystem(string airLoopName, List<string> zoneNames, string coolingEfficiencyMetric, double coolingCapacity, bool economizerRequired, bool ervRequired)
            {
                AirLoopName = airLoopName;
                ZoneNames = zoneNames;
                CoolingEfficiencyMetric = coolingEfficiencyMetric;
                CoolingCapacity = coolingCapacity;
                EconomizerRequired = economizerRequired;
                ERVRequired = ervRequired;
                FanSchedule = "On 24/7";
            }

            public void UpdateFan(Dictionary<string, IdfObject> FanIdfObjectsDict)
            {
                string fanName = this.AirLoopName + " AHU Supply Fan";
                IdfObject fanIdfObject = FanIdfObjectsDict[fanName];
                fanIdfObject["Fan Total Efficiency"].Number = 0.4;
                fanIdfObject["Pressure Rise"].Number = 640;
                fanIdfObject["Motor Efficiency"].Number = 1;
            }
            public void UpdateDXCoolingCoil(Dictionary<string, IdfObject> DXCoilIdfObjectsDict)
            {
                IdfObject DXCoilIdfObject = DXCoilIdfObjectsDict[this.AirLoopName + " AHU Unitary Cooling Coil DX Cooling Coil"];
                DXCoilIdfObject["Total Cooling Capacity Function of Temperature Curve Name"].Value = "NECB 2020 - DXClgCoilTotalClgCapFuncTemperature";
                DXCoilIdfObject["Energy Input Ratio Function of Temperature Curve Name"].Value = "NECB 2020 - DXClgCoilEnergyInputRatioFuncTemperature";
                if (this.CoolingEfficiencyMetric == "SEER")
                {
                    DXCoilIdfObject["Gross Rated Cooling COP"].Number = 4.306107953835;
                    DXCoilIdfObject["2017 Rated Evaporator Fan Power Per Volume Flow Rate"].Number = 131.499549687332;
                }
                else
                {
                    if (this.CoolingCapacity >= 19000 && this.CoolingCapacity < 40000)
                    {
                        DXCoilIdfObject["Gross Rated Cooling COP"].Number = 3.22378177201532;
                        DXCoilIdfObject["2017 Rated Evaporator Fan Power Per Volume Flow Rate"].Number = 16.7780000025444;
                    }
                    else if (this.CoolingCapacity >= 40000 && this.CoolingCapacity < 70000)
                    {
                        DXCoilIdfObject["Gross Rated Cooling COP"].Number = 3.16516755797868;
                        DXCoilIdfObject["2017 Rated Evaporator Fan Power Per Volume Flow Rate"].Number = 63.5430000024894;
                    }
                    else if (this.CoolingCapacity >= 70000 && this.CoolingCapacity < 223000)
                    {
                        DXCoilIdfObject["Gross Rated Cooling COP"].Number = 2.87209648779547;
                        DXCoilIdfObject["2017 Rated Evaporator Fan Power Per Volume Flow Rate"].Number = 1.41800000254678;
                    }
                    else
                    {
                        DXCoilIdfObject["Gross Rated Cooling COP"].Number = 2.78417516674051;
                        DXCoilIdfObject["2017 Rated Evaporator Fan Power Per Volume Flow Rate"].Number = 12.2850000025455;
                    }
                }
            }
            public void UpdateZoneHVACEquipmentList(Dictionary<string, IdfObject> ZoneHVACEquipmentListDict)
            {
                foreach (string zoneName in this.ZoneNames)
                {
                    IdfObject zoneHVACEquipmentListObject = ZoneHVACEquipmentListDict[zoneName + " Equipment"];
                    zoneHVACEquipmentListObject["Zone Equipment 2 Object Type"].Value = "ZoneHVAC:Baseboard:Convective:Water";
                    zoneHVACEquipmentListObject["Zone Equipment 2 Name"].Value = zoneName + " Water Convector";
                    zoneHVACEquipmentListObject["Zone Equipment 2 Cooling Sequence"].Value = "2";
                    zoneHVACEquipmentListObject["Zone Equipment 2 Heating or No-Load Sequence"].Value = "2";
                    zoneHVACEquipmentListObject["Zone Equipment 2 Sequential Cooling Fraction Schedule Name"].Value = "Cooling fraction schedule";
                    zoneHVACEquipmentListObject["Zone Equipment 2 Sequential Heating Fraction Schedule Name"].Value = "Heating fraction schedule";
                }
            }
            public void UpdateControllerOutdoorAir(Dictionary<string, IdfObject> ControllerOutdoorAirDict)
            {
                IdfObject controllerOutdoorAirObject = ControllerOutdoorAirDict[this.AirLoopName + " AHU Outdoor Air Controller"];
                if (this.EconomizerRequired)
                {
                    controllerOutdoorAirObject["Economizer Control Type"].Value = "FixedDryBulb";
                    controllerOutdoorAirObject["Economizer Maximum Limit Dry-Bulb Temperature"].Number = 23.88889;
                    controllerOutdoorAirObject["Lockout Type"].Value = "LockoutWithHeating";
                }
                else
                {
                    controllerOutdoorAirObject["Economizer Control Type"].Value = "NoEconomizer";
                }
            }
            public void UpdateOutdoorAirSystemEquipmentList(Dictionary<string, IdfObject> OutdoorAirSystemEquipmentListDict)
            {
                IdfObject OutdoorAirSystemEquipmentListObject = OutdoorAirSystemEquipmentListDict[this.AirLoopName + " AHU Outdoor air Equipment List"];
                OutdoorAirSystemEquipmentListObject.InsertFields(1, "HeatExchanger:AirToAir:SensibleAndLatent", this.AirLoopName + " AHU Heat Recovery Device");
            }
            public void UpdateAvailabilityManagerAssignmentList(Dictionary<string, IdfObject> AvailabilityManagerAssignmentListDict)
            {
                IdfObject AvailabilityManagerAssignmentListObject = AvailabilityManagerAssignmentListDict[this.AirLoopName + " AvailabilityManager List"];
                AvailabilityManagerAssignmentListObject["Availability Manager 1 Object Type"].Value = "AvailabilityManager:Scheduled";
                AvailabilityManagerAssignmentListObject["Availability Manager 1 Name"].Value = this.AirLoopName + " Availability";
            }
            public void UpdateOutdoorAirMixer(Dictionary<string, IdfObject> OutdoorAirMixerDict)
            {
                IdfObject OutdoorAirMixerObject = OutdoorAirMixerDict[this.AirLoopName + " AHU Outdoor Air Mixer"];
                OutdoorAirMixerObject["Outdoor Air Stream Node Name"].Value = this.AirLoopName + " AHU Heat Recovery Supply Outlet";
            }
            public void UpdateAvailabilityManagerNightCycleAndFanSchedule(Dictionary<string, IdfObject> NightCycleDict)
            {
                IdfObject NightCycleObject = NightCycleDict[this.AirLoopName + " AHU Night Cycle Operation"];
                NightCycleObject["Applicability Schedule Name"].Value = "Off 24/7";
                this.FanSchedule = NightCycleObject["Fan Schedule Name"].Value;
            }

            public string AddIDFObjects()
            {
                string baseboardSchedule = this.FanSchedule.Substring(0, this.FanSchedule.IndexOf("_Fans")) + "_Baseboards";

                string IDFObjects = @"AvailabilityManager:Scheduled,
{0} Availability,                             !- Name
{1};                                          !- Schedule Name
";
                if (this.ERVRequired)
                {
                    IDFObjects += @"HeatExchanger:AirToAir:SensibleAndLatent,
{0} AHU Heat Recovery Device,            !- Name            
On 24/7,                                 !- Availability Schedule Name
autosize,                                !- Nominal Supply Air Flow Rate
0.5,                                     !- Sensible Effectiveness at 100% Heating Air Flow
0,                                       !- Latent Effectiveness at 100% Heating Air Flow
0.5,                                     !- Sensible Effectiveness at 75% Heating Air Flow
0,                                       !- Latent Effectiveness at 75% Heating Air Flow
0.5,                                     !- Sensible Effectiveness at 100% Cooling Air Flow
0,                                       !- Latent Effectiveness at 100% Cooling Air Flow
0.5,                                     !- Sensible Effectiveness at 75% Cooling Air Flow
0,                                       !- Latent Effectiveness at 100% Cooling Air Flow
{0} AHU Outdoor Air Inlet,               !- Supply Air Inlet Node Name
{0} AHU Heat Recovery Supply Outlet,     !- Supply Air Outlet Node Name
{0} AHU Relief Air Outlet,               !- Exhaust Air Inlet Node Name
{0} AHU Heat Recovery Relief Outlet,     !- Exhaust Air Outlet Node Name
0,                                       !- Nominal Electric Power
No,                                      !- Supply Air Outlet Temperature Control
Plate,                                   !- Heat Exchanger Type
None,                                    !- Frost Control Type
1.7,                                     !- Threshold Temperature
0.083,                                   !- Initial Defrost Time Fraction
0.012,                                   !- Rate of Defrost Time Fraction
Yes;                                     !- Economizer Lockout
";
                }
                string baseboardIDFObjects = "";
                foreach (string zone in this.ZoneNames)
                {
                    baseboardIDFObjects += String.Format(@"ZoneHVAC:Baseboard:Convective:Water,
{0} Water Convector,                          !- Name
{1},                                          !- Availability Schedule Name
{0} Water Convector Hot Water Inlet Node,     !- Inlet Node Name
{0} Water Convector Hot Water Outlet Node,    !- Outlet Node Name
HeatingDesignCapacity,                        !- Heating Design Capacity Method
autosize,                                     !- Heating Design Capacity
,                                             !- Heating Design Capacity Per Floor Area
,                                             !- Fraction of Autosized Heating Design Capacity
autosize,                                     !- U-Factor Times Area Value
autosize,                                     !- Maximum Water Flow Rate
0.001;                                        !- Convergence Tolerance
", zone, baseboardSchedule);
                }
                return String.Format(IDFObjects + baseboardIDFObjects, this.AirLoopName, this.FanSchedule);
            }
        }

        public class AirSystemsManager
        {
            private List<HVACSystem> HVACSystems { get; set; }
            public AirSystemsManager(List<HVACSystem> hvacSystems)
            {
                HVACSystems = hvacSystems;
            }

            public void ModifyIDFObjects(IdfReader idfFile)
            {
                Dictionary<string, IdfObject> dxCoolingCoilsDict = idfFile["Coil:Cooling:DX:SingleSpeed"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> fansDict = idfFile["Fan:ConstantVolume"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> zoneHVACEquipmentListDict = idfFile["ZoneHVAC:EquipmentList"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> controllerOutdoorAirDict = idfFile["Controller:OutdoorAir"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> outdoorAirSystemEquipmentListDict = idfFile["AirLoopHVAC:OutdoorAirSystem:EquipmentList"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> availabilityManagerAssignmentListDict = idfFile["AvailabilityManagerAssignmentList"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> outdoorAirMixerDict = idfFile["OutdoorAir:Mixer"].ToDictionary(x => x["Name"].Value, x => x);
                Dictionary<string, IdfObject> nightCycleDict = idfFile["AvailabilityManager:NightCycle"].ToDictionary(x => x["Name"].Value, x => x);
                foreach (HVACSystem airSystem in this.HVACSystems)
                {
                    airSystem.UpdateFan(fansDict);
                    airSystem.UpdateDXCoolingCoil(dxCoolingCoilsDict);
                    airSystem.UpdateZoneHVACEquipmentList(zoneHVACEquipmentListDict);
                    airSystem.UpdateControllerOutdoorAir(controllerOutdoorAirDict);
                    airSystem.UpdateAvailabilityManagerAssignmentList(availabilityManagerAssignmentListDict);
                    airSystem.UpdateAvailabilityManagerNightCycleAndFanSchedule(nightCycleDict);
                    if (airSystem.ERVRequired)
                    {
                        airSystem.UpdateOutdoorAirSystemEquipmentList(outdoorAirSystemEquipmentListDict);
                        airSystem.UpdateOutdoorAirMixer(outdoorAirMixerDict);
                    }
                }
            }
            public void AddNewObjectsToIdf(IdfReader idfFile)
            {
                string newObjectsToIdf = @"Curve:Biquadratic,
NECB 2020 - DXClgCoilTotalClgCapFuncTemperature,    !-Name
0.86790540000,                                      !- Coefficient1 Constant
0.01424592000,                                      !- Coefficient2 x
0.00055436400,                                      !- Coefficient3 x**2
-0.00755748000,                                     !- Coefficient4 y
0.00003304800,                                      !- Coefficient5 y**2
-0.00019180800,                                     !- Coefficient6 x*y
12.77778,                                           !- Minimum Value of x
23.88889,                                           !- Maximum Value of x
18,                                                 !- Minimum Value of y
46.11111,                                           !- Maximum Value of y
,                                                   !- Minimum Curve Output
,                                                   !- Maximum Curve Output
Temperature,                                        !- Input Unit Type for X
Temperature,                                        !- Input Unit Type for Y
Dimensionless;                                      !- Output Unit Type

Curve:Biquadratic,
NECB 2020 - DXClgCoilEnergyInputRatioFuncTemperature,    !-Name
0.1169362000,                                            !- Coefficient1 Constant
0.0284932800,                                            !- Coefficient2 x
-0.0004111560,                                           !- Coefficient3 x**2
0.0214108200,                                            !- Coefficient4 y
0.0001610280,                                            !- Coefficient5 y**2
-0.0006791040,                                           !- Coefficient6 x*y
12.77778,                                                !- Minimum Value of x
23.88889,                                                !- Maximum Value of x
18,                                                      !- Minimum Value of y
46.11111,                                                !- Maximum Value of y
,                                                        !- Minimum Curve Output
,                                                        !- Maximum Curve Output
Temperature,                                             !- Input Unit Type for X
Temperature,                                             !- Input Unit Type for Y
Dimensionless;                                           !- Output Unit Type

Schedule:Compact,
NECB Schedule A_Baseboards,
Any Number,
Through: 31 Dec,
For: Weekdays,
Until: 06:00,
1,
Until: 21:00,
0,
Until: 24:00,
1,
For: SummerDesignDay WinterDesignDay,
Until: 24:00,
0,
For: AllOtherDays,
Until: 24:00,
1;

Schedule:Compact,
NECB Schedule B_Baseboards,
Any Number,
Through: 31 Dec,
For: Weekdays Saturday,
Until: 01:00,
0,
Until: 08:00,
1,
Until: 24:00,
0,
For:Sunday,
Until: 01:00,
0,
Until: 09:00,
1,
Until: 23:00,
0,
Until: 24:00,
1,
For: SummerDesignDay WinterDesignDay,
Until: 24:00,
0;

Schedule:Compact,
NECB Schedule C_Baseboards,
Any Number,
Through: 31 Dec,
For: Weekdays Saturday,
Until: 07:00,
1,
Until: 21:00,
0,
Until: 24:00,
1,
For: Sunday,
Until: 07:00,
1,
Until: 19:00,
0,
Until: 24:00,
1,
For: SummerDesignDay WinterDesignDay,
Until: 24:00,
0;

Schedule:Compact,
NECB Schedule D_Baseboards,
Any Number,
Through: 31 Dec,
For: Weekdays Saturday,
Until: 07:00,
1,
Until: 23:00,
0,
Until: 24:00,
1,
For: SummerDesignDay WinterDesignDay,
Until: 24:00,
0,
For: AllOtherDays,
Until: 24:00,
1;

Schedule:Compact,
NECB Schedule E_Baseboards,
Any Number,
Through: 31 Dec,
For: Weekdays,
Until: 07:00,
1,
Until: 19:00,
0,
Until: 24:00,
1,
For: Saturday,
Until: 08:00,
1,
Until: 17:00,
0,
Until: 24:00,
1,
For: SummerDesignDay WinterDesignDay,
Until: 24:00,
0,
For: AllOtherDays,
Until: 24:00,
1;
";
                foreach (HVACSystem airSystem in this.HVACSystems)
                {
                    newObjectsToIdf += airSystem.AddIDFObjects();
                }
                idfFile.Load(newObjectsToIdf);
            }
        }

        public class HotWaterPlantManager
        {
            private List<string> ZoneNames;
            private int ZonesNumber { get; set; }
            private List<HVACSystem> HVACSystems { get; set; }
            public HotWaterPlantManager(List<HVACSystem> hvacSystems)
            {
                HVACSystems = hvacSystems;
                ZoneNames = new List<string>();
                foreach (HVACSystem hvacSystem in this.HVACSystems)
                {
                    foreach (string zoneName in hvacSystem.ZoneNames)
                    {
                        ZoneNames.Add(zoneName);
                    }
                }
                this.ZonesNumber = ZoneNames.Count;

            }
            public void AddNewObjectsToIdf(IdfReader idfFile)
            {
                string severalObjectsToIdf = @"Sizing:Plant,
HW Loop,     !- Plant or Condenser Loop Name
Heating,     !- Loop Type
82,          !- Design Loop Exit Temperature
16;          !- Loop Design Temperature Difference

Branch,
HW Loop Demand Side Inlet Branch,
,
Pipe:Adiabatic,
HW Loop Demand Side Inlet Branch Pipe,
HW Loop Demand Side Inlet,
HW Loop Demand Side Inlet Branch Pipe Outlet;

Branch,
HW Loop Demand Side Bypass Branch,
,
Pipe:Adiabatic,
HW Loop Demand Side Bypass Pipe,
HW Loop Demand Side Bypass Pipe Inlet Node,
HW Loop Demand Side Bypass Pipe Outlet Node;

Branch,
HW Loop Demand Side Outlet Branch,
,
Pipe:Adiabatic,
HW Loop Demand Side Outlet Branch Pipe,
HW Loop Demand Side Outlet Branch Pipe Inlet,
HW Loop Demand Side Outlet;

Branch,
HW Loop Supply Side Inlet Branch,
,
Pump:ConstantSpeed,
HW Loop Supply Pump,
HW Loop Supply Side Inlet,
HW Loop Supply Pump Water Outlet Node;

Branch,
HW Loop Supply Side Bypass Branch,
,
Pipe:Adiabatic,
HW Loop Supply Side Bypass Pipe,
HW Loop Supply Side Bypass Pipe Inlet Node,
HW Loop Supply Side Bypass Pipe Outlet Node;

Branch,
Boiler HW Loop Supply Side Branch,
,
Boiler:HotWater,
Boiler,
Boiler Water Inlet Node,
Boiler Water Outlet Node;

Branch,
HW Loop Supply Side Outlet Branch,
,
Pipe:Adiabatic,
HW Loop Supply Side Outlet Branch Pipe,
HW Loop Supply Side Outlet Branch Pipe Inlet,
HW Loop Supply Side Outlet;

BranchList,
HW Loop Supply Side Branches,
HW Loop Supply Side Inlet Branch,
HW Loop Supply Side Bypass Branch,
Boiler HW Loop Supply Side Branch,
HW Loop Supply Side Outlet Branch;

Connector:Splitter,
HW Loop Supply Splitter,
HW Loop Supply Side Inlet Branch,
HW Loop Supply Side Bypass Branch,
Boiler HW Loop Supply Side Branch;

Connector:Mixer,
HW Loop Supply Mixer,
HW Loop Supply Side Outlet Branch,
Boiler HW Loop Supply Side Branch,
HW Loop Supply Side Bypass Branch;

ConnectorList,
HW Loop Demand Side Connectors,
Connector:Splitter,
HW Loop Demand Splitter,
Connector:Mixer,
HW Loop Demand Mixer;

ConnectorList,
HW Loop Supply Side Connectors,
Connector:Splitter,
HW Loop Supply Splitter,
Connector:Mixer,
HW Loop Supply Mixer;

NodeList,
HW Loop Setpoint Manager Node List,
HW Loop Supply Side Outlet;

Pipe:Adiabatic,
HW Loop Demand Side Inlet Branch Pipe,
HW Loop Demand Side Inlet,
HW Loop Demand Side Inlet Branch Pipe Outlet;

Pipe:Adiabatic,
HW Loop Demand Side Bypass Pipe,
HW Loop Demand Side Bypass Pipe Inlet Node,
HW Loop Demand Side Bypass Pipe Outlet Node;

Pipe:Adiabatic,
HW Loop Demand Side Outlet Branch Pipe,
HW Loop Demand Side Outlet Branch Pipe Inlet,
HW Loop Demand Side Outlet;

Pipe:Adiabatic,
HW Loop Supply Side Bypass Pipe,
HW Loop Supply Side Bypass Pipe Inlet Node,
HW Loop Supply Side Bypass Pipe Outlet Node;

Pipe:Adiabatic,
HW Loop Supply Side Outlet Branch Pipe,
HW Loop Supply Side Outlet Branch Pipe Inlet,
HW Loop Supply Side Outlet;

Pump:ConstantSpeed,
HW Loop Supply Pump,
HW Loop Supply Side Inlet,
HW Loop Supply Pump Water Outlet Node,
autosize,
207593.31,
autosize,
0.9,
0,
Intermittent;

Boiler:HotWater,
Boiler,
NaturalGas,
autosize,
0.9,
LeavingBoiler,
NECB 2020 - Non Condensing Boiler,
autosize,
0,
1,
1,
Boiler Water Inlet Node,
Boiler Water Outlet Node,
100,
NotModulated,
0,
1;

PlantLoop,
HW Loop,
Water,
,
HW Loop Operation,
HW Loop Supply Side Outlet,
100,
10,
autosize,
0,
autocalculate,
HW Loop Supply Side Inlet,
HW Loop Supply Side Outlet,
HW Loop Supply Side Branches,
HW Loop Supply Side Connectors,
HW Loop Demand Side Inlet,
HW Loop Demand Side Outlet,
HW Loop Demand Side Branches,
HW Loop Demand Side Connectors,
UniformLoad,
HW Loop AvailabilityManager List,
SingleSetpoint,
None,
None;

AvailabilityManager:Scheduled,
HW Loop Availability,
On 24/7;

AvailabilityManagerAssignmentList,
HW Loop AvailabilityManager List,
AvailabilityManager:Scheduled,
HW Loop Availability;

PlantEquipmentList,
HW Loop Scheme 1 Range 1 Equipment List,
Boiler:HotWater,
Boiler;

PlantEquipmentOperation:HeatingLoad,
HW Loop Scheme 1,
0,
1000000000000000,
HW Loop Scheme 1 Range 1 Equipment List;

PlantEquipmentOperationSchemes,
HW Loop Operation,
PlantEquipmentOperation:HeatingLoad,
HW Loop Scheme 1,
On 24/7;

SetpointManager:OutdoorAirReset,
HW Loop Setpoint Manager,
Temperature,
82,
-16,
60,
0,
HW Loop Setpoint Manager Node List;

Curve:Cubic,
NECB 2020 - Non Condensing Boiler,
0.38305305,
2.05668035,
-2.64690465,
1.21476776,
0,
1,
,
,
Dimensionless,
Dimensionless;
";

                string branches = "";
                string branchList = @"BranchList,
HW Loop Demand Side Branches,
HW Loop Demand Side Inlet Branch,
HW Loop Demand Side Bypass Branch,
";
                string connectorSplitter = @"Connector:Splitter,
HW Loop Demand Splitter,
HW Loop Demand Side Inlet Branch,
HW Loop Demand Side Bypass Branch,
";
                string connectorMixer = @"Connector:Mixer,
HW Loop Demand Mixer,
HW Loop Demand Side Outlet Branch,
HW Loop Demand Side Bypass Branch,
";

                int i = 1;
                foreach (string zoneName in this.ZoneNames)
                {
                    branches += String.Format(@"Branch,
{0} Water Convector HW Loop Demand Side Branch,
,
ZoneHVAC:Baseboard:Convective:Water,
{0} Water Convector,
{0} Water Convector Hot Water Inlet Node,
{0} Water Convector Hot Water Outlet Node;
", zoneName);

                    string delimiter = i < this.ZonesNumber ? "," : ";";
                    branchList += String.Format(@"{0} Water Convector HW Loop Demand Side Branch,
", zoneName);
                    connectorSplitter += String.Format(@"{0} Water Convector HW Loop Demand Side Branch{1}
", zoneName, delimiter);
                    connectorMixer += String.Format(@"{0} Water Convector HW Loop Demand Side Branch{1}
", zoneName, delimiter);
                    i++;
                }
                branchList += @"HW Loop Demand Side Outlet Branch;
";
                severalObjectsToIdf += branches + branchList + connectorSplitter + connectorMixer;
                idfFile.Load(severalObjectsToIdf);
            }
        }


        public class NECBSystem3
        {
            public AirSystemsManager AirSystemsHandler;
            public HotWaterPlantManager HotWaterPlantHandler;
            public NECBSystem3(List<HVACSystem> hvacSystems)
            {
                AirSystemsHandler = new AirSystemsManager(hvacSystems);
                HotWaterPlantHandler = new HotWaterPlantManager(hvacSystems);
            }
            public void ModifyAirLoops(IdfReader idfFile)
            {
                this.AirSystemsHandler.ModifyIDFObjects(idfFile);
                this.AirSystemsHandler.AddNewObjectsToIdf(idfFile);
            }

            public void CreateHotWaterPlant(IdfReader idfFile)
            {

                HotWaterPlantHandler.AddNewObjectsToIdf(idfFile);
            }
        }
        public class CSVResultsGenerator
        {
            private Dictionary<string, double> coolingCapacities;
            private Dictionary<string, double> supplyAirflows;
            private Dictionary<string, double> outdoorAirflows;

            private string GetEplusOutTblFile()
            {
                string defaultCsvPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "DesignBuilder", "EnergyPlus", "eplustbl.csv");

                if (File.Exists(defaultCsvPath))
                {
                    return defaultCsvPath;
                }

                string usersFolder = @"C:\ProgramData\DesignBuilder\JobServer\Users\User";
                List<string> jobFolders = Directory.GetDirectories(usersFolder)
                                          .Where(f =>
                                          {
                                              int temp;
                                              return int.TryParse(Path.GetFileName(f), out temp);
                                          })
                                          .ToList();
                string latestJobFolder = jobFolders
                                   .Select(f =>
                                   {
                                       int id;
                                       int.TryParse(Path.GetFileName(f), out id);
                                       return new { Path = f, Id = id };
                                   })
                                   .OrderByDescending(f => f.Id)
                                   .First()
                                   .Path;

                return Path.Combine(latestJobFolder, "eplustbl.csv");
            }
            public CSVResultsGenerator()
            {
                TableReader tableReader = new TableReader(GetEplusOutTblFile());
                coolingCapacities = tableReader.GetTable("DX Cooling Coils").GetData(0, 2);
                supplyAirflows = tableReader.GetTable("Controller:OutdoorAir").GetData(0, 1);
                outdoorAirflows = tableReader.GetTable("Controller:OutdoorAir").GetData(0, 2);
            }
            public void GenerateResultsCSV(List<HVACSystem> hvacSystemList)
            {
                using (var writer = new StreamWriter(GlobalVariables.ResultsSummaryPath))
                {
                    writer.WriteLine("Air Loop Name,Cooling Capacity (W),Supply Airflow (m3/s),Outdoor Airflow (m3/s),Requires Economizer?,Requires ERV?, Cooling Metric, " +
                        "Economizer Status, ERV Status, Cooling Metric Status");
                    foreach (HVACSystem airSystem in hvacSystemList)
                    {
                        double coolingCapacity = coolingCapacities[airSystem.AirLoopName + " AHU UNITARY COOLING COIL DX COOLING COIL"];
                        double supplyAirflow = supplyAirflows[airSystem.AirLoopName + " AHU OUTDOOR AIR CONTROLLER"];
                        double outdoorAirflow = outdoorAirflows[airSystem.AirLoopName + " AHU OUTDOOR AIR CONTROLLER"];
                        bool ervRequired = VerifyERV(supplyAirflow, outdoorAirflow);
                        string efficiencyMetric = coolingCapacity < 19000 ? "SEER" : "IEER";
                        string eRVStatus = ervRequired == airSystem.ERVRequired ? "Ok" : "Missing";
                        string coolingMetricStatus = efficiencyMetric == airSystem.CoolingEfficiencyMetric ? "Ok" : "Missing";
                        bool economizerRequired = false;
                        if (coolingCapacity > 20000)
                            economizerRequired = true;
                        if (supplyAirflow > 1.5)
                            economizerRequired = true;
                        string economizerStatus = economizerRequired == airSystem.EconomizerRequired ? "Ok" : "Missing";
                        writer.WriteLine("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}", airSystem.AirLoopName, coolingCapacity, supplyAirflow, outdoorAirflow, economizerRequired, ervRequired, efficiencyMetric, economizerStatus, eRVStatus, coolingMetricStatus);
                    }
                }
            }
            private bool VerifyERV(double supplyflow, double outdoorflow)
            {
                double outdoorAirPercent = 100 * outdoorflow / supplyflow;
                if (outdoorAirPercent < 10)
                {
                    return false;
                }
                Dictionary<string, List<double[]>> thresholds = new Dictionary<string, List<double[]>>();

                thresholds["7"] = new List<double[]> {
        new double[] {10, 20, 2.12},
        new double[] {20, 30, 1.89},
        new double[] {30, 40, 1.18},
        new double[] {40, 50, 0.47}
    };

                thresholds["8"] = new List<double[]> {
        new double[] {10, 20, 2.12},
        new double[] {20, 30, 1.89},
        new double[] {30, 40, 1.18},
        new double[] {40, 50, 0.47}
    };

                thresholds["5"] = new List<double[]> {
        new double[] {10, 20, 12.27},
        new double[] {20, 30, 7.55},
        new double[] {30, 40, 2.6},
        new double[] {40, 50, 2.12},
        new double[] {50, 60, 1.65},
        new double[] {60, 70, 0.94},
        new double[] {70, 80, 0.47}
    };

                thresholds["6"] = new List<double[]> {
        new double[] {10, 20, 12.27},
        new double[] {20, 30, 7.55},
        new double[] {30, 40, 2.6},
        new double[] {40, 50, 2.12},
        new double[] {50, 60, 1.65},
        new double[] {60, 70, 0.94},
        new double[] {70, 80, 0.47}
    };

                foreach (double[] t in thresholds[GlobalVariables.ClimateZone])
                {
                    double min = t[0];
                    double max = t[1];
                    double limit = t[2];

                    if (outdoorAirPercent >= min && outdoorAirPercent < max && supplyflow < limit)
                    {
                        return false;
                    }
                }
                return true;
            }
        }
    }    
}