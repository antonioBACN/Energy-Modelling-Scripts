using DB.Extensibility.Contracts;
using EpNet;
using System;
using System.Collections.Generic;
using System.Linq;


namespace DB.Extensibility.Scripts
{
    public class FanCoilsControl : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(ApiEnvironment.EnergyPlusInputIdfPath, ApiEnvironment.EnergyPlusInputIddPath);
            List<IdfObject> fancoilListObject = idf["ZoneHVAC:FourPipeFanCoil"];
            foreach (IdfObject fancoil in fancoilListObject)
            {
                fancoil["Capacity Control Method"].Value = "ASHRAE90VariableFan";
                fancoil["Low Speed Supply Air Flow Ratio"].Value = "0.5";
                fancoil["Minimum Supply Air Temperature in Cooling Mode"].Number = 12.7;
                fancoil["Maximum Supply Air Temperature in Heating Mode"].Number = 35;
            }
            idf.Save();
        }
    }
}
