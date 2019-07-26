using WRLDCWarehouse.Core.Entities;
using WRLDCWarehouse.ETL.Extracts;
using WRLDCWarehouse.ETL.Loads;
using WRLDCWarehouse.ETL.Enums;
using WRLDCWarehouse.Core.ForiegnEntities;
using Microsoft.Extensions.Logging;

namespace WRLDCWarehouse.ETL.Jobs
{
    public class JobReadForeignSsVoltage
    {
        public async Task ImportForeignSsVoltage(WRLDCWarehouseDbContext _context, ILogger _log, string oracleConnStr, EntityWriteOption opt)
        {
            VolSubStationExtract volExtract = new VolSubStationExtract();
            List<VoltageForeign> ssVoltageList = volExtract.ExtractSubStationVoltages(oracleConnStr);

            LoadSsVoltage loadSsVoltage = new LoadSsVoltage();  
            foreach(VoltageForeign voltageForeign in ssVoltageList)
            {
                Voltage insertedVoltage = await loadSsVoltage.LoadSingleAsync(_context, _log, stateForeign, opt);
            }

        }
    }
}