using System.Threading.Tasks;
using WRLDCWarehouse.Data;
using WRLDCWarehouse.Core.Entities;
using WRLDCWarehouse.Core.ForiegnEntities;
using Microsoft.EntityFrameworkCore;
using WRLDCWarehouse.ETL.Enums;
using System;
using Microsoft.Extensions.Logging;

namespace WRLDCWarehouse.ETL.Loads
{
    public class LoadSsVoltage
    {
        public async Task<Voltage> LoadSingleAsync(WRLDCWarehouseDbContext _context, ILogger _log, VoltageForeign volForeign, EntityWriteOption opt)
        {
            // check if entity already exists
            Voltage existingSSVoltage = await _context.Voltage.SingleOrDefaultAsync(ss => ss.WebUatId == volForeign.WebUatId);
            //Entity name which is Voltage here should be decide

            // check if we should not modify existing entities
            if (opt == EntityWriteOption.DontReplace && existingSS != null)
            {
                return existingSSVoltage;
            }

            //Find the associated SubStation by SubstationId 
            int assSubStationId = volForeign.AssSubstationWebUatId;
            Substation assSubstation = await _context.Substation.SingleOrDefaultAsync(ass => ass.WebUatId == assSubStationId);

            // if Associated Substation doesnot exist, skip the import. Ideally, there should not be such case
            if (assSubstation == null)
            {
                _log.LogCritical($"Unable to find AssociatedSubstation with webUatId {assSubStationId} while inserting Substation with webUatId {volForeign.WebUatId} ");
                return null;
            }

            // if entity is not present, then insert
            if (existingSSVoltage == null)
            {
                Voltage newVoltage = new Voltage();
                newVoltage.VoltageId = volForeign.WebUatId;
                newVoltage.SubstationId = volForeign.AssSubstationWebUatId;
                newVoltage.Date = volForeign.Date;
                newVoltage.Time = volForeign.Time;
                newVoltage.VoltageValue = volForeign.Voltage;
                newVoltage.Substation = assSubstation;
                _context.Voltage.Add(newVoltage);
                await _context.SaveChangesAsync();
                return newVoltage;
            }

            // check if we have to replace the entity completely
            if (opt == EntityWriteOption.Replace && existingSSVoltage != null)
            {
                _context.Voltage.Remove(existingSSVoltage);
                Voltage newVoltage = new Voltage();
                newVoltage.VoltageId = volForeign.WebUatId;
                newVoltage.SubstationId = volForeign.AssSubstationWebUatId;
                newVoltage.Date = volForeign.Date;
                newVoltage.Time = volForeign.Time;
                newVoltage.VoltageValue = volForeign.Voltage;
                newVoltage.Substation = assSubstation;

                _context.Voltage.Add(newVoltage);
                await _context.SaveChangesAsync();
                return newVoltage;
            }

              // check if we have to modify the entity completely
            if (opt == EntityWriteOption.Modify && existingSSVoltage != null)
            {
                existingSSVoltage.VoltageId = volForeign.WebUatId;
                existingSSVoltage.SubstationId = volForeign.AssSubstationWebUatId;
                existingSSVoltage.Date = volForeign.Date;
                existingSSVoltage.Time = volForeign.Time;
                existingSSVoltage.VoltageValue = volForeign.Voltage;
                existingSSVoltage.Substation = assSubstation;
                await _context.SaveChangesAsync();
                return existingSS;
            }
            return null;

        }
    }
}