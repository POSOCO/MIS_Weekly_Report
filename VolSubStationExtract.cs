using WRLDCWarehouse.Core.ForiegnEntities;
using Oracle.ManagedDataAccess.Client;

namespace WRLDCWarehouse.ETL.Extracts
{
    public class VolSubStationExtract
    {
        public List<Voltage> ExtractSubStationVoltages(string oracleConnString)
        {
            using (OracleConnection con = new OracleConnection(oracleConnString))
            {
                using (OracleCommand cmd = con.CreateCommand())
                {
                    try
                    {
                        con.Open();
                        cmd.BindByName = true;

                        cmd.CommandText = "SELECT ID,ELEMENT_KEY, DATE_KEY,TIME_KEY,VOLTAGE_VALUE FROM STG_SCADA_VOLTAGE ";

                        OracleParameter id = new OracleParameter("id", 1);
                        cmd.Parameters.Add(id);

                        //Execute the command and use DataReader to display the data
                        OracleDataReader reader = cmd.ExecuteReader();

                        List<VoltageForeign> Voltages = new List<VoltageForeign>();
                        while(reader.Read())
                        {
                            VoltageForeign voltage = new Voltage();
                            voltage.WebUatId = reader.GetInt32(0);
                            voltage.AssSubstationWebUatId = reader.GetInt32(1);
                            voltage.Date = reader.GetString(2);
                            voltage.Time = reader.GetString(3);
                            voltage.Voltage = reader.GetDouble(4);

                            Voltages.Add(voltage);

                        }

                        reader.Dispose();

                        return Voltages;

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        return null;
                    }
                }
            }
}