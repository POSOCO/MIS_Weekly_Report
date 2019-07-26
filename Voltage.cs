namespace WRLDCWarehouse.Core.Entities
{
    public class Voltage
    {
        public int VoltageId { get; set; }

        public Substation Substation { get; set; }
        public int SubstationId { get; set; }

        public string Date { get; set; }

        public string Time { get; set; }


        public double VoltageValue { get; set; }



    }
}