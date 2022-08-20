namespace ReportExcel.Class
{
    class ReportData
    {
        private string _datetime { get; set; }
        private double _max { get; set; }
        private double _avg { get; set; }
        private double _min { get; set; }
        internal ReportData( string datetime,double max, double avg,double min)
        {
            _datetime = datetime;
            _max = max;
            _avg = avg;
            _min = min;
        }
    }
}
