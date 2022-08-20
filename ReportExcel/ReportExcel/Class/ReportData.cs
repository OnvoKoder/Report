namespace ReportExcel.Class
{
    class ReportData
    {
        public string _datetime { get; set; }
        public double _max { get; set; }
        public double _avg { get; set; }
        public double _min { get; set; }
        public ReportData( string datetime,double max, double avg,double min)
        {
            _datetime = datetime;
            _max = max;
            _avg = avg;
            _min = min;
        }
    }
}
