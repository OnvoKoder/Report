namespace ReportExcel.Class
{
    class ReportData
    {
        internal string _datetime { get; set; }
        internal double _max { get; set; }
        internal double _avg { get; set; }
        internal double _min { get; set; }
        internal ReportData( string datetime,double max, double avg,double min)
        {
            _datetime = datetime;
            _max = max;
            _avg = avg;
            _min = min;
        }
    }
}
