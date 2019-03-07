using System;

namespace TimecampConverter
{
    internal class UniteeTask
    {
        public UniteeTask()
        {
        }

        public string Description { get; internal set; }
        public string Project { get; internal set; }
        public decimal TimeSpent { get; internal set; }
        public DateTime Date { get; internal set; }
    }
}