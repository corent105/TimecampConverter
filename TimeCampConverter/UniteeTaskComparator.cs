using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace TimecampConverter
{
    class UniteeTaskComparator : IComparer<UniteeTask>
    {
        public int Compare(UniteeTask x, UniteeTask y)
        {
            return x.Date.CompareTo(y.Date);
        }
    }
}
