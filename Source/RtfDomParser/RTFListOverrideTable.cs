﻿using System.Collections.Generic;

namespace RtfDomParser
{
    public class RTFListOverrideTable : List<RTFListOverride>
    {
        public RTFListOverrideTable()
        {
        }

        public RTFListOverride GetByID(int id)
        {
            foreach (var item in this)
            {
                if (item.ID == id)
                {
                    return item;
                }
            }
            return null;
        }
    }

    public class RTFListOverride
    {
        private int _ListID = 0;

        public int ListID
        {
            get { return _ListID; }
            set { _ListID = value; }
        }

        private int _ListOverriedCount = 0;

        public int ListOverriedCount
        {
            get { return _ListOverriedCount; }
            set { _ListOverriedCount = value; }
        }

        private int _ID = 1;

        public int ID
        {
            get { return _ID; }
            set { _ID = value; }
        }

        public override string ToString()
        {
            return "ID:" + ID + " ListID:" + ListID + " Count:" + ListOverriedCount;
        }
    }
}
