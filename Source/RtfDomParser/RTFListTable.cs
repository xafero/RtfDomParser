﻿using System;
using System.Collections.Generic;

namespace RtfDomParser
{
    public class RTFListTable : List<RTFList>
    {
        public RTFListTable()
        {
        }

        public RTFList GetByID(int id)
        {
            foreach (var list in this)
            {
                if (list.ListID == id)
                {
                    return list;
                }
            }
            return null;
        }
    }

    public class RTFList
    {
        private int _ListID = 0;

        public int ListID
        {
            get { return _ListID; }
            set { _ListID = value; }
        }

        private int _ListTemplateID = 0;

        public int ListTemplateID
        {
            get { return _ListTemplateID; }
            set { _ListTemplateID = value; }
        }

        private bool _ListSimple = false;

        public bool ListSimple
        {
            get { return _ListSimple; }
            set { _ListSimple = value; }
        }

        private bool _ListHybrid = false;

        public bool ListHybrid
        {
            get { return _ListHybrid; }
            set { _ListHybrid = value; }
        }

        private string _ListName = null;

        public string ListName
        {
            get { return _ListName; }
            set { _ListName = value; }
        }

        private string _ListStyleName = null;

        public string ListStyleName
        {
            get { return _ListStyleName; }
            set { _ListStyleName = value; }
        }

        private int _LevelStartAt = 1;

        public int LevelStartAt
        {
            get { return _LevelStartAt; }
            set { _LevelStartAt = value; }
        }

        private LevelNumberType _LevelNfc =  LevelNumberType.None  ;

        public LevelNumberType LevelNfc
        {
            get { return _LevelNfc; }
            set { _LevelNfc = value; }
        }

        private int _LevelJc = 0;

        public int LevelJc
        {
            get { return _LevelJc; }
            set { _LevelJc = value; }
        }

        private int _LevelFollow = 0;

        public int LevelFollow
        {
            get { return _LevelFollow; }
            set { _LevelFollow = value; }
        }

        private string _FontName = null;
        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName
        {
            get { return _FontName; }
            set { _FontName = value; }
        }

        private string _LevelText = null;

        public string LevelText
        {
            get { return _LevelText; }
            set { _LevelText = value; }
        }

        public override string ToString()
        {
            if (LevelNfc == LevelNumberType.Bullet)
            {
                var text = "ID:" + ListID + "   Bullet:";
                if (string.IsNullOrEmpty(LevelText) == false)
                {
                    text = text + "(" + Convert.ToString((short)LevelText[0]) + ")";
                }
                return text;
            }
            else
            {
                return  "ID:" + ListID + " " + LevelNfc.ToString() + " Start:" + LevelStartAt;
            }
        }
    }
}
