/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */

using System;

namespace RtfDomParser
{
    /// <summary>
    /// string attribute
    /// </summary>
    [Serializable()]
    public class StringAttribute
    {
        /// <summary>
        /// initialize instance
        /// </summary>
        public StringAttribute()
        {
        }

        private string strName = null;
        /// <summary>
        /// name
        /// </summary>
        [System.ComponentModel.DefaultValue( null )]
        public string Name
        {
            get
            {
                return strName; 
            }
            set 
            {
                strName = value; 
            }
        }

        private string strValue = null;
        /// <summary>
        /// value
        /// </summary>
        [System.ComponentModel.DefaultValue( null)]
        public string Value
        {
            get
            {
                return strValue; 
            }
            set
            {
                strValue = value; 
            }
        }
        public override string ToString()
        {
            return strName + "=" + strValue;
        }
    }

    [Serializable()]
    [System.Diagnostics.DebuggerTypeProxy(typeof(RTFInstanceDebugView))]
    public class StringAttributeCollection : System.Collections.CollectionBase
    {
        public StringAttributeCollection()
        {
        }

        public string this[string name]
        {
            get
            {
                foreach (StringAttribute attr in this)
                {
                    if (attr.Name == name)
                    {
                        return attr.Value;
                    }
                }
                return null;
            }
            set
            {
                foreach (StringAttribute item in this)
                {
                    if (item.Name == name)
                    {
                        if (value == null)
                            List.Remove(item);
                        else
                            item.Value = value;
                        return;
                    }
                }
                if (value != null)
                {
                    var newItem = new StringAttribute();
                    newItem.Name = name;
                    newItem.Value = value;
                    List.Add(newItem);
                }
            }
        }

        public int Add(StringAttribute item)
        {
            return List.Add(item);
        }

        public void Remove(StringAttribute item)
        {
            List.Remove(item);
        }
    }
}
