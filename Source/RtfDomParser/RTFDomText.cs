/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */



using System;
using System.Text;

namespace RtfDomParser
{
    /// <summary>
    /// text element
    /// </summary>
    [Serializable()]
    public class RTFDomText : RTFDomElement
    {
        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFDomText()
        {
            // text element can not contains any child element
            Locked = true;
        }

        private DocumentFormatInfo myFormat = new DocumentFormatInfo();
        /// <summary>
        /// format
        /// </summary>
        public DocumentFormatInfo Format
        {
            get
            {
                return myFormat;
            }
            set
            {
                myFormat = value;
            }
        }

        private string strText = null;
        /// <summary>
        /// text
        /// </summary>
        [System.ComponentModel.DefaultValue( null)]
        public string Text
        {
            get
            {
                return strText;
            }
            set
            {
                strText = value;
            }
        }
        public override string InnerText
        {
            get
            {
                return strText;
            }
        }
        public override string ToString()
        {
            var str = new StringBuilder();
            str.Append("Text");
            if (Format != null)
            {
                if (Format.Hidden)
                {
                    str.Append("(Hidden)");
                }
            }
            str.Append(":" + strText);
            return str.ToString();
        }
    }
}
