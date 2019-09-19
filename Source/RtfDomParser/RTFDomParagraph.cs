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
    /// paragraph element
    /// </summary>
    [Serializable()]
    public class RTFDomParagraph : RTFDomElement
    {
        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFDomParagraph()
        {
        }

        internal bool TemplateGenerated = false;
        /// <summary>
        /// �Ƿ�����ʱ���ɵĶ������
        /// </summary>
        public bool IsTemplateGenerated
        {
            get
            {
                return TemplateGenerated;
            }
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
        public override string InnerText
        {
            get
            {
                return base.InnerText + Environment.NewLine ;
            }
        }
        public override string ToString()
        {
            var str = new StringBuilder();
            str.Append("Paragraph");
            if (Format != null)
            {
                str.Append("(" + Format.Align + ")");
                if (Format.ListID >= 0)
                {
                    str.Append("ListID:" + Format.ListID);
                }
                //if (this.Format.NumberedList)
                //{
                //    str.Append("(NumberedList)");
                //}
                //else if (this.Format.BulletedList)
                //{
                //    str.Append("(BulletedList)");
                //}
            }
            
            return str.ToString();
        }
    }
}
