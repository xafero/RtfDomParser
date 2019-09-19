﻿/*
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
    /// 强制分页符
    /// </summary>
    [Serializable()]
    public class RTFDomPageBreak:RTFDomElement
    {
        /// <summary>
        /// 初始化对象
        /// </summary>
        public RTFDomPageBreak()
        {
            //对象不能有子元素
            Locked = true;
        }

        public override string InnerText
        {
            get
            {
                return "";
            }
        }
        public override string ToString()
        {
            return "page";
        }
    }


}
