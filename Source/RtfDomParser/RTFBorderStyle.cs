/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */

using System.Drawing ;
using System.Drawing.Drawing2D;
using System.ComponentModel;

namespace RtfDomParser
{
    public class RTFBorderStyle
    {
        private bool _Left = false;
        [DefaultValue( false )]
        public bool Left
        {
            get { return _Left; }
            set { _Left = value; }
        }

        private bool _Top = false;
        [DefaultValue(false)]
        public bool Top
        {
            get { return _Top; }
            set { _Top = value; }
        }

        private bool _Right = false;
        [DefaultValue(false)]
        public bool Right
        {
            get { return _Right; }
            set { _Right = value; }
        }

        private bool _Bottom = false;
        [DefaultValue(false)]
        public bool Bottom
        {
            get { return _Bottom; }
            set { _Bottom = value; }
        }

        private DashStyle _Style = DashStyle.Solid;
        [DefaultValue(DashStyle.Solid)]
        public DashStyle Style
        {
            get { return _Style; }
            set { _Style = value; }
        }

        private Color _Color = Color.Black;
        [DefaultValue(typeof( Color ) , "Black")]
        public Color Color
        {
            get { return _Color; }
            set { _Color = value; }
        }

        private bool _Thickness = false;
        /// <summary>
        /// 粗边框样式
        /// </summary>
        [DefaultValue( false)]
        public bool Thickness
        {
            get { return _Thickness; }
            set { _Thickness = value; }
        }
        /// <summary>
        /// 复制对象
        /// </summary>
        /// <returns>复制品</returns>
        public RTFBorderStyle Clone()
        {
            var b = new RTFBorderStyle();
            b._Bottom = _Bottom;
            b._Color = _Color;
            b._Left = _Left;
            b._Right = _Right;
            b._Style = _Style;
            b._Top = _Top;
            b._Thickness = _Thickness;
            return b;
        }

        public bool EqualsValue(RTFBorderStyle b)
        {
            if (b == this)
            {
                return true;
            }
            if (b == null)
            {
                return false;
            }
            if (b._Bottom != _Bottom
                || b._Color != _Color
                || b._Left != _Left
                || b._Right != _Right
                || b._Style != _Style
                || b._Top != _Top
                || b._Thickness != _Thickness )
            {
                return false;
            }
            else
            {
                return true;
            }
        }
    }
}
