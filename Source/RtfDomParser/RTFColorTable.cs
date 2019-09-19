/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System.Collections ;

namespace RtfDomParser
{
	/// <summary>
	/// rtf color table
	/// </summary>
    [System.Diagnostics.DebuggerDisplay("Count={Count}")]
    [System.Diagnostics.DebuggerTypeProxy(typeof(RTFInstanceDebugView))]
	public class RTFColorTable
	{
		/// <summary>
		/// initialize instance
		/// </summary>
		public RTFColorTable()
		{
		}

		private ArrayList myItems = new ArrayList();
		/// <summary>
		/// get color at special index
		/// </summary>
		public System.Drawing.Color this[ int index ]
		{
			get
            {
                return ( System.Drawing.Color ) myItems[ index ] ;
            }
		}

		/// <summary>
		/// get color at special index , if index out of range , return default color
		/// </summary>
		/// <param name="index">index</param>
		/// <param name="DefaultValue">default value</param>
		/// <returns>color value</returns>
		public System.Drawing.Color GetColor( int index , System.Drawing.Color DefaultValue )
		{
			index -- ;
            if (index >= 0 && index < myItems.Count)
            {
                return (System.Drawing.Color)myItems[index];
            }
            else
            {
                return DefaultValue;
            }
		}

        private bool bolCheckValueExistWhenAdd = true ;
        /// <summary>
        /// check color value exist when add color to list
        /// </summary>
        public bool CheckValueExistWhenAdd
        {
            get
            {
                return bolCheckValueExistWhenAdd; 
            }
            set
            {
                bolCheckValueExistWhenAdd = value; 
            }
        }

		/// <summary>
		/// add color to list
		/// </summary>
		/// <param name="c">new color value</param>
		public void Add( System.Drawing.Color c )
		{
			if( c.IsEmpty )
				return ;
			if( c.A == 0 )
				return ;
			
			if( c.A != 255 )
			{
				c = System.Drawing.Color.FromArgb( 255 , c );
			}

            if (bolCheckValueExistWhenAdd)
            {
                if (IndexOf(c) < 0)
                {
                    myItems.Add(c);
                }
            }
            else
            {
                myItems.Add(c);
            }
		}
		/// <summary>
		/// delete special color
		/// </summary>
		/// <param name="c">color value</param>
		public void Remove( System.Drawing.Color c )
		{
			var index = IndexOf( c );
			if( index >= 0 )
				myItems.RemoveAt( index );
		}
		/// <summary>
		/// get color index
		/// </summary>
		/// <param name="c">color</param>
		/// <returns>index , if not found , return -1</returns>
		public int IndexOf( System.Drawing.Color c )
		{
            if (c.A == 0)
            {
                return -1;
            }
			if( c.A != 255 )
			{
				c = System.Drawing.Color.FromArgb( 255 , c );
			}
            for (var iCount = 0; iCount < myItems.Count; iCount++)
            {
                var color = (System.Drawing.Color)myItems[iCount];
                if (color.ToArgb() == c.ToArgb())
                {
                    return iCount;
                }
            }
			return -1 ;
		}
		/// <summary>
		/// ����б�
		/// </summary>
		public void Clear()
		{
			myItems.Clear();
		}
		/// <summary>
		/// Ԫ�ظ���
		/// </summary>
		public int Count
		{
			get{ return myItems.Count ; }
		}

		/// <summary>
		/// �����ɫ��
		/// </summary>
		/// <param name="writer">RTF�ĵ���д��</param>
		public void Write( RTFWriter writer )
		{
			writer.WriteStartGroup();
			writer.WriteKeyword( RTFConsts._colortbl );
			writer.WriteRaw(";");
			for( var iCount = 0 ; iCount < myItems.Count ; iCount ++ )
			{
				var c = ( System.Drawing.Color ) myItems[ iCount ] ;
				writer.WriteKeyword( "red" + c.R );
				writer.WriteKeyword( "green" + c.G );
				writer.WriteKeyword( "blue" + c.B );
				writer.WriteRaw(";");
			}
			writer.WriteEndGroup();
		}

        /// <summary>
        /// ���ƶ���
        /// </summary>
        /// <returns>����Ʒ</returns>
        public RTFColorTable Clone()
        {
            var table = new RTFColorTable();
            for (var iCount = 0; iCount < myItems.Count; iCount++)
            {
                var c = ( System.Drawing.Color ) myItems[ iCount ] ;
                table.myItems.Add(c);
            }
            return table;
        }
    }
}