/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System;
using System.Collections ;
using System.Collections.Generic ;
using System.Text ;

namespace RtfDomParser
{
    

    /// <summary>
    /// font table
    /// </summary>
    [System.Diagnostics.DebuggerDisplay("Count={ Count }")]
    [System.Diagnostics.DebuggerTypeProxy(typeof(RTFInstanceDebugView))]
	public class RTFFontTable : CollectionBase
	{

		/// <summary>
		/// initialize instance
		/// </summary>
		public RTFFontTable()
		{
		}

		//private ArrayList myItems = new ArrayList();

		/// <summary>
		/// get font information special index
		/// </summary>
		public RTFFont this[ int fontIndex ]
		{
			get
			{
                foreach (RTFFont item in this)
                {
                    if (item.Index == fontIndex)
                        return item;
                }
                return null;
			}
		}

        /// <summary>
        /// get font object special name
        /// </summary>
        /// <param name="fontName">font name</param>
        /// <returns>font object</returns>
        public RTFFont this[string fontName]
        {
            get
            {
                foreach (RTFFont item in this)
                {
                    if (item.Name == fontName)
                    {
                        return item;
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// get font object special font index
        /// </summary>
        /// <param name="fontIndex">font index</param>
        /// <returns>font object</returns>
        public string GetFontName(int fontIndex)
        {
            var font = this[fontIndex];
            if (font != null)
            {
                return font.Name;
            }
            else
            {
                return null;
            }
        } 

		/// <summary>
		/// add font
		/// </summary>
		/// <param name="f">font name</param>
		public RTFFont Add( string f )
		{
            return Add(Count, f, Encoding.Default);
		}

        /// <summary>
        /// add font
        /// </summary>
        /// <param name="f">font name</param>
        public RTFFont Add(string f , Encoding encoding )
        {
            return Add(Count, f, encoding);
        }

		/// <summary>
		/// add font
		/// </summary>
		/// <param name="index">special font index</param>
		/// <param name="f">font name</param>
        public RTFFont Add(int index, string f, Encoding encoding)
        {
            if (this[f] == null)
            {
                var font = new RTFFont(index, f);
                if (encoding != null)
                {
                    font.Charset = RTFFont.GetCharset(encoding);
                }
                List.Add(font);
                return font ;
            }
            return this[f];
        }

        public void Add(RTFFont f)
        {
            List.Add(f);
        }

		/// <summary>
		/// Remove font
		/// </summary>
		/// <param name="f">font name</param>
		public void Remove( string f )
		{
            var item = this[f];
            if (item != null)
                List.Remove( item );
		}

		/// <summary>
		/// Get font index special font name
		/// </summary>
		/// <param name="f">font name</param>
		/// <returns>font index</returns>
		public int IndexOf( string f )
		{
            foreach (RTFFont item in this)
            {
                if (item.Name == f)
                {
                    return item.Index;
                }
            }
			return -1 ;
		}
		  
		/// <summary>
		/// Write font table rtf
		/// </summary>
		/// <param name="writer">rtf text writer</param>
		public void Write( RTFWriter writer )
		{
			writer.WriteStartGroup();
			writer.WriteKeyword( RTFConsts._fonttbl );
			foreach( RTFFont item in this )
			{
				writer.WriteStartGroup();
				writer.WriteKeyword( "f" + item.Index );
                if (item.Charset != 0)
                {
                    writer.WriteKeyword("fcharset" + item.Charset);
                }
				writer.WriteText( item.Name );
				writer.WriteEndGroup();
			}
			writer.WriteEndGroup();
		}

		public override string ToString()
		{
			var str = new StringBuilder();
			foreach( RTFFont item in this )
			{
				str.Append( Environment.NewLine );
				str.Append( "Index " + item.Index + "   Name:" + item.Name );
			}
			return str.ToString();
		}

        /// <summary>
        /// close object
        /// </summary>
        /// <returns>new object</returns>
        public RTFFontTable Clone()
        {
            var table = new RTFFontTable();
            foreach (RTFFont item in this )
            {
                var newItem = item.Clone();
                table.List.Add(newItem);
            }
            return table;
        }
	}

    /// <summary>
    /// rtf font information
    /// </summary>
    public class RTFFont
    {
        /// <summary>
        /// initialize instance
        /// </summary>
        /// <param name="index">font index</param>
        /// <param name="name">font name</param>
        public RTFFont(int index, string name)
        {
            intIndex = index;
            strName = name;
        }

        private int intIndex = 0;
        /// <summary>
        /// font index
        /// </summary>
        public int Index
        {
            get
            {
                return intIndex; 
            }
            set
            {
                intIndex = value; 
            }
        }

        private bool _NilFlag = false;

        public bool NilFlag
        {
            get { return _NilFlag; }
            set { _NilFlag = value; }
        }

        private string strName = null;
        /// <summary>
        /// font name
        /// </summary>
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

        private int intCharset = 1;
        /// <summary>
        /// charset 
        /// </summary>
        public int Charset
        {
            get
            {
                return intCharset;
            }
            set
            {
                intCharset = value;
                myEncoding = GetRTFEncoding(intCharset);
            }
        }

        private static Dictionary<int, Encoding> _EncodingCharsets = null;
        private static void CheckEncodingCharsets()
        {
            if (_EncodingCharsets == null)
            {
                _EncodingCharsets = new Dictionary<int, Encoding>();
                //_EncodingCharsets[0] = ANSIEncoding.Instance;
                //_EncodingCharsets[1] = Encoding.Default;
                _EncodingCharsets[77] = Encoding.GetEncoding(10000);//Mac ,macintosh ��ŷ�ַ�(Mac)
                _EncodingCharsets[128] = Encoding.GetEncoding(932);//Shift Jis ;ANSI/OEM - Japanese, Shift-JIS 
                _EncodingCharsets[130] = Encoding.GetEncoding(1361);//Johab;Korean (Johab) 
                _EncodingCharsets[134] = Encoding.GetEncoding(936);//GB2312
                _EncodingCharsets[136] = Encoding.GetEncoding(10002);//Big5
                _EncodingCharsets[161] = Encoding.GetEncoding(1253);//Greek
                _EncodingCharsets[162] = Encoding.GetEncoding(1254);//Turkish
                _EncodingCharsets[163] = Encoding.GetEncoding(1258);//Vietnamese;ANSI/OEM - Vietnamese 
                _EncodingCharsets[177] = Encoding.GetEncoding(1255);//Hebrw
                _EncodingCharsets[178] = Encoding.GetEncoding(864);//Arabic
                _EncodingCharsets[179] = Encoding.GetEncoding(864);//Arabic Traditional
                _EncodingCharsets[180] = Encoding.GetEncoding(864);//Arabic user
                _EncodingCharsets[181] = Encoding.GetEncoding(864);//Hebrew user
                _EncodingCharsets[186] = Encoding.GetEncoding(775);//Baltic
                _EncodingCharsets[204] = Encoding.GetEncoding(866);//Russian
                _EncodingCharsets[222] = Encoding.GetEncoding(874);//Thai
                _EncodingCharsets[255] = Encoding.GetEncoding(437);//OEM



            }
        }

        internal static int GetCharset(Encoding encoding)
        {
            CheckEncodingCharsets();
            foreach (var key in _EncodingCharsets.Keys)
            {
                if (_EncodingCharsets[key] == encoding)
                {
                    return key;
                }
            }
            return 1;
        }

        internal static Encoding GetRTFEncoding(int fchartset)
        {
            if (fchartset == 0)
            {
                return ANSIEncoding.Instance;
            }
            else if (fchartset == 1)
            {
                return Encoding.Default;
            }
            else
            {
                CheckEncodingCharsets();
                if (_EncodingCharsets.ContainsKey(fchartset))
                {
                    return _EncodingCharsets[fchartset];
                }
                else
                {
                    return null;
                }
            }

            //switch (fchartset)
            //{
            //    case 0: 	// ANSI
            //        return ANSIEncoding.Instance;
            //    case 1:	// Default
            //        return System.Text.Encoding.Default;
                    
            //    //case 2:	// Symbol
            //    //case 3:	// Invalid
            //    case 77:   // Mac
            //        return System.Text.Encoding.GetEncoding(10000); //macintosh ��ŷ�ַ�(Mac)
                    
            //    case 128:	// Shift Jis
            //        return System.Text.Encoding.GetEncoding(932);// ANSI/OEM - Japanese, Shift-JIS 
                    
            //    //case 129:	// Hangul
            //    case 130:	// Johab
            //        return System.Text.Encoding.GetEncoding(1361);// Korean (Johab) 
                    
            //    case 134:	// GB2312
            //        return System.Text.Encoding.GetEncoding(936);
                    
            //    case 136:	// Big5
            //        return System.Text.Encoding.GetEncoding(10002);// MAC - Traditional Chinese (Big5) 
                    
            //    case 161:	// Greek
            //        return System.Text.Encoding.GetEncoding(1253);// ANSI - Greek 
                    
            //    case 162:	// Turkish
            //        return System.Text.Encoding.GetEncoding(1254);//ANSI - Turkish 
                    
            //    case 163:	// Vietnamese
            //        return System.Text.Encoding.GetEncoding(1258);// ANSI/OEM - Vietnamese 
                    
            //    case 177:	// Hebrew
            //        return System.Text.Encoding.GetEncoding(1255);// ANSI - Hebrew 
                    
            //    case 178:	// Arabic
            //        return System.Text.Encoding.GetEncoding(864);//OEM - Arabic 
                    
            //    case 179:	// Arabic Traditional
            //        return System.Text.Encoding.GetEncoding(864);//OEM - Arabic 
                    
            //    case 180:	// Arabic user
            //        return System.Text.Encoding.GetEncoding(864);//OEM - Arabic 
                    
            //    case 181:	// Hebrew user
            //        return System.Text.Encoding.GetEncoding(864);//OEM - Arabic 
                    
            //    case 186:	// Baltic
            //        return System.Text.Encoding.GetEncoding(775);//OEM - Baltic 
                    
            //    case 204:	// Russian
            //        return System.Text.Encoding.GetEncoding(866);//OEM - Russian 
                    
            //    case 222:	// Thai
            //        return System.Text.Encoding.GetEncoding(874);//ANSI/OEM - Thai (same as 28605, ISO 8859-15) 
                    
            //    //case 238:	// Eastern European
            //    //case 254:	// PC 437
            //    case 255:	// OEM
            //        return System.Text.Encoding.GetEncoding(437);//OEM - United States 
                    
            //    default:
            //        return null;
            //}
        }

        private Encoding myEncoding = null ;
        /// <summary>
        /// encoding
        /// </summary>
        public Encoding Encoding
        {
            get 
            {
                return myEncoding; 
            }
        }

        public RTFFont Clone()
        {
            var f = new RTFFont( intIndex , strName );
            f.intCharset = intCharset;
            f.intIndex = intIndex;
            f.myEncoding = myEncoding;
            f.strName = strName;
            return f;
        }

        public override string ToString()
        {
            return intIndex + ":" + strName + " Charset:" + intCharset;
        }
    }

    
    /// <summary>
    /// internal encoding for ansi
    /// </summary>
    internal class ANSIEncoding : Encoding
    {
        public static ANSIEncoding Instance = new ANSIEncoding();
        public override string GetString(byte[] bytes, int index, int count)
        {
            var str = new StringBuilder();
            var endIndex = Math.Min(bytes.Length-1, index + count -1);
            for (var iCount = index ; iCount <= endIndex ; iCount++)
            {
                str.Append(System.Convert.ToChar(bytes[iCount]));
            }
            return str.ToString();
        }

        public override int GetByteCount(char[] chars, int index, int count)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        public override int GetBytes(char[] chars, int charIndex, int charCount, byte[] bytes, int byteIndex)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        public override int GetCharCount(byte[] bytes, int index, int count)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        public override int GetChars(byte[] bytes, int byteIndex, int byteCount, char[] chars, int charIndex)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        public override int GetMaxByteCount(int charCount)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        public override int GetMaxCharCount(int byteCount)
        {
            throw new Exception("The method or operation is not implemented.");
        }

    }
}