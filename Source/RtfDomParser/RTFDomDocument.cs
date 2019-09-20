/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */

using System;
using System.Collections;
using System.Text;
using System.ComponentModel ;
using System.Collections.Generic;
using System.Drawing ;
using System.Drawing.Drawing2D ;

namespace RtfDomParser
{
    /// <summary>
    /// RTF Document
    /// </summary>
    /// <remarks>
    /// This type is the root of RTF Dom tree structure
    /// </remarks>
    public partial class RTFDomDocument : RTFDomElement
    {
        static RTFDomDocument()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFDomDocument()
        {
            OwnerDocument = this;
        }
         
        private string strFollowingChars = null;
        /// <summary>
        /// following characters
        /// </summary>
        [DefaultValue( null)]
        public string FollowingChars
        {
            get
            {
                return strFollowingChars; 
            }
            set
            {
                strFollowingChars = value; 
            }
        }

        private string strLeadingChars = null;
        /// <summary>
        /// leading characters
        /// </summary>
        [DefaultValue(null)]
        public string LeadingChars
        {
            get
            {
                return strLeadingChars; 
            }
            set
            {
                strLeadingChars = value; 
            }
        }

        private Encoding myDefaultEncoding = Encoding.Default ;
        /// <summary>
        /// text encoding of current font
        /// </summary>
        private Encoding myFontChartset = null;
        /// <summary>
        /// text encoding of associate font 
        /// </summary>
        private Encoding myAssociateFontChartset = null;
        /// <summary>
        /// text encoding
        /// </summary>
        internal Encoding RuntimeEncoding
        {
            get
            {
                if (myFontChartset != null)
                {
                    return myFontChartset;
                }
                if (myAssociateFontChartset != null)
                {
                    return myAssociateFontChartset;
                }
                return myDefaultEncoding;
            }
        }

        /// <summary>
        /// default font name
        /// </summary>
        private static string DefaultFontName = Constants.DefaultFontName;


        private RTFFontTable myFontTable = new RTFFontTable();
        /// <summary>
        /// font table
        /// </summary>
        public RTFFontTable FontTable
        {
            get
            {
                return myFontTable;
            }
            set
            {
                myFontTable = value;
            }
        }

        private RTFColorTable myColorTable = new RTFColorTable();
        /// <summary>
        /// color table
        /// </summary>
        public RTFColorTable ColorTable
        {
            get
            {
                return myColorTable;
            }
            set
            {
                myColorTable = value;
            }
        }

        private RTFListTable _ListTable = new RTFListTable();

        public RTFListTable ListTable
        {
            get
            {
                return _ListTable; 
            }
            set
            {
                _ListTable = value; 
            }
        }

        private RTFListOverrideTable _ListOverrideTable = new RTFListOverrideTable();
        
        public RTFListOverrideTable ListOverrideTable
        {
            get { return _ListOverrideTable; }
            set { _ListOverrideTable = value; }
        }

        private RTFDocumentInfo myInfo = new RTFDocumentInfo();
        /// <summary>
        /// document information
        /// </summary>
        public RTFDocumentInfo Info
        {
            get
            {
                return myInfo;
            }
            set
            {
                myInfo = value;
            }
        }

        private string strGenerator = null;
        /// <summary>
        /// document generator
        /// </summary>
        [DefaultValue( null )]
        public string Generator
        {
            get
            {
                return strGenerator;
            }
            set
            {
                strGenerator = value;
            }
        }

        private int intPaperWidth = 12240;
        /// <summary>
        /// paper width,unit twips
        /// </summary>
        [DefaultValue( 12240 )]
        public int PaperWidth
        {
            get
            {
                return intPaperWidth; 
            }
            set
            {
                intPaperWidth = value; 
            }
        }

        private int intPaperHeight = 15840;
        /// <summary>
        /// paper height,unit twips
        /// </summary>
        [DefaultValue( 15840 )]
        public int PaperHeight
        {
            get
            {
                return intPaperHeight; 
            }
            set
            {
                intPaperHeight = value; 
            }
        }

        private int intLeftMargin = 1800;
        /// <summary>
        /// left margin,unit twips
        /// </summary>
        [DefaultValue( 1800 )]
        public int LeftMargin
        {
            get
            {
                return intLeftMargin; 
            }
            set
            {
                intLeftMargin = value; 
            }
        }

        private int intTopMargin = 1440;
        /// <summary>
        /// top margin,unit twips
        /// </summary>
        [DefaultValue( 1440 )]
        public int TopMargin
        {
            get
            {
                return intTopMargin; 
            }
            set
            {
                intTopMargin = value; 
            }
        }

        private int intRightMargin = 1800;
        /// <summary>
        /// right margin,unit twips
        /// </summary>
        [DefaultValue(1800)]
        public int RightMargin
        {
            get
            {
                return intRightMargin;
            }
            set
            {
                intRightMargin = value;
            }
        }

        private int intBottomMargin = 1440;
        /// <summary>
        /// bottom margin,unit twips
        /// </summary>
        [DefaultValue(1440)]
        public int BottomMargin
        {
            get
            {
                return intBottomMargin;
            }
            set
            {
                intBottomMargin = value;
            }
        }

        private bool bolLandscape = false;
        /// <summary>
        /// landscape
        /// </summary>
        [DefaultValue(false)]
        public bool Landscape
        {
            get
            {
                return bolLandscape;
            }
            set
            {
                bolLandscape = value;
            }
        }

        private int _HeaderDistance = 720;
        /// <summary>
        /// Header's distance from the top of the page( Twips)
        /// </summary>
        [DefaultValue( 720)]
        public int HeaderDistance
        {
            get
            {
                return _HeaderDistance; 
            }
            set
            {
                _HeaderDistance = value; 
            }
        }

        private int _FooterDistance = 720;
        /// <summary>
        /// Footer's distance from the bottom of the page( twips)
        /// </summary>
        [DefaultValue( 720 )]
        public int FooterDistance
        {
            get
            {
                return _FooterDistance; 
            }
            set
            {
                _FooterDistance = value; 
            }
        }
        /// <summary>
        /// client area width,unit twips
        /// </summary>
        [Browsable( false )]
        public int ClientWidth
        {
            get
            {
                if (bolLandscape)
                {
                    return intPaperHeight - intLeftMargin - intRightMargin;
                }
                else
                {
                    return intPaperWidth - intLeftMargin - intRightMargin;
                }
            }
        }

        private bool bolChangeTimesNewRoman = false ;
        /// <summary>
        /// convert "Times new roman" to default font when parse rtf content
        /// </summary>
        [DefaultValue( true )]
        public bool ChangeTimesNewRoman
        {
            get
            {
                return bolChangeTimesNewRoman;
            }
            set
            {
                bolChangeTimesNewRoman = value;
            }
        }
        //private Stack myElements = new Stack();

        /// <summary>
        /// progress event
        /// </summary>
        public event ProgressEventHandler Progress = null;

        /// <summary>
        /// raise progress event
        /// </summary>
        /// <param name="max">progress max value</param>
        /// <param name="Value">progress value</param>
        /// <param name="message">progress message</param>
        /// <returns>user cancel</returns>
        protected bool OnProgress(int max, int Value, string message)
        {
            if (Progress != null)
            {
                var args = new ProgressEventArgs(max, Value, message);
                Progress(this, args);
                return args.Cancel;
            }
            return false;
        }


        /// <summary>
        /// load a rtf file and parse
        /// </summary>
        /// <param name="fileName">file name</param>
        public void Load(string fileName)
        {
            using (var stream = new System.IO.FileStream(
                fileName,
                System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                Load(stream);
            }
        }

        /// <summary>
        /// load a rtf document from a stream and parse content
        /// </summary>
        /// <param name="stream">stream</param>
        public void Load(System.IO.Stream stream)
        {
            //_HtmlContentBuilder = new StringBuilder();
            //_RTFHtmlState = true ;
            _HtmlContent = null;
            Elements.Clear();
            bolStartContent = false;
            var reader = new RTFReader(stream);
            var format = new DocumentFormatInfo();
            _ParagraphFormat = null;
            Load(reader, format );
            // combination table rows to table
            CombinTable(this);
            FixElements(this);
            FixRTFHtml();
        }


        /// <summary>
        /// load a rtf document from a text reader and parse content
        /// </summary>
        /// <param name="reader">text reader</param>
        public void Load(System.IO.TextReader reader )
        {
            //_HtmlContentBuilder = new StringBuilder();
            //_RTFHtmlState = true;
            _HtmlContent = null;
            Elements.Clear();
            bolStartContent = false;
            var r = new RTFReader(reader);
            var format = new DocumentFormatInfo();
            _ParagraphFormat = null;
            Load(r, format );
            // combination table rows to table
            CombinTable(this);
            FixElements(this);
            FixRTFHtml();
        }

        /// <summary>
        /// load a rtf document from a string in rtf format and parse content
        /// </summary>
        /// <param name="rtfText">text</param>
        public void LoadRTFText(string rtfText)
        {
            var reader = new System.IO.StringReader(rtfText);
            //_HtmlContentBuilder = new StringBuilder();
            //_RTFHtmlState = true;
            _HtmlContent = null;
            Elements.Clear();
            bolStartContent = false;
            var rtfReader = new RTFReader(reader);
            var format = new DocumentFormatInfo();
            _ParagraphFormat = null;
            Load(rtfReader, format );
            CombinTable(this);
            FixElements(this);
            FixRTFHtml();
        }

        private void FixRTFHtml()
        {
            //if (this._HtmlContentBuilder != null)
            //{
            //    this._HtmlContent = this._HtmlContentBuilder.ToString();
            //    this._HtmlContentBuilder = null;
            //}
            //this._RTFHtmlState = false;
        }

        /// <summary>
        /// ���ĵ������е����ݶ������ڶ����С�
        /// </summary>
        public void FixForParagraphs( RTFDomElement parentElement )
        {
            RTFDomParagraph lastParagraph = null;
            var list = new RTFDomElementList();
            foreach (RTFDomElement element in parentElement.Elements)
            {
                if (element is RTFDomHeader
                    || element is RTFDomFooter)
                {
                    FixForParagraphs(element);
                    lastParagraph = null;
                    list.Add(element);
                    continue;
                }
                if ( element is RTFDomParagraph 
                    || element is RTFDomTableRow
                    || element is RTFDomTable 
                    || element is RTFDomTableCell )
                {
                    lastParagraph = null;
                    list.Add(element);
                    continue;
                }
                if (lastParagraph == null)
                {
                    lastParagraph = new RTFDomParagraph();
                    list.Add(lastParagraph);
                    if( element is RTFDomText )
                    {
                        lastParagraph.Format = ( ( RTFDomText ) element ).Format.Clone();
                    }
                }
                lastParagraph.Elements.Add(element);
            }//foreach
            parentElement.Elements.Clear();
            foreach (RTFDomElement element in list)
            {
                //if (element is RTFDomHeader
                //    || element is RTFDomFooter)
                //{
                //    FixForParagraphs(element);
                //}
                parentElement.Elements.Add(element);
            }
        }

        private void FixElements(RTFDomElement parentElement)
        {
            // combin text element , decrease number of RTFDomText instance
            var result = new ArrayList();
            foreach (RTFDomElement element in parentElement.Elements)
            {
                if (element is RTFDomParagraph)
                {
                    var p = (RTFDomParagraph)element;
                    if (p.Format.PageBreak)
                    {
                        p.Format.PageBreak = false;
                        result.Add(new RTFDomPageBreak());
                    }
                }
                if (element is RTFDomText)
                {
                    if (result.Count > 0 && result[result.Count - 1] is RTFDomText)
                    {
                        var lastText = (RTFDomText)result[result.Count - 1];
                        var txt = (RTFDomText)element;
                        if (lastText.Text.Length == 0 || txt.Text.Length == 0)
                        {
                            if (lastText.Text.Length == 0)
                            {
                                // close text format
                                lastText.Format = txt.Format.Clone();
                            }
                            lastText.Text = lastText.Text + txt.Text;
                        }
                        else
                        {
                            if (lastText.Format.EqualsSettings(txt.Format))
                            {
                                lastText.Text = lastText.Text + txt.Text;
                            }
                            else
                            {
                                result.Add(txt);
                            }
                        }
                    }
                    else
                    {
                        result.Add(element);
                    }
                }
                else
                {
                    result.Add(element);
                }
            }//foreach
            parentElement.Elements.Clear();
            parentElement.Locked = false;
            foreach (RTFDomElement element in result)
            {
                parentElement.AppendChild(element);
            }

            foreach (var element in parentElement.Elements.ToArray())
            {
                if (element is RTFDomTable)
                {
                    UpdateTableCells((RTFDomTable)element, true);
                }
            }


            //// ɾ����ʱ���ɵĶ������󣬽������������ϴ����ƶ�һλ
            //RTFDomParagraph tempP = null;
            //RTFDomParagraph lastP = null;
            //foreach (RTFDomElement element in parentElement.Elements)
            //{
            //    if (element is RTFDomParagraph)
            //    {
            //        RTFDomParagraph p = (RTFDomParagraph)element;
            //        if (p.TemplateGenerated)
            //        {
            //            tempP = p;
            //        }
            //        else
            //        {
            //            if (tempP != null)
            //            {
            //                tempP.TemplateGenerated = false;
            //                tempP.Format = p.Format;
            //                tempP = p;
            //            }
            //        }
            //        lastP = p;
            //    }
            //}//foreach
            //if (tempP != null && lastP != null)
            //{
            //    // �������������ϴ�λ�ƶ����������һ������Ϊ�գ���ɾ�����һ�����䡣
            //    if (lastP.Elements.Count == 0)
            //    {
            //        parentElement.Elements.Remove(lastP);
            //    }
            //}

            foreach (RTFDomElement element in parentElement.Elements)
            {
                FixElements(element);
            }
        }

        private RTFDomElement[] GetLastElements( bool checkLockState)
        {
            var result = new List<RTFDomElement>();
            RTFDomElement element = this;
            while (element != null)
            {
                if (checkLockState)
                {
                    if (element.Locked)
                    {
                        break;
                    }
                }
                result.Add(element);
                element = element.Elements.LastElement;
            }
            if (checkLockState)
            {
                for (var iCount = result.Count - 1; iCount >= 0; iCount--)
                {
                    if (result[iCount].Locked)
                    {
                        result.RemoveAt(iCount);
                    }
                }
            }
            return result.ToArray();
        }

        public RTFDomElement GetLastElement(Type elementType)
        {
            var elements = GetLastElements(true );
            for (var iCount = elements.Length - 1; iCount >= 0; iCount--)
            {
                if (elementType.IsInstanceOfType(elements[iCount]))
                    return elements[iCount];
            }
            return null;
        }

        public RTFDomElement GetLastElement(Type elementType, bool lockStatus)
        {
            var elements = GetLastElements( true );
            for (var iCount = elements.Length - 1; iCount >= 0; iCount--)
            {
                if (elementType.IsInstanceOfType(elements[iCount]))
                {
                    if (elements[iCount].Locked == lockStatus)
                    {
                        return elements[iCount];
                    }
                }
            }
            return null;
        }

        public RTFDomElement GetLastElement()
        {
            var elements = GetLastElements( true );
            return elements[elements.Length - 1];
        }

        private void CompleteParagraph()
        {
            var lastElement = GetLastElement();
            while ( lastElement != null )
            {
                if (lastElement is RTFDomParagraph)
                {
                    var p = (RTFDomParagraph)lastElement;
                    p.Locked = true;
                    if (_ParagraphFormat != null)
                    {
                        p.Format = _ParagraphFormat;
                        _ParagraphFormat = _ParagraphFormat.Clone();
                    }
                    else
                    {
                        _ParagraphFormat = new DocumentFormatInfo();
                    }
                    break;
                }
                lastElement = lastElement.Parent;
            }
        }

        private void AddContentElement(RTFDomElement newElement)
        {
            var elements = GetLastElements( true );
            RTFDomElement lastElement = null;
            if (elements.Length > 0)
            {
                lastElement = elements[elements.Length - 1];
            }
            if ( lastElement is RTFDomDocument 
                || lastElement is RTFDomHeader
                || lastElement is RTFDomFooter ) 
            {
                if (newElement is RTFDomText
                    || newElement is RTFDomImage
                    || newElement is RTFDomObject
                    || newElement is RTFDomShape
                    || newElement is RTFDomShapeGroup)
                {
                    var p = new RTFDomParagraph();
                    if (lastElement.Elements.Count > 0)
                    {
                        p.TemplateGenerated = true;
                    }
                    if (_ParagraphFormat != null)
                    {
                        p.Format = _ParagraphFormat;
                    }
                    lastElement.AppendChild(p);
                    p.Elements.Add(newElement);
                    return;
                }
            }
            var element = elements[elements.Length - 1];
            //if ( newElement is RTFDomTableRow)
            //{
            //    System.Diagnostics.Debugger.Break();
            //}
            if (newElement != null && newElement.NativeLevel > 0)
            {
                for (var iCount = elements.Length - 1; iCount >= 0; iCount--)
                {
                    if (elements[iCount].NativeLevel == newElement.NativeLevel)
                    {
                        for (var iCount2 = iCount; iCount2 < elements.Length; iCount2++)
                        {
                            var e2 = elements[iCount2];
                            //if (newElement.GetType().Equals(e2.GetType()))
                            //{
                            //}
                            if (newElement is RTFDomText
                                || newElement is RTFDomImage
                                || newElement is RTFDomObject
                                || newElement is RTFDomShape
                                || newElement is RTFDomShapeGroup
                                || newElement is RTFDomField
                                || newElement is RTFDomBookmark
                                || newElement is RTFDomLineBreak)
                            {
                                if (newElement.NativeLevel == e2.NativeLevel)
                                {
                                    if (e2 is RTFDomTableRow
                                        || e2 is RTFDomTableCell
                                        || e2 is RTFDomField
                                        || e2 is RTFDomParagraph )
                                    {
                                        continue;
                                    }
                                }
                            }

                            elements[iCount2].Locked = true;
                        }
                        break;
                    }
                }
            }
            
            for (var iCount = elements.Length - 1; iCount >= 0; iCount--)
            {
                if (elements[iCount].Locked == false)
                {
                    element = elements[iCount];
                    if (element is RTFDomImage)
                    {
                        element.Locked = true;
                    }
                    else
                    {
                        break;
                    }
                }
            }
            if (element is RTFDomTableRow)
            {
                // If the last element is table row 
                // can not contains any element , 
                // so need create a cell element.
                var cell = new RTFDomTableCell();
                cell.NativeLevel = element.NativeLevel;
                element.AppendChild(cell);
                if (newElement is RTFDomTableRow)
                {
                    cell.Elements.Add(newElement);
                }
                else
                {
                    var cellP = new RTFDomParagraph();
                    cellP.Format = _ParagraphFormat.Clone();
                    cellP.NativeLevel = cell.NativeLevel;
                    cell.AppendChild(cellP);
                    if (newElement != null)
                    {
                        cellP.AppendChild(newElement);
                    }
                }
            }
            else
            {
                if (newElement != null)
                {
                    if (element is RTFDomParagraph &&
                        ( newElement is RTFDomParagraph
                        || newElement is RTFDomTableRow ) )
                    {
                        // If both is paragraph , append new paragraph to the parent of old paragraph
                        element.Locked = true;
                        element.Parent.AppendChild(newElement);
                    }
                    else
                    {
                        element.AppendChild(newElement);
                    }
                }
            }
        }



        private int ListTextFlag = 0;
        private bool bolStartContent = false;

        /// <summary>
        /// convert a hex string to a byte array
        /// </summary>
        /// <param name="hex">hex string</param>
        /// <returns>byte array</returns>
        private byte[] HexToBytes(string hex)
        {
            var chars = "0123456789abcdef";

            var index = 0;
            var Value = 0;
            var CharCount = 0;
            var buffer = new ByteBuffer();
            for (var iCount = 0; iCount < hex.Length; iCount++)
            {
                var c = hex[iCount];
                c = char.ToLower(c);
                index = chars.IndexOf(c);
                if (index >= 0)
                {
                    CharCount++;
                    Value = Value * 16 + index;
                    if (CharCount > 0 && (CharCount % 2) == 0)
                    {
                        buffer.Add((byte)Value);
                        Value = 0;
                    }
                }
            }
            return buffer.ToArray();
        }

        /// <summary>
        /// ��������Ԫ�غϲ��ɱ���Ԫ��
        /// </summary>
        /// <param name="parentElement">��Ԫ�ض���</param>
        private void CombinTable(RTFDomElement parentElement)
        {
            var result = new ArrayList();
            var rows = new ArrayList();
            var lastRowWidth = -1;
            RTFDomTableRow lastRow = null;
            foreach (RTFDomElement element in parentElement.Elements)
            {
                if (element is RTFDomTableRow)
                {
                    var row = (RTFDomTableRow)element;
                    row.Locked = false;
                    var cellSettings = row.CellSettings;
                    if (cellSettings.Count == 0)
                    {
                        if (lastRow != null && lastRow.CellSettings.Count == row.Elements.Count)
                        {
                            cellSettings = lastRow.CellSettings;
                        }
                    }
                    if (cellSettings.Count == row.Elements.Count)
                    {
                        for (var iCount = 0; iCount < row.Elements.Count; iCount++)
                        {
                            row.Elements[iCount].Attributes = (RTFAttributeList)cellSettings[iCount];
                        }
                    }
                    var isLastRow = row.HasAttribute(RTFConsts._lastrow);
                    if (isLastRow == false)
                    {
                        var index = parentElement.Elements.IndexOf(element);
                        if (index == parentElement.Elements.Count - 1)
                        {
                            // this element is the last element
                            // then this row is the last row
                            isLastRow = true;
                        }
                        else
                        {
                            var e2 = parentElement.Elements[index + 1];
                            if ( !(e2 is RTFDomTableRow))
                            {
                                // next element is not row 
                                isLastRow = true;
                            }
                        }
                    }
                    // split to table
                    if (isLastRow)
                    {
                        // if current row mark the last row , then 
                        // generate a new table
                        rows.Add(row);
                        result.Add(CreateTable(rows));
                        lastRowWidth = -1;
                    }
                    else
                    {
                        var width = 0;
                        if (row.HasAttribute(RTFConsts._trwWidth))
                        {
                            width = row.Attributes[RTFConsts._trwWidth];
                            if (row.HasAttribute(RTFConsts._trwWidthA))
                            {
                                width = width - row.Attributes[RTFConsts._trwWidthA];
                            }
                        }
                        else
                        {
                            foreach (RTFDomTableCell cell in row.Elements)
                            {
                                if (cell.HasAttribute(RTFConsts._cellx))
                                {
                                    width = Math.Max(width, cell.Attributes[RTFConsts._cellx]);
                                }
                            }
                        }
                        if (lastRowWidth > 0 && lastRowWidth != width)
                        {
                            // If row's width is change , then can consider multi-table combin
                            // then split and generate new table
                            if (rows.Count > 0)
                            {
                                result.Add(CreateTable(rows));
                            }
                        }
                        lastRowWidth = width;
                        rows.Add(row);
                    }
                    lastRow = row;
                }
                else if( element is RTFDomTableCell )
                {
                    lastRow = null;
                    CombinTable(element);
                    if (rows.Count > 0)
                    {
                        result.Add(CreateTable(rows));
                    }
                    result.Add(element);
                    lastRowWidth = -1;
                }
                else
                {
                    lastRow = null;
                    CombinTable(element);
                    if (rows.Count > 0)
                    {
                        result.Add(CreateTable(rows));
                    }
                    result.Add(element);
                    lastRowWidth = -1;
                }
            }//foreach
            if (rows.Count > 0)
            {
                result.Add(CreateTable(rows));
            }
            parentElement.Locked = false;
            parentElement.Elements.Clear();
            foreach (RTFDomElement element in result)
            {
                parentElement.AppendChild(element);
            }

        }
        /// <summary>
        /// create table
        /// </summary>
        /// <param name="rows">table rows</param>
        /// <returns>new table</returns>
        private RTFDomTable CreateTable(ArrayList rows)
        {
            if (rows.Count > 0)
            {
                var table = new RTFDomTable();
                var index = 0;
                foreach (RTFDomTableRow row in rows)
                {
                    row.RowIndex = index;
                    index++;
                    table.AppendChild(row);
                }
                rows.Clear();
                foreach (RTFDomTableRow row in table.Elements)
                {
                    foreach (RTFDomTableCell cell in row.Elements)
                    {
                        CombinTable(cell);
                    }
                }
                return table;
            }
            else
            {
                throw new ArgumentException("rows");
            }
        }

        private int intDefaultRowHeight = 400;
        /// <summary>
        /// default row's height, in twips.
        /// </summary>
        public int DefaultRowHeight
        {
            get
            {
                return intDefaultRowHeight;
            }
            set
            {
                intDefaultRowHeight = value;
            }
        }
         
        private void UpdateTableCells(RTFDomTable table, bool fixTableCellSize)
        {
            // number of table column
            var columns = 0;
            // flag of cell merge
            var merge = false;
            // right position of all cells
            var rights = new ArrayList();

            // right position of table
            var tableLeft = 0;
            for (var iCount = table.Elements.Count - 1; iCount >= 0; iCount--)
            {
                var row = (RTFDomTableRow)table.Elements[iCount];
                if (row.Elements.Count == 0)
                {
                    // ɾ��û�����ݵı�����
                    table.Elements.RemoveAt(iCount);
                }
            }
            if (table.Elements.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("");
            }
            foreach (RTFDomTableRow row in table.Elements)
            {
                var lastCellX = 0;

                columns = Math.Max(columns, row.Elements.Count);
                if (row.HasAttribute(RTFConsts._irow))
                {
                    row.RowIndex = row.Attributes[RTFConsts._irow];
                }
                row.IsLastRow = row.HasAttribute(RTFConsts._lastrow);
                row.Header = row.HasAttribute(RTFConsts._trhdr);
                // read row height
                if (row.HasAttribute(RTFConsts._trrh))
                {
                    row.Height = row.Attributes[RTFConsts._trrh];
                    if (row.Height == 0)
                    {
                        row.Height = DefaultRowHeight;
                    }
                    else if (row.Height < 0)
                    {
                        row.Height = -row.Height;
                    }
                }
                else
                {
                    row.Height = DefaultRowHeight;
                }
                // read default padding of cell
                if (row.HasAttribute(RTFConsts._trpaddl))
                {
                    row.PaddingLeft = row.Attributes[RTFConsts._trpaddl];
                }
                else
                {
                    row.PaddingLeft = int.MinValue;
                }
                if (row.HasAttribute(RTFConsts._trpaddt))
                {
                    row.PaddingTop = row.Attributes[RTFConsts._trpaddt];
                }
                else
                {
                    row.PaddingTop = int.MinValue;
                }

                if (row.HasAttribute(RTFConsts._trpaddr))
                {
                    row.PaddingRight = row.Attributes[RTFConsts._trpaddr];
                }
                else
                {
                    row.PaddingRight = int.MinValue;
                }

                if (row.HasAttribute(RTFConsts._trpaddb))
                {
                    row.PaddingBottom = row.Attributes[RTFConsts._trpaddb];
                }
                else
                {
                    row.PaddingBottom = int.MinValue;
                }

                if (row.HasAttribute(RTFConsts._trleft))
                {
                    tableLeft = row.Attributes[RTFConsts._trleft];
                }
                if (row.HasAttribute(RTFConsts._trcbpat))
                {
                    row.Format.BackColor = ColorTable.GetColor(
                        row.Attributes[RTFConsts._trcbpat] ,
                        Color.Transparent );
                }
                var widthCount = 0;
                foreach (RTFDomTableCell cell in row.Elements)
                {
                    // set cell's dispaly format

                    if (cell.HasAttribute(RTFConsts._clvmgf))
                    {
                        // cell vertically merge
                        merge = true;
                    }
                    if (cell.HasAttribute(RTFConsts._clvmrg))
                    {
                        // cell vertically merge by another cell
                        merge = true;
                    }
                    if (cell.HasAttribute(RTFConsts._clpadl))
                    {
                        cell.PaddingLeft = cell.Attributes[RTFConsts._clpadl];
                    }
                    else
                    {
                        cell.PaddingLeft = int.MinValue;
                    }
                    if (cell.HasAttribute(RTFConsts._clpadr))
                    {
                        cell.PaddingRight = cell.Attributes[RTFConsts._clpadr];
                    }
                    else
                    {
                        cell.PaddingRight = int.MinValue;
                    }
                    if (cell.HasAttribute(RTFConsts._clpadt))
                    {
                        cell.PaddingTop = cell.Attributes[RTFConsts._clpadt];
                    }
                    else
                    {
                        cell.PaddingTop = int.MinValue;
                    }
                    if (cell.HasAttribute(RTFConsts._clpadb))
                    {
                        cell.PaddingBottom = cell.Attributes[RTFConsts._clpadb];
                    }
                    else
                    {
                        cell.PaddingBottom = int.MinValue;
                    }

                    // whether dispaly border line
                    cell.Format.LeftBorder = cell.HasAttribute(RTFConsts._clbrdrl);
                    cell.Format.TopBorder = cell.HasAttribute(RTFConsts._clbrdrt);
                    cell.Format.RightBorder = cell.HasAttribute(RTFConsts._clbrdrr);
                    cell.Format.BottomBorder = cell.HasAttribute(RTFConsts._clbrdrb);
                    if (cell.HasAttribute(RTFConsts._brdrcf))
                    {
                        cell.Format.BorderColor = ColorTable.GetColor(
                            cell.GetAttributeValue(RTFConsts._brdrcf, 1) ,
                            Color.Black );
                    }
                    for (var iCount = cell.Attributes.Count - 1; iCount >= 0; iCount--)
                    {
                        // ���� brdrtbl ָ��������ĳ����Ԫ��߿���
                        var name3 = cell.Attributes.GetItem(iCount).Name;
                        if ( name3 == RTFConsts._brdrtbl 
                            || name3 == RTFConsts._brdrnone 
                            || name3 == RTFConsts._brdrnil )
                        {
                            // ĳ���߿���ʾ
                            for (var iCount2 = iCount - 1; iCount2 >= 0; iCount2--)
                            {
                                var name2 = cell.Attributes.GetItem(iCount2).Name;
                                if (name2 == RTFConsts._clbrdrl)
                                {
                                    cell.Format.LeftBorder = false;
                                    break;
                                }
                                else if (name2 == RTFConsts._clbrdrt)
                                {
                                    cell.Format.TopBorder = false;
                                    break;
                                }
                                else if (name2 == RTFConsts._clbrdrr)
                                {
                                    cell.Format.RightBorder = false;
                                    break;
                                }
                                else if (name2 == RTFConsts._clbrdrb)
                                {
                                    cell.Format.BottomBorder = false;
                                    break;
                                }
                            }//for
                        }
                    }
                    // vertial alignment
                    if (cell.HasAttribute(RTFConsts._clvertalt))
                    {
                        cell.VerticalAlignment = RTFVerticalAlignment.Top;
                    }
                    else if (cell.HasAttribute(RTFConsts._clvertalc))
                    {
                        cell.VerticalAlignment = RTFVerticalAlignment.Middle;
                    }
                    else if (cell.HasAttribute(RTFConsts._clvertalb))
                    {
                        cell.VerticalAlignment = RTFVerticalAlignment.Bottom;
                    }
                    // background color
                    if (cell.HasAttribute(RTFConsts._clcbpat))
                    {
                        cell.Format.BackColor = ColorTable.GetColor(cell.Attributes[RTFConsts._clcbpat] , Color.Transparent );
                    }
                    else
                    {
                        cell.Format.BackColor = Color.Transparent;
                    }
                    if (cell.HasAttribute(RTFConsts._clcfpat))
                    {
                        cell.Format.BorderColor = ColorTable.GetColor(cell.Attributes[RTFConsts._clcfpat]  , Color.Black );
                    }

                    // cell's width
                    var cellWidth = 2763;// cell's default with is 2763 Twips(570 Document)
                    if (cell.HasAttribute(RTFConsts._cellx))
                    {
                        cellWidth = cell.Attributes[RTFConsts._cellx] - lastCellX;
                        if (cellWidth < 100)
                        {
                            cellWidth = 100;
                        }
                    }
                    var right = lastCellX + cellWidth;
                    // fix cell's right position , if this position is very near with another cell's 
                    // right position( less then 45 twips or 3 pixel), then consider these two position
                    // is the same , this can decrease number of table columns
                    for (var iCount = 0; iCount < rights.Count; iCount++)
                    {
                        if (Math.Abs(right - (int)rights[iCount]) < 45)
                        {
                            right = (int)rights[iCount];
                            cellWidth = right - lastCellX;
                            break;
                        }
                    }

                    cell.Left = lastCellX;
                    cell.Width = cellWidth;
                    if (cell.HasAttribute(RTFConsts._cellx) == false)
                    {

                    }
                    widthCount += cellWidth;
                    //int right = cell.Left + cell.Width;
                    if (rights.Contains(right) == false)
                    {
                        // becase of convert twips to unit of document may cause truncation error.
                        // This may cause rights.Contains mistake . so scale cell's with with 
                        // native twips unit , after all computing , convert to unit of document.
                        rights.Add(right);
                    }
                    lastCellX = lastCellX + cellWidth;
                }//foreach
                row.Width = widthCount;
            }//foreach
            if (rights.Count == 0)
            {
                // can not detect cell's width , so consider set cell's width
                // automatic, then set cell's default width.
                var cols = 1;
                foreach (RTFDomTableRow row in table.Elements)
                {
                    cols = Math.Max(cols, row.Elements.Count);
                }
                var w = (int)(ClientWidth / cols);
                for (var iCount = 0; iCount < cols; iCount++)
                {
                    rights.Add(iCount * w + w);
                }
            }
            // computing cell's rowspan and colspan , number of rights array is the number of table columns.

            rights.Add(0);
            rights.Sort();
            // add table column instance
            for (var iCount = 1; iCount < rights.Count; iCount++)
            {
                var col = new RTFDomTableColumn();
                col.Width = (int)rights[iCount] - (int)rights[iCount - 1];
                table.Columns.Add(col);
            }

            for (var rowIndex = 1; rowIndex < table.Elements.Count; rowIndex++)
            {
                var row = (RTFDomTableRow)table.Elements[rowIndex];
                for (var colIndex = 0; colIndex < row.Elements.Count; colIndex++)
                {
                    var cell = (RTFDomTableCell)row.Elements[colIndex];
                    if (cell.Width == 0)
                    {
                        // If current cell not special width , then use the width of cell which 
                        // in the same colum and in the last row
                        var preRow = (RTFDomTableRow)table.Elements[rowIndex - 1];
                        if (preRow.Elements.Count > colIndex)
                        {
                            var preCell = (RTFDomTableCell)preRow.Elements[colIndex];
                            cell.Left = preCell.Left;
                            cell.Width = preCell.Width;
                            CopyStyleAttribute(cell, preCell.Attributes);
                        }
                    }
                }
            }
            if (merge == false)
            {
                // If not detect cell merge , maby exist cell merge in the same row
                foreach (RTFDomTableRow row in table.Elements)
                {
                    if (row.Elements.Count < table.Columns.Count)
                    {
                        // if number of row's cells not equals the number of table's columns
                        // then exist cell merge.
                        merge = true;
                        break;
                    }
                }
            }
            if (merge)
            {
                // detect cell merge , begin merge operation

                // Becase of in rtf format,cell which merged by another cell in the same row , 
                // does no written in rtf text , so delay create those cell instance .
                foreach (RTFDomTableRow row in table.Elements)
                {
                    if (row.Elements.Count != table.Columns.Count)
                    {
                        // If number of row's cells not equals number of table's columns ,
                        // then consider there are hanppend  horizontal merge.
                        var cells = row.Elements.ToArray();
                        foreach (RTFDomTableCell cell in cells)
                        {
                            var index = rights.IndexOf(cell.Left);
                            var index2 = rights.IndexOf(cell.Left + cell.Width);
                            var intColSpan = index2 - index;
                            // detect vertical merge
                            var verticalMerge = cell.HasAttribute(RTFConsts._clvmrg);
                            if (verticalMerge == false)
                            {
                                // If this cell does not merged by another cell abover , 
                                // then set colspan
                                cell.ColSpan = intColSpan;
                            }
                            if (row.Elements.LastElement == cell)
                            {
                                cell.ColSpan = table.Columns.Count - row.Elements.Count + 1;
                                intColSpan = cell.ColSpan;
                            }
                            for (var iCount = 0; iCount < intColSpan - 1; iCount++)
                            {
                                var newCell = new RTFDomTableCell();
                                newCell.Attributes = cell.Attributes.Clone();
                                row.Elements.Insert(row.Elements.IndexOf(cell) + 1, newCell);
                                if (verticalMerge)
                                {
                                    // This cell has been merged.
                                    newCell.Attributes[RTFConsts._clvmrg] = 1;
                                    newCell.OverrideCell = cell;
                                }
                            }//for
                        }
                        if (row.Elements.Count != table.Columns.Count)
                        {
                            // If the last cell has been merged. then supply new cells.
                            var lastCell = (RTFDomTableCell)row.Elements.LastElement;
                            if (lastCell == null)
                            {
                                Console.WriteLine("");
                            }
                            //if (lastCell.OverrideCell == null && lastCell.ColSpan > 1)
                            //{
                            //    lastCell.ColSpan = table.Columns.Count - row.Elements.IndexOf(lastCell);
                            //}
                            for (var iCount = row.Elements.Count; iCount < rights.Count; iCount++)
                            {
                                var newCell = new RTFDomTableCell();
                                CopyStyleAttribute(newCell, lastCell.Attributes);
                                row.Elements.Add(newCell);
                            }
                        }//if
                    }//if
                }//foreach

                // set cell's vertial merge.
                foreach (RTFDomTableRow row in table.Elements)
                {
                    foreach (RTFDomTableCell cell in row.Elements)
                    {
                        if (cell.HasAttribute(RTFConsts._clvmgf) == false)
                        {
                            //if this cell does not mark vertial merge , then next cell
                            continue;
                        }
                        // if this cell mark vertial merge.
                        var colIndex = row.Elements.IndexOf(cell);
                        for (var rowIndex = table.Elements.IndexOf(row) + 1;
                            rowIndex < table.Elements.Count;
                            rowIndex++)
                        {
                            var row2 = (RTFDomTableRow)table.Elements[rowIndex];
                            if (colIndex >= row2.Elements.Count)
                            {
                                Console.Write("");
                            }
                            var cell2 = (RTFDomTableCell)row2.Elements[colIndex];
                            if (cell2.HasAttribute(RTFConsts._clvmrg))
                            {
                                if (cell2.OverrideCell != null)
                                {
                                    // If this cell has been merge by another cell( must in the same row )
                                    // then break the circle
                                    break;
                                }
                                // increase vertial merge.
                                cell.RowSpan++;
                                cell2.OverrideCell = cell;
                            }//if
                            else
                            {
                                // if this cell not mark merged by another cell , then break the circel
                                break;
                            }
                        }//for
                    }
                }

                // set cell's OverridedCell information
                foreach (RTFDomTableRow row in table.Elements)
                {
                    foreach (RTFDomTableCell cell in row.Elements)
                    {
                        if (cell.RowSpan > 1 || cell.ColSpan > 1)
                        {
                            for (var rowIndex = 1; rowIndex <= cell.RowSpan; rowIndex++)
                            {
                                for (var colIndex = 1; colIndex <= cell.ColSpan; colIndex++)
                                {
                                    var r = table.Elements.IndexOf(row) + rowIndex - 1;
                                    var c = row.Elements.IndexOf(cell) + colIndex - 1;
                                    var cell2 = (RTFDomTableCell)table.Elements[r].Elements[c];
                                    if (cell != cell2)
                                    {
                                        cell2.OverrideCell = cell;
                                    }
                                }//for
                            }//for
                        }//if
                    }//foreach
                }//foreach

            }//if

            if (fixTableCellSize)
            {

                // Fix table's left position use the first table column
                if (table.Columns.Count > 0)
                {
                    ((RTFDomTableColumn)table.Columns[0]).Width -= tableLeft;
                }
            }

        }

        private void CopyStyleAttribute(RTFDomTableCell cell, RTFAttributeList table)
        {
            var attrs = table.Clone();
            attrs.Remove(RTFConsts._clvmgf);
            attrs.Remove(RTFConsts._clvmrg);
            cell.Attributes = attrs;
        }

        public override string ToString()
        {
            return "RTFDocument:" + myInfo.Title;
        }

        private bool ApplyText(
            RTFTextContainer myText, 
            RTFReader reader,
            DocumentFormatInfo format )
        {
            if (myText.HasContent)
            {
                var strText = myText.Text;
                myText.Clear();
                //if (this._RTFHtmlState == false)
                //{
                //    _HtmlContentBuilder.Append(strText);
                //    return false ;
                //}
                // if current element is image element , then finish handle image element
                var img = (RTFDomImage)GetLastElement( typeof( RTFDomImage )) ;
                if( img != null && img.Locked == false )
                {
                    img.Data = HexToBytes(strText);
                    img.Format = format.Clone();
                    img.Width = (int)(img.DesiredWidth * img.ScaleX / 100);
                    img.Height = (int)(img.DesiredHeight * img.ScaleY / 100);
                    img.Locked = true;
                    if (reader.TokenType != RTFTokenType.GroupEnd)
                    {
                        ReadToEndGround(reader);
                    }
                    return true;
                }
                else if (format.ReadText && bolStartContent)
                {
                    var txt = new RTFDomText();
                    txt.NativeLevel = myText.Level;
                    txt.Format = format.Clone();
                    if (txt.Format.Align == RTFAlignment.Justify)
                        txt.Format.Align = RTFAlignment.Left;
                    txt.Text = strText;
                    AddContentElement(txt);
                }
            }
            return false;
        }


        private int intTokenCount = 0;
        private DocumentFormatInfo _ParagraphFormat = null;
        private void Load(RTFReader reader , DocumentFormatInfo parentFormat)
        {
            var ForbitPard = false;
            DocumentFormatInfo format = null;
            if (_ParagraphFormat == null)
            {
                _ParagraphFormat = new DocumentFormatInfo();
            }
            if (parentFormat == null)
            {
                format = new DocumentFormatInfo();
            }
            else
            {
                format = parentFormat.Clone();
                format.NativeLevel = parentFormat.NativeLevel + 1;
            }
            var myText = new RTFTextContainer( this );
            var levelBack = reader.Level;
            while (reader.ReadToken() != null)
            {
                if (reader.TokenCount - intTokenCount > 100)
                {
                    intTokenCount = reader.TokenCount;
                    OnProgress(reader.ContentLength, reader.ContentPosition, null);
                }
                if (bolStartContent)
                {
                    if (myText.Accept(reader.CurrentToken , reader ))
                    {
                        myText.Level = reader.Level;
                        continue;
                    }
                    else if (myText.HasContent)
                    {
                        if (ApplyText(myText, reader, format))
                        {
                            break;
                        }
                    }
                }

                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    var elements = GetLastElements( true );
                    for (var iCount = 0 ; iCount < elements.Length ; iCount ++ )
                    {
                        var element = elements[iCount];
                        if (element.NativeLevel >= 0 && element.NativeLevel > reader.Level)
                        {
                            for (var iCount2 = iCount ; iCount2 < elements.Length; iCount2++)
                            {
                                elements[iCount2].Locked = true;
                            }
                            break;
                        }
                    }//for

                    break;
                }

                if (reader.Level < levelBack)
                {
                    break;
                }

                if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    //level++;
                    Load(reader, format );
                    if (reader.Level < levelBack)
                    {
                        break;
                    }
                    //continue;
                }

                if (reader.TokenType == RTFTokenType.Control 
                    || reader.TokenType == RTFTokenType.Keyword 
                    || reader.TokenType == RTFTokenType.ExtKeyword)
                {
                     

                    switch (reader.Keyword)
                    {
                        case "fromhtml":
                            // ��HTML��ʽ��ȡ�ĵ�����
                            ReadHtmlContent(reader);
                            return;
                        case RTFConsts._listtable:
                            ReadListTable(reader);
                            return;
                        case RTFConsts._listoverride:
                            // unknow keyword
                            ReadToEndGround(reader);
                            break;
                        case RTFConsts._ansi :
                            break;
                        case RTFConsts._ansicpg :
                            // read default encoding
                            myDefaultEncoding = Encoding.GetEncoding(reader.Parameter);
                            break;
                        case RTFConsts._fonttbl :
                            // read font table
                            ReadFontTable(reader);
                            break;
                        case "listoverridetable":
                            ReadListOverrideTable(reader);
                            break;
                        case "filetbl":
                            // unsupport file list
                            ReadToEndGround( reader );
                            break ;// finish current level
                            //break;
                        case RTFConsts._colortbl :
                            // read color table
                            ReadColorTable(reader);
                            return;// finish current level
                            //break;
                        case "stylesheet":
                            // unsupport style sheet list
                            ReadToEndGround(reader);
                            break;
                        case RTFConsts._generator :
                            // read document generator
                            Generator = ReadInnerText(reader, true );
                            break;
                        case RTFConsts._info :
                            // read document information
                            ReadDocumentInfo(reader);
                            return ;
                        case RTFConsts._headery:
                            {
                                if (reader.HasParam)
                                {
                                    HeaderDistance = reader.Parameter;
                                }
                            }
                            break;
                        case RTFConsts._footery :
                            {
                                if (reader.HasParam)
                                {
                                    FooterDistance = reader.Parameter;
                                }
                            }
                            break;
                        case RTFConsts._header:
                            {
                                // analyse header
                                var header = new RTFDomHeader();
                                header.Style = HeaderFooterStyle.AllPages ;
                                AppendChild(header);
                                Load(reader, parentFormat);
                                header.Locked = true;
                                _ParagraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts._headerl:
                            {
                                // analyse header
                                var header = new RTFDomHeader();
                                header.Style = HeaderFooterStyle.LeftPages ;
                                AppendChild(header);
                                Load(reader, parentFormat);
                                header.Locked = true;
                                _ParagraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts._headerr:
                            {
                                // analyse header
                                var header = new RTFDomHeader();
                                header.Style = HeaderFooterStyle.RightPages;
                                AppendChild(header);
                                Load(reader, parentFormat);
                                header.Locked = true;
                                _ParagraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts._headerf:
                            {
                                // analyse header
                                var header = new RTFDomHeader();
                                header.Style = HeaderFooterStyle.FirstPage;
                                AppendChild(header);
                                Load(reader, parentFormat);
                                header.Locked = true;
                                _ParagraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts._footer:
                            {
                                // analyse footer
                                var footer = new RTFDomFooter();
                                footer.Style = HeaderFooterStyle.AllPages;
                                AppendChild(footer);
                                Load(reader, parentFormat);
                                footer.Locked = true;
                                _ParagraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts._footerl:
                            {
                                // analyse footer
                                var footer = new RTFDomFooter();
                                footer.Style = HeaderFooterStyle.LeftPages ;
                                AppendChild(footer);
                                Load(reader, parentFormat);
                                footer.Locked = true;
                                _ParagraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts._footerr:
                            {
                                // analyse footer
                                var footer = new RTFDomFooter();
                                footer.Style = HeaderFooterStyle.RightPages ;
                                AppendChild(footer);
                                Load(reader, parentFormat);
                                footer.Locked = true;
                                _ParagraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts._footerf:
                            {
                                // analyse footer
                                var footer = new RTFDomFooter();
                                footer.Style = HeaderFooterStyle.FirstPage ;
                                AppendChild(footer);
                                Load(reader, parentFormat);
                                footer.Locked = true;
                                _ParagraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts._xmlns:
                            {
                                // unsupport xml namespace
                                ReadToEndGround(reader);
                                break;
                            }
                        case  RTFConsts._nonesttables:
                            {
                                // I support nest table , then ignore this keyword
                                ReadToEndGround(reader);
                                break;
                            }
                        case  RTFConsts._xmlopen:
                            {
                                // unsupport xmlopen keyword
                                break;
                            }
                        case RTFConsts._revtbl:
                            {
                                //ReadToEndGround(reader);
                                break;
                            }
                        
                       
                        //**************** read document information ***********************
                        case RTFConsts._paperw:
                            {
                                // read paper width
                                intPaperWidth = reader.Parameter;
                                break;
                            }
                        case RTFConsts._paperh:
                            {
                                // read paper height
                                intPaperHeight = reader.Parameter;
                                break;
                            }
                        case RTFConsts._margl:
                            {
                                // read left margin
                                intLeftMargin = reader.Parameter;
                                break;
                            }
                        case RTFConsts._margr:
                            {
                                // read right margin
                                intRightMargin = reader.Parameter;
                                break;
                            }
                        case RTFConsts._margb:
                            {
                                // read bottom margin
                                intBottomMargin = reader.Parameter;
                                break;
                            }
                        case RTFConsts._margt:
                            {
                                // read top margin 
                                intTopMargin = reader.Parameter;
                                break;
                            }
                        case RTFConsts._landscape:
                            {
                                // set landscape
                                bolLandscape = true;
                                break;
                            }
                        case RTFConsts._fchars :
                            FollowingChars = ReadInnerText(reader, true);
                            break;
                        case RTFConsts._lchars :
                            LeadingChars = ReadInnerText(reader, true);
                            break;
                        case "pnseclvl":
                            // ignore this keyword
                            ReadToEndGround(reader);
                            break; ;
                        ////**************** read html content ***************************
                        //case "htmlrtf":
                        //    {
                        //        if (reader.HasParam && reader.Parameter == 0)
                        //        {
                        //            _RTFHtmlState = false;
                        //        }
                        //        else
                        //        {
                        //            _RTFHtmlState = true;
                        //            //while ( reader.PeekTokenType() == RTFTokenType.Text )
                        //            //{
                        //            //    reader.ReadToken();
                        //            //}
                        //            //string text = ReadInnerText(reader, null, false, true, false);
                        //        }
                        //    }
                        //    break;
                        //case "htmltag":
                        //    {
                        //        if (reader.InnerReader.Peek() == (int)' ')
                        //        {
                        //            reader.InnerReader.Read();
                        //        }
                        //        string text = ReadInnerText(reader, null, true, false, true);
                        //        if (string.IsNullOrEmpty(text) == false)
                        //        {
                        //            _HtmlContentBuilder.Append(text);
                        //        }
                        //        //while (true)
                        //        //{
                        //        //    int c = reader.InnerReader.Peek();
                        //        //    if (c < 0)
                        //        //    {
                        //        //        break;
                        //        //    }
                        //        //    if (c == '}')
                        //        //    {
                        //        //        break;
                        //        //    }
                        //        //    _HtmlContentBuilder.Append((char)c);
                        //        //    reader.InnerReader.Read();
                        //        //}
                        //    }
                        //    break;
                        //**************** read paragraph format ***********************
                        case RTFConsts._pard:
                            {
                                bolStartContent = true;
                                if (ForbitPard)
                                    continue;
                                // clear paragraph format
                                _ParagraphFormat.ResetParagraph();
                                //format.ResetParagraph();
                                break;
                            }
                        case RTFConsts._par:
                            {
                                bolStartContent = true;
                                // new paragraph
                                if (GetLastElement(typeof(RTFDomParagraph)) == null)
                                {
                                    var p = new RTFDomParagraph();
                                    p.Format = _ParagraphFormat ;
                                    _ParagraphFormat = _ParagraphFormat.Clone();
                                    AddContentElement(p);
                                    p.Locked = true;
                                }
                                else
                                {
                                    // ������ǰ����
                                    CompleteParagraph();
                                    // �����µĶ���
                                    var p = new RTFDomParagraph();
                                    p.Format = _ParagraphFormat;
                                    AddContentElement(p);
                                }
                                bolStartContent = true;
                                break;
                            }
                        case RTFConsts._page:
                            {
                                // ǿ�Ʒ�ҳ
                                bolStartContent = true;
                                CompleteParagraph();
                                AddContentElement(new RTFDomPageBreak());
                                break;
                            }
                        case RTFConsts._pagebb:
                            {
                                // �ڶ���ǰǿ�Ʒ�ҳ
                                bolStartContent = true;
                                _ParagraphFormat.PageBreak = true;
                                break;
                            }
                        case RTFConsts._ql:
                            {
                                // left alignment
                                bolStartContent = true;
                                _ParagraphFormat.Align = RTFAlignment.Left;
                                break;
                            }
                        case RTFConsts._qc:
                            {
                                // center alignment
                                bolStartContent = true;
                                _ParagraphFormat.Align = RTFAlignment.Center;
                                break;
                            }
                        case RTFConsts._qr:
                            {
                                // right alignment
                                bolStartContent = true;
                                _ParagraphFormat.Align = RTFAlignment.Right;
                                break;
                            }
                        case RTFConsts._qj:
                            {
                                // jusitify alignment
                                bolStartContent = true;
                                _ParagraphFormat.Align = RTFAlignment.Justify;
                                break;
                            }
                        case RTFConsts._sl:
                            {
                                // line spacing
                                bolStartContent = true;
                                if (reader.Parameter >= 0)
                                {
                                    _ParagraphFormat.LineSpacing = reader.Parameter;
                                }
                            }
                            break;
                        case RTFConsts._slmult:
                            {
                                bolStartContent = true;
                                _ParagraphFormat.MultipleLineSpacing = (reader.Parameter == 1);
                            }
                            break;
                        case RTFConsts._sb:
                            {
                                // spacing before paragraph
                                bolStartContent = true;
                                _ParagraphFormat.SpacingBefore = reader.Parameter;
                            }
                            break;
                        case RTFConsts._sa:
                            {
                                // spacing after paragraph
                                bolStartContent = true;
                                _ParagraphFormat.SpacingAfter = reader.Parameter;
                            }
                            break;
                        case RTFConsts._fi:
                            {
                                // indent first line
                                bolStartContent = true;
                                _ParagraphFormat.ParagraphFirstLineIndent = reader.Parameter;
                                //if (reader.Parameter >= 400)
                                //{
                                //    _ParagraphFormat.ParagraphFirstLineIndent = reader.Parameter; //doc.StandTabWidth;
                                //}
                                //else
                                //{
                                //    _ParagraphFormat.ParagraphFirstLineIndent = 0;
                                //}
                                ////UpdateParagraph( CurrentParagraphEOF , format );
                                break;
                            }
                        case RTFConsts._brdrw:
                            {
                                bolStartContent = true;
                                if (reader.HasParam)
                                {
                                    _ParagraphFormat.BorderWidth = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._pn :
                            {
                                bolStartContent = true;
                                _ParagraphFormat.ListID = -1;
                                break;
                            }

                        case RTFConsts._pntext:
                            break;
                        case RTFConsts._pntxtb:
                            break;
                        case RTFConsts._pntxta:
                            break;

                        case RTFConsts._pnlvlbody:
                            {
                                // numbered list style
                                bolStartContent = true; 
                                //_ParagraphFormat.NumberedList = true;
                                //_ParagraphFormat.BulletedList = false;
                                //if (_ParagraphFormat.Parent != null)
                                //{
                                //    _ParagraphFormat.Parent.NumberedList = format.NumberedList;
                                //    _ParagraphFormat.Parent.BulletedList = format.BulletedList;
                                //}
                                break;
                            }
                        case RTFConsts._pnlvlblt:
                            {
                                // bulleted list style
                                bolStartContent = true; 
                                //_ParagraphFormat.NumberedList = false;
                                //_ParagraphFormat.BulletedList = true;
                                //if (_ParagraphFormat.Parent != null)
                                //{
                                //    _ParagraphFormat.Parent.NumberedList = format.NumberedList;
                                //    _ParagraphFormat.Parent.BulletedList = format.BulletedList;
                                //}
                                break;
                            }
                        case RTFConsts._listtext:
                            {
                                bolStartContent = true; 
                                var txt = ReadInnerText(reader, true);
                                if (txt != null)
                                {
                                    txt = txt.Trim();
                                    if (txt.StartsWith("l"))
                                    {
                                        ListTextFlag = 1;
                                    }
                                    else
                                    {
                                        ListTextFlag = 2;
                                    }
                                }
                                break;
                            }
                        case RTFConsts._ls:
                            {
                                bolStartContent = true;
                                _ParagraphFormat.ListID = reader.Parameter;

                                //if (ListTextFlag == 1)
                                //{
                                //    _ParagraphFormat.NumberedList = false;
                                //    _ParagraphFormat.BulletedList = true;
                                //}
                                //else if (ListTextFlag == 2)
                                //{
                                //    _ParagraphFormat.NumberedList = true;
                                //    _ParagraphFormat.BulletedList = false;
                                //}
                                ListTextFlag = 0;
                                break;
                            }
                        case RTFConsts._li:
                            {
                                bolStartContent = true; 
                                if (reader.HasParam)
                                {
                                    _ParagraphFormat.LeftIndent = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._line:
                            {
                                bolStartContent = true;
                                // break line
                                //if (this._RTFHtmlState == false)
                                //{
                                //    this._HtmlContentBuilder.Append(Environment.NewLine);
                                //}
                                //else
                                {
                                    if (format.ReadText)
                                    {
                                        var line = new RTFDomLineBreak();
                                        line.NativeLevel = reader.Level;
                                        AddContentElement(line);
                                    }
                                }
                                break;
                            }
                        // ****************** read text format ******************************
                        case RTFConsts._insrsid :
                            break;
                        case RTFConsts._plain:
                            {
                                // clear text format
                                bolStartContent = true;
                                format.ResetText();
                                break;
                            }
                        case RTFConsts._f:
                            {
                                // font name
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    var FontName = FontTable.GetFontName(reader.Parameter);
                                    if (FontName != null)
                                        FontName = FontName.Trim();
                                    if (FontName == null || FontName.Length == 0)
                                        FontName = DefaultFontName;

                                    if (ChangeTimesNewRoman)
                                    {
                                        if (FontName == "Times New Roman")
                                        {
                                            FontName = DefaultFontName;
                                        }
                                    }
                                    format.FontName = FontName;
                                }
                                myFontChartset = FontTable[reader.Parameter].Encoding;
                                break;
                            }
                        case RTFConsts._af:
                            {
                                myAssociateFontChartset = FontTable[reader.Parameter].Encoding;
                                break;
                            }
                        case RTFConsts._fs:
                            {
                                // font size
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    if ( reader.HasParam)
                                    {
                                        format.FontSize = reader.Parameter / 2.0f;
                                    }
                                }
                                break;
                            }
                        case RTFConsts._cf:
                            {
                                // font color
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    if (reader.HasParam)
                                    {
                                        format.TextColor = ColorTable.GetColor( reader.Parameter, Color.Black);
                                    }
                                }
                                break;
                            }
                       
                        case RTFConsts._cb:
                        case RTFConsts._chcbpat:
                            {
                                // background color
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    if ( reader.HasParam )
                                    {
                                        format.BackColor = ColorTable.GetColor(reader.Parameter, Color.Empty);
                                    }
                                }
                                break;
                            }
                        case RTFConsts._b:
                            {
                                // bold
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Bold = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts._v :
                            {
                                // hidden text
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    if (reader.HasParam && reader.Parameter == 0)
                                    {
                                        format.Hidden = false;
                                    }
                                    else
                                    {
                                        format.Hidden = true;
                                    }
                                }
                                break;
                            }
                        case RTFConsts._highlight:
                            {
                                // highlight content
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    if (reader.HasParam)
                                    {
                                        format.BackColor = ColorTable.GetColor(
                                            reader.Parameter ,
                                            Color.Empty );
                                    }
                                }
                                break;
                            }
                        case RTFConsts._i:
                            {
                                // italic
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Italic = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts._ul:
                            {
                                // under line
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Underline = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts._noul:
                        {
                            // no underline
                            bolStartContent = true;
                            format.Underline = false;
                            
                            break;
                        }

                        case RTFConsts._strike:
                            {
                                // strikeout
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Strikeout = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts._sub:
                            {
                                // subscript
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Subscript = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts._super:
                            {
                                // superscript
                                bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Superscript = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts._nosupersub:
                            {
                                // nosupersub
                                bolStartContent = true;
                                format.Subscript = false;
                                format.Superscript = false;
                                break;
                            }
                        case RTFConsts._brdrb:
                            {
                                bolStartContent = true;
                                //format.ParagraphBorder.Bottom = true;
                                _ParagraphFormat.BottomBorder = true;
                                break;
                            }
                        case RTFConsts._brdrl:
                            {
                                bolStartContent = true;
                                //format.ParagraphBorder.Left = true ;
                                _ParagraphFormat.LeftBorder = true;
                                break;
                            }
                        case RTFConsts._brdrr:
                            {
                                bolStartContent = true;
                                //format.ParagraphBorder.Right = true ;
                                _ParagraphFormat.RightBorder = true;
                                break;
                            }
                        case RTFConsts._brdrt:
                            {
                                bolStartContent = true;
                                //format.ParagraphBorder.Top = true;
                                _ParagraphFormat.BottomBorder = true;
                                break;
                            }
                        case RTFConsts._brdrcf:
                            {
                                bolStartContent = true;
                                var element = GetLastElement( typeof( RTFDomTableRow ) , false );
                                if (element is RTFDomTableRow)
                                {
                                    // reading a table row
                                    var row = (RTFDomTableRow)element;
                                    RTFAttributeList style = null;
                                    if (row.CellSettings.Count > 0)
                                    {
                                        style = (RTFAttributeList)row.CellSettings[row.CellSettings.Count - 1];
                                        style.Add(reader.Keyword, reader.Parameter);
                                    }
                                    //else
                                    //{
                                    //    style = new RTFAttributeList();
                                    //    row.CellSettings.Add(style);
                                    //}
                                    
                                }
                                else
                                {
                                    _ParagraphFormat.BorderColor = ColorTable.GetColor(reader.Parameter , Color.Black);
                                    format.BorderColor = format.BorderColor;
                                }
                                break;
                            }
                        case RTFConsts._brdrs:
                            {
                                bolStartContent = true;
                                _ParagraphFormat.BorderThickness = false;
                                format.BorderThickness = false;
                                break;
                            }
                        case RTFConsts._brdrth:
                            {
                                bolStartContent = true;
                                _ParagraphFormat.BorderThickness = true;
                                format.BorderThickness = true;
                                break;
                            }
                        case RTFConsts._brdrdot:
                            {
                                bolStartContent = true;
                                _ParagraphFormat.BorderStyle= DashStyle.Dot;
                                format.BorderStyle = DashStyle.Dot;
                                break;
                            }
                        case RTFConsts._brdrdash:
                            {
                                bolStartContent = true;
                                _ParagraphFormat.BorderStyle = DashStyle.Dash;
                                format.BorderStyle = DashStyle.Dash;
                                break;
                            }
                        case RTFConsts._brdrdashd:
                            {
                                bolStartContent = true;
                                _ParagraphFormat.BorderStyle = DashStyle.DashDot;
                                format.BorderStyle = DashStyle.DashDot;
                                break;
                            }
                        case RTFConsts._brdrdashdd:
                            {
                                bolStartContent = true;
                                _ParagraphFormat.BorderStyle = DashStyle.DashDotDot;
                                format.BorderStyle = DashStyle.DashDotDot;
                                break;
                            }
                        case RTFConsts._brdrnil:
                            {
                                bolStartContent = true;
                                _ParagraphFormat.LeftBorder = false;
                                _ParagraphFormat.TopBorder = false;
                                _ParagraphFormat.RightBorder = false;
                                _ParagraphFormat.BottomBorder = false;

                                format.LeftBorder = false;
                                format.TopBorder = false;
                                format.RightBorder = false;
                                format.BottomBorder = false;
                                break;
                            }
                        case RTFConsts._brsp:
                            {
                                bolStartContent = true;
                                if (reader.HasParam)
                                {
                                    _ParagraphFormat.BorderSpacing = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._chbrdr:
                            {
                                bolStartContent = true;
                                format.LeftBorder = true;
                                format.TopBorder = true;
                                format.RightBorder = true;
                                format.BottomBorder = true;
                                break;
                            }
                        case RTFConsts._bkmkstart:
                            {
                                // book mark
                                bolStartContent = true;
                                if (format.ReadText && bolStartContent)
                                {
                                    var bk = new RTFDomBookmark();
                                    bk.Name = ReadInnerText(reader, true);
                                    bk.Locked = true;
                                    AddContentElement(bk);
                                }
                                break;
                            }
                        case RTFConsts._bkmkend:
                            {
                                ForbitPard = true;
                                format.ReadText = false;
                                break;
                            }
                        case RTFConsts._field:
                            {
                                // field
                                bolStartContent = true;
                                ReadDomField(reader, format);
                                return; // finish current level
                                //break;
                            }
                        
                        //case RTFConsts._objdata:
                        //case RTFConsts._objclass:
                        //    {
                        //        ReadToEndGround(reader);
                        //        break;
                        //    }

                        #region read object *********************************

                        case RTFConsts._object:
                            {
                                // object
                                bolStartContent = true;
                                ReadDomObject(reader, format);
                                return;// finish current level
                            }

                        #endregion

                        #region read image **********************************

                        case RTFConsts._shppict :
                            // continue the following token
                            break;
                        case RTFConsts._nonshppict :
                            // unsupport keyword
                            ReadToEndGround(reader);
                            break;
                        case RTFConsts._pict:
                            {
                                // read image data
                                //ReadDomImage(reader, format);
                                bolStartContent = true;
                                var img = new RTFDomImage();
                                img.NativeLevel = reader.Level;
                                AddContentElement(img);
                            }
                            break;
                        case RTFConsts._picscalex:
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.ScaleX = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._picscaley:
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.ScaleY = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._picwgoal:
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.DesiredWidth = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._pichgoal:
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.DesiredHeight = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._blipuid:
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.ID = ReadInnerText(reader, true);
                                }
                                break;
                            }
                        case RTFConsts._emfblip :
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Emfblip;
                                }
                                break;
                            }
                        case RTFConsts._pngblip :
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Pngblip;
                                }
                                break;
                            }
                        case RTFConsts._jpegblip :
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Jpegblip;
                                }
                                break;
                            }
                        case RTFConsts._macpict :
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Macpict;
                                }
                                break;
                            }
                        case RTFConsts._pmmetafile:
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Pmmetafile;
                                }
                                break;
                            }
                        case RTFConsts._wmetafile :
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Wmetafile;
                                }
                                break;
                            }
                        case RTFConsts._dibitmap :
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Dibitmap;
                                }
                                break;
                            }
                        case RTFConsts._wbitmap :
                            {
                                var img = (RTFDomImage)GetLastElement(typeof(RTFDomImage));
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Wbitmap;
                                }
                                break;
                            }
                        #endregion

                        #region read shape ************************************************
                        case RTFConsts._sp:
                            {
                                // begin read shape property
                                var level = 0;
                                string vName = null;
                                string vValue = null;
                                while (reader.ReadToken() != null)
                                {
                                    if (reader.TokenType == RTFTokenType.GroupStart)
                                    {
                                        level++;
                                    }
                                    else if (reader.TokenType == RTFTokenType.GroupEnd)
                                    {
                                        level--;
                                        if (level < 0)
                                        {
                                            break;
                                        }
                                    }
                                    else if (reader.Keyword == RTFConsts._sn)
                                    {
                                        vName = ReadInnerText(reader, true);
                                    }
                                    else if (reader.Keyword == RTFConsts._sv)
                                    {
                                        vValue = ReadInnerText(reader, true);
                                    }
                                }//while
                                var shape = (RTFDomShape)GetLastElement(typeof(RTFDomShape));
                                if (shape != null)
                                {
                                    shape.ExtAttrbutes[vName] = vValue;
                                }
                                else
                                {
                                    var g = (RTFDomShapeGroup)GetLastElement(typeof(RTFDomShapeGroup));
                                    if (g != null)
                                    {
                                        g.ExtAttrbutes[vName] = vValue;
                                    }
                                }
                                break;
                            }
                        case RTFConsts._shptxt:
                            {
                                // handle following token
                                break;
                            }
                        case RTFConsts._shprslt:
                            {
                                // ignore this level
                                ReadToEndGround( reader );
                                break  ;
                            }
                        case RTFConsts._shp:
                            {
                                bolStartContent = true;
                                var shape = new RTFDomShape();
                                shape.NativeLevel = reader.Level;
                                AddContentElement(shape);
                                break;
                            }
                        case RTFConsts._shpleft:
                            {
                                var shape = (RTFDomShape)GetLastElement(typeof(RTFDomShape));
                                if (shape != null)
                                {
                                    shape.Left = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._shptop:
                            {
                                var shape = (RTFDomShape)GetLastElement(typeof(RTFDomShape));
                                if (shape != null)
                                {
                                    shape.Top = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._shpright:
                            {
                                var shape = (RTFDomShape)GetLastElement(typeof(RTFDomShape));
                                if (shape != null)
                                {
                                    shape.Width = reader.Parameter - shape.Left;
                                }
                                break;
                            }
                        case RTFConsts._shpbottom:
                            {
                                var shape = (RTFDomShape)GetLastElement(typeof(RTFDomShape));
                                if (shape != null)
                                {
                                    shape.Height = reader.Parameter - shape.Top;
                                }
                                break;
                            }
                        case RTFConsts._shplid:
                            {
                                var shape = (RTFDomShape)GetLastElement(typeof(RTFDomShape));
                                if (shape != null)
                                {
                                    shape.ShapeID = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._shpz:
                            {
                                var shape = (RTFDomShape)GetLastElement(typeof(RTFDomShape));
                                if (shape != null)
                                {
                                    shape.ZIndex = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts._shpgrp:
                            {
                                var group = new RTFDomShapeGroup();
                                group.NativeLevel = reader.Level;
                                AddContentElement(group);
                                break;
                            }
                        case RTFConsts._shpinst:
                            {
                                break;
                            }
                        #endregion

                        #region read table ************************************************
                        case RTFConsts._intbl:
                        case RTFConsts._trowd:
                        case RTFConsts._itap:
                            {
                                // these keyword said than current paragraph is table row
                                bolStartContent = true;
                                var es = GetLastElements(true);
                                RTFDomElement lastUnlockElement = null;
                                RTFDomElement lastTableElement = null;
                                for (var iCount = es.Length - 1; iCount >= 0; iCount--)
                                {
                                    var e = es[iCount];
                                    if (e.Locked == false)
                                    {
                                        if (lastUnlockElement == null && !(e is RTFDomParagraph))
                                        {
                                            lastUnlockElement = e;
                                        }
                                        if (e is RTFDomTableRow || e is RTFDomTableCell)
                                        {
                                            lastTableElement = e;
                                            break;
                                        }
                                    }
                                }
                                if ( reader.Keyword == RTFConsts._intbl)
                                {
                                    if (lastTableElement == null)
                                    {
                                        // if can not find unlocked row 
                                        // then new row
                                        var row = new RTFDomTableRow();
                                        row.NativeLevel = reader.Level;
                                        lastUnlockElement.AppendChild(row);
                                    }
                                }
                                else if ( reader.Keyword == RTFConsts._trowd)
                                {
                                    // clear row format
                                    RTFDomTableRow row = null;
                                    if (lastTableElement == null)
                                    {
                                        row = new RTFDomTableRow();
                                        row.NativeLevel = reader.Level;
                                        lastUnlockElement.AppendChild(row);
                                    }
                                    else
                                    {
                                        row = lastTableElement as RTFDomTableRow;
                                        if (row == null)
                                        {
                                            row = (RTFDomTableRow)lastTableElement.Parent;
                                        }
                                    }
                                    row.Attributes.Clear();
                                    row.CellSettings.Clear();
                                    _ParagraphFormat.ResetParagraph();
                                }
                                else if (reader.Keyword == RTFConsts._itap)
                                {
                                    // set nested level
                                    RTFDomTableRow row = null;

                                    if (reader.Parameter == 0)
                                    {
                                        // is the 0 level , belong to document , not to a table
                                        // ������������ \itap0���ܲ��������������� \cellָ��Ϊ׼
                                        //foreach (RTFDomElement element in es)
                                        //{
                                        //    if (element is RTFDomTableRow || element is RTFDomTableCell)
                                        //    {
                                        //        element.Locked = true;
                                        //    }
                                        //}
                                    }
                                    else
                                    {
                                        // in a row
                                        if (lastTableElement == null)
                                        {
                                            row = new RTFDomTableRow();
                                            row.NativeLevel = reader.Level;
                                            lastUnlockElement.AppendChild(row);
                                        }
                                        else
                                        {
                                            row = lastTableElement as RTFDomTableRow;
                                            if (row == null)
                                            {
                                                row = (RTFDomTableRow)lastTableElement.Parent;
                                            }
                                            //row.Attributes.Clear();
                                            //row.CellSettings = new ArrayList();
                                        }
                                        if ( reader.Parameter == row.Level)
                                        {
                                        }
                                        else if ( reader.Parameter > row.Level)
                                        {
                                            // nested row
                                            var newRow = new RTFDomTableRow();
                                            newRow.Level = reader.Parameter;
                                            var parentCell = (RTFDomTableCell)GetLastElement(typeof(RTFDomTableCell), false);
                                            if (parentCell == null)
                                                AddContentElement(newRow);
                                            else
                                                parentCell.AppendChild(newRow);
                                        }
                                        else if (reader.Parameter < row.Level)
                                        {
                                            // exit nested row
                                        }
                                    }
                                }//else if
                                break;
                            }
                        case RTFConsts._nesttableprops:
                            {
                                // ignore
                                break;
                            }
                        case RTFConsts._row:
                            {
                                // finish read row
                                bolStartContent = true;
                                var es = GetLastElements(true);
                                for (var iCount = es.Length - 1; iCount >= 0; iCount--)
                                {
                                    es[iCount].Locked = true;
                                    if (es[iCount] is RTFDomTableRow)
                                    {
                                        break;
                                    }
                                }
                                break;
                            }
                        case RTFConsts._nestrow:
                            {
                                // finish nested row
                                bolStartContent = true;
                                var es = GetLastElements(true);
                                for (var iCount = es.Length - 1; iCount >= 0; iCount--)
                                {
                                    es[iCount].Locked = true;
                                    if (es[iCount] is RTFDomTableRow)
                                    {
                                        break;
                                    }
                                }
                                break;
                            }
                        case RTFConsts._trrh:
                        case RTFConsts._trautofit :
                        case RTFConsts._irowband :
                        case RTFConsts._trhdr:
                        case RTFConsts._trkeep:
                        case RTFConsts._trkeepfollow :
                        case RTFConsts._trleft:
                        case RTFConsts._trqc :
                        case RTFConsts._trql :
                        case RTFConsts._trqr :
                        case RTFConsts._trcbpat:
                        case RTFConsts._trcfpat:
                        case RTFConsts._trpat:
                        case RTFConsts._trshdng:
                        case RTFConsts._trwWidth:
                        case RTFConsts._trwWidthA:
                        case RTFConsts._irow:
                        case RTFConsts._trpaddb:
                        case RTFConsts._trpaddl:
                        case RTFConsts._trpaddr:
                        case RTFConsts._trpaddt:
                        case RTFConsts._trpaddfb :
                        case RTFConsts._trpaddfl :
                        case RTFConsts._trpaddfr :
                        case RTFConsts._trpaddft :
                        case RTFConsts._lastrow:
                            {
                                // meet row control word , not parse at first , just save it 
                                bolStartContent = true;
                                var row = (RTFDomTableRow)GetLastElement(typeof(RTFDomTableRow), false);
                                if (row != null)
                                {
                                    row.Attributes.Add( reader.Keyword , reader.Parameter );
                                }
                                break;
                            }
                        case RTFConsts._clvmgf:
                        case RTFConsts._clvmrg:
                        case RTFConsts._cellx:
                        case RTFConsts._clvertalt:
                        case RTFConsts._clvertalc:
                        case RTFConsts._clvertalb:
                        case RTFConsts._clNoWrap:
                        case RTFConsts._clcbpat:
                        case RTFConsts._clcfpat:
                        case RTFConsts._clpadl:
                        case RTFConsts._clpadt:
                        case RTFConsts._clpadr:
                        case RTFConsts._clpadb:
                        case RTFConsts._clbrdrl:
                        case RTFConsts._clbrdrt:
                        case RTFConsts._clbrdrr:
                        case RTFConsts._clbrdrb:
                        case RTFConsts._brdrtbl :
                        case RTFConsts._brdrnone :
                            {
                                // meet cell control word , no parse at first , just save it
                                bolStartContent = true;
                                var row = (RTFDomTableRow)GetLastElement(typeof(RTFDomTableRow), false);
                                //if (row != null && row.Locked == false )
                                {
                                    RTFAttributeList style = null;
                                    if (row.CellSettings.Count > 0)
                                    {
                                        style = (RTFAttributeList)row.CellSettings[row.CellSettings.Count - 1];
                                        if (style.Contains(RTFConsts._cellx))
                                        {
                                            // if find repeat control word , then can consider this control word
                                            // belong to the next cell . userly cellx is the last control word of 
                                            // a cell , when meet cellx , the current cell defind is finished.
                                            style = new RTFAttributeList();
                                            row.CellSettings.Add(style);
                                        }
                                    }
                                    if (style == null)
                                    {
                                        style = new RTFAttributeList();
                                        row.CellSettings.Add(style);
                                    }
                                    style.Add(reader.Keyword, reader.Parameter);
                                    
                                }
                                break;
                            }
                        case RTFConsts._cell:
                            {
                                // finish cell content
                                bolStartContent = true;
                                AddContentElement(null);
                                CompleteParagraph();
                                _ParagraphFormat.Reset();
                                format.Reset();
                                var es = GetLastElements(true );
                                for (var iCount = es.Length - 1; iCount >= 0; iCount--)
                                {
                                    if (es[iCount].Locked == false)
                                    {
                                        es[iCount].Locked = true;
                                        if (es[iCount] is RTFDomTableCell)
                                        {
                                            //((RTFDomTableCell)es[iCount]).Format = format.Clone();
                                            break;
                                        }
                                    }
                                }
                                break;
                            }
                        case RTFConsts._nestcell:
                            {
                                // finish nested cell content
                                bolStartContent = true;
                                AddContentElement(null);
                                CompleteParagraph();
                                var es = GetLastElements( false );
                                for (var iCount = es.Length - 1; iCount >= 0; iCount--)
                                {
                                    es[iCount].Locked = true;
                                    if (es[iCount] is RTFDomTableCell)
                                    {
                                        ((RTFDomTableCell)es[iCount]).Format = format;
                                        break;
                                    }
                                }
                                break;
                            }
                        #endregion
                        default :
                            // unsupport keyword
                            if (reader.TokenType == RTFTokenType.ExtKeyword
                                && reader.FirstTokenInGroup )
                            {
                                // if meet unsupport extern keyword , and this token is the first token in 
                                // current group , then ingore whole group.
                                ReadToEndGround(reader);
                                break ;
                            }
                            break;
                    }//switch
                }
                
            }//while
            if (myText.HasContent)
            {
                ApplyText(myText, reader, format);
            }
        }
         

        /// <summary>
        /// read data , until at the front of the end token belong the current level.
        /// </summary>
        /// <param name="reader"></param>
        private void ReadToEndGround( RTFReader reader )
        {
            reader.ReadToEndGround( );
        }


        private void ReadListOverrideTable(RTFReader reader)
        {
            _ListOverrideTable = new RTFListOverrideTable();
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    break;
                }
                else if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    // ��ȡһ����¼
                    var level = reader.Level;
                    RTFListOverride record = null;
                    while (reader.ReadToken() != null)
                    {
                        if (reader.TokenType == RTFTokenType.GroupEnd)
                        {
                            break;
                        }
                        if (reader.CurrentToken.Key == "listoverride")
                        {
                            record = new RTFListOverride();
                            _ListOverrideTable.Add(record);
                            continue;
                        }
                        if (record == null)
                        {
                            continue;
                        }
                        switch (reader.CurrentToken.Key)
                        {
                            case "listid":
                                record.ListID = reader.CurrentToken.Param;
                                break;
                            case "listoverridecount":
                                record.ListOverriedCount = reader.CurrentToken.Param;
                                break;
                            case "ls":
                                record.ID = reader.CurrentToken.Param;
                                break;
                        }
                    }
                }
            }//while
        }

        #region HTML RTF 

        
        private string _HtmlContent = null;

        /// <summary>
        /// HTML content in RTF
        /// </summary>
        public string HtmlContent
        {
            get { return _HtmlContent; }
            set { _HtmlContent = value; }
        }

        private void ReadHtmlContent(RTFReader reader)
        {
            var htmlStr = new StringBuilder();
            var htmlState = true;
            while (reader.ReadToken() != null)
            {
                if (reader.Keyword == "htmlrtf")
                {
                    if (reader.HasParam && reader.Parameter == 0)
                    {
                        htmlState = false;
                    }
                    else
                    {
                        htmlState = true;
                    }
                }
                else if (reader.Keyword == "htmltag")
                {
                    if (reader.InnerReader.Peek() == (int)' ')
                    {
                        reader.InnerReader.Read();
                    }
                    var text = ReadInnerText(reader, null, true, false, true);
                    if (string.IsNullOrEmpty(text) == false)
                    {
                        htmlStr.Append(text);
                    }
                }
                else if (reader.TokenType == RTFTokenType.Keyword
                    || reader.TokenType == RTFTokenType.ExtKeyword)
                {
                    if (htmlState == false)
                    {
                        switch (reader.Keyword)
                        {
                            case "par":
                                htmlStr.Append(Environment.NewLine);
                                break;
                            case "line":
                                htmlStr.Append(Environment.NewLine);
                                break;
                            case "tab":
                                htmlStr.Append("\t");
                                break;
                            case "lquote":
                                htmlStr.Append("&lsquo;");
                                break;
                            case "rquote":
                                htmlStr.Append("&rsquo;");
                                break;
                            case "ldblquote":
                                htmlStr.Append("&ldquo;");
                                break;
                            case "rdblquote":
                                htmlStr.Append("&rdquo;");
                                break;
                            case "bullet":
                                htmlStr.Append("&bull;");
                                break;
                            case "endash":
                                htmlStr.Append("&ndash;");
                                break;
                            case "emdash":
                                htmlStr.Append("&mdash;");
                                break;
                            case "~":
                                htmlStr.Append("&nbsp;");
                                break;
                            case "_":
                                htmlStr.Append("&shy;");
                                break;
                        }
                    }
                }
                else if (reader.TokenType == RTFTokenType.Text)
                {
                    if (htmlState == false)
                    {
                        htmlStr.Append(reader.Keyword);
                    }
                }
            }//while
            HtmlContent = htmlStr.ToString();
        }

        #endregion

        private void ReadListTable(RTFReader reader)
        {
            _ListTable = new RTFListTable();
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    break;
                }
                else if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    var firstRead = true;
                    RTFList currentList = null;
                    var level = reader.Level ;
                    while (reader.ReadToken() != null)
                    {
                        if (reader.TokenType == RTFTokenType.GroupEnd)
                        {
                            if( reader.Level < level )
                            {
                                break;
                            }
                        }
                        else if (reader.TokenType == RTFTokenType.GroupStart)
                        {
                            // if meet nested level , then ignore
                            //reader.ReadToken();
                            //ReadToEndGround(reader);
                            //reader.ReadToken();
                        }
                        if (firstRead)
                        {
                            if (reader.CurrentToken.Key != "list")
                            {
                                // ������list��ͷ�����Ե�
                                ReadToEndGround(reader);
                                reader.ReadToken();
                                break;
                            }
                            currentList = new RTFList();
                            _ListTable.Add(currentList);
                            firstRead = false;
                        }
                        switch (reader.CurrentToken.Key)
                        {
                            case "listtemplateid":
                                currentList.ListTemplateID = reader.CurrentToken.Param;
                                break;
                            case "listid":
                                currentList.ListID = reader.CurrentToken.Param;
                                break;
                            case "listhybrid":
                                currentList.ListHybrid = true;
                                break;
                            case "levelfollow":
                                currentList.LevelFollow = reader.CurrentToken.Param;
                                break;
                            case "levelstartat":
                                currentList.LevelStartAt = reader.CurrentToken.Param;
                                break;
                            case "levelnfc":
                                if (currentList.LevelNfc == LevelNumberType.None)
                                {
                                    currentList.LevelNfc = (LevelNumberType)reader.CurrentToken.Param;
                                }
                                break;
                            case "levelnfcn":
                                 if (currentList.LevelNfc == LevelNumberType.None)
                                {
                                    currentList.LevelNfc = (LevelNumberType)reader.CurrentToken.Param;
                                }
                                break;
                            case "leveljc":
                                currentList.LevelJc = reader.CurrentToken.Param;
                                break;
                            case "leveltext":
                                //if (currentList.LevelNfc == LevelNumberType.Bullet)
                                {
                                    if (string.IsNullOrEmpty(currentList.LevelText))
                                    {
                                        var text = ReadInnerText(reader, true);
                                        if( text != null && text.Length > 2 )
                                        {
                                            var len = (int)text[0];
                                            len = Math.Min(len, text.Length - 1);
                                            text = text.Substring(1, len);
                                        }
                                        currentList.LevelText = text ;
                                    }
                                }
                                break;
                            case "f":
                                currentList.FontName = FontTable.GetFontName(reader.CurrentToken.Param);
                                break;
                        }
                    }//while
                }
            }//while
        }
         

        /// <summary>
        /// read font table
        /// </summary>
        /// <param name="group"></param>
        private void ReadFontTable(RTFReader reader)
        {
            myFontTable.Clear();
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    break;
                }
                else if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    var index = -1;
                    string name = null;
                    var charset = 1;
                    var nilFlag = false;
                    while (reader.ReadToken() != null)
                    {
                        if (reader.TokenType == RTFTokenType.GroupEnd)
                        {
                            break;
                        }
                        else if (reader.TokenType == RTFTokenType.GroupStart)
                        {
                            // if meet nested level , then ignore
                            reader.ReadToken();
                            ReadToEndGround(reader);
                            reader.ReadToken();
                        }
                        else if (reader.Keyword == "f" && reader.HasParam)
                        {
                            index = reader.Parameter;
                        }
                        else if (reader.Keyword == "fnil")
                        {
                            name = Constants.DefaultFontName;
                            nilFlag = true;
                        }
                        else if (reader.Keyword == RTFConsts._fcharset)
                        {
                            charset = reader.Parameter;
                        }
                        else if (reader.CurrentToken.IsTextToken)
                        {
                            //if (defaultFont == false)
                            {
                                name = ReadInnerText(
                                    reader,
                                    reader.CurrentToken, 
                                    false, 
                                    false , 
                                    false );
                                if (name != null)
                                {
                                    name = name.Trim();

                                    if (name.EndsWith(";"))
                                    {
                                        name = name.Substring(0, name.Length - 1);
                                    }
                                }
                            }
                        }
                    }
                    if (index >= 0 && name != null)
                    {
                        if (name.EndsWith(";"))
                        {
                            name = name.Substring(0, name.Length - 1);
                        }
                        name = name.Trim();
                        if (string.IsNullOrEmpty(name))
                        {
                            name = Constants.DefaultFontName;
                        }
                        //System.Console.WriteLine( "Index:" + index + "  Name:" + name );
                        var font = new RTFFont(index, name);
                        font.Charset = charset;
                        font.NilFlag = nilFlag;
                        myFontTable.Add(font);
                    }
                }//else
            }//while
        }

        /// <summary>
        /// read color table
        /// </summary>
        /// <param name="group"></param>
        private void ReadColorTable(RTFReader reader)
        {
            myColorTable.Clear();
            myColorTable.CheckValueExistWhenAdd = false;
            var r = -1;
            var g = -1;
            var b = -1;
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    break;
                }
                switch (reader.Keyword)
                {
                    case "red":
                        r = reader.Parameter;
                        break;
                    case "green":
                        g = reader.Parameter;
                        break;
                    case "blue":
                        b = reader.Parameter;
                        break;
                    case ";":
                        if (r >= 0 && g >= 0 && b >= 0)
                        {
                            var c = Color.FromArgb(255, r, g, b);
                            myColorTable.Add(c);
                            r = -1;
                            g = -1;
                            b = -1;
                        }
                        break;
                }
            }
            if (r >= 0 && g >= 0 && b >= 0)
            {
                // read the last color
                var c = Color.FromArgb(255, r, g, b);
                myColorTable.Add(c);
            }
        }

        /// <summary>
        /// read document information
        /// </summary>
        /// <param name="group"></param>
        private void ReadDocumentInfo(RTFReader reader)
        {
            myInfo.Clear();
            var level = 0;
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    level++;
                }
                else if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    level--;
                    if (level < 0)
                    {
                        break;
                    }
                }
                else
                {
                    switch (reader.Keyword)
                    {
                        case "creatim":
                            myInfo.Creatim = ReadDateTime(reader);
                            level--;
                            break;
                        case "revtim":
                            myInfo.Revtim = ReadDateTime(reader);
                            level--;
                            break;
                        case "printim":
                            myInfo.Printim = ReadDateTime(reader);
                            level--;
                            break;
                        case "buptim":
                            myInfo.Buptim = ReadDateTime(reader);
                            level--;
                            break;
                        default:
                            if (reader.Keyword != null)
                            {
                                if (reader.HasParam)
                                {
                                    myInfo.SetInfo(reader.Keyword, reader.Parameter.ToString());
                                }
                                else
                                {
                                    myInfo.SetInfo(reader.Keyword, ReadInnerText(reader, true));
                                }
                            }
                            break;
                    }
                }
            }//while
        }

        /// <summary>
        /// read datetime
        /// </summary>
        /// <param name="reader">reader</param>
        /// <returns>datetime value</returns>
        private DateTime ReadDateTime(RTFReader reader)
        {
            var yr = 1900;
            var mo = 1;
            var dy = 1;
            var hr = 0;
            var min = 0;
            var sec = 0;
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    break;
                }
                switch (reader.Keyword)
                {
                    case "yr":
                        yr = reader.Parameter;
                        break;
                    case "mo":
                        mo = reader.Parameter;
                        break;
                    case "dy":
                        dy = reader.Parameter;
                        break;
                    case "hr":
                        hr = reader.Parameter;
                        break;
                    case "min":
                        min = reader.Parameter;
                        break;
                    case "sec":
                        sec = reader.Parameter;
                        break;
                }//switch
            }//while

            if (yr < 1 || mo < 1 || dy < 1 || hr < 0 || min < 0 || sec < 0)
                return DateTime.MinValue;
            
            return new DateTime(yr, mo, dy, hr, min, sec);
        }

        //private RTFDomImage ReadDomImage(RTFReader reader, DocumentFormatInfo format)
        //{
        //    bolStartContent = true;
        //    RTFDomImage img = new RTFDomImage();
        //    img.NativeLevel = reader.Level;
        //    this.AddContentElement(img);
        //    RTFTextContainer txt = new RTFTextContainer( this );
        //    while (reader.ReadToken() != null)
        //    {
        //        if (reader.Level < img.NativeLevel)
        //        {
        //            break;
        //        }
        //        if (reader.TokenType == RTFTokenType.GroupStart)
        //        {
        //            continue;
        //        }
        //        if (reader.TokenType == RTFTokenType.GroupEnd)
        //        {
        //            continue;
        //        }
        //        if (reader.TokenType == RTFTokenType.Text)
        //        {
        //            txt.Accept(reader.CurrentToken, reader);
        //            continue;
        //        }
        //        switch (reader.Keyword)
        //        {
        //            case RTFConsts._nonshppict :
        //                // ignore group
        //                this.ReadToEndGround(reader);
        //                break;
        //            case RTFConsts._picscalex:
        //                img.ScaleX = reader.Parameter;
        //                break;
        //            case RTFConsts._picscaley:
        //                img.ScaleY = reader.Parameter;
        //                break;
        //            case RTFConsts._picwgoal:
        //                img.DesiredWidth = reader.Parameter;
        //                break;
        //            case RTFConsts._pichgoal:
        //                img.DesiredHeight = reader.Parameter;
        //                break;
        //            case RTFConsts._blipuid:
        //                img.ID = ReadInnerText(reader, true);
        //                break;
        //            case RTFConsts._emfblip:
        //                img.PicType = RTFPicType.Emfblip;
        //                break;
        //            case RTFConsts._pngblip:
        //                img.PicType = RTFPicType.Pngblip;
        //                break;
        //            case RTFConsts._jpegblip:
        //                img.PicType = RTFPicType.Jpegblip;
        //                break;
        //            case RTFConsts._macpict:
        //                img.PicType = RTFPicType.Macpict;
        //                break;
        //            case RTFConsts._pmmetafile:
        //                img.PicType = RTFPicType.Pmmetafile;
        //                break;
        //            case RTFConsts._wmetafile:
        //                img.PicType = RTFPicType.Wmetafile;
        //                break;
        //            case RTFConsts._dibitmap:
        //                img.PicType = RTFPicType.Dibitmap;
        //                break;
        //            case RTFConsts._wbitmap:
        //                img.PicType = RTFPicType.Wbitmap;
        //                break;
        //        }//switch
        //    }//while
        //    if (txt.HasContent)
        //    {
        //        img.Data = HexToBytes(txt.Text);
        //    }
        //    img.Format = format.Clone();
        //    img.Width = (int)(img.DesiredWidth * img.ScaleX / 100);
        //    img.Height = (int)(img.DesiredHeight * img.ScaleY / 100);
        //    img.Locked = true;
        //    return img;
        //}
            
        /// <summary>
        /// Read a rtf emb object
        /// </summary>
        /// <param name="reader">reader</param>
        /// <param name="format">format</param>
        /// <returns>rtf emb object instance</returns>
        private RTFDomObject ReadDomObject(RTFReader reader, DocumentFormatInfo format)
        {
            var obj = new RTFDomObject();
            obj.NativeLevel = reader.Level;
            AddContentElement(obj);
            var levelBack = reader.Level;
            while (reader.ReadToken() != null)
            {
                if (reader.Level < levelBack)
                {
                    break;
                }
                if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    continue;
                }
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    continue;
                }
                if (reader.Level == obj.NativeLevel + 1
                    && reader.Keyword.StartsWith("attribute_"))
                {
                    obj.CustomAttributes[reader.Keyword] = ReadInnerText(reader, true);
                }
                switch (reader.Keyword)
                {
                    case RTFConsts._objautlink:
                        obj.Type = RTFObjectType.AutLink;
                        break;
                    case RTFConsts._objclass:
                        obj.ClassName = ReadInnerText(reader, true);
                        break;
                    case RTFConsts._objdata:
                        var data = ReadInnerText(reader, true);
                        obj.Content = HexToBytes(data);
                        break;
                    case RTFConsts._objemb:
                        obj.Type = RTFObjectType.EMB;
                        break;
                    case RTFConsts._objh:
                        obj.Height = reader.Parameter;
                        break;
                    case RTFConsts._objhtml:
                        obj.Type = RTFObjectType.Html;
                        break;
                    case RTFConsts._objicemb:
                        obj.Type = RTFObjectType.Icemb;
                        break;
                    case RTFConsts._objlink:
                        obj.Type = RTFObjectType.Link;
                        break;
                    case RTFConsts._objname:
                        obj.Name = ReadInnerText(reader, true);
                        break;
                    case RTFConsts._objocx:
                        obj.Type = RTFObjectType.Ocx;
                        break;
                    case RTFConsts._objpub:
                        obj.Type = RTFObjectType.Pub;
                        break;
                    case RTFConsts._objsub:
                        obj.Type = RTFObjectType.Sub;
                        break;
                    case RTFConsts._objtime:
                        break;
                    case RTFConsts._objw:
                        obj.Width = reader.Parameter;
                        break;
                    case RTFConsts._objscalex:
                        obj.ScaleX = reader.Parameter;
                        break;
                    case RTFConsts._objscaley:
                        obj.ScaleY = reader.Parameter;
                        break;
                    case RTFConsts._result :
                        // ��ȡ��������������
                        var result = new RTFDomElementContainer();
                        result.Name = RTFConsts._result;
                        obj.AppendChild(result);
                        Load(reader, format);
                        result.Locked = true;
                        break;
                }
            }//while
            obj.Locked = true;
            return obj;
        }

        /// <summary>
        /// read field
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        private RTFDomField ReadDomField(
            RTFReader reader,
            DocumentFormatInfo format  )
        {
            var field = new RTFDomField();
            field.NativeLevel = reader.Level;
            AddContentElement(field);
            var levelBack = reader.Level;
            while (reader.ReadToken() != null)
            {
                if (reader.Level < levelBack)
                {
                    break;
                }

                if (reader.TokenType == RTFTokenType.GroupStart)
                {
                }
                else if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    
                }
                else
                {
                    switch (reader.Keyword)
                    {
                        case RTFConsts._flddirty:
                            {
                                field.Method = RTFDomFieldMethod.Dirty;
                                break;
                            }
                        case RTFConsts._fldedit:
                            {
                                field.Method = RTFDomFieldMethod.Edit;
                                break;
                            }
                        case RTFConsts._fldlock:
                            {
                                field.Method = RTFDomFieldMethod.Lock;
                                break;
                            }
                        case RTFConsts._fldpriv:
                            {
                                field.Method = RTFDomFieldMethod.Priv;
                                break;
                            }
                        case RTFConsts._fldrslt:
                            {
                                var result = new RTFDomElementContainer();
                                result.Name = RTFConsts._fldrslt;
                                field.AppendChild(result);
                                Load(reader, format );
                                result.Locked = true;
                                //field.Result = ReadInnerText(reader, true);
                                break;
                            }
                        case RTFConsts._fldinst:
                            {
                                var inst = new RTFDomElementContainer();
                                inst.Name = RTFConsts._fldinst;
                                field.AppendChild(inst);
                                Load(reader, format );
                                inst.Locked = true;
                                var txt = inst.InnerText;
                                if (txt != null)
                                {
                                    var index = txt.IndexOf(RTFConsts._HYPERLINK);
                                    if (index >= 0)
                                    {
                                        string link = null;
                                        var index1 = txt.IndexOf('\"', index);
                                        if (index1 > 0 && txt.Length > index1 + 2)
                                        {
                                            var index2 = txt.IndexOf('\"', index1 + 2);
                                            if (index2 > index1)
                                            {
                                                link = txt.Substring(index1 + 1, index2 - index1 - 1);
                                                if (format.Parent != null)
                                                {
                                                    if (link.StartsWith("_Toc"))
                                                        link = "#" + link;
                                                    format.Parent.Link = link;
                                                }
                                                break;
                                            }
                                        }//if
                                    }//if
                                }

                                break;
                            }
                    }//switch
                }
            }//while
            field.Locked = true;
            return field;

        }//private RTFDomField ReadDomField(RTFReader reader)


        private string ReadInnerText(RTFReader reader, bool deeply)
        {
            return ReadInnerText(
                reader,
                null,
                deeply, 
                false , 
                false );
        }

        /// <summary>
        /// read the following plain text in the current level
        /// </summary>
        /// <param name="reader">RTF reader</param>
        /// <param name="deeply">whether read the text in the sub level</param>
        /// <returns>text</returns>
        private string ReadInnerText(
            RTFReader reader,
            RTFToken firstToken,
            bool deeply, 
            bool breakMeetControlWord,
            bool htmlMode )
        {
            var level = 0;
            var container = new RTFTextContainer(this);
            container.Accept(firstToken , reader );
            while (true)
            {
                var type = reader.PeekTokenType();
                if (type == RTFTokenType.Eof)
                    break;
                if (type == RTFTokenType.GroupStart)
                {
                    level++;
                }
                else if (type == RTFTokenType.GroupEnd)
                {
                    level--;
                    if (level < 0)
                    {
                        break;
                    }
                }
                reader.ReadToken();

                if (deeply || level == 0)
                {
                    if ( htmlMode )
                    {
                        if (reader.Keyword == "par")
                        {
                            container.Append(Environment.NewLine);
                            continue;
                        }
                    }
                    if (container.Accept(reader.CurrentToken , reader ))
                    {
                        continue;
                    }
                    else
                    {
                        if (breakMeetControlWord)
                        {
                            break;
                        }
                    }
                }

            }//while
            return container.Text;
        }

        public override string ToDomString()
        {
            var builder = new StringBuilder();
            builder.Append(ToString());
            builder.Append( Environment.NewLine + "   Info");
            foreach (var item in myInfo.StringItems)
            {
                builder.Append(Environment.NewLine + "      " + item);
            }
            builder.Append( Environment.NewLine + "   ColorTable(" + myColorTable.Count + ")");
            for (var iCount = 0; iCount < myColorTable.Count; iCount ++ )
            {
                var c = myColorTable[iCount];
                builder.Append(Environment.NewLine + "      " + iCount + ":" + c.R + " " + c.G + " " + c.B );
            }
            builder.Append(Environment.NewLine + "   FontTable(" + myFontTable.Count + ")");
            foreach (RTFFont font in myFontTable)
            {
                builder.Append(Environment.NewLine + "      " + font.ToString());
            }
            if (_ListTable.Count > 0)
            {
                builder.Append(Environment.NewLine + "   ListTable(" + _ListTable.Count + ")");
                foreach (var list in _ListTable)
                {
                    builder.Append(Environment.NewLine + "      " + list.ToString());
                }
            }
            if (ListOverrideTable.Count > 0)
            {
                builder.Append(Environment.NewLine + "   ListOverrideTable(" + ListOverrideTable.Count + ")");
                foreach (var list in ListOverrideTable)
                {
                    builder.Append(Environment.NewLine + "      " + list.ToString());
                }
            }
            builder.Append(Environment.NewLine + "   -----------------------");
            if (string.IsNullOrEmpty(HtmlContent) == false)
            {
                builder.Append(Environment.NewLine + "   HTMLContent:" + HtmlContent);
                builder.Append(Environment.NewLine + "   -----------------------");
            }
            ToDomString(Elements, builder, 1);
            return builder.ToString();
        }
    }
}
