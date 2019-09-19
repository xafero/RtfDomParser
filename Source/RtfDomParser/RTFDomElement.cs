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
using System.ComponentModel;

namespace RtfDomParser
{
    /// <summary>
    /// RTF dom element
    /// </summary>
    /// <remarks>this is the most base element type</remarks>
    public abstract class RTFDomElement
    {
        private RTFAttributeList myAttributes = new RTFAttributeList();
        /// <summary>
        /// RTF native attribute
        /// </summary>
        public RTFAttributeList Attributes
        {
            get
            {
                return myAttributes;
            }
            set
            {
                myAttributes = value;
            }
        }

        public bool HasAttribute(string name)
        {
            return myAttributes.Contains(name);
        }

        public int GetAttributeValue( string name , int defaultValue )
        {
            if( myAttributes.Contains( name ))
                return myAttributes[ name ] ;
            else
                return defaultValue ;
        }

        private RTFDomElementList myElements = new RTFDomElementList();
        /// <summary>
        /// child elements list
        /// </summary>
        public RTFDomElementList Elements
        {
            get
            {
                return myElements;
            }
        }

        private RTFDomDocument myOwnerDocument = null;
        /// <summary>
        /// the docuemnt which owned this element
        /// </summary>
        [Browsable( false )]
        [System.Xml.Serialization.XmlIgnore()]
        public RTFDomDocument OwnerDocument
        {
            get
            {
                return myOwnerDocument;
            }
            set
            {
                myOwnerDocument = value;
                foreach (RTFDomElement element in Elements)
                {
                    element.OwnerDocument = value;
                }
            }
        }
        /// <summary>
        /// append child element
        /// </summary>
        /// <param name="element">child element</param>
        /// <returns>index of element</returns>
        public int AppendChild(RTFDomElement element)
        {
            CheckLocked();
            element.myParent = this;
            element.OwnerDocument = myOwnerDocument;
            return myElements.Add(element);
        }

        /// <summary>
        /// set attribute
        /// </summary>
        /// <param name="name">name</param>
        /// <param name="Value">value</param>
        public void SetAttribute(string name, int Value)
        {
            CheckLocked();
            myAttributes[name] = Value;
        }

        private RTFDomElement myParent = null;
        /// <summary>
        /// parent element
        /// </summary>
        [Browsable( false )]
        public RTFDomElement Parent
        {
            get
            {
                return myParent;
            }
        }

        [Browsable( false )]
        public virtual string InnerText
        {
            get
            {
                var str = new StringBuilder();
                if (myElements != null)
                {
                    foreach (RTFDomElement element in myElements)
                    {
                        str.Append(element.InnerText);
                    }
                }
                return str.ToString();
            }
        }


        private void CheckLocked()
        {
            if (bolLocked)
            {
                throw new InvalidOperationException("Element locked");
            }
        }

        private bool bolLocked = false;
        /// <summary>
        /// Whether element is locked , if element is lock , it can not append chidl element
        /// </summary>
        [System.Xml.Serialization.XmlIgnore( )]
        [Browsable( false )]
        public bool Locked
        {
            get 
            {
                return bolLocked; 
            }
            set
            {
                bolLocked = value; 
            }
        }

        public void SetLockedDeeply( bool locked )
        {
            bolLocked = locked;
            if (myElements != null)
            {
                foreach (RTFDomElement element in myElements)
                {
                    element.SetLockedDeeply(locked);
                }
            }
        }

        public void PrintDomString()
        {
            Console.WriteLine(ToDomString());
        }

        public virtual string ToDomString()
        {
            var builder = new StringBuilder();
            builder.Append(ToString());
            ToDomString(Elements, builder, 1);
            return builder.ToString();
        }

        protected void ToDomString(RTFDomElementList elements, StringBuilder builder, int level)
        {
            foreach (RTFDomElement element in elements)
            {
                builder.Append(Environment.NewLine);
                builder.Append( new string( ' ' , level * 4 ));
                builder.Append(element.ToString());
                ToDomString(element.Elements, builder, level + 1);
            }
        }

        /// <summary>
        /// Native level in RTF document
        /// </summary>
        [NonSerialized()]
        internal int NativeLevel = -1;
    }
}
