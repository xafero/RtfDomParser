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
    /// debug information dispalyer at design time
    /// </summary>
    public class RTFInstanceDebugView
    {
        /// <summary>
        /// initialize instance
        /// </summary>
        /// <param name="instance"></param>
        public RTFInstanceDebugView(object instance)
        {
            if (instance == null)
                throw new ArgumentNullException("instance");
            myInstance = instance;
        }

        private object myInstance = null;

        /// <summary>
        /// output debug item at design time
        /// </summary>
        [System.Diagnostics.DebuggerBrowsable(System.Diagnostics.DebuggerBrowsableState.RootHidden)]
        public object Items
        {
            get
            {
                if (myInstance is System.Collections.IEnumerable)
                {
                    var list = (System.Collections.CollectionBase)myInstance;
                    var items = new object[list.Count];
                    var iCount = 0;
                    foreach (var obj in list)
                    {
                        items[iCount] = obj;
                        iCount++;
                    }
                    return items;
                }
                else if (myInstance is RTFColorTable)
                {
                    var table = (RTFColorTable)myInstance;
                    var items = new object[table.Count];
                    for (var iCount = 0; iCount < table.Count; iCount++)
                    {
                        items[iCount] = table[iCount];
                    }
                    return items;
                }
                else if (myInstance is RTFDocumentInfo)
                {
                    var info = (RTFDocumentInfo)myInstance;
                    return info.StringItems;
                }
                else
                {
                    return myInstance;
                }
            }
        }
    }
}
