/*
 * Created by Ranorex
 * User: tom
 * Date: 13/12/2018
 * Time: 15:35
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

using iText.Kernel.Pdf;
using iText.Forms.Xfa;
using System.Xml.Linq;
using System.Linq;

namespace pdftest
{
    /// <summary>
    /// Ranorex user code collection. A collection is used to publish user code methods to the user code library.
    /// </summary>
    [UserCodeCollection]
    public class FuncLib
    {
        // You can use the "Insert New User Code Method" functionality from the context menu,
        // to add a new method with the attribute [UserCodeMethod].
        
        /// <summary>
        /// This is a placeholder text. Please describe the purpose of the
        /// user code method here. The method is published to the user code library
        /// within a user code collection.
        /// </summary>
        [UserCodeMethod]
        public static void GetXFAfieldValue(String path, String FieldName)
        {
        	PdfReader newreader = new PdfReader(path);
            PdfDocument pdfDoc = new PdfDocument(newreader);
            
            XfaForm xfa = new XfaForm (pdfDoc);
            if (xfa.IsXfaPresent()){
            	
            	XDocument myXMLDoc = xfa.GetDomDocument();
            	IEnumerable<XElement> mylist =myXMLDoc.Descendants(FieldName);
            	List<XElement> mylistt = mylist.ToList<XElement>();
            	Report.Info (mylistt[0].Value);
            	// or if more than one, you can cycle through:
            	/***
            	foreach (XElement theEl in mylist){
            		Report.Info(theEl.Value);
            	}
				***/
            }
            else{
            	Report.Warn("This is not an XFA document");
            }
        }
    }
}
