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
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Forms.Xfa;
using System.Xml.Linq;
using System.Linq;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace pdftest
{
    /// <summary>
    /// Ranorex user code collection. A collection is used to publish user code methods to the user code library.
    /// </summary>
    [UserCodeCollection]
    public class FuncLib
    {
        /// <summary>
        /// Returns a value from a field in an XFA PDF Dynamic Form document
        /// params: path to PDF, Field to get value from
        /// returns true or false for success or not
        /// </summary>
        [UserCodeMethod]
        public static bool GetXFAfieldValue(String path, String FieldName)
        {
        	if (!File.Exists(path)){
        		Report.Warn("File " + path + " does not exist, please check file path!");
        		return false;
        	}
        	
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
				return true;
            }
            else{
            	Report.Warn("This is not an XFA document");
            	return false;
            }
        }
        
        
        /// <summary>
        /// Takes a path to a PDF document and extracts the text
        /// Returns: Text from PDF or Null if file does not exist
        /// Parameters: full file path to PDF doc
        /// </summary>
        [UserCodeMethod]
        public static String GetTextFromPDF(String src)
        {
        	//string src = @"C:\Users\tom\Documents\Property_contract_sample.pdf";
        	if (!File.Exists(src)){
        		Report.Warn("File " + src + " does not exist, please check file path!");
        		return null;
        	}
            PdfReader reader = new PdfReader(src);
            PdfDocument pdf = new PdfDocument(reader);
            string text = string.Empty;
            for(int page = 1; page <= pdf.GetNumberOfPages(); page++)
            {
            	PdfPage myPage = pdf.GetPage(page);
        		text += PdfTextExtractor.GetTextFromPage(myPage); 
            }
            reader.Close();
            return text;    
        }
        
        
        /// <summary>
        /// Converts a PDF to a Text File
        /// params: path to pdf doc, path to text file
        /// returns true on success
        /// </summary>
        [UserCodeMethod]
        public static bool SavePDFasTextFile(String src, String dest)
        {
        	if (!File.Exists(src)){
        		Report.Warn("File " + src + " does not exist, please check file path!");
        		return false;
        	}
        	String extText = GetTextFromPDF(src);
        	
        	TextWriter textWriter = new StreamWriter(dest, false, System.Text.Encoding.UTF8);
        	textWriter.Write(extText);
        	textWriter.Close();
        	return true;
        }
        
        
        /// <summary>
        /// Find text in a String, then return number of characters after position
        /// Parameters: text string to search in, text to search for, number of characters after position to return
        /// Returns: substring
        /// </summary>
        [UserCodeMethod]
        public static String FindValueInText(String text, String searchText, int charLen)
        {
        	int pos = text.IndexOf(searchText) + searchText.Length;
            return text.Substring(pos, charLen);
        }
        
        
        /// <summary>
        /// Convert a PDF to Text, then find a string of text in it and return a text from that position
        /// Params: path to PDF, text to search for, length of chars to return
        /// Returns: substring from doc
        /// </summary>
        [UserCodeMethod]
        public static String GetTextStringFromPDF(String src, String searchText, int charLen)
        {
        	String PDFText = GetTextFromPDF(src);
        	return FindValueInText (PDFText, searchText, charLen);
        }
        
        
        
        /// <summary>
        /// Converts a Word document to a text file
        /// params: path to word file, path to text file output
        /// returns: nothing
        /// </summary>
        [UserCodeMethod]
        public static void SaveWordAsTextFile(String src, String dest)
        {
        	Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Document document = null; 
    		application.Visible = false;
    		object missing = System.Reflection.Missing.Value;
            object readOnly = false;
            document = application.Documents.Open(src, ref missing, ref readOnly);
            //document.ExportAsFixedFormat(@"C:\Users\tom\Documents\Property_policy_sample.pdf", WdExportFormat.wdExportFormatPDF);
            document.SaveAs(dest,2);
            //Ranorex.Report.Info(document.Range(1,document.Characters.Count).Text);
            application.ActiveDocument.Close();
            application.Quit();
        }
        
        
    }
}
