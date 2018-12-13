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
using iText.Forms.Fields;
using iText.Forms;
using iText.Forms.Util;
using iText.Forms.Xfa;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Linq;

using Microsoft.Office.Interop.Word;


namespace pdftest
{
    /// <summary>
    /// Description of UserCodeModule1.
    /// </summary>
    [TestModule("CBCBA0F2-CC01-474A-B1F6-EB67BC17AB56", ModuleType.UserCode, 1)]
    public class UserCodeModule1 : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public UserCodeModule1()
        {
            // Do not delete - a parameterless constructor is required!
        }

        /// <summary>
        /// Performs the playback of actions in this module.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
            
            // This is dependent on the itext 7 community NuGet package reference
            
            string src = @"C:\Users\tom\Documents\Property_contract_sample.pdf";
           
            PdfReader reader = new PdfReader(src);
            PdfDocument pdf = new PdfDocument(reader);
            string text = string.Empty;
            for(int page = 1; page <= pdf.GetNumberOfPages(); page++)
            {
            	PdfPage myPage = pdf.GetPage(page);
        		text += PdfTextExtractor.GetTextFromPage(myPage); 
            }
            reader.Close();
            int pos = text.IndexOf("Spoluúčast")+10;
            Ranorex.Report.Info (text.Substring(pos,10));
            
            //*************************//
            // Now for the word document. Open the word document, then save as text only, you could then open it and read from it
           	// using normal System.IO library
        	Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Document document = null; 
    		application.Visible = false;
    		object missing = System.Reflection.Missing.Value;
            object readOnly = false;
            document = application.Documents.Open(@"C:\Users\tom\Documents\Property_policy_sample.doc", ref missing, ref readOnly);
            //document.ExportAsFixedFormat(@"C:\Users\tom\Documents\Property_policy_sample.pdf", WdExportFormat.wdExportFormatPDF);
            document.SaveAs(@"C:\Users\tom\Documents\Property_policy_sample.txt",2);
            Ranorex.Report.Info(document.Range(1,document.Characters.Count).Text);
            application.ActiveDocument.Close();
            application.Quit();
                


			// Try to open the Xfa dynamic acrobat form as xml document
			
			// https://csharp.hotexamples.com/examples/iTextSharp.text.pdf/XfaForm/-/php-xfaform-class-examples.html
			
		
            PdfReader newreader = new PdfReader(@"C:\Users\tom\Documents\eForm_example.pdf");
            PdfDocument pdfDoc = new PdfDocument(newreader);
            
            XfaForm xfa = new XfaForm (pdfDoc);
            if (xfa.IsXfaPresent()){
            	
            	XDocument myXMLDoc = xfa.GetDomDocument();
            	IEnumerable<XElement> mylist =myXMLDoc.Descendants("CONTRACT_ID");
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
            
	            PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, false);
				IDictionary<String, PdfFormField> fields = form.GetFormFields();
				PdfFormField toSet;
				fields.TryGetValue("CONTRACT_ID", out toSet);
				
				Report.Info(toSet.GetValueAsString());
            }
            
		      //reader.Close();

        }
    }
}
