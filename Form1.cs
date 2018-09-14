using System.Windows.Forms;
using System.IO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

using V = DocumentFormat.OpenXml.Vml;
using System.Diagnostics;
// using System.Xml.Linq;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using OpenXmlPowerTools;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Xml.Linq;
using System.Xml;



namespace OpenXMLPractice
{
    public partial class Form1 : Form
    {
        public static string valueglobal = "0";  //abtracnumber del inciso ubicado en el archivo numeric xml del documento
        public static bool bandera = true;
        static int acu = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void cmdCreateNew_Click(object sender, EventArgs e)
        {
           // createDocument();
            // SearchAndReplace("D:\\test11.docx");
            documentProtocol("D:\\test11.docx");
          //  RemoveSectionBreaks("D:\\test11.docx");
        }

        public void createDocument() {
            using (WordprocessingDocument doc = WordprocessingDocument.Create("D:\\test11.docx", DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());

                // String msg contains the text, "Hello, Word!"
                run.AppendChild(new Text("New text in document"));

                MessageBox.Show("New word file created successfully");
            }
        }

        public void eliminar() {
          
        }

        public static void SearchAndReplace(string document)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {


                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex(@"^\s+$[\r\n]*");
                docText = regexText.Replace(docText, "*");

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }

                MessageBox.Show("Edited successfully");
            }
        }
        public static Paragraph GetText(string cellText)
        {
            var run = new Run(new Text(cellText));

            return new Paragraph(run);
        }
        public static string valorInciso(int indice, string tipo , string idvalue) {

            string valor = " ";
            if (idvalue != valueglobal)
            {
                acu = 0;
                indice = acu;
                bandera = true;
            }

           
            if (tipo.Equals("lowerLetter"))
            {
                valor = getLowerLetter(indice);
               // newPara.AppendChild(new Run(new Text("a) ") { Space = SpaceProcessingModeValues.Preserve }));
            }
            if (tipo.Equals("decimal"))
            {
                valor = geDecimal(indice);
               //newPara.AppendChild(new Run(new Text("1) ") { Space = SpaceProcessingModeValues.Preserve }));
            }
            if (tipo.Equals("upperLetter"))
            {
                valor = getLowerLetter(indice).ToUpper();
                // newPara.AppendChild(new Run(new Text("A) ") { Space = SpaceProcessingModeValues.Preserve }));
            }

            return valor;
        }

        public static string getLowerLetter (int indice){

            string valor = "";
            switch (indice)
            {
                case 0:
                    valor = "a) ";
                    break;
                case 1:
                    valor = "b) ";
                    break;
                case 2:
                    valor = "c) ";
                    break;
                case 3:
                    valor = "d) ";
                    break;
                case 4:
                    valor = "e) ";
                    break;
                case 5:
                    valor = "f) ";
                    break;
                case 6:
                    valor = "g) ";
                    break;
                case 7:
                    valor = "h) ";
                    break;
                case 8:
                    valor = "i) ";
                    break;
                case 9:
                    valor = "j) ";
                    break;
                case 10:
                    valor = "k) ";
                    break;
                case 11:
                    valor = "l) ";
                    break;
                case 12:
                    valor = "m) ";
                    break;
                case 13:
                    valor = "n) ";
                    break;
                case 14:
                    valor = "o) ";
                    break;
                case 15:
                    valor = "o) ";
                    break;
                case 16:
                    valor = "q) ";
                    break;
                case 17:
                    valor = "r) ";
                    break;
                case 18:
                    valor = "s) ";
                    break;
                case 19:
                    valor = "t) ";
                    break;
                case 20:
                    valor = "u) ";
                    break;
                case 21:
                    valor = "v) ";
                    break;
                case 22:
                    valor = "w) ";
                    break;
                case 23:
                    valor = "x) ";
                    break;
                case 24:
                    valor = "y) ";
                    break;
                case 25:
                    valor = "z) ";
                    break;
                default:
                    Console.WriteLine("Default case");
                    break;
            }
            return valor;
        }
        public static string geDecimal(int indice)
        {

            string valor = (indice + 1).ToString() +") ";
            return valor;
        }

        public static void documentProtocol(string filename)
        {
      
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {

                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                // XDocument xDoc = mainPart.GetXDocument();

                Numbering numbering = myDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering;

                //obtener los number
                // var abstractNum = numbering.ChildElements.FirstOrDefault(x => x.LocalName.Equals("abstractNum"));
                //OpenXmlElement abstractNum = numbering.ChildElements.FirstOrDefault(x => x.LocalName.Equals("abstractNum"));
                List<OpenXmlElement> ablist = numbering.ChildElements.Where(x => x.LocalName.Equals("abstractNum")).ToList();
                List<OpenXmlElement> item = numbering.ChildElements.Where(x => x.LocalName.Equals("num")).ToList();
                // obtengo todos los parrafos del documento
                List<Paragraph> paragraphs = mainPart.Document.Descendants<Paragraph>().ToList();
                int i = 1;

                Paragraph newPara = new Paragraph();
                // newPara.AppendChild(new Run(new Text("This sentence has spacing between the gg in to")));
                bool bold = false;
                foreach (Paragraph p in paragraphs)
                {
                    bold = false;
                    List<OpenXmlElement> elems = p.Descendants<OpenXmlElement>().Where(x => x.LocalName == "numId").ToList();
                    // List<OpenXmlElement> elemsall = p.Descendants<OpenXmlElement>().ToList();
                    List<OpenXmlElement> elemsbold = p.Descendants<OpenXmlElement>().Where(x => x.LocalName == "b").ToList();
                    if( elemsbold.Count> 0){
                        bold = true;
                    }
                    // int acu = 0;
                    foreach (var child in elems)
                    {

                    

                        string value = "";
                        if (child.OuterXml.Contains("<w:numId"))
                         {

                             XmlDocument levelXml = new XmlDocument();
                             levelXml.LoadXml(child.OuterXml);
                             XmlNamespaceManager namespaceManager = new XmlNamespaceManager(levelXml.NameTable);
                             namespaceManager.AddNamespace("w",
                                 "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                             XmlNode runProperty = levelXml.SelectSingleNode("w:numId", namespaceManager);
                             value = runProperty.Attributes["w:val"].Value;
                             if (bandera)
                             {
                                 valueglobal = value;
                                 bandera = false;
                             }
                             
                         }

                        //buscar el abstractNum en numbering
                        string auxvalue = "";
                        string tipo = "";
                        foreach (var it in item) {
                            XmlDocument levelXml = new XmlDocument();
                            levelXml.LoadXml(it.OuterXml);
                            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(levelXml.NameTable);
                            namespaceManager.AddNamespace("w",
                                "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                            XmlNode runProperty = levelXml.SelectSingleNode("w:num", namespaceManager);
                            auxvalue = runProperty.Attributes["w:numId"].Value;
                            if (value == auxvalue) {
                                XmlNode ab = runProperty.SelectSingleNode("/w:num/w:abstractNumId", namespaceManager);
                                string abnum = ab.Attributes["w:val"].Value;

                                foreach (var abitem in ablist) {
                                    XmlDocument levelXmlab = new XmlDocument();
                                    levelXmlab.LoadXml(abitem.OuterXml);
                                    XmlNode abProperty = levelXmlab.SelectSingleNode("w:abstractNum", namespaceManager);
                                    string abvalue = abProperty.Attributes["w:abstractNumId"].Value;
                                    if (abnum == abvalue) {
                                        Console.Write("obtener tipo list");
                                        XmlNode numtype = abProperty.SelectSingleNode("/w:abstractNum/w:lvl/w:numFmt", namespaceManager);
                                         tipo = numtype.Attributes["w:val"].Value;
                                    }
                                }
                            }
                            
                        }
                        //insertar segun el tipo
                        string inciso = valorInciso(acu, tipo, value);

                        if (bold)
                        {
                            //create a run for the bold text
                            Run run = new Run();
                            run.Append(new Text(inciso) { Space = SpaceProcessingModeValues.Preserve });
                            //create runproperties and append a "Bold" to them
                            RunProperties runProperties = new RunProperties();
                            runProperties.Append(new Bold());
                            //set the first runs RunProperties to the RunProperties containing the bold
                            run.RunProperties = runProperties;

                            newPara.AppendChild(run);
                        }
                        else {

                            newPara.AppendChild(new Run(new Text(inciso) { Space = SpaceProcessingModeValues.Preserve }));
                        }
                      
                        
                        //subir acumulador
                        acu++;
                       // newPara.AppendChild(new Run(new Text(value+ " ") { Space = SpaceProcessingModeValues.Preserve }));
                    }
                    //foreach (OpenXmlElement el in elems) {
  
                    //    Console.Write("ele");
                    //}

                   //  List<OpenXmlElement> childelems = p.ChildElements.ToList();
                    List<OpenXmlElement> childelems = p.ChildElements.Where(x => x.LocalName == "r").ToList();
                    foreach (OpenXmlElement elem in childelems) {
                        newPara.AppendChild(elem.CloneNode(true));
                    }
                    //for (int k = 0; k < p.ChildElements.Count; k++) {
                    //    string texto = p.ElementAt(k).InnerText;
                    //    if (texto.Length > 0) {
                    //        // newPara.AppendChild((Run)p.ElementAt(k).CloneNode(true));
                    //        newPara.AppendChild(new Run(p.ElementAt(k).CloneNode(true)));
                            
                    //    }
                       
                    //    Console.Write("text");
                    //}

                    newPara.AppendChild(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
                        
                   // newPara.AppendChild((Run)b.GetFirstChild<Run>().CloneNode(true));
                    // newPara.AppendChild(new Run(new Text("  ")));
                    // newPara.AppendChild(new Run(new Text(b.InnerText)));
                     // newPara.AppendChild(new Run(b.);
                    
                   // b.Remove();
                    Console.Write("text");
                    i++;
                }

                //var linea = newPara.Descendants<Break>().ToList();
                //foreach (Break br in linea)
                //{
                //    br.Remove();
                //}
                // adicionar Paragraph
                var body = myDoc.MainDocumentPart.Document.Body;
               
                // adiciona el parrafo final
                // body.Append(newPara);

                // creat nuevo documento

                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create("D:\\protocolo.docx", WordprocessingDocumentType.Document))
                {
                    // Insert other code here. 
                    // Add a main document part. 
                    MainDocumentPart mainPartnew  = wordDocument.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPartnew.Document = new Document();
                    Body bodynew = mainPartnew.Document.AppendChild(new Body());
                    bodynew.Append(newPara);
                    mainPart.Document.Save();
                }

                mainPart.Document.Save();

            }

        }
        static void RemoveSectionBreaks(string filename)
        {

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {

                MainDocumentPart mainPart = myDoc.MainDocumentPart;

                List<ParagraphProperties> paraProps = mainPart.Document.Descendants<ParagraphProperties>()

                .Where(pPr => IsSectionProps(pPr)).ToList();

                foreach (ParagraphProperties pPr in paraProps)
                {

                    pPr.RemoveChild<SectionProperties>(pPr.GetFirstChild<SectionProperties>());

                }

                mainPart.Document.Save();

            }

        }

        static bool IsSectionProps(ParagraphProperties pPr)
        {

            SectionProperties sectPr = pPr.GetFirstChild<SectionProperties>();

            if (sectPr == null)

                return false;

            else

                return true;

        }

        public void remover() {
            Dictionary<String, BookmarkStart> bookMap = new Dictionary<String, BookmarkStart>(); //a dictionary so we can lookup a bookmarkStart with it's name
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open("D:\\test11.docx", true))
            {
                var mainPart = wordDoc.MainDocumentPart; //get mainpart
                var bookmarks = mainPart.Document.Body.Descendants<BookmarkStart>(); //get all bookmarks
                foreach (BookmarkStart bookmarkStart in bookmarks) //save them to dictionary
                {
                    bookMap[bookmarkStart.Name] = bookmarkStart;
                }
                // ReplaceInBookmark(bookmarks[@"\n"], "text to insert");
               
            }
        }
        public void ReplaceInBookmark(BookmarkStart bookmarkStart, string text)
        {
            OpenXmlElement elem = bookmarkStart.NextSibling();
            while (elem != null && !(elem is BookmarkEnd))
            {
                OpenXmlElement nextElem = elem.NextSibling();
                elem.Remove();
                elem = nextElem;
            }
            bookmarkStart.Parent.InsertAfter<Run>(new Run(new Text(text)), bookmarkStart);
        }
    }
}
