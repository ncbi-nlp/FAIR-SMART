package FormatConverterLib;

import bioc.BioCAnnotation;
import bioc.BioCCollection;
import bioc.BioCDocument;
import bioc.BioCLocation;
import bioc.BioCPassage;
import bioc.BioCRelation;
import bioc.BioCNode;

import bioc.io.BioCDocumentWriter;
import bioc.io.BioCFactory;
import bioc.io.woodstox.ConnectorWoodstox;
import nu.xom.Nodes;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.ZoneId;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLStreamException;

import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.io.RandomAccessFile;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import java.awt.geom.Rectangle2D;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;

import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTable;
import org.apache.poi.hslf.usermodel.HSLFTableCell;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFTextShape;
import org.apache.poi.sl.usermodel.ShapeType;

//import org.apache.xmlbeans.ResourceLoader;
//import net.sourceforge.tess4j.ITesseract;
//import net.sourceforge.tess4j.Tesseract;
//import net.sourceforge.tess4j.TesseractException;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.google.gson.Gson;

import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.PdfTextExtractor;

import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.compress.archivers.tar.TarArchiveEntry;
import org.apache.commons.compress.archivers.tar.TarArchiveInputStream;
import org.apache.commons.compress.archivers.tar.TarArchiveOutputStream;
import org.apache.commons.compress.compressors.gzip.GzipCompressorInputStream;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.commons.lang3.StringEscapeUtils;

import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;
import net.sourceforge.tess4j.ITessAPI.TessPageIteratorLevel;
import net.sourceforge.tess4j.*;

import java.util.logging.Logger;
import org.slf4j.bridge.SLF4JBridgeHandler;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.Rectangle;
import java.awt.*;

import javax.swing.text.*;
import javax.swing.text.rtf.RTFEditorKit;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;

/**
 * 
 * using http://pdfbox.apache.org/download.cgi#20x for pdf2text
 *
 */

public class FormatConverter 
{
	//private static final Logger logger = LoggerFactory.getLogger(FormatConverter.class);
	
	/*
	 * Contexts in BioC file
	 */
	public ArrayList<String> PMIDs=new ArrayList<String>(); // Type: PMIDs
	public ArrayList<ArrayList<String>> PassageNames = new ArrayList(); // PassageName
	public ArrayList<ArrayList<Integer>> PassageOffsets = new ArrayList(); // PassageOffset
	public ArrayList<ArrayList<String>> PassageContexts = new ArrayList(); // PassageContext
	public ArrayList<ArrayList<ArrayList<String>>> Annotations = new ArrayList(); // Annotation - GNormPlus
	public static ArrayList<String> XML_names = new ArrayList<String>();
	public static ArrayList<String> XML_contents = new ArrayList<String>();
	public static HashMap<String, String> oafile_pmcid2pmid_hash = new HashMap<String, String>();
	
	public static String string_refine (String inputtext) // for refining the strings in Word, Excel and PPT files
	{
		
		if (inputtext == null) 
		{
			return null;
		}

		StringBuilder refinedText = new StringBuilder();

		for (char c : inputtext.toCharArray()) 
		{
			// Remove control characters (except tab, line feed, and carriage return)
			if (c == 0x03B4) { // δ
			    refinedText.append("&delta;");
			} else if (c == 0x03C9) { // ω
			    refinedText.append("&omega;");
			} else if (c == 0x03BC) { // μ
			    refinedText.append("&mu;");
			} else if (c == 0x03BA) { // κ
			    refinedText.append("&kappa;");
			} else if (c == 0x03B1) { // α
			    refinedText.append("&alpha;");
			} else if (c == 0x03B3) { // γ
			    refinedText.append("&gamma;");
			} else if (c == 0x0263) { // ɣ
			    refinedText.append("&#x263;"); // HTML code for ɣ
			} else if (c == 0x03B2) { // β
			    refinedText.append("&beta;");
			} else if (c == 0x00D7) { // ×
			    refinedText.append("&times;");
			} else if (c == 0x2011) { // ‑
			    refinedText.append("&#x2011;"); // HTML code for non-breaking hyphen
			} else if (c == 0x00B9) { // ¹
			    refinedText.append("&sup1;");
			} else if (c == 0x00B2) { // ²
			    refinedText.append("&sup2;");
			} else if (c == 0x00B0) { // °
			    refinedText.append("&deg;");
			} else if (c == 0x00F6) { // ö
			    refinedText.append("&ouml;");
			} else if (c == 0x00E9) { // é
			    refinedText.append("&eacute;");
			} else if (c == 0x00E0) { // à
			    refinedText.append("&agrave;");
			} else if (c == 0x00C1) { // Á
			    refinedText.append("&Aacute;");
			} else if (c == 0x03B5) { // ε
			    refinedText.append("&epsilon;");
			} else if (c == 0x03B8) { // θ
			    refinedText.append("&theta;");
			} else if (c == 0x2022) { // •
			    refinedText.append("&bull;");
			} else if (c == 0x00B5) { // µ
			    refinedText.append("&micro;");
			} else if (c == 0x03BB) { // λ
			    refinedText.append("&lambda;");
			} else if (c == 0x207A) { // ⁺
			    refinedText.append("&#x207A;"); // HTML code for superscript plus
			} else if (c == 0x03BD) { // ν
			    refinedText.append("&nu;");
			} else if (c == 0x00EF) { // ï
			    refinedText.append("&iuml;");
			} else if (c == 0x00E3) { // ã
			    refinedText.append("&atilde;");
			} else if (c == 0x2261) { // ≡
			    refinedText.append("&equiv;");
			} else if (c == 0x00F3) { // ó
			    refinedText.append("&oacute;");
			} else if (c == 0x00B3) { // ³
			    refinedText.append("&sup3;");
			} else if (c == 0x3010) { // 〖
			    refinedText.append("&#x3010;"); // HTML code for left black lenticular bracket
			} else if (c == 0x3011) { // 〗
			    refinedText.append("&#x3011;"); // HTML code for right black lenticular bracket
			} else if (c == 0x00C5) { // Å
			    refinedText.append("&Aring;");
			} else if (c == 0x03C1) { // ρ
			    refinedText.append("&rho;");
			} else if (c == 0x00FC) { // ü
			    refinedText.append("&uuml;");
			} else if (c == 0x025B) { // ɛ
			    refinedText.append("&#x025B;"); // HTML code for Latin small letter open e
			} else if (c == 0x010D) { // č
			    refinedText.append("&#x010D;"); // HTML code for Latin small letter c with caron
			} else if (c == 0x0161) { // š
			    refinedText.append("&#x0161;"); // HTML code for Latin small letter s with caron
			} else if (c == 0x00DF) { // ß
			    refinedText.append("&szlig;");
			} else if (c == 0x2550) { // ═
			    refinedText.append("&#x2550;"); // HTML code for box drawings double horizontal
			} else if (c == 0x00A3) { // £
			    refinedText.append("&pound;");
			} else if (c == 0x0141) { // Ł
			    refinedText.append("&#x0141;"); // HTML code for Latin capital letter L with stroke
			} else if (c == 0x0192) { // ƒ
			    refinedText.append("&fnof;");
			} else if (c == 0x00E4) { // ä
			    refinedText.append("&auml;");
			} else if (c == 0x2013) { // –
			    refinedText.append("&ndash;");
			} else if (c == 0x207B) { // ⁻
			    refinedText.append("&#x207B;"); // HTML code for superscript minus
			} else if (c == 0x3008) { // 〈
			    refinedText.append("&#x3008;"); // HTML code for left angle bracket
			} else if (c == 0x3009) { // 〉
			    refinedText.append("&#x3009;"); // HTML code for right angle bracket
			} else if (c == 0x03C7) { // χ
			    refinedText.append("&chi;");
			} else if (c == 0x0110) { // Đ
			    refinedText.append("&#x0110;"); // HTML code for Latin capital letter D with stroke
			} else if (c == 0x2030) { // ‰
			    refinedText.append("&permil;");
			} else if (c == 0x00B7) { // ·
			    refinedText.append("&middot;");
			} else if (c == 0x2192) { // →
			    refinedText.append("&rarr;");
			} else if (c == 0x2190) { // ←
			    refinedText.append("&larr;");
			} else if (c == 0x03B6) { // ζ
			    refinedText.append("&zeta;");
			} else if (c == 0x03C0) { // π
			    refinedText.append("&pi;");
			} else if (c == 0x03C4) { // τ
			    refinedText.append("&tau;");
			} else if (c == 0x03BE) { // ξ
			    refinedText.append("&xi;");
			} else if (c == 0x03B7) { // η
			    refinedText.append("&eta;");
			} else if (c == 0x00F8) { // ø
			    refinedText.append("&oslash;");
			} else if (c == 0x0394) { // Δ
			    refinedText.append("&Delta;");
			} else if (c == 0x2206) { // ∆
			    refinedText.append("&#x2206;"); // HTML code for increment (different from Greek Delta)
			} else if (c == 0x2211) { // ∑
			    refinedText.append("&sum;");
			} else if (c == 0x03A9) { // Ω
			    refinedText.append("&Omega;");
			} else if (c == 0x03B4) { // δ
			    refinedText.append("&delta;");
			} else if (c == 0x03C3) { // σ
			    refinedText.append("&sigma;");
			} else if (c == 0x03A6) { // Φ
			    refinedText.append("&Phi;");
			}  
		    else if ((c >= 0x20 && c <= 0xD7FF) || (c >= 0xE000 && c <= 0xFFFD) || (c >= 0x10000 && c <= 0x10FFFF) || c == 0x9 || c == 0xA || c == 0xD) 
			{
				refinedText.append(c);
			} 
			else 
			{
				// Replace invalid characters with a space or any other valid character
				refinedText.append(' ');
			}
		}
	    
        // First, escape using StringEscapeUtils
        String escapedHtml = StringEscapeUtils.escapeHtml4(refinedText.toString()).replace("\n", "").trim();
        return escapedHtml;
	}
	private static String bytesToHex(byte[] bytes) 
	{
	    StringBuilder sb = new StringBuilder();
	    for (byte b : bytes) {
	        sb.append(String.format("%02X", b));
	    }
	    return sb.toString();
	}
	public static String BioCFormatCheck(String InputFile) throws IOException
	{
		String headerStr = "";
		try (FileInputStream fis = new FileInputStream(new File(InputFile))) 
		{
            // Check the first 8 bytes for the OLE2 signature
            byte[] header = new byte[8];
            if (fis.read(header) != -1) 
            {
                headerStr = bytesToHex(header);
            }
        } catch (IOException e) 
		{
            e.printStackTrace();
        }
		
		if(InputFile.toLowerCase().endsWith(".txt"))
		{
			return "TXT";

		}
		else if(InputFile.toLowerCase().endsWith(".tsv"))
		{
			return "TSV";

		}
		else if(InputFile.toLowerCase().endsWith(".csv"))
		{
			return "CSV";

		}
		else if(InputFile.toLowerCase().endsWith(".xlsx"))
		{
			if(headerStr.equalsIgnoreCase("504B0304"))
			{
				return "Excelx";
			}
			else if(headerStr.equalsIgnoreCase("504B030414000600"))
			{
				return "Excelx";
			}
			else if(headerStr.equalsIgnoreCase("D0CF11E0A1B11AE1"))
			{
				return "Excel";
			}
			else
			{
				return "Excelx";
			}
		}
		else if(InputFile.toLowerCase().endsWith(".xls"))
		{
			if(headerStr.equalsIgnoreCase("504B0304"))
			{
				return "Excelx";
			}
			else if(headerStr.equalsIgnoreCase("D0CF11E0A1B11AE1"))
			{
				return "Excel";
			}
			else if(headerStr.equalsIgnoreCase("504B030414000600"))
			{
				return "Excelx";
			}
			else
			{
				return "Excel";
			}
		}
		else if(InputFile.toLowerCase().matches(".*\\.pptx"))
		{
			if(headerStr.equalsIgnoreCase("504B0304"))
			{
				return "PPTx";
			}
			else if(headerStr.equalsIgnoreCase("504B030414000600"))
			{
				return "PPTx";
			}
			else if(headerStr.equalsIgnoreCase("D0CF11E0A1B11AE1"))
			{
				return "PPT";
			}
			else
			{
				return "PPTx";
			}
		}
		else if(InputFile.toLowerCase().matches(".*\\.ppt"))
		{
			if(headerStr.equalsIgnoreCase("504B0304"))
			{
				return "PPTx";
			}
			else if(headerStr.equalsIgnoreCase("D0CF11E0A1B11AE1"))
			{
				return "PPT";
			}
			else if(headerStr.equalsIgnoreCase("504B030414000600"))
			{
				return "PPTx";
			}
			else
			{
				return "PPT";
			}
		}
		else if(InputFile.toLowerCase().endsWith(".docx"))
		{
			if(headerStr.equalsIgnoreCase("504B0304"))
			{
				return "Wordx";
			}
			else if(headerStr.equalsIgnoreCase("504B030414000600"))
			{
				return "Wordx";
			}
			else if(headerStr.equalsIgnoreCase("D0CF11E0A1B11AE1"))
			{
				return "Word";
			}
			else if(headerStr.equalsIgnoreCase("7B5C727466") || headerStr.equalsIgnoreCase("5C72746631"))
			{
				return "RTF";
			}
			else
			{
				return "Wordx";
			}
		}
		else if(InputFile.toLowerCase().endsWith(".doc"))
		{
			if(headerStr.equalsIgnoreCase("504B0304"))
			{
				return "Wordx";
			}
			else if(headerStr.equalsIgnoreCase("D0CF11E0A1B11AE1"))
			{
				return "Word";
			}
			else if(headerStr.equalsIgnoreCase("504B030414000600"))
			{
				return "Wordx";
			}
			else if(headerStr.equalsIgnoreCase("7B5C727466") || headerStr.equalsIgnoreCase("5C72746631"))
			{
				return "RTF";
			}
			else
			{
				return "Word";
			}
		}
		else if(InputFile.toLowerCase().matches(".*\\.(gif|jpg|png|jpeg|tif|tiff)"))
		{
			return "IMG";
		}
		else if(InputFile.toLowerCase().endsWith(".tar.gz"))
		{
			return "tar.gz";
		}
		else
		{
			File file = new File(InputFile);
			PDFParser pdfparser = new PDFParser(new RandomAccessFile(file,"r"));
			try
			{
				pdfparser.parse();
				return "PDF";
			}
			catch (IOException notpdf)
			{
				ConnectorWoodstox connector = new ConnectorWoodstox();
				BioCCollection collection = new BioCCollection();
				try
				{
					collection = connector.startRead(new InputStreamReader(new FileInputStream(InputFile), "UTF-8"));
				}
				catch (UnsupportedEncodingException | FileNotFoundException | XMLStreamException e) //if not BioC
				{
					BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(InputFile), "UTF-8"));
					String line="";
					String status="";
					String Pmid = "";
					boolean tiabs=false;
					Pattern patt = Pattern.compile("^([^\\|\\t]+)\\|([^\\|\\t]+)\\|([^\\|\\t]*)$");
					while ((line = br.readLine()) != null)  
					{
						Matcher mat = patt.matcher(line);
						if(mat.find()) //Title|Abstract
			        	{
							if(Pmid.equals(""))
							{
								Pmid = mat.group(1);
							}
							else if(!Pmid.equals(mat.group(1)))
							{
								return "[Error of PubTator format]: "+InputFile+" - A blank is needed between "+Pmid+" and "+mat.group(1)+".";
							}
							status = "tiabs";
							tiabs = true;
			        	}
						else if (line.contains("\t")) //Annotation
			        	{
			        	}
						else if(line.length()==0) //Processing
						{
							if(status.equals(""))
							{
								if(Pmid.equals(""))
								{
									return "[Error 1.0]: "+InputFile+" - It's neither BioC nor PubTator format.";
								}
								else
								{
									return "[Error of PubTator format]: "+InputFile+" - A redundant blank is after "+Pmid+".";
								}
							}
							Pmid="";
							status="";
						}
					}
					br.close();
					if(tiabs == false)
					{
						return "[Error 1.1]: "+InputFile+" - It's neither BioC nor PubTator format.";
					}
					
					if(status.equals(""))
					{
						return "PubTator";
					}
					else
					{
						return "[Error of PubTator format]: "+InputFile+" - The last column missed a blank.";
					}
				}
				return headerStr;
			}
		}
	}
	public static void PubTator2BioC(String input,String output) throws IOException, XMLStreamException
	{
		String parser = BioCFactory.WOODSTOX;
		BioCFactory factory = BioCFactory.newFactory(parser);
		BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		BioCCollection biocCollection = new BioCCollection();
		
		//time
		ZoneId zonedId = ZoneId.of( "America/Montreal" );
		LocalDate today = LocalDate.now( zonedId );
		biocCollection.setDate(today.toString());
		
		biocCollection.setKey("BioC.key");//key
		biocCollection.setSource("PubTator");//source
		
		BioCOutputFormat.writeCollectionInfo(biocCollection);
		BufferedReader inputfile = new BufferedReader(new InputStreamReader(new FileInputStream(input), "UTF-8"));
		ArrayList<String> ParagraphType=new ArrayList<String>(); // Type: Title|Abstract
		ArrayList<String> ParagraphContent = new ArrayList<String>(); // Text
		ArrayList<String> annotations = new ArrayList<String>(); // Annotation
		ArrayList<String> relations = new ArrayList<String>(); // relation
		HashMap <String,HashMap<Integer,Integer>> id2Aid = new HashMap <String,HashMap<Integer,Integer>>();
		String line;
		String Pmid="";
		int count_mention=0;
		while ((line = inputfile.readLine()) != null)  
		{
			if(line.contains("|") && !line.contains("\t")) //Title|Abstract
        	{
				String str[]=line.split("\\|",-1);
				Pmid=str[0];
				if(str[1].equals("t"))
				{
					str[1]="title";
				}
				if(str[1].equals("a"))
				{
					str[1]="abstract";
				}
				ParagraphType.add(str[1]);
				if(str.length==3)
				{
					ParagraphContent.add(str[2]);
				}
				else
				{
					ParagraphContent.add("");
				}
        	}
			else if (line.contains("\t")) //Annotation
        	{
				String anno[]=line.split("\t",-1);
				if(anno.length==7 && anno[1].matches("[0-9]+"))
				{
					annotations.add(anno[1]+"\t"+anno[2]+"\t"+anno[3]+"\t"+anno[4]+"\t"+anno[5]+"\t"+anno[6]);
					String ids[]=anno[5].split(",",-1);
					for(int ids_i=0;ids_i<ids.length;ids_i++)
					{
						if(!id2Aid.containsKey(ids[ids_i]))
						{
							id2Aid.put(ids[ids_i],new HashMap<Integer,Integer>());
						}
						id2Aid.get(ids[ids_i]).put(count_mention,Integer.parseInt(anno[1]));
					}
					count_mention++;
				}
				else if(anno.length==6 && anno[1].matches("[0-9]+"))
				{
					annotations.add(anno[1]+"\t"+anno[2]+"\t"+anno[3]+"\t"+anno[4]+"\t"+anno[5]);
					String ids[]=anno[5].split(",",-1);
					for(int ids_i=0;ids_i<ids.length;ids_i++)
					{
						if(!id2Aid.containsKey(ids[ids_i]))
						{
							id2Aid.put(ids[ids_i],new HashMap<Integer,Integer>());
						}
						id2Aid.get(ids[ids_i]).put(count_mention,Integer.parseInt(anno[1]));
					}
					count_mention++;
				}
				else if(anno.length==5 && anno[1].matches("[0-9]+"))
				{
					annotations.add(anno[1]+"\t"+anno[2]+"\t"+anno[3]+"\t"+anno[4]);
					count_mention++;
				}
				else if(anno.length>=4)
				{
					relations.add(anno[1]+"\t"+anno[2]+"\t"+anno[3]+"\t"+anno[4]);
				}
        	}
			else if(line.length()==0) //Processing
			{
				BioCDocument biocDocument = new BioCDocument();
				biocDocument.setID(Pmid);
				int startoffset=0;
				for(int i=0;i<ParagraphType.size();i++)
				{
					BioCPassage biocPassage = new BioCPassage();
					Map<String, String> Infons = new HashMap<String, String>();
					Infons.put("type", ParagraphType.get(i));
					biocPassage.setInfons(Infons);
					biocPassage.setText(ParagraphContent.get(i));
					biocPassage.setOffset(startoffset);
					startoffset=startoffset+ParagraphContent.get(i).length()+1;
					for(int j=0;j<annotations.size();j++)
					{
						String anno[]=annotations.get(j).split("\t");
						if((Integer.parseInt(anno[0])<startoffset || Integer.parseInt(anno[0])==0) && Integer.parseInt(anno[0])>=startoffset-(ParagraphContent.get(i).length()+1))
						{
							BioCAnnotation biocAnnotation = new BioCAnnotation();
							Map<String, String> AnnoInfons = new HashMap<String, String>();
							if(anno.length>=5)
							{
								AnnoInfons.put("identifier", anno[4]);
							}
							if(anno.length>=6)
							{
								AnnoInfons.put("note", anno[5]);
							}
							AnnoInfons.put("type", anno[3]);
							biocAnnotation.setInfons(AnnoInfons);
							BioCLocation location = new BioCLocation();
							location.setOffset(Integer.parseInt(anno[0]));
							location.setLength(Integer.parseInt(anno[1])-Integer.parseInt(anno[0]));
							biocAnnotation.setLocation(location);
							biocAnnotation.setText(anno[2]);
							biocAnnotation.setID(""+j);
							biocPassage.addAnnotation(biocAnnotation);
						}
					}
					biocDocument.addPassage(biocPassage);
				}
				for(int j=0;j<relations.size();j++)
				{
					String rel[]=relations.get(j).split("\t");
					BioCRelation biocrelation = new BioCRelation();
					Map<String, String> relationtype = new HashMap<String, String>();
					String type=rel[0];
					String entity1=rel[1];
					String entity2=rel[2];
					String novelty=rel[3];
					
					HashMap<Integer,Integer> entity1_Aid2start=id2Aid.get(entity1);
					HashMap<Integer,Integer> entity2_Aid2start=id2Aid.get(entity2);
					
					int min=10000;
					String target_id1="";
					String target_id2="";
					for(int id1 : entity1_Aid2start.keySet())
					{
						int start1=entity1_Aid2start.get(id1);
						for(int id2 : entity2_Aid2start.keySet())
						{
							int start2=entity2_Aid2start.get(id2);
							if(start1>=start2 && (start1-start2)<min)
							{
								target_id1=Integer.toString(id1);
								target_id2=Integer.toString(id2);
								min=start1-start2;
							}
							else if(start2>start1 && (start2-start1)<min)
							{
								target_id1=Integer.toString(id1);
								target_id2=Integer.toString(id2);
								min=start2-start1;
							}
						}
					}
					relationtype.put("annotator", "BioRED");
					relationtype.put("type", type);
					relationtype.put("note", novelty);
					//relationtype.put("entity1", entity1);
					//relationtype.put("entity2", entity2);
					
					biocrelation.setID("R"+j);
					biocrelation.setInfons(relationtype);
					biocrelation.addNode(target_id1, "");
					biocrelation.addNode(target_id2, "");
					biocDocument.addRelation(biocrelation);
				}
				biocCollection.addDocument(biocDocument);
				ParagraphType.clear();
				ParagraphContent.clear();
				annotations.clear();
				relations.clear();
				id2Aid.clear();
				BioCOutputFormat.writeDocument(biocDocument);
				count_mention=0;
			}
		}
		BioCOutputFormat.close();
		inputfile.close();
	}
	public static void BioC2PubTator(String input,String output) throws IOException, XMLStreamException
	{
		HashMap<String, String> pmidlist = new HashMap<String, String>(); // check if appear duplicate pmids
		boolean duplicate = false;
		BufferedWriter PubTatorOutputFormat = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		ConnectorWoodstox connector = new ConnectorWoodstox();
		BioCCollection collection = new BioCCollection();
		collection = connector.startRead(new InputStreamReader(new FileInputStream(input), "UTF-8"));
		while (connector.hasNext()) 
		{
			BioCDocument document = connector.next();
			String PMID = document.getID();
			if(pmidlist.containsKey(PMID)){System.out.println("\n[Warning]: duplicate pmid-"+PMID);duplicate = true;}
			else{pmidlist.put(PMID,"");}
			String Anno="";
			int count_passage=1;
			for (BioCPassage passage : document.getPassages()) 
			{
				if((!passage.getInfons().isEmpty()) && passage.getInfon("type").equals("title"))
				{
					PubTatorOutputFormat.write(PMID+"|t|"+passage.getText()+"\n");
				}
				else if((!passage.getInfons().isEmpty()) && passage.getInfon("type").equals("abstract"))
				{
					PubTatorOutputFormat.write(PMID+"|a|"+passage.getText()+"\n");
				}
				else if((!passage.getInfons().isEmpty()))
				{
					PubTatorOutputFormat.write(PMID+"|"+passage.getInfon("type")+"|"+passage.getText()+"\n");
				}
				else
				{
					PubTatorOutputFormat.write(PMID+"|Passage_"+count_passage+"|"+passage.getText()+"\n");
				}
				
				for (BioCAnnotation annotation : passage.getAnnotations()) 
				{
					String Annotype = annotation.getInfon("type");
					String Annoid="";
					Map<String,String> Infons = annotation.getInfons();
					for(String InfonType : Infons.keySet())
					{
						if(!InfonType.equals("type") && !InfonType.equals("NCBI Homologene"))
						{
							if(Annoid.equals(""))
							{
								Annoid=Infons.get(InfonType);
							}
							else
							{
								Annoid=Annoid+"|"+Infons.get(InfonType);
							}
						}
					}
					int start = annotation.getLocations().get(0).getOffset();
					int last = start + annotation.getLocations().get(0).getLength();
					String AnnoMention=annotation.getText();
					Anno=Anno+PMID+"\t"+start+"\t"+last+"\t"+AnnoMention+"\t"+Annotype+"\t"+Annoid+"\n";
				}
				count_passage++;
			}
			PubTatorOutputFormat.write(Anno);
			
			//relation
			String Rel="";
			for (BioCRelation biocrelation : document.getRelations()) 
			{
				Rel=Rel+PMID+"\t"+biocrelation.getInfon("relation")+"\t"+biocrelation.getInfon("Gene1")+"\t"+biocrelation.getInfon("Gene2")+"\n";
			}
			PubTatorOutputFormat.write(Rel+"\n");
		}
		PubTatorOutputFormat.close();
		if(duplicate == true){System.exit(0);}
	}
	public static void BioC2SciLite(String input,String output) throws IOException, XMLStreamException
	{
		BufferedWriter outputfile = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		HashMap<String, String> pmidlist = new HashMap<String, String>(); // check if appear duplicate pmids
		boolean duplicate = false;
		ConnectorWoodstox connector = new ConnectorWoodstox();
		BioCCollection collection = new BioCCollection();
		collection = connector.startRead(new InputStreamReader(new FileInputStream(input), "UTF-8"));
		while (connector.hasNext()) 
		{
			BioCDocument document = connector.next();
			String PMID = document.getID();
			if(pmidlist.containsKey(PMID)){System.out.println("\nWarning: duplicate pmid-"+PMID);duplicate = true;}
			else{pmidlist.put(PMID,"");}
			String Anno="";
			int count_passage=0;
			for (BioCPassage passage : document.getPassages()) 
			{
				count_passage++;
				
				int count_ann=0;
				/* Annotation */
				for (BioCAnnotation annotation : passage.getAnnotations()) 
				{
					count_ann++;
					String Annotype = annotation.getInfon("type");
					String Annoid="";
					Map<String,String> Infons = annotation.getInfons();
					for(String InfonType : Infons.keySet()) // check all Infontype
					{
						if(!InfonType.equals("type"))
						{
							if(Annoid.equals(""))
							{
								Annoid=Infons.get(InfonType);
							}
							else
							{
								Annoid=Annoid+"|"+Infons.get(InfonType);
							}
						}
					}
					Annoid=Annoid.replace("RS#:", "");
					int start = annotation.getLocations().get(0).getOffset()-passage.getOffset();
					int last = start + annotation.getLocations().get(0).getLength();
					String AnnoMention=annotation.getText();
					String prefix="";
					String postfix="";
					
					if(start>20)
					{
						prefix=passage.getText().substring(start-20,start);
					}
					else
					{
						prefix=passage.getText().substring(0,start);
					}
					
					if(passage.getText().length()-last>20)
					{
						postfix=passage.getText().substring(last,last+20);
					}
					else
					{
						postfix=passage.getText().substring(last,passage.getText().length());
					}
					
					HashMap <String,String> jo = new HashMap <String,String>();
					jo.put("ann", "http://rdf.ebi.ac.uk/resource/europepmc/annotations/PMC"+PMID+"#"+count_passage+"-"+count_ann);
					jo.put("position",count_passage+"."+count_ann);
					if(Annoid.matches("[0-9]+"))
					{
						jo.put("tag", "http://identifiers.org/dbsnp/rs"+Annoid);
					}
					jo.put("prefix", prefix);
					jo.put("exact", AnnoMention);
					jo.put("postfix", postfix);
					jo.put("pmcid", "PMC"+PMID);
					Gson gson = new Gson(); 
					String json = gson.toJson(jo); 
					outputfile.write(json+"\n");
				}
			}
		}
		if(duplicate == true){System.exit(0);}
		outputfile.close();
	}
	public static void FreeText2PubTator(String input,String output) throws IOException
	{
		File file = new File(input);
        String input_filename = file.getName();
    	
        BufferedReader inputfile = new BufferedReader(new InputStreamReader(new FileInputStream(input), "UTF-8"));
		BufferedWriter outputfile = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		
		String line;
		int count_line=0;
		while ((line = inputfile.readLine()) != null)  
		{
			line=line.replaceAll("^[\t ]+", "");
			if(!line.equals(""))
			{
				if(count_line==0)
				{
					outputfile.write(input_filename+"|t|"+line+"\n");
				}
				else if(count_line==1)
				{
					outputfile.write(input_filename+"|a|"+line);
				}
				else
				{
					outputfile.write(" "+line);
				}
				count_line++;
			}
		}
		outputfile.write("\n");
		inputfile.close();
		outputfile.close();
	}
	public static void FreeText2BioC(String input,String output) throws IOException, XMLStreamException
	{
		String parser = BioCFactory.WOODSTOX;
        BioCFactory factory = BioCFactory.newFactory(parser);
        BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), StandardCharsets.UTF_8));
        BioCCollection biocCollection = new BioCCollection();
        ZoneId zonedId = ZoneId.of("America/Montreal");
        LocalDate today = LocalDate.now(zonedId);
        biocCollection.setDate(today.toString());
        biocCollection.setKey("BioC.key");
        biocCollection.setSource("BioC");

        BioCOutputFormat.writeCollectionInfo(biocCollection);
        BioCDocument biocDocument = new BioCDocument();
        biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));

        int startoffset = 0;
        BioCPassage biocPassage;
        Map<String, String> Infons;
        int tableCount = 1;
        int textCount = 1;

        // Reading the .txt file line by line
        BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(input), "UTF-8"));
		 
        String line;
        StringBuilder tableBuilder = new StringBuilder();
        StringBuilder plainTextBuilder = new StringBuilder();
        boolean inTable = false;

        while ((line = reader.readLine()) != null) 
        {
            line = line.trim();  // Trim whitespace around the line

            if (line.contains("\t")) {  // Check if it's a table line
                if (!inTable) {
                    // Starting a new table
                    inTable = true;
                    tableBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?> <table border=1><tbody>");
                }
                // Adding table rows and cells
                tableBuilder.append("<tr>");
                String[] cells = line.split("\t");
                for (String cell : cells) {
                    String refinedCell = string_refine(cell);
                    tableBuilder.append("<td>").append(refinedCell).append("</td>");
                    plainTextBuilder.append(refinedCell).append(" \u2028"); // Add line separator
                }
                tableBuilder.append("</tr>");
                plainTextBuilder.append(" \u2029"); // Add paragraph separator
            } else {
                // If a new non-table line is encountered while in a table, close the table
                if (inTable) {
                    tableBuilder.append("</tbody></table>");
                    // Create a new BioCPassage for the table
                    biocPassage = new BioCPassage();
                    Infons = new HashMap<>();
                    Infons.put("xml", tableBuilder.toString());
                    Infons.put("id", "Table_" + tableCount);
                    Infons.put("section_type", "TABLE");
                    Infons.put("type", "table");
                    biocPassage.setInfons(Infons);
                    biocPassage.setOffset(startoffset);
                    biocPassage.setText(plainTextBuilder.toString());
                    startoffset += plainTextBuilder.length() + 1;

                    biocDocument.addPassage(biocPassage);
                    tableCount++;

                    // Reset the table builders for the next table
                    tableBuilder.setLength(0);
                    plainTextBuilder.setLength(0);
                    inTable = false;
                }

                // Add non-table text as a new BioCPassage
                if (!line.isEmpty()) {
                    biocPassage = new BioCPassage();
                    Infons = new HashMap<>();
                    Infons.put("id", "Text_" + textCount);
                    Infons.put("section_type", "TEXT");
                    Infons.put("type", "Text");
                    biocPassage.setInfons(Infons);
                    biocPassage.setText(string_refine(line));
                    biocPassage.setOffset(startoffset);
                    startoffset += line.length() + 1;

                    biocDocument.addPassage(biocPassage);
                    textCount++;
                }
            }
        }

        // If the file ends with a table, make sure to close it
        if (inTable) {
            tableBuilder.append("</tbody></table>");
            biocPassage = new BioCPassage();
            Infons = new HashMap<>();
            Infons.put("xml", tableBuilder.toString());
            Infons.put("id", "Table_" + tableCount);
            Infons.put("section_type", "TABLE");
            Infons.put("type", "table");
            biocPassage.setInfons(Infons);
            biocPassage.setOffset(startoffset);
            biocPassage.setText(plainTextBuilder.toString());
            startoffset += plainTextBuilder.length() + 1;

            biocDocument.addPassage(biocPassage);
        }
        
        biocCollection.addDocument(biocDocument);
        BioCOutputFormat.writeDocument(biocDocument);
        BioCOutputFormat.close();
	}
	public static void PubTator2HTML(String input,String output) throws IOException, XMLStreamException
	{
		ArrayList<String> color_arr = new ArrayList<String>();
		int color_arr_count=0;
		color_arr.add("255,153,0");color_arr.add("102,204,0");color_arr.add("200,64,240");color_arr.add("0,208,255");color_arr.add("130,210,170");color_arr.add("250,150,150");color_arr.add("150,150,250");color_arr.add("150,250,250");color_arr.add("250,150,250");color_arr.add("180,80,180");color_arr.add("250,220,180");color_arr.add("180,180,80");color_arr.add("230,230,230");color_arr.add("230,230,130");color_arr.add("230,230,30");color_arr.add("230,180,230");color_arr.add("230,180,130");color_arr.add("230,180,30");color_arr.add("230,130,230");color_arr.add("230,130,130");color_arr.add("230,130,30");color_arr.add("230,80,230");color_arr.add("230,80,130");color_arr.add("230,80,30");color_arr.add("230,30,230");color_arr.add("230,30,130");color_arr.add("230,30,30");color_arr.add("180,230,230");color_arr.add("180,230,130");color_arr.add("180,230,30");color_arr.add("180,180,230");color_arr.add("180,180,130");color_arr.add("180,180,30");color_arr.add("180,130,224");color_arr.add("180,130,130");color_arr.add("180,130,30");color_arr.add("180,80,230");color_arr.add("180,80,130");color_arr.add("180,80,30");color_arr.add("180,30,230");color_arr.add("180,30,130");color_arr.add("180,30,30");color_arr.add("130,230,230");color_arr.add("130,230,130");color_arr.add("130,230,30");color_arr.add("130,180,230");color_arr.add("130,180,130");color_arr.add("130,180,30");color_arr.add("130,130,230");color_arr.add("130,130,130");color_arr.add("130,130,30");color_arr.add("130,80,230");color_arr.add("130,80,130");color_arr.add("130,80,30");color_arr.add("130,30,230");color_arr.add("130,30,130");color_arr.add("130,30,30");color_arr.add("80,230,230");color_arr.add("80,230,130");color_arr.add("80,230,30");color_arr.add("80,180,230");color_arr.add("80,180,130");color_arr.add("80,180,30");color_arr.add("80,130,230");color_arr.add("80,130,130");color_arr.add("80,130,30");color_arr.add("80,80,230");color_arr.add("80,80,130");color_arr.add("80,80,30");color_arr.add("80,30,230");color_arr.add("80,30,130");color_arr.add("80,30,30");
		HashMap<String,String> color_hash = new HashMap<String,String> ();
		
		BufferedReader inputfile = new BufferedReader(new InputStreamReader(new FileInputStream(input), "UTF-8"));
		HashMap<String, String> annotation_hash = new HashMap<String, String>();
		HashMap<String, Integer> annotation_count_hash = new HashMap<String, Integer>();
		ArrayList<String> ParagraphType=new ArrayList<String>(); // Type: Title|Abstract
		ArrayList<String> ParagraphContent = new ArrayList<String>(); // Text
		ArrayList<String> annotation_arr = new ArrayList<String>(); // Annotation
		HashMap<Integer, String> annotation_mention_hash = new HashMap<Integer, String>();
		String line;
		String Pmid="";
		String output_STR="";
		int count_anno=0;
		while ((line = inputfile.readLine()) != null)  
		{
			if(line.contains("|") && !line.contains("\t")) //Title|Abstract
        	{
				String str[]=line.split("\\|",-1);
				Pmid=str[0];
				if(str[1].equals("t"))
				{
					str[1]="title";
				}
				if(str[1].equals("a"))
				{
					str[1]="abstract";
				}
				ParagraphType.add(str[1]);
				if(str.length==3)
				{
					ParagraphContent.add(str[2]);
				}
				else
				{
					ParagraphContent.add("");
				}
        	}
			else if (line.contains("\t")) //Annotation
        	{
				String anno[]=line.split("\t");
				String start=anno[1];
				String last=anno[2];
				String AnnoMention=anno[3];
				String Annotype=anno[4];
				
				if(anno.length==6)
				{
					String Annoid=anno[5];
					annotation_mention_hash.put(Integer.parseInt(start), start+"\t"+last+"\t"+AnnoMention+"\t"+Annotype+"\t"+Annoid);
				}
				else if(anno.length==5)
				{
					annotation_mention_hash.put(Integer.parseInt(start), start+"\t"+last+"\t"+AnnoMention+"\t"+Annotype);
				}
				count_anno++;
        	}
			else if(line.length()==0) //Processing
			{
				String Paragraphs="";
				for(int i=0;i<ParagraphContent.size();i++)
				{
					Paragraphs=Paragraphs+ParagraphContent.get(i)+" ";
				}
				
				while(count_anno>0)
				{
					int max_start=0;
					for(Integer start : annotation_mention_hash.keySet())
					{
						if(start>max_start)
						{
							max_start=start;
						}
					}
					annotation_arr.add(annotation_mention_hash.get(max_start));
					annotation_mention_hash.remove(max_start);
					count_anno--;
				}
				for(int x=0;x<annotation_arr.size();x++)
				{
					String str[]=annotation_arr.get(x).split("\\t");
					int start = Integer.parseInt(str[0]);
					int last = Integer.parseInt(str[1]);
					String mention=str[2];
					String type=str[3];
					String id="";
					if(str.length==5)
					{
						id=str[4];
					}
					annotation_hash.put(type+"\t"+id,mention);
					if(!annotation_count_hash.containsKey(type+"\t"+id))
					{
						annotation_count_hash.put(type+"\t"+id,1);
					}
					else
					{
						annotation_count_hash.put(type+"\t"+id,annotation_count_hash.get(type+"\t"+id)+1);
					}
					String pre=Paragraphs.substring(0, start);
					String post=Paragraphs.substring(last, Paragraphs.length());
					if(!color_hash.containsKey(type))
					{
						color_hash.put(type, color_arr.get(color_arr_count));
						color_arr_count++;
					}
					Paragraphs=pre+"<font style=\"background-color: rgb("+color_hash.get(type)+")\" title='"+id+"'>"+mention+"</font>"+post;
				}
				output_STR=output_STR+"PMID:"+Pmid+"<BR />"+Paragraphs+"<BR /><BR />\n";
				
				ParagraphType.clear();
				ParagraphContent.clear();
				annotation_arr.clear();
				annotation_mention_hash.clear();
				count_anno=0;
			}
		}
		inputfile.close();
		
		BufferedWriter HTMLOutputFormat = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		HTMLOutputFormat.write("<!DOCTYPE html>\n<html><head>\n<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>\n<title>BioC Documents</title>\n</head><body>");
		HTMLOutputFormat.write("<table border=1><tr><td>Type</td><td>concepts (identifiers) - mentioned frequency</td></tr>");
		for(String type: color_hash.keySet())
		{
			HTMLOutputFormat.write("<tr style=\"background-color: rgb("+color_hash.get(type)+")\">");
			HTMLOutputFormat.write("<td>"+type+"</td>");
			HTMLOutputFormat.write("<td>");
			
			for(String typeid: annotation_hash.keySet())
			{
				String type_id[]=typeid.split("\\t");
				if(type_id[0].equals(type))
				{
					if(type_id[0].equals(type_id[1]))
					{
						HTMLOutputFormat.write(annotation_hash.get(typeid)+" - "+annotation_count_hash.get(typeid)+"<BR />");
					}
					else
					{
						HTMLOutputFormat.write(annotation_hash.get(typeid)+" ("+type_id[1]+") - "+annotation_count_hash.get(typeid)+"<BR />");
					}
				}
			}
			HTMLOutputFormat.write("</td>");
			HTMLOutputFormat.write("</tr>");
		}
		HTMLOutputFormat.write("</table><BR />");
		HTMLOutputFormat.write(output_STR);
		HTMLOutputFormat.write("</body></html>");
		HTMLOutputFormat.close();
		
	}
	public static void BioC2HTML(String input,String output) throws IOException, XMLStreamException
	{
		ArrayList<String> color_arr = new ArrayList<String>();
		int color_arr_count=0;
		color_arr.add("255,153,0");color_arr.add("102,204,0");color_arr.add("200,64,240");color_arr.add("0,208,255");color_arr.add("130,210,170");color_arr.add("250,150,150");color_arr.add("150,150,250");color_arr.add("150,250,250");color_arr.add("250,150,250");color_arr.add("180,80,180");color_arr.add("250,220,180");color_arr.add("180,180,80");color_arr.add("230,230,230");color_arr.add("230,230,130");color_arr.add("230,230,30");color_arr.add("230,180,230");color_arr.add("230,180,130");color_arr.add("230,180,30");color_arr.add("230,130,230");color_arr.add("230,130,130");color_arr.add("230,130,30");color_arr.add("230,80,230");color_arr.add("230,80,130");color_arr.add("230,80,30");color_arr.add("230,30,230");color_arr.add("230,30,130");color_arr.add("230,30,30");color_arr.add("180,230,230");color_arr.add("180,230,130");color_arr.add("180,230,30");color_arr.add("180,180,230");color_arr.add("180,180,130");color_arr.add("180,180,30");color_arr.add("180,130,224");color_arr.add("180,130,130");color_arr.add("180,130,30");color_arr.add("180,80,230");color_arr.add("180,80,130");color_arr.add("180,80,30");color_arr.add("180,30,230");color_arr.add("180,30,130");color_arr.add("180,30,30");color_arr.add("130,230,230");color_arr.add("130,230,130");color_arr.add("130,230,30");color_arr.add("130,180,230");color_arr.add("130,180,130");color_arr.add("130,180,30");color_arr.add("130,130,230");color_arr.add("130,130,130");color_arr.add("130,130,30");color_arr.add("130,80,230");color_arr.add("130,80,130");color_arr.add("130,80,30");color_arr.add("130,30,230");color_arr.add("130,30,130");color_arr.add("130,30,30");color_arr.add("80,230,230");color_arr.add("80,230,130");color_arr.add("80,230,30");color_arr.add("80,180,230");color_arr.add("80,180,130");color_arr.add("80,180,30");color_arr.add("80,130,230");color_arr.add("80,130,130");color_arr.add("80,130,30");color_arr.add("80,80,230");color_arr.add("80,80,130");color_arr.add("80,80,30");color_arr.add("80,30,230");color_arr.add("80,30,130");color_arr.add("80,30,30");
		HashMap<String,String> color_hash = new HashMap<String,String> ();
		
		HashMap<String, String> pmidlist = new HashMap<String, String>(); // check if appear duplicate pmids
		boolean duplicate = false;
		ConnectorWoodstox connector = new ConnectorWoodstox();
		BioCCollection collection = new BioCCollection();
		collection = connector.startRead(new InputStreamReader(new FileInputStream(input), "UTF-8"));
		String output_STR="";
		HashMap<String, String> annotation_hash = new HashMap<String, String>();
		HashMap<String, Integer> annotation_count_hash = new HashMap<String, Integer>();
		while (connector.hasNext()) 
		{
			BioCDocument document = connector.next();
			String PMID = document.getID();
			if(pmidlist.containsKey(PMID)){System.out.println("\n[Warning]: duplicate pmid-"+PMID);duplicate = true;}
			else{pmidlist.put(PMID,"");}
			for (BioCPassage passage : document.getPassages()) 
			{
				String passage_text=passage.getText();
				HashMap<Integer, String> annotation_mention_hash = new HashMap<Integer, String>();
				ArrayList<String> annotation_arr = new ArrayList<String>();
				int count_anno=0;
				for (BioCAnnotation annotation : passage.getAnnotations()) 
				{
					String Annotype = annotation.getInfon("type");
					int start = annotation.getLocations().get(0).getOffset();
					int last = start + annotation.getLocations().get(0).getLength();
					String AnnoMention=annotation.getText();
					Map<String,String> Infons = annotation.getInfons();
					String Annoid = "";
					for(String InfonType : Infons.keySet())
					{
						if(!InfonType.equals("type"))
						{
							if(Annoid.equals(""))
							{
								Annoid=Infons.get(InfonType);
							}
							else
							{
								Annoid=Annoid+"|"+Infons.get(InfonType);
							}
						}
					}
					if(Annoid.equals(""))
					{
						annotation_mention_hash.put(start, start+"\t"+last+"\t"+AnnoMention+"\t"+Annotype);
					}
					else
					{
						annotation_mention_hash.put(start, start+"\t"+last+"\t"+AnnoMention+"\t"+Annotype+"\t"+Annoid);
					}
						
					count_anno++;
				}
				while(count_anno>0)
				{
					int max_start=0;
					for(Integer start : annotation_mention_hash.keySet())
					{
						if(start>max_start)
						{
							max_start=start;
						}
					}
					annotation_arr.add(annotation_mention_hash.get(max_start));
					annotation_mention_hash.remove(max_start);
					count_anno--;
				}
				for(int x=0;x<annotation_arr.size();x++)
				{
					String str[]=annotation_arr.get(x).split("\\t");
					int start = Integer.parseInt(str[0])-passage.getOffset();
					int last = Integer.parseInt(str[1])-passage.getOffset();
					String mention=str[2];
					String type=str[3];
					String id="";
					if(str.length==5)
					{
						id=str[4];
					}
					annotation_hash.put(type+"\t"+id,mention);
					if(!annotation_count_hash.containsKey(type+"\t"+id))
					{
						annotation_count_hash.put(type+"\t"+id,1);
					}
					else
					{
						annotation_count_hash.put(type+"\t"+id,annotation_count_hash.get(type+"\t"+id)+1);
					}
					String pre=passage_text.substring(0, start);
					String post=passage_text.substring(last, passage_text.length());
					if(!color_hash.containsKey(type))
					{
						color_hash.put(type, color_arr.get(color_arr_count));
						color_arr_count++;
					}
					passage_text=pre+"<font style=\"background-color: rgb("+color_hash.get(type)+")\" title='"+id+"'>"+mention+"</font>"+post;
				}
				output_STR=output_STR+passage_text+"<BR /><BR />\n";
			}
		}
		if(duplicate == true){System.exit(0);}
		
		BufferedWriter HTMLOutputFormat = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		HTMLOutputFormat.write("<!DOCTYPE html>\n<html><head>\n<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>\n<title>BioC Documents</title>\n</head><body>");
		HTMLOutputFormat.write("<table border=1><tr><td>Type</td><td>concepts (identifiers) - mentioned frequency</td></tr>");
		for(String type: color_hash.keySet())
		{
			HTMLOutputFormat.write("<tr style=\"background-color: rgb("+color_hash.get(type)+")\">");
			HTMLOutputFormat.write("<td>"+type+"</td>");
			HTMLOutputFormat.write("<td>");
			
			for(String typeid: annotation_hash.keySet())
			{
				String type_id[]=typeid.split("\\t");
				if(type_id[0].equals(type))
				{
					if(type_id[0].equals(type_id[1]))
					{
						HTMLOutputFormat.write(annotation_hash.get(typeid)+" - "+annotation_count_hash.get(typeid)+"<BR />");
					}
					else
					{
						HTMLOutputFormat.write(annotation_hash.get(typeid)+" ("+type_id[1]+") - "+annotation_count_hash.get(typeid)+"<BR />");
					}
				}
			}
			HTMLOutputFormat.write("</td>");
			HTMLOutputFormat.write("</tr>");
		}
		HTMLOutputFormat.write("</table><BR />");
		HTMLOutputFormat.write(output_STR);
		HTMLOutputFormat.write("</body></html>");
		HTMLOutputFormat.close();
	}
	public static void XML2BioC_recursive(Node node, String name) throws IOException, XMLStreamException, ParserConfigurationException, SAXException // for XML2BioC
	{
		if(!node.hasChildNodes())
		{
			String content = node.getTextContent();
			content = content.replaceAll("[\t \r\n]+", " ");
			if(content.length()>5)
			{
				XML_names.add(name);
				XML_contents.add(content);
			}
		}
		else
		{
			//NamedNodeMap test= node.getAttributes(); /*extract attributes*/
			NodeList children=node.getChildNodes();
			if(name.equals(""))
			{
				name=node.getNodeName();
			}
			else
			{
				name=name+"/"+node.getNodeName();
			}
			for(int i=0;i<children.getLength();i++)
			{
				XML2BioC_recursive(children.item(i),name);
			}
		}
	}
	public static void XML2BioC(String input,String output) throws IOException, XMLStreamException, ParserConfigurationException, SAXException
	{
		String parser = BioCFactory.WOODSTOX;
		BioCFactory factory = BioCFactory.newFactory(parser);
		BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		BioCCollection biocCollection = new BioCCollection();
		ZoneId zonedId = ZoneId.of( "America/Montreal" );
		LocalDate today = LocalDate.now( zonedId );
		biocCollection.setDate(today.toString());
		biocCollection.setKey("BioC.key");//key
		biocCollection.setSource("BioC");//source
		
		File fXmlFile = new File(input);
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(fXmlFile);
		NodeList nodes=doc.getChildNodes();
		
		BioCOutputFormat.writeCollectionInfo(biocCollection);
		BioCDocument biocDocument = new BioCDocument();
		int startoffset=0;
		int count_line=0;
		String ID=input.replaceAll("^.*/", "");
		ID=ID.replaceAll(".xml", "");
		if(ID.equals("")){ID="1";}
		biocDocument.setID(ID);
		
		for(int i=0;i<nodes.getLength();i++)
		{	
			XML2BioC_recursive(nodes.item(i),"");
		}
		for (int i=0;i<XML_names.size();i++)
		{
			String name = XML_names.get(i);
			String content = XML_contents.get(i);
			
			count_line++;
			BioCPassage biocPassage = new BioCPassage();
			Map<String, String> Infons = new HashMap<String, String>();
			Infons.put("type", name);
			biocPassage.setInfons(Infons);
			biocPassage.setText(string_refine(content));
			biocPassage.setOffset(startoffset);
			startoffset=startoffset+content.length()+1;
			biocDocument.addPassage(biocPassage);
		}
		
		biocCollection.addDocument(biocDocument);
		BioCOutputFormat.writeDocument(biocDocument);
		BioCOutputFormat.close();
	}
	
	/*
	 * For supplementary materials
	 */
	public static void PDF2BioC(String input,String output) throws IOException, XMLStreamException
	{
		/**
		 * pdfbox
		 */
		File file = new File(input);
		PDFParser pdfparser = new PDFParser(new RandomAccessFile(file,"r"));  
		pdfparser.parse();
		COSDocument cosDoc = pdfparser.getDocument();
		PDFTextStripper pdfStripper = new PDFTextStripper();
		PDDocument pdDoc = new PDDocument(cosDoc);
		
		String parser = BioCFactory.WOODSTOX;
		BioCFactory factory = BioCFactory.newFactory(parser);
		BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		BioCCollection biocCollection = new BioCCollection();
		ZoneId zonedId = ZoneId.of( "America/Montreal" );
		LocalDate today = LocalDate.now( zonedId );
		biocCollection.setDate(today.toString());
		biocCollection.setKey("BioC.key");//key
		biocCollection.setSource("BioC");//source
		
		BioCOutputFormat.writeCollectionInfo(biocCollection);
		BioCDocument biocDocument = new BioCDocument();
		biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));
		
		int count_line=1;
		int startoffset=0;
		
		BioCPassage biocPassage = new BioCPassage();
		Map<String, String> Infons = new HashMap<String, String>();
		Infons.put("id", "Front_1");
        Infons.put("section_type", "TITLE");
        Infons.put("type", "front");
        //Infons.put("license", "CC BY");
        biocPassage.setInfons(Infons);
        biocPassage.setText("Supplementary Material");
		biocPassage.setOffset(startoffset);
		startoffset=startoffset+"Supplementary Material".length()+1;
		biocDocument.addPassage(biocPassage);
		
		for(int i=1;i<=pdDoc.getNumberOfPages();i++)
		{
			pdfStripper.setStartPage(i);
			pdfStripper.setEndPage(i);
			String line = "";
			
			try {
		        line = pdfStripper.getText(pdDoc);
		    } catch (NullPointerException e) {
		        //System.out.println("Line_"+count_line+":"+line);
		        continue;
		    }
			
			line=line.replaceAll("[^\\x09\\x0A\\x0D\\x20-\\xD7FF\\xE000-\\xFFFD\\x10000-x10FFFF]"," ");
			line=line.replaceAll("[\n\r\t]+", " ");
			line=line.replaceAll(" [ ]+", " ");
			if(!line.equals(""))
			{
				count_line++;
				biocPassage = new BioCPassage();
				Infons = new HashMap<String, String>();
				Infons.put("id", "Line_" + count_line);
	            Infons.put("section_type", "TEXT");
	            Infons.put("type", "text");
	            biocPassage.setInfons(Infons);
	            line=string_refine(line);
				biocPassage.setText(line);
				biocPassage.setOffset(startoffset);
				startoffset=startoffset+line.length()+1;
				biocDocument.addPassage(biocPassage);
			}
		}
		
		biocCollection.addDocument(biocDocument);
		BioCOutputFormat.writeDocument(biocDocument);
		BioCOutputFormat.close();
		
		cosDoc.close();
		pdDoc.close();
	}
	public static void PDF2PubTator(String input,String output) throws IOException, XMLStreamException
	{
		/**
		 * pdfbox
		 */
		
		BufferedWriter outputfile  = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		
		File file = new File(input);
		PDFParser pdfparser = new PDFParser(new RandomAccessFile(file,"r"));  
		pdfparser.parse();
		COSDocument cosDoc = pdfparser.getDocument();
		PDFTextStripper pdfStripper = new PDFTextStripper();
		PDDocument pdDoc = new PDDocument(cosDoc);
		
		input=input.replaceAll("[^0-9]","");
		if(input.equals("")){input="1";}
		int count_line=0;
		for(int i=1;i<=pdDoc.getNumberOfPages();i++)
		{
			pdfStripper.setStartPage(i);
			pdfStripper.setEndPage(i);
			String line = pdfStripper.getText(pdDoc);
			line=line.replaceAll("[^\\x09\\x0A\\x0D\\x20-\\xD7FF\\xE000-\\xFFFD\\x10000-x10FFFF]"," ");
			line=line.replaceAll("[\n\r\t]+", " ");
			line=line.replaceAll(" [ ]+", " ");
			if(!line.equals(""))
			{
				outputfile.write(input+"|Line_"+count_line+"|"+line+"\n");
				count_line++;
			}
		}
		
		outputfile.write("\n");
		outputfile.close();
		cosDoc.close();
		pdDoc.close();
	}
	public static void pptx2BioC(String input,String output) throws IOException, XMLStreamException, ParserConfigurationException, SAXException // for both ppt and pptx
	{
		String parser = BioCFactory.WOODSTOX;
	    BioCFactory factory = BioCFactory.newFactory(parser);
	    BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
	    BioCCollection biocCollection = new BioCCollection();
	    ZoneId zonedId = ZoneId.of("America/Montreal");
	    LocalDate today = LocalDate.now(zonedId);
	    biocCollection.setDate(today.toString());
	    biocCollection.setKey("BioC.key");
	    biocCollection.setSource("BioC");

	    BioCOutputFormat.writeCollectionInfo(biocCollection);
	    BioCDocument biocDocument = new BioCDocument();
	    biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));
		
	    int startoffset=0;
		BioCPassage biocPassage = new BioCPassage();
		Map<String, String> Infons = new HashMap<String, String>();
		Infons.put("id", "Front_1");
        Infons.put("section_type", "TITLE");
        Infons.put("type", "front");
        //Infons.put("license", "CC BY");
        biocPassage.setInfons(Infons);
        biocPassage.setText("Supplementary Material");
		biocPassage.setOffset(startoffset);
		startoffset=startoffset+"Supplementary Material".length()+1;
		biocDocument.addPassage(biocPassage);
		
		try (FileInputStream fis = new FileInputStream(input);
	         XMLSlideShow ppt = new XMLSlideShow(fis)) 
	    {

	        int count_slide = 1;
	        for (XSLFSlide slide : ppt.getSlides()) 
	        {
	            StringBuilder xmlStringBuilder = new StringBuilder();
	            StringBuilder plainTextBuilder = new StringBuilder();
	            xmlStringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
	            boolean table_boolean=false;
	            for (XSLFShape shape : slide.getShapes()) 
	            {
	                if (shape instanceof org.apache.poi.xslf.usermodel.XSLFTextShape) 
	                {
	                    org.apache.poi.xslf.usermodel.XSLFTextShape textShape = (org.apache.poi.xslf.usermodel.XSLFTextShape) shape;
	                    String line = textShape.getText();
	                    line=string_refine(line);
	                    xmlStringBuilder.append("<text>").append(line).append("</text>");
	                    plainTextBuilder.append(line).append(" \u2028"); // Add line separator
	                } 
	                else if (shape instanceof XSLFTable) 
	                {
	                    XSLFTable table = (XSLFTable) shape;
	                    xmlStringBuilder.append("<table border=1><tbody>");
	                    for (XSLFTableRow row : table.getRows()) 
	                    {
	                        xmlStringBuilder.append("<tr>");
	                        for (XSLFTableCell cell : row.getCells()) 
	                        {
	                            String cellText = cell.getText();
	                            cellText=string_refine(cellText);
	                            xmlStringBuilder.append("<td>").append(cellText).append("</td>");
	                            plainTextBuilder.append(cellText).append(" \u2028"); // Add line separator
	                        }
	                        xmlStringBuilder.append("</tr>");
	                        plainTextBuilder.append(" \u2029"); // Add paragraph separator
	                    }
	                    xmlStringBuilder.append("</tbody></table>");
	                    table_boolean=true;
	                }
	            }
	            
	            biocPassage = new BioCPassage();
	            Infons = new HashMap<>();
	            if(table_boolean==true)
	            {
		            Infons.put("xml", xmlStringBuilder.toString());
		            Infons.put("id", "Slide_" + count_slide);
		            Infons.put("section_type", "TABLE");
		            Infons.put("type", "table");
		        }
	            else
            	{
		            Infons.put("id", "Slide_" + count_slide);
		            Infons.put("section_type", "SLIDE");
		            Infons.put("type", "slide");
		        }
	            biocPassage.setInfons(Infons);
	            biocPassage.setOffset(startoffset);
	            biocPassage.setText(plainTextBuilder.toString());
	            startoffset += plainTextBuilder.length() + 1;
	            biocDocument.addPassage(biocPassage);
	            count_slide++;
	        }
	    } catch (IOException e) {
	        e.printStackTrace();
	    }

	    biocCollection.addDocument(biocDocument);
	    BioCOutputFormat.writeDocument(biocDocument);
	    BioCOutputFormat.close();
	}
	public static void ppt2BioC(String input,String output) throws IOException, XMLStreamException, ParserConfigurationException, SAXException // for both ppt and pptx
	{
		String parser = BioCFactory.WOODSTOX;
	    BioCFactory factory = BioCFactory.newFactory(parser);
	    BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
	    BioCCollection biocCollection = new BioCCollection();
	    ZoneId zonedId = ZoneId.of("America/Montreal");
	    LocalDate today = LocalDate.now(zonedId);
	    biocCollection.setDate(today.toString());
	    biocCollection.setKey("BioC.key");
	    biocCollection.setSource("BioC");

	    BioCOutputFormat.writeCollectionInfo(biocCollection);
	    BioCDocument biocDocument = new BioCDocument();
	    biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));
	    
	    int startoffset=0;
		BioCPassage biocPassage = new BioCPassage();
		Map<String, String> Infons = new HashMap<String, String>();
		Infons.put("id", "Front_1");
        Infons.put("section_type", "TITLE");
        Infons.put("type", "front");
        //Infons.put("license", "CC BY");
        biocPassage.setInfons(Infons);
        biocPassage.setText("Supplementary Material");
		biocPassage.setOffset(startoffset);
		startoffset=startoffset+"Supplementary Material".length()+1;
		biocDocument.addPassage(biocPassage);
		
		try (FileInputStream fis = new FileInputStream(input);
	         HSLFSlideShow ppt = new HSLFSlideShow(fis)) {

	        int count_slide = 1;
	        for (HSLFSlide slide : ppt.getSlides()) 
	        {
	            StringBuilder xmlStringBuilder = new StringBuilder();
	            StringBuilder plainTextBuilder = new StringBuilder();
	            xmlStringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
	            boolean table_boolean=false;
	            for (HSLFShape shape : slide.getShapes())
	            {
	                if (shape instanceof HSLFTextShape) 
	                {
	                    HSLFTextShape textShape = (HSLFTextShape) shape;
	                    String line = textShape.getText();
	                    line=string_refine(line);
	                    xmlStringBuilder.append("<text>").append(line).append("</text>");
	                    plainTextBuilder.append(line).append(" \u2028"); // Add line separator
	                } 
	                else if (shape instanceof HSLFTable) 
	                {
	                	HSLFTable table = (HSLFTable) shape;
	                    xmlStringBuilder.append("<table border=1><tbody>");
	                    
	                    for(int x=0;x<table.getNumberOfRows();x++)
	                    {
	                    	xmlStringBuilder.append("<tr>");
	                    	for(int y=0;y<table.getNumberOfColumns();y++)
	 	                    {
	                    		String cellText_str="";
	                    		if(table.getCell(x, y) != null)
	                    		{
	                    			cellText_str = table.getCell(x, y).getText();
	                    		}
	                    		cellText_str=string_refine(cellText_str);
	    	                    xmlStringBuilder.append("<td>").append(cellText_str).append("</td>");
	                            plainTextBuilder.append(cellText_str).append(" \u2028"); // Add line separator
	 	                    }
	                    	xmlStringBuilder.append("</tr>");
	                        plainTextBuilder.append(" \u2029"); // Add paragraph separator
	                    }
	                    xmlStringBuilder.append("</tbody></table>");
	                    table_boolean=true;
	                }
	            }
	            plainTextBuilder.append(" \u2029"); // Add paragraph separator

	            biocPassage = new BioCPassage();
	            Infons = new HashMap<>();
	            if(table_boolean==true)
	            {
		            Infons.put("xml", xmlStringBuilder.toString());
		            Infons.put("id", "Slide_" + count_slide);
		            Infons.put("section_type", "TABLE");
		            Infons.put("type", "table");
		        }
	            else
            	{
		            Infons.put("id", "Slide_" + count_slide);
		            Infons.put("section_type", "SLIDE");
		            Infons.put("type", "slide");
		        }
	            biocPassage.setInfons(Infons);
	            biocPassage.setOffset(startoffset);
	            biocPassage.setText(plainTextBuilder.toString());
	            startoffset += plainTextBuilder.length() + 1;
	            biocDocument.addPassage(biocPassage);
	            count_slide++;
	        }
	    } catch (IOException e) {
	        e.printStackTrace();
	    }

	    biocCollection.addDocument(biocDocument);
	    BioCOutputFormat.writeDocument(biocDocument);
	    BioCOutputFormat.close();
	}
	public static void ppt2PubTator(String input,String output) throws IOException, XMLStreamException, ParserConfigurationException, SAXException
	{
		BufferedWriter outputfile  = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		
		
		try (FileInputStream fis = new FileInputStream(input);
	             XMLSlideShow ppt = new XMLSlideShow(fis)) 
		{

				int count_slide=1;
				for (XSLFSlide slide : ppt.getSlides()) 
	            {
					int count_shape=0;
					for (XSLFShape shape : slide.getShapes()) 
                    {
                    	int count_line=0;
        				count_shape++;
        				
        				if (shape instanceof org.apache.poi.xslf.usermodel.XSLFTextShape) 
                        {
        					count_line++;
            				
        					org.apache.poi.xslf.usermodel.XSLFTextShape textShape = (org.apache.poi.xslf.usermodel.XSLFTextShape) shape;
                            String line=textShape.getText();
                        	line=string_refine(line);
            				outputfile.write(input+"|slide"+count_slide+"_shape"+count_shape+"_line"+count_line+"|"+line+"\n");
                            count_line++;
                        }
                    }
	                System.out.println();
	                count_slide++;
	            }
        } catch (IOException e) {
            e.printStackTrace();
        }
		
		outputfile.write("\n");
		outputfile.close();
	}
	public static void Excel2BioC(String input,String output) throws IOException, XMLStreamException
	{
		String parser = BioCFactory.WOODSTOX;
        BioCFactory factory = BioCFactory.newFactory(parser);
        BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
        BioCCollection biocCollection = new BioCCollection();
        ZoneId zonedId = ZoneId.of("America/Montreal");
        LocalDate today = LocalDate.now(zonedId);
        biocCollection.setDate(today.toString());
        biocCollection.setKey("BioC.key");
        biocCollection.setSource("BioC");

        BioCOutputFormat.writeCollectionInfo(biocCollection);
        BioCDocument biocDocument = new BioCDocument();
        biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));
	    
        int startoffset=0;
		BioCPassage biocPassage = new BioCPassage();
		Map<String, String> Infons = new HashMap<String, String>();
		Infons.put("id", "Front_1");
        Infons.put("section_type", "TITLE");
        Infons.put("type", "front");
        //Infons.put("license", "CC BY");
        biocPassage.setInfons(Infons);
        biocPassage.setText("Supplementary Material");
		biocPassage.setOffset(startoffset);
		startoffset=startoffset+"Supplementary Material".length()+1;
		biocDocument.addPassage(biocPassage);
		
		try (FileInputStream file = new FileInputStream(new File(input));
             Workbook workbook = new HSSFWorkbook(file)) 
        {
            for (int count_table = 1; count_table <= workbook.getNumberOfSheets(); count_table++) 
            {
            	org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(count_table-1);
                StringBuilder xmlStringBuilder = new StringBuilder();
                xmlStringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?> <table border=1><tbody>");
                StringBuilder plainTextBuilder = new StringBuilder();
                
                // Get merged regions for the current sheet
                Map<String, CellRangeAddress> mergedCells = new HashMap<>();
                for (int i = 0; i < sheet.getNumMergedRegions(); i++) 
                {
                    CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                    for (int r = mergedRegion.getFirstRow(); r <= mergedRegion.getLastRow(); r++) 
                    {
                        for (int c = mergedRegion.getFirstColumn(); c <= mergedRegion.getLastColumn(); c++) 
                        {
                            mergedCells.put(r + "," + c, mergedRegion);
                        }
                    }
                }
                
                for (Row row : sheet) {
                    xmlStringBuilder.append("<tr>");
                    for (org.apache.poi.ss.usermodel.Cell cell : row) 
                    {
                        String cellKey = cell.getRowIndex() + "," + cell.getColumnIndex();
                        if (mergedCells.containsKey(cellKey)) 
                        {
                            CellRangeAddress mergedRegion = mergedCells.get(cellKey);
                            if (cell.getRowIndex() == mergedRegion.getFirstRow() && cell.getColumnIndex() == mergedRegion.getFirstColumn()) 
                            {
                                int colspan = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
                                int rowspan = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
                                String cellText_str=string_refine(cell.toString());
                                xmlStringBuilder.append("<td colspan=\"").append(colspan).append("\" rowspan=\"").append(rowspan).append("\">");
                                xmlStringBuilder.append(cellText_str);
                                xmlStringBuilder.append("</td>");
                                plainTextBuilder.append(cellText_str).append(" \u2028"); // Add line separator
                            }
                        }
                        else 
                        {
                        	String cellText_str=string_refine(cell.toString());
                            xmlStringBuilder.append("<td>");
                            xmlStringBuilder.append(cellText_str);
                            xmlStringBuilder.append("</td>");
                            plainTextBuilder.append(cellText_str).append(" \u2028"); // Add line separator
                        }
                    }
                    xmlStringBuilder.append("</tr>");
                    plainTextBuilder.append(" \u2029"); // Add paragraph separator
                }
                xmlStringBuilder.append("</tbody></table>");

                biocPassage = new BioCPassage();
                Infons = new HashMap<>();
                Infons.put("xml", xmlStringBuilder.toString());
                Infons.put("id", "Table_"+count_table);
                Infons.put("section_type", "TABLE");
                Infons.put("type", "table");
                biocPassage.setInfons(Infons);
                biocPassage.setOffset(startoffset);
                biocPassage.setText(plainTextBuilder.toString());
                startoffset += plainTextBuilder.length() + 1;
                biocDocument.addPassage(biocPassage);
            }
        }
        biocCollection.addDocument(biocDocument);
        BioCOutputFormat.writeDocument(biocDocument);
        BioCOutputFormat.close();
	}
	public static void Excel2PubTator(String input,String output) throws IOException, XMLStreamException
	{
		BufferedWriter outputfile  = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));

		try (FileInputStream file = new FileInputStream(new File(input));
				Workbook workbook = new HSSFWorkbook(file);) 
	    {
			for(int count_table=1;count_table<=workbook.getNumberOfSheets();count_table++)
			{
				count_table++;
				int count_row=0;
				org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(count_table-1);
				for (Row row : sheet)
				{
					count_row++;
					String cells="";
					for (org.apache.poi.ss.usermodel.Cell cell : row) 
				    {
				    	cells=cells+cell.toString()+"; ";
				    }
				    outputfile.write((count_table+100000)+"|Table"+count_table+"_Row"+count_row+"|"+cells+"\n");
				}
				if(count_row>0)
				{
					outputfile.write("\n");
				}
			}
			outputfile.close();
		}
	}
	public static void Excelx2BioC(String input,String output) throws IOException, XMLStreamException
	{
		String parser = BioCFactory.WOODSTOX;
        BioCFactory factory = BioCFactory.newFactory(parser);
        BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
        BioCCollection biocCollection = new BioCCollection();
        ZoneId zonedId = ZoneId.of("America/Montreal");
        LocalDate today = LocalDate.now(zonedId);
        biocCollection.setDate(today.toString());
        biocCollection.setKey("BioC.key");
        biocCollection.setSource("BioC");

        BioCOutputFormat.writeCollectionInfo(biocCollection);
        BioCDocument biocDocument = new BioCDocument();
        biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));
	    
        int startoffset=0;
		BioCPassage biocPassage = new BioCPassage();
		Map<String, String> Infons = new HashMap<String, String>();
		Infons.put("id", "Front_1");
        Infons.put("section_type", "TITLE");
        Infons.put("type", "front");
        //Infons.put("license", "CC BY");
        biocPassage.setInfons(Infons);
        biocPassage.setText("Supplementary Material");
		biocPassage.setOffset(startoffset);
		startoffset=startoffset+"Supplementary Material".length()+1;
		biocDocument.addPassage(biocPassage);
		
		try (FileInputStream file = new FileInputStream(new File(input));
             Workbook workbook = new XSSFWorkbook(file)) 
        {
            for (int count_table = 1; count_table <= workbook.getNumberOfSheets(); count_table++) 
            {
                org.apache.poi.ss.usermodel.Sheet  sheet = workbook.getSheetAt(count_table-1);
                StringBuilder xmlStringBuilder = new StringBuilder();
                xmlStringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?> <table border=1><tbody>");
                StringBuilder plainTextBuilder = new StringBuilder();

                // Get merged regions for the current sheet
                Map<String, CellRangeAddress> mergedCells = new HashMap<>();
                for (int i = 0; i < sheet.getNumMergedRegions(); i++) 
                {
                    CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                    for (int r = mergedRegion.getFirstRow(); r <= mergedRegion.getLastRow(); r++) 
                    {
                        for (int c = mergedRegion.getFirstColumn(); c <= mergedRegion.getLastColumn(); c++) 
                        {
                            mergedCells.put(r + "," + c, mergedRegion);
                        }
                    }
                }

                for (Row row : sheet) 
                {
                    xmlStringBuilder.append("<tr>");
                    for (org.apache.poi.ss.usermodel.Cell cell : row) 
                    {
                        String cellKey = cell.getRowIndex() + "," + cell.getColumnIndex();
                        if (mergedCells.containsKey(cellKey)) 
                        {
                            CellRangeAddress mergedRegion = mergedCells.get(cellKey);
                            if (cell.getRowIndex() == mergedRegion.getFirstRow() && cell.getColumnIndex() == mergedRegion.getFirstColumn()) 
                            {
                                int colspan = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
                                int rowspan = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
                                String cellText_str=string_refine(cell.toString());
                                xmlStringBuilder.append("<td colspan=\"").append(colspan).append("\" rowspan=\"").append(rowspan).append("\">");
                                xmlStringBuilder.append(cellText_str);
                                xmlStringBuilder.append("</td>");
                                plainTextBuilder.append(cellText_str).append(" \u2028"); // Add line separator
                            }
                        } 
                        else 
                        {
                        	String cellText_str=string_refine(cell.toString());
                            xmlStringBuilder.append("<td>");
                            xmlStringBuilder.append(cellText_str);
                            xmlStringBuilder.append("</td>");
                            plainTextBuilder.append(cellText_str).append(" \u2028");// Add line separator
                        }
                    }
                    xmlStringBuilder.append("</tr>");
                    plainTextBuilder.append(" \u2029"); // Add paragraph separator
                }
                xmlStringBuilder.append("</tbody></table>");

                biocPassage = new BioCPassage();
                Infons = new HashMap<>();
                Infons.put("xml", xmlStringBuilder.toString());
                Infons.put("id", "Table_" + count_table);
                Infons.put("section_type", "TABLE");
                Infons.put("type", "table");
                biocPassage.setInfons(Infons);
                biocPassage.setOffset(startoffset);
                biocPassage.setText(plainTextBuilder.toString());
                startoffset += plainTextBuilder.length() + 1;
                biocDocument.addPassage(biocPassage);
            }
        }
        biocCollection.addDocument(biocDocument);
        BioCOutputFormat.writeDocument(biocDocument);
        BioCOutputFormat.close();
	}
	public static void Excelx2PubTator(String input,String output) throws IOException, XMLStreamException
	{
		BufferedWriter outputfile  = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		
		try (FileInputStream file = new FileInputStream(new File(input));
				Workbook workbook = new XSSFWorkbook(file);) 
	    {
			for(int i=0;i<workbook.getNumberOfSheets();i++)
			{
				org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(i);
				int j = 0;
				for (Row row : sheet)
				{
					String cells="";
					for (org.apache.poi.ss.usermodel.Cell cell : row) 
				    {
				    	cells=cells+cell.toString()+"; ";
				    }
				    j++;
				    outputfile.write((i+100000)+"|"+j+"|"+cells+"\n");
				}
				if(j>0)
				{
					outputfile.write("\n");
				}
			}
			outputfile.close();
	    }
	}
	public static void CSV2BioC(String input, String output) throws IOException, XMLStreamException, CsvValidationException 
	{
        String parser = BioCFactory.WOODSTOX;
        BioCFactory factory = BioCFactory.newFactory(parser);
        BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), StandardCharsets.UTF_8));
        BioCCollection biocCollection = new BioCCollection();
        ZoneId zonedId = ZoneId.of("America/Montreal");
        LocalDate today = LocalDate.now(zonedId);
        biocCollection.setDate(today.toString());
        biocCollection.setKey("BioC.key");
        biocCollection.setSource("BioC");

        BioCOutputFormat.writeCollectionInfo(biocCollection);
        BioCDocument biocDocument = new BioCDocument();
        biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));

        int startoffset = 0;
        BioCPassage biocPassage;
        Map<String, String> Infons;
        int tableCount = 1;

        // Parse the CSV file using OpenCSV's CSVReader
        try (CSVReader reader = new CSVReader(new InputStreamReader(Files.newInputStream(Paths.get(input)), StandardCharsets.UTF_8))) 
        {
            StringBuilder xmlStringBuilder = new StringBuilder();
            xmlStringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?> <table border=1><tbody>");
            StringBuilder plainTextBuilder = new StringBuilder();

            String[] nextLine;
            while ((nextLine = reader.readNext()) != null) 
            {
                xmlStringBuilder.append("<tr>");
                for (String field : nextLine) 
                {
                    String refinedField = string_refine(field);
                    xmlStringBuilder.append("<td>").append(refinedField).append("</td>");
                    plainTextBuilder.append(refinedField).append(" \u2028"); // Add line separator
                }
                xmlStringBuilder.append("</tr>");
                plainTextBuilder.append(" \u2029"); // Add paragraph separator
            }
            xmlStringBuilder.append("</tbody></table>");

            biocPassage = new BioCPassage();
            Infons = new HashMap<>();
            Infons.put("xml", xmlStringBuilder.toString());
            Infons.put("id", "Table_" + tableCount);
            Infons.put("section_type", "TABLE");
            Infons.put("type", "table");
            biocPassage.setInfons(Infons);
            biocPassage.setOffset(startoffset);
            biocPassage.setText(plainTextBuilder.toString());
            startoffset += plainTextBuilder.length() + 1;

            biocDocument.addPassage(biocPassage);
            tableCount++;
        }

        biocCollection.addDocument(biocDocument);
        BioCOutputFormat.writeDocument(biocDocument);
        BioCOutputFormat.close();
    }
	public static void TSV2BioC(String input, String output) throws IOException, XMLStreamException 
	{
        String parser = BioCFactory.WOODSTOX;
        BioCFactory factory = BioCFactory.newFactory(parser);
        BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
        BioCCollection biocCollection = new BioCCollection();
        ZoneId zonedId = ZoneId.of("America/Montreal");
        LocalDate today = LocalDate.now(zonedId);
        biocCollection.setDate(today.toString());
        biocCollection.setKey("BioC.key");
        biocCollection.setSource("BioC");

        BioCOutputFormat.writeCollectionInfo(biocCollection);
        BioCDocument biocDocument = new BioCDocument();
        biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));

        int startoffset = 0;
        BioCPassage biocPassage = new BioCPassage();
        Map<String, String> Infons = new HashMap<>();
        Infons.put("id", "Front_1");
        Infons.put("section_type", "TITLE");
        Infons.put("type", "front");
        biocPassage.setInfons(Infons);
        biocPassage.setText("Supplementary Material");
        biocPassage.setOffset(startoffset);
        startoffset = startoffset + "Supplementary Material".length() + 1;
        biocDocument.addPassage(biocPassage);

        // Reading CSV or TSV file
        try (BufferedReader reader = Files.newBufferedReader(Paths.get(input), StandardCharsets.UTF_8)) 
        {
            StringBuilder xmlStringBuilder = new StringBuilder();
            xmlStringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?> <table border=1><tbody>");
            StringBuilder plainTextBuilder = new StringBuilder();

            String line;
            int rowNumber = 0;
            while ((line = reader.readLine()) != null) {
                xmlStringBuilder.append("<tr>");
                String[] cells = line.split("\t"); // TSV uses tab as delimiter
                for (String cell : cells) 
                {
                    String refinedCell = string_refine(cell);
                    xmlStringBuilder.append("<td>").append(refinedCell).append("</td>");
                    plainTextBuilder.append(refinedCell).append(" \u2028"); // Add line separator
                }
                xmlStringBuilder.append("</tr>");
                plainTextBuilder.append(" \u2029"); // Add paragraph separator
                rowNumber++;
            }
            xmlStringBuilder.append("</tbody></table>");

            biocPassage = new BioCPassage();
            Infons = new HashMap<>();
            Infons.put("xml", xmlStringBuilder.toString());
            Infons.put("id", "Table_1");
            Infons.put("section_type", "TABLE");
            Infons.put("type", "table");
            biocPassage.setInfons(Infons);
            biocPassage.setOffset(startoffset);
            biocPassage.setText(plainTextBuilder.toString());
            startoffset += plainTextBuilder.length() + 1;
            biocDocument.addPassage(biocPassage);
        }

        biocCollection.addDocument(biocDocument);
        BioCOutputFormat.writeDocument(biocDocument);
        BioCOutputFormat.close();
    }
	public static void RTF2BioC(String input, String output) throws IOException, XMLStreamException 
	{
        FileInputStream file = new FileInputStream(new File(input));

        RTFEditorKit rtfEditorKit = new RTFEditorKit();
        javax.swing.text.Document document = rtfEditorKit.createDefaultDocument();
        try {
            rtfEditorKit.read(file, document, 0);
        } catch (Exception e) {
            e.printStackTrace();
        }
        file.close();

        String parser = BioCFactory.WOODSTOX;
        BioCFactory factory = BioCFactory.newFactory(parser);
        BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
        BioCCollection biocCollection = new BioCCollection();
        ZoneId zonedId = ZoneId.of("America/Montreal");
        LocalDate today = LocalDate.now(zonedId);
        biocCollection.setDate(today.toString());
        biocCollection.setKey("BioC.key");
        biocCollection.setSource("BioC");

        BioCOutputFormat.writeCollectionInfo(biocCollection);
        BioCDocument biocDocument = new BioCDocument();
        biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));

        int startoffset = 0;
        BioCPassage biocPassage = new BioCPassage();
        Map<String, String> Infons = new HashMap<>();
        Infons.put("id", "Front_1");
        Infons.put("section_type", "TITLE");
        Infons.put("type", "front");
        biocPassage.setInfons(Infons);
        biocPassage.setText("Supplementary Material");
        biocPassage.setOffset(startoffset);
        startoffset = startoffset + "Supplementary Material".length() + 1;
        biocDocument.addPassage(biocPassage);

        // Process RTF content as text paragraphs
        String content;
        try {
            content = document.getText(0, document.getLength());
        } catch (BadLocationException e) {
            e.printStackTrace();
            return;
        }

        String[] paragraphs = content.split("\\n\\r?"); // Split by new lines
        int para_count = 1;

        for (String paragraph : paragraphs) {
            if (!paragraph.trim().isEmpty()) {
                biocPassage = new BioCPassage();
                Infons = new HashMap<>();
                Infons.put("id", "Text_" + para_count);
                Infons.put("section_type", "TEXT");
                Infons.put("type", "Text");
                biocPassage.setInfons(Infons);
                String refinedLine = string_refine(paragraph);
                biocPassage.setText(refinedLine);
                biocPassage.setOffset(startoffset);
                startoffset = startoffset + refinedLine.length() + 1;
                biocDocument.addPassage(biocPassage);
                para_count++;
            }
        }

        // If there are tables in the RTF content, you would need to parse and extract them similarly.

        biocCollection.addDocument(biocDocument);
        BioCOutputFormat.writeDocument(biocDocument);
        BioCOutputFormat.close();
    }
	public static void Word2BioC(String input,String output) throws IOException, XMLStreamException
	{
		String parser = BioCFactory.WOODSTOX;
        BioCFactory factory = BioCFactory.newFactory(parser);
        BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
        BioCCollection biocCollection = new BioCCollection();
        ZoneId zonedId = ZoneId.of("America/Montreal");
        LocalDate today = LocalDate.now(zonedId);
        biocCollection.setDate(today.toString());
        biocCollection.setKey("BioC.key");
        biocCollection.setSource("BioC");

        BioCOutputFormat.writeCollectionInfo(biocCollection);
        BioCDocument biocDocument = new BioCDocument();
        biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));
	    
        int startoffset=0;
		BioCPassage biocPassage = new BioCPassage();
		Map<String, String> Infons = new HashMap<String, String>();
		Infons.put("id", "Front_1");
        Infons.put("section_type", "TITLE");
        Infons.put("type", "front");
        //Infons.put("license", "CC BY");
        biocPassage.setInfons(Infons);
        biocPassage.setText("Supplementary Material");
		biocPassage.setOffset(startoffset);
		startoffset=startoffset+"Supplementary Material".length()+1;
		biocDocument.addPassage(biocPassage);
		
		try (FileInputStream file = new FileInputStream(new File(input));
		        HWPFDocument document = new HWPFDocument(file);) 
	    {
	        Range range = document.getRange();
	        int para_count = 1;
	        int table_count = 1;
	
	        for (int i = 0; i < range.numParagraphs(); i++) 
	        {
	            Paragraph para = range.getParagraph(i);
	            if (para.isInTable()) 
	            {
	                Table table = range.getTable(para);
	                StringBuilder xmlStringBuilder = new StringBuilder();
	                xmlStringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?> <table border=1><tbody>");
	                StringBuilder plainTextBuilder = new StringBuilder();
	
	                for (int j = 0; j < table.numRows(); j++) 
	                {
	                    TableRow row = table.getRow(j);
	                    xmlStringBuilder.append("<tr>");
	                    for (int k = 0; k < row.numCells(); k++) 
	                    {
	                        TableCell cell = row.getCell(k);
	                        String cellText = cell.text().replaceAll("[\\n\\r]+", "").trim();
	                        cellText=string_refine(cellText);
	                        xmlStringBuilder.append("<td>")
	                                        .append(cellText)
	                                        .append("</td>");
	                        plainTextBuilder.append(cellText).append(" \u2028");// Add line separator
	                    }
	                    xmlStringBuilder.append("</tr>");
	                    plainTextBuilder.append(" \u2029"); // Add paragraph separator
	                }
	                xmlStringBuilder.append("</tbody></table>");
	
	                biocPassage = new BioCPassage();
	                Infons = new HashMap<>();
	                Infons.put("xml", xmlStringBuilder.toString());
	                Infons.put("id", "Table_" + table_count);
	                Infons.put("section_type", "TABLE");
	                Infons.put("type", "table");
	                biocPassage.setInfons(Infons);
	                biocPassage.setOffset(startoffset);
	                biocPassage.setText(plainTextBuilder.toString());
	                startoffset = startoffset + plainTextBuilder.length() + 1;
	                biocDocument.addPassage(biocPassage);
	                table_count++;
	                i += table.numParagraphs() - 1; // Skip paragraphs that are part of this table
	            }
	            else
	            {
	                String line = para.text().replaceAll("[\\n\\r]+", "").trim();
	                if (!line.equals("")) 
	                {
	                    biocPassage = new BioCPassage();
	                    Infons = new HashMap<>();
	                    Infons.put("id", "Text_" + para_count);
	                    Infons.put("section_type", "TEXT");
	                    Infons.put("type", "Text");
	                    biocPassage.setInfons(Infons);
	                    biocPassage.setText(string_refine(line));
	                    biocPassage.setOffset(startoffset);
	                    startoffset = startoffset + line.length() + 1;
	                    biocDocument.addPassage(biocPassage);
	                    para_count++;
	                }
	            }
	        }
	
	    }
	    biocCollection.addDocument(biocDocument);
        BioCOutputFormat.writeDocument(biocDocument);
        BioCOutputFormat.close();
	}
	public static void Word2PubTator(String input,String output) throws IOException, XMLStreamException
	{
		BufferedWriter outputfile  = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		try 
		{
			FileInputStream file = new FileInputStream(new File(input));
			HWPFDocument doc = new HWPFDocument(file);
			WordExtractor we = new WordExtractor(doc);
			String[] paragraphs = we.getParagraphText();
			
			int count_para=1;
			for (String para : paragraphs) 
			{
				para=para.replaceAll("", "");
				para=para.replaceAll("[\\n\\r]+", "");
				if(!para.equals(""))
				{
					outputfile.write("1000001|"+count_para+"|"+para+"\n");
					count_para++;
				}
			}
			outputfile.write("\n");
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		outputfile.close();
	}
	public static void Wordx2BioC(String input,String output) throws IOException, XMLStreamException
	{
		String parser = BioCFactory.WOODSTOX;
		BioCFactory factory = BioCFactory.newFactory(parser);
		BioCDocumentWriter BioCOutputFormat = factory.createBioCDocumentWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		BioCCollection biocCollection = new BioCCollection();
		ZoneId zonedId = ZoneId.of( "America/Montreal" );
		LocalDate today = LocalDate.now( zonedId );
		biocCollection.setDate(today.toString());
		biocCollection.setKey("BioC.key");//key
		biocCollection.setSource("BioC");//source
		
		BioCOutputFormat.writeCollectionInfo(biocCollection);
		BioCDocument biocDocument = new BioCDocument();
		biocDocument.setID(input.substring(input.lastIndexOf('/') + 1));
	    
		int startoffset=0;
		BioCPassage biocPassage = new BioCPassage();
		Map<String, String> Infons = new HashMap<String, String>();
		Infons.put("id", "Front_1");
        Infons.put("section_type", "TITLE");
        Infons.put("type", "front");
        //Infons.put("license", "CC BY");
        biocPassage.setInfons(Infons);
        biocPassage.setText("Supplementary Material");
		biocPassage.setOffset(startoffset);
		startoffset=startoffset+"Supplementary Material".length()+1;
		biocDocument.addPassage(biocPassage);
		
		try (FileInputStream file = new FileInputStream(new File(input));
			     XWPFDocument document = new XWPFDocument(file)) 
		{

			    // Iterator to maintain the natural order of elements in the document
		    Iterator<IBodyElement> bodyElementsIterator = document.getBodyElementsIterator();

		    int count_para = 1;
		    while (bodyElementsIterator.hasNext()) {
		        IBodyElement element = bodyElementsIterator.next();

		        if (element.getElementType() == BodyElementType.PARAGRAPH) 
		        {
		            XWPFParagraph para = (XWPFParagraph) element;
		            String line = para.getText().replaceAll("[\\n\\r]+", "");
		            line = line.replaceAll("\u0007", "");

		            if (!line.isEmpty()) 
		            {
		                BioCPassage biocPassage1 = new BioCPassage();
		                Map<String, String> Infons1 = new HashMap<>();
		                Infons1.put("id", "Text_" + count_para);
		                Infons1.put("section_type", "TEXT");
		                Infons1.put("type", "text");
		                biocPassage1.setInfons(Infons1);
		                biocPassage1.setText(string_refine(line));
		                biocPassage1.setOffset(startoffset);
		                startoffset += line.length() + 1;
		                biocDocument.addPassage(biocPassage1);
		                count_para++;
		            }
		        } 
		        else if (element.getElementType() == BodyElementType.TABLE) 
		        {
		            XWPFTable table = (XWPFTable) element;
		            StringBuilder xmlStringBuilder = new StringBuilder();
		            xmlStringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?> <table border=1><tbody>");
		            StringBuilder plainTextBuilder = new StringBuilder();

		            for (int i = 0; i < table.getNumberOfRows(); i++) 
		            {
		                List<XWPFTableCell> cells = table.getRow(i).getTableCells();
		                xmlStringBuilder.append("<tr>");
		                for (XWPFTableCell cell : cells) 
		                {
		                    xmlStringBuilder.append("<td");
		                    if (cell.getCTTc().getTcPr().isSetGridSpan()) 
		                    {
		                        int colspan = cell.getCTTc().getTcPr().getGridSpan().getVal().intValue();
		                        xmlStringBuilder.append(" colspan=\"" + colspan + "\"");
		                    }
		                    String cellText = string_refine(cell.getText());
		                    xmlStringBuilder.append(">");
		                    xmlStringBuilder.append(cellText);
		                    xmlStringBuilder.append("</td>");
		                    plainTextBuilder.append(cellText).append(" \u2028"); // Add line separator
		                }
		                xmlStringBuilder.append("</tr>");
		                plainTextBuilder.append(" \u2029"); // Add paragraph separator
		            }
		            xmlStringBuilder.append("</tbody></table>");

		            BioCPassage biocPassage1 = new BioCPassage();
		            Map<String, String> Infons1 = new HashMap<>();
		            Infons1.put("xml", xmlStringBuilder.toString());
		            Infons1.put("id", "Table_" + count_para);
		            Infons1.put("section_type", "TABLE");
		            Infons1.put("type", "table");
		            biocPassage1.setInfons(Infons1);
		            biocPassage1.setOffset(startoffset);
		            biocPassage1.setText(plainTextBuilder.toString());
		            startoffset += plainTextBuilder.length() + 1;
		            biocDocument.addPassage(biocPassage1);
		            count_para++;
		        }
		    }
		}

        biocCollection.addDocument(biocDocument);
        BioCOutputFormat.writeDocument(biocDocument);
        BioCOutputFormat.close();
	}
	public static void Wordx2PubTator(String input,String output) throws IOException, XMLStreamException
	{
		BufferedWriter outputfile  = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output), "UTF-8"));
		try 
		{
			FileInputStream file = new FileInputStream(new File(input));
			XWPFDocument document = new XWPFDocument(file);
			
			List<XWPFParagraph> paragraphs = document.getParagraphs();
			List<XWPFTable> tables = document.getTables();
			int count_para=1;
			for (XWPFParagraph para : paragraphs) 
			{
				String pa=para.getText().replaceAll("[\\n\\r]+", "");
				if(!pa.equals(""))
				{
					outputfile.write("1000001|"+count_para+"|"+pa+"\n");
					count_para++;
				}
			}
			count_para=1;
			for (XWPFTable ta : tables) 
			{
				for(int i=0;i<ta.getNumberOfRows();i++)
				{
					if (count_para == 1){outputfile.write("\n");}
					List<XWPFTableCell> cells= ta.getRow(i).getTableCells();
					String cells_str="";
					for(XWPFTableCell c : cells)
					{
						cells_str=cells_str+c.getText()+"; ";
					}
					outputfile.write("1000002|"+count_para+"|"+cells_str+"\n");
					count_para++;
				}
			}
			outputfile.write("\n");
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		outputfile.close();
	}
	
	/*
	 * Support functions
	 */
	public static void Decompression(String input,String output) throws IOException, XMLStreamException
	{
		try (TarArchiveInputStream fin = new TarArchiveInputStream(new GzipCompressorInputStream(new FileInputStream(input))))
		{
            TarArchiveEntry entry;
            while ((entry = fin.getNextTarEntry()) != null) {
                if (entry.isDirectory()) {
                    continue;
                }
                File curfile = new File(output, entry.getName());
                File parent = curfile.getParentFile();
                if (!parent.exists()) {
                    parent.mkdirs();
                }
                IOUtils.copy(fin, new FileOutputStream(curfile));
            }
        }
	}
	private static TarArchiveOutputStream getTarArchiveOutputStream(String name) throws IOException 
	{
		TarArchiveOutputStream taos = new TarArchiveOutputStream(new FileOutputStream(name));
		// TAR has an 8 gig file limit by default, this gets around that
		taos.setBigNumberMode(TarArchiveOutputStream.BIGNUMBER_STAR);
		// TAR originally didn't support long file names, so enable the support for it
		taos.setLongFileMode(TarArchiveOutputStream.LONGFILE_GNU);
		taos.setAddPaxHeadersForNonAsciiNames(true);
		return taos;
	}
	private static void addToArchiveCompression(TarArchiveOutputStream out, File file, String dir) throws IOException 
	{
        String entry = dir + File.separator + file.getName();
        if (file.isFile()){
            out.putArchiveEntry(new TarArchiveEntry(file, entry));
            try (FileInputStream in = new FileInputStream(file)){
                IOUtils.copy(in, out);
            }
            out.closeArchiveEntry();
        } else if (file.isDirectory()) {
            File[] children = file.listFiles();
            if (children != null){
                for (File child : children){
                    addToArchiveCompression(out, child, entry);
                }
            }
        } else {
            System.out.println(file.getName() + " is not supported");
        }
    }
	
	/*
	 * Main ()
	 */
	public static void main(String [] args) throws IOException, InterruptedException, XMLStreamException, ParserConfigurationException, SAXException, TesseractException, CsvValidationException 
	{
		String input="input";
		String output="output";
		if(args.length<2)
		{
			System.out.println("\n$ java -jar FormatConverter.jar [inputfile] [outputfile] [output format:BioC|PubTator] [input format:BioC|PubTator] [fold]\n");
			System.out.println("* [inputfile] and [outputfile] can be file or folder. [input] and [output] folders are the defaults.");
			System.out.println("* BioC-XML|PubTator|FreeText|PDF|MSWord|MSExcel formats are allowed in [input format]. The format is auto detected if not assigned.");
			System.out.println("* BioC-XML is the default [output format], if not assigned.\n");
			System.out.println("* BioC-XML is the default format of the output, if not assigned.\n");
		}
		else
		{
			input = args[0];
			output= args[1];
		}
		
		String format="";
		String FormatCheck="";
		if(args.length<3)
		{
			format="BioC";
		}
		else
		{
			format= args[2];
		}
		if(args.length>=4)
		{
			FormatCheck= args[3];
		}
		
		int Num=10000;
		if(args.length>=5)
		{
			Num= Integer.parseInt(args[4]);
		}
		if((!input.matches("/net/intdev/pubtator/Regular_Update_PMC_suppl\\.BioC/output\\.original/.*")) && (!output.equals("/net/intdev/pubtator/Regular_Update_PMC_suppl\\.BioC/output\\.bioc/.*"))) //the input/output means it's \\intdev\pubtator\Regular_Update_PMC_suppl.BioC
		{
			System.out.println("--------------Parameters--------------------");
			System.out.println("[inputfile]: "+input);
			System.out.println("[outputfile]: "+output);
			System.out.println("[output format]: "+format);
			if(FormatCheck.equals(""))
			{
				System.out.println("[input format]: Auto detected.");
			}
			else
			{
				System.out.println("[input format]: "+FormatCheck);
			}
			if(Num==10000)
			{
				System.out.println("[fold]: N/A");
			}
			else
			{
				System.out.println("[fold]: "+Num);
			}
			System.out.println("--------------Start-------------------------\n");
		}
		
		File file = new File(input);

		boolean isDirectory = file.isDirectory(); // Check if it's a directory
		boolean isFile =      file.isFile();      // Check if it's a regular file
		
		ArrayList<String> inputfiles = new ArrayList<String>();
		ArrayList<String> outputfiles = new ArrayList<String>();
		if(isFile)
		{
			inputfiles.add(input);
			outputfiles.add(output);
		}
		else if(isDirectory)
		{
			File[] listOfFiles = file.listFiles();
			for (int i = 0; i < listOfFiles.length; i++)
			{
				if (listOfFiles[i].isFile()) 
				{
					String filename = listOfFiles[i].getName();
					
					Pattern patt = Pattern.compile("^([0-9]+).txt");
					Matcher mat = patt.matcher(filename);
					if(mat.find()) 
		        	{
						if(Integer.parseInt(mat.group(1))%100==Num || Num==10000)
						{
							File f = new File(input+"/"+filename);
							if(f.exists() && !f.isDirectory()) 
							{ 
								inputfiles.add(input+"/"+filename);
								outputfiles.add(output+"/"+filename);
							}
						}
		        	}
					else
					{
						File f = new File(input+"/"+filename);
						if(f.exists() && !f.isDirectory()) 
						{ 
							inputfiles.add(input+"/"+filename);
							outputfiles.add(output+"/"+filename);
						}
					}
				}
			}
		}
		else
		{
			System.out.println("[Error]: Input file is not exist.");
		}
		
		for(int file_i=0;file_i<inputfiles.size();file_i++)
		{	
			String inputfile=inputfiles.get(file_i);
			String outputfile=outputfiles.get(file_i); //update-20240729
			 
			File outputf = new File(outputfile);
			if(outputf.exists() && (!outputf.isDirectory()))
			{
				System.out.println(outputfile+" - Done. (The output file exists in output folder)");
			}
			else  //update-20240729
			{
				if(FormatCheck.equals(""))
				{
					FormatCheck = BioCFormatCheck(inputfile);
				}
				
				System.out.println("Input Format: " + FormatCheck);
				
				if(FormatCheck.equals("PDF"))
				{
					if(format.equals("BioC"))
					{
						System.out.println("Format convert from PDF to BioC(XML): "+inputfile+" -> "+outputfile);
						PDF2BioC(inputfile,outputfile);
					}
					else if(format.equals("PubTator"))
					{
						System.out.println("Format convert from PDF to PubTator: "+inputfile+" -> "+outputfile);
						PDF2PubTator(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.0]: Current output format options for PDF are : PubTator|BioC(xml)");
					}
				}
				else if(FormatCheck.equals("BioC"))
				{
					if(format.equals("PubTator"))
					{
						System.out.println("Format convert from BioC(XML) to PubTator: "+inputfile+" -> "+outputfile);
						BioC2PubTator(inputfile,outputfile);
					}
					else if(format.equals("HTML")) //with annotation only
					{
						System.out.println("Format convert from BioC(XML) to HTML: "+inputfile+" -> "+outputfile);
						BioC2HTML(inputfile,outputfile);
					}
					else if(format.equals("SciLite")) //with annotation only
					{
						System.out.println("Format convert from BioC(XML) to SciLite: "+inputfile+" -> "+outputfile);
						BioC2SciLite(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.1]: Current output format options for BioC(XML) are : PubTator|HTML");
					}
				}
				else if(FormatCheck.equals("PubTator"))
				{
					if(format.equals("BioC"))
					{
						System.out.println("Format convert from PubTator to BioC(XML): "+inputfile+" -> "+outputfile);
						PubTator2BioC(inputfile,outputfile+".XML");
					}
					else if(format.equals("HTML")) //with annotation only
					{
						System.out.println("Format convert from PubTator to HTML: "+inputfile+" -> "+outputfile);
						PubTator2HTML(inputfile,outputfile);
					}
					else
					{
						//System.out.println("Current format options are : PubTator|BioC|HTML");
						System.out.println("\n[Error 2.2]: Current output format options for PubTator are : BioC(XML)|HTML");
					}
				}
				else if(FormatCheck.equals("PPT"))
				{
					if(format.equals("PubTator"))
					{
						System.out.println("Format convert from Word to PubTator: "+inputfile+" -> "+outputfile);
						ppt2PubTator(inputfile,outputfile);
					}
					else if(format.equals("BioC"))
					{
						System.out.println("Format convert from Excel to BioC: "+inputfile+" -> "+outputfile);
						ppt2BioC(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : PubTator|BioC");
					}
				}
				else if(FormatCheck.equals("PPTx"))
				{
					if(format.equals("PubTator"))
					{
						System.out.println("Format convert from Word to PubTator: "+inputfile+" -> "+outputfile);
						ppt2PubTator(inputfile,outputfile);
					}
					else if(format.equals("BioC"))
					{
						System.out.println("Format convert from Excel to BioC: "+inputfile+" -> "+outputfile);
						pptx2BioC(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : PubTator|BioC");
					}
				}
				else if(FormatCheck.equals("Word"))
				{
					if(format.equals("PubTator"))
					{
						System.out.println("Format convert from Word to PubTator: "+inputfile+" -> "+outputfile);
						Word2PubTator(inputfile,outputfile);
					}
					else if(format.equals("BioC"))
					{
						System.out.println("Format convert from Excel to BioC: "+inputfile+" -> "+outputfile);
						Word2BioC(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : PubTator|BioC");
					}
				}
				else if(FormatCheck.equals("RTF"))
				{
					if(format.equals("BioC"))
					{
						System.out.println("Format convert from RTF to BioC: "+inputfile+" -> "+outputfile);
						RTF2BioC(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : BioC");
					}
				}
				else if(FormatCheck.equals("Wordx"))
				{
					if(format.equals("PubTator"))
					{
						System.out.println("Format convert from Word (docx) to PubTator: "+inputfile+" -> "+outputfile);
						Wordx2PubTator(inputfile,outputfile);
					}
					else if(format.equals("BioC"))
					{
						System.out.println("Format convert from Excel to BioC: "+inputfile+" -> "+outputfile);
						Wordx2BioC(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : PubTator|BioC");
					}
				}
				else if(FormatCheck.equals("Excel"))
				{
					if(format.equals("PubTator"))
					{
						System.out.println("Format convert from Excelx to PubTator: "+inputfile+" -> "+outputfile);
						Excel2PubTator(inputfile,outputfile);
					}
					else if(format.equals("BioC"))
					{
						System.out.println("Format convert from Excelx to BioC: "+inputfile+" -> "+outputfile);
						Excel2BioC(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : PubTator|BioC");
					}
				}
				else if(FormatCheck.equals("Excelx"))
				{
					if(format.equals("PubTator"))
					{
						System.out.println("Format convert from Excel to PubTator: "+inputfile+" -> "+outputfile);
						Excelx2PubTator(inputfile,outputfile);
					}
					else if(format.equals("BioC"))
					{
						System.out.println("Format convert from Excel to BioC: "+inputfile+" -> "+outputfile);
						Excelx2BioC(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : PubTator|BioC");
					}
				}
				else if(FormatCheck.equals("TSV"))
				{
					if(format.equals("BioC"))
					{
						System.out.println("Format convert from TSV to BioC: "+inputfile+" -> "+outputfile);
						TSV2BioC(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : BioC");
					}
				}
				else if(FormatCheck.equals("CSV"))
				{
					if(format.equals("BioC"))
					{
						System.out.println("Format convert from CSV to BioC: "+inputfile+" -> "+outputfile);
						CSV2BioC(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : BioC");
					}
				}
				else if(FormatCheck.equals("TXT"))
				{
					if(format.equals("PubTator"))
					{
						System.out.println("Format convert from TXT to PubTator: "+inputfile+" -> "+outputfile);
						FreeText2PubTator(inputfile,outputfile);
					}
					else if(format.equals("BioC"))
					{
						System.out.println("Format convert from TXT to BioC: "+inputfile+" -> "+outputfile);
						FreeText2BioC(inputfile,outputfile);
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : BioC");
					}
				}
				else if(FormatCheck.equals("XML"))
				{
					if(format.equals("BioC"))
					{
						System.out.println("Format convert from XML to BioC: "+inputfile+" -> "+outputfile);
						try {
							XML_names = new ArrayList<String>();
							XML_contents = new ArrayList<String>();
							XML2BioC(inputfile,outputfile);
						} catch (Exception e) {
							e.printStackTrace();
						}
					}
					else
					{
						System.out.println("\n[Error 2.4]: Current output format options for free text are : PubTator");
					}
				}
				else if(FormatCheck.equals("tar.gz"))
				{
					System.out.println("Decompression : "+inputfile+" -> "+outputfile);
					Decompression(inputfile,outputfile);
				}
				else
				{
					System.out.println("\n[Error 2.3]: the file '"+inputfile+"' (Format Code:"+FormatCheck+") is skipped.");
				}
			}
			FormatCheck="";
		}
	}	
}