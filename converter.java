import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;
import java.io.FileNotFoundException;


/**
 * Convert New York City Administrative Code XML files to Excel Spreadsheet
 * @imports Apache POI, Apache Commons, W3C libraries
 * @author Rebecca Ramnauth
 * Date: 4-22-2017
 */
public class converter {
    private static Workbook workbook;
    private static int rowNum;
	
	private final static int FILE_NAME = 0;
	
	//LEVEL 1 : TITLE	
	private final static int RECORD_ID = 1; 		//TITLE attribute id
	private final static int RECORD_NUMBER = 2;		//TITLE attribute number
	private final static int RECORD_VERSION = 3;	//TITLE attribute version
	private final static int TITLE_NUM = 4; 		//HEADING, Regex excluding alpha characters
	private final static int TITLE_SUBJECT = 5;		//HEADING, substring of everything beyond the ': '
	
	//LEVEL 2 : CHAPTER
	private final static int C_RECORD_ID = 6;		//RECORD attribute id
	private final static int C_RECORD_NUMBER = 7;	//RECORD attribute number
	private final static int C_RECORD_VERSION = 8;	//RECORD attribute version
	private final static int CHAPTER_NUM = 9;		//HEADING, Regex excluding alpha characters
	private final static int CHAPTER_SUBJECT = 10;	//HEADING, substring of everything beyond the ': '
	
	//LEVEL 3 : SUBCHAPTER
	private final static int SUBCHAPTER_NUM = 11;
	
	//LEVEL 4 : APPENDIX
	private final static int HAS_APPENDIX = 12;
	
	//LEVEL 5 : PART
	private final static int HAS_PART = 13;
	
	//LEVEL 6 : ARTICLE
	private final static int HAS_ARTICLE = 14;
	
	//LEVEL 7 : SUBARTICLE
	private final static int HAS_SUBARTICLE = 15;
	
	//LEVEL 8 : SECTION
	private final static int S_RECORD_ID = 16;		//RECORD attribute id
	private final static int S_RECORD_NUMBER = 17;	//RECORD attribute number
	private final static int S_RECORD_VERSION = 18;	//RECORD attribute version
	private final static int SECTION_NUM = 19;		//HEADING, Regex excluding alpha characters
	private final static int SECTION_SUBJECT = 20;	//HEADING, substring of everything beyond the ': '
	
	//LEVEL 9 : PARAGRAPH
	private final static int PARAGRAPH = 21;		//PARA
	
	//NEXT STEPS
	private final static int CITATION = 22;

    public static void main(String[] args) throws Exception {
        //retrieveAndReadXml(); //test: read the first file and extract data.
        read();
        copy();
        //format();
    }
    
    /**
     * Defines possibilities of the file name given the first
     * Calls retrieveAndReadXml() to actually undergo the conversion process
     * @param
     
	private static void possibility() throws Exception{
		File home = new File ("C:\\Users\\Shahendra\\Documents\\JCreator Pro\\MyProjects\\excel-xml-nycac\\admin_Test");
		File[] files = home.listFiles();
		int total = 0;
		for (int i = 0; i < files.length; i++){
			read(files[i]);
			total++;
		}
		System.out.println("TOTAL " + total);
	}
    */
	
    private static void read() throws Exception{
    	Cell cell;
    	    	
    	initXls();
    	Sheet sheet = workbook.getSheetAt(0);
    	
    	File origin = new File ("C:\\Users\\Shahendra\\Documents\\JCreator Pro\\MyProjects\\excel-xml-nycac\\admin_Test");
    	File[] files = origin.listFiles();
    	int count = 0;
    	
    	DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
    	DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
    	
    	for (File file : files){	
    		System.out.println("Reading " + file.getAbsolutePath());
    		Document doc = dBuilder.parse(file);
    		/*
	    	NodeList title = doc.getElementsByTagName("LEVEL");
	    	for (int a = 0; a < title.getLength(); a++){
	    		Node home = title.item(a);
	    		if (home.getNodeType() == Node.ELEMENT_NODE){
	    			Element name = (Element) home;
	    			String chapter;
	    			try{
	    				chapter = name.getElementsByTagName("HEADING").item(0).getTextContent(); //pulls all headings (i.e., title, chapter, sub., section) causing duplicates in paragraphing
	    			}
	    			catch(NullPointerException e){
	    				chapter = "NULL";
	    			}
	    	*/		
	    			NodeList para = doc.getElementsByTagName("PARA");			
	    			for (int b = 0; b < para.getLength(); b++){
	    				Row row = sheet.createRow(rowNum++);;
	    				Node parag = para.item(b);
	    				if (parag.getNodeType() == Node.ELEMENT_NODE){
	    					Element paragrapf = (Element) parag;
	    					String requirement = paragrapf.getTextContent();
	    					//TITLE
	    					cell = row.createCell(TITLE_NUM);
	    					if ((coalesce(requirement, "Title")).equals("Title"))
								cell.setCellValue(requirement);
	    					else
	    						cell.setCellValue("");
	    					//CHAPTER
	    					cell = row.createCell(CHAPTER_NUM);
	    					if ((coalesce(requirement, "Chapter")).equals("Chapter"))
	    						cell.setCellValue(requirement);
	    					else
	    						cell.setCellValue("");
	    					//SUBCHAPTER
	    					cell = row.createCell(SUBCHAPTER_NUM);
	    					if ((coalesce(requirement, "Subchapter")).equals("Subchapter"))
	    						cell.setCellValue(requirement);
	    					else
	    						cell.setCellValue("");
	    					//SECTION
	    					cell = row.createCell(SECTION_NUM);
	    					if ((coalesce(requirement, "§")).equals("§"))
	    						cell.setCellValue(requirement);
	    					else
	    						cell.setCellValue("");
	    						
	    					//Row row = sheet.createRow(rowNum++);;
	    					
	    					//cell = row.createCell(CHAPTER_NUM);
			    			//cell.setCellValue(chapter);
			    						    			
			    			cell = row.createCell(FILE_NAME);
			    			cell.setCellValue(file.getAbsolutePath());
			    			
			    			cell = row.createCell(PARAGRAPH);
			    			cell.setCellValue(requirement);
	    				}
	    			}
	    		//}
	    	//}
	    	count++;
    	}
       	System.out.println("TOTAL: " + count);
       	
        FileOutputStream fileOut = new FileOutputStream("example.xlsx");
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();
    }
        
    public static String coalesce(String rem, String value) {
	    /*
	    for(int i = 0; i < obj.length(); i++)
	    	if(obj.substring(i, i+1) != null && !(obj.substring(i, i+1)).equals(" ")){
	    		//return obj.substring(i, i+1);
	    		boolean equal = (obj.substring(i, i+1)).equals("§");
	    		if (equal)
	    			System.out.println("FOUND " + equal + " is: " + obj.substring(i, i+1));
	    	}
	    return false;
	   */
	   String obj = rem.trim();
	   int len = value.length();
	   //System.out.println(obj);
	   if (obj.length() == 0 || obj.length() < len){
	   		return " ";
	   }
	   //System.out.println("	SUBSTRING: " + obj.substring(i, i+1));
	   return obj.substring(0, len);
	}
	
	public static void copy() throws Exception{
		Sheet sheet = workbook.getSheetAt(0);
		//TITLE
		for (Cell cell : sheet.getRow(TITLE_NUM)){
			System.out.println(cell.getStringCellValue()); 		//INCOMPLETED!!!!
		}
	}

    /**
     * Initializes the POI workbook and writes the header row
     */
    private static void initXls() {
        workbook = new XSSFWorkbook();

        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        style.setAlignment(CellStyle.ALIGN_CENTER);

        Sheet sheet = workbook.createSheet();
        rowNum = 0;
        Row row = sheet.createRow(rowNum++);
        
        //private final static int FILE_NAME = 0;
        Cell cell = row.createCell(FILE_NAME);
        cell.setCellValue("File");
        cell.setCellStyle(style);
	
		//LEVEL 1 : TITLE	
		//private final static int RECORD_ID = 0; 		//TITLE attribute id
		cell = row.createCell(RECORD_ID);
        cell.setCellValue("Title ID");
        cell.setCellStyle(style);
		//private final static int RECORD_NUMBER = 0;	//TITLE attribute number
		cell = row.createCell(RECORD_NUMBER);
        cell.setCellValue("Title Record");
        cell.setCellStyle(style);
		//private final static int RECORD_VERSION = 0;	//TITLE attribute version
		cell = row.createCell(RECORD_VERSION);
        cell.setCellValue("Title Version");
        cell.setCellStyle(style);
		//private final static int TITLE_NUM = 1; 		//HEADING, Regex excluding alpha characters
		cell = row.createCell(TITLE_NUM);
        cell.setCellValue("Title #");
        cell.setCellStyle(style);
		//private final static int TITLE_SUBJECT = 1;		//HEADING, substring of everything beyond the ': '
		cell = row.createCell(TITLE_SUBJECT);
        cell.setCellValue("Title Subject");
        cell.setCellStyle(style);
		
		//LEVEL 2 : CHAPTER
		//private final static int C_RECORD_ID = 0;		//RECORD attribute id
		cell = row.createCell(C_RECORD_ID);
        cell.setCellValue("Chapter ID");
        cell.setCellStyle(style);
		//private final static int C_RECORD_NUMBER = 0;	//RECORD attribute number
		cell = row.createCell(C_RECORD_NUMBER);
        cell.setCellValue("Chapter Record");
        cell.setCellStyle(style);
		//private final static int C_RECORD_VERSION = 0;	//RECORD attribute version
		cell = row.createCell(C_RECORD_VERSION);
        cell.setCellValue("Chapter Version");
        cell.setCellStyle(style);
		//private final static int CHAPTER_NUM = 0;		//HEADING, Regex excluding alpha characters
		cell = row.createCell(CHAPTER_NUM);
        cell.setCellValue("Chapter #");
        cell.setCellStyle(style);
		//private final static int CHAPTER_SUBJECT = 0;	//HEADING, substring of everything beyond the ': '
		cell = row.createCell(CHAPTER_SUBJECT);
        cell.setCellValue("Chapter Subject");
        cell.setCellStyle(style);
		
		//LEVEL 3 : SUBCHAPTER
		//private final static int HAS_SUBCHAPTER = 0;
		cell = row.createCell(SUBCHAPTER_NUM);
        cell.setCellValue("Subchapter #");
        cell.setCellStyle(style);
		
		//LEVEL 4 : APPENDIX
		//private final static int HAS_APPENDIX = 0;
		cell = row.createCell(HAS_APPENDIX);
        cell.setCellValue("Appendix Exists?");
        cell.setCellStyle(style);
		
		//LEVEL 5 : PART
		//private final static int HAS_PART = 0;
		cell = row.createCell(HAS_PART);
        cell.setCellValue("Part Exists?");
        cell.setCellStyle(style);
		
		//LEVEL 6 : ARTICLE
		//private final static int HAS_ARTICLE = 0;
		cell = row.createCell(HAS_ARTICLE);
        cell.setCellValue("Article Exists?");
        cell.setCellStyle(style);
		
		//LEVEL 7 : SUBARTICLE
		//private final static int HAS_SUBARTICLE = 0;
		cell = row.createCell(HAS_SUBARTICLE);
        cell.setCellValue("Subarticle Exists?");
        cell.setCellStyle(style);
		
		//LEVEL 8 : SECTION
		//private final static int S_RECORD_ID = 0;		//RECORD attribute id
		cell = row.createCell(S_RECORD_ID);
        cell.setCellValue("Section ID");
        cell.setCellStyle(style);
		//private final static int S_RECORD_NUMBER = 0;	//RECORD attribute number
		cell = row.createCell(S_RECORD_NUMBER);
        cell.setCellValue("Section Record");
        cell.setCellStyle(style);
		//private final static int S_RECORD_VERSION = 0; 	//RECORD attribute version
		cell = row.createCell(S_RECORD_VERSION);
        cell.setCellValue("Section Version");
        cell.setCellStyle(style);
		//private final static int SECTION_NUM = 0;		//HEADING, Regex excluding alpha characters
		cell = row.createCell(SECTION_NUM);
        cell.setCellValue("Section #");
        cell.setCellStyle(style);
		//private final static int SECTION_SUBJECT = 0;	//HEADING, substring of everything beyond the ': '
		cell = row.createCell(SECTION_SUBJECT);
        cell.setCellValue("Section Subject");
        cell.setCellStyle(style);
		
		//LEVEL 9 : PARAGRAPH
		//private final static int PARAGRAPH = 0;			//PARA
		cell = row.createCell(PARAGRAPH);
        cell.setCellValue("Subrequirement");
        cell.setCellStyle(style);
		
		//NEXT STEPS
		//private final static int CITATION = 0;
	    cell = row.createCell(CITATION);
        cell.setCellValue("Citation");
        cell.setCellStyle(style);
    }
}
