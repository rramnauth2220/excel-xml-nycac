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

import java.util.*;

//import java.util.Scanner;
//import java.util.Arrays;
//import java.util.Comparator;

/**
 * Convert New York City Administrative Code XML files to Excel Spreadsheet
 * @imports Apache POI, Apache Commons, W3C libraries
 * @author Rebecca Ramnauth
 * Date: 4-22-2017
 */
public class converter {
    private static Workbook workbook;
    private static Workbook workbook2;
    
    private static int rowNum;
	
	private final static int FILE_NAME = 0;
	
	//LEVEL 1 : TITLE	
	//private final static int RECORD_ID = 1; 		//TITLE attribute id
	//private final static int RECORD_NUMBER = 2;		//TITLE attribute number
	//private final static int RECORD_VERSION = 3;	//TITLE attribute version
	//private final static int TITLE = 4; 		//HEADING, Regex excluding alpha characters
	//private final static int TITLE_SUBJECT = 5;		//HEADING, substring of everything beyond the ': '
	private final static int TITLE = 1;
	
	//LEVEL 2 : CHAPTER
	//private final static int C_RECORD_ID = 6;		//RECORD attribute id
	//private final static int C_RECORD_NUMBER = 7;	//RECORD attribute number
	//private final static int C_RECORD_VERSION = 8;	//RECORD attribute version
	//private final static int CHAPTER = 9;		//HEADING, Regex excluding alpha characters
	//private final static int CHAPTER_SUBJECT = 10;	//HEADING, substring of everything beyond the ': '
	private final static int CHAPTER = 2;
	
	//LEVEL 3 : SUBCHAPTER
	private final static int SUBCHAPTER = 3;
	
	private final static int ARTICLE = 4;
	
	//LEVEL 4 : APPENDIX
	//private final static int HAS_APPENDIX = 12;
	
	//LEVEL 5 : PART
	//private final static int HAS_PART = 13;
	
	//LEVEL 6 : ARTICLE
	//private final static int HAS_ARTICLE = 14;
	
	//LEVEL 7 : SUBARTICLE
	//private final static int HAS_SUBARTICLE = 15;
	
	//LEVEL 8 : SECTION
	//private final static int S_RECORD_ID = 16;		//RECORD attribute id
	//private final static int S_RECORD_NUMBER = 17;	//RECORD attribute number
	//private final static int S_RECORD_VERSION = 18;	//RECORD attribute version
	//private final static int SECTION = 19;		//HEADING, Regex excluding alpha characters
	//private final static int SECTION_SUBJECT = 20;	//HEADING, substring of everything beyond the ': '
	private final static int SECTION = 5;
	
	//LEVEL 9 : PARAGRAPH
	private final static int PARAGRAPH = 6;		//PARA
	
	//NEXT STEPS
	private final static int CITATION = 7;

    public static void main(String[] args) throws Exception {
        //retrieveAndReadXml(); //test: read the first file and extract data.
        String xlsx = "example.xlsx";
        read(xlsx);
        //copy(xlsx);
        //format();
    }
    	
    private static void read(String xlsx) throws Exception{
    	Cell cell;
    	    	
    	initXls();
    	Sheet sheet = workbook.getSheetAt(0);
    	
    	File origin = new File ("C:\\Users\\Shahendra\\Documents\\JCreator Pro\\MyProjects\\excel-xml-nycac\\admin_Test\\NYCAC 1321 - 1814");
    	File[] files = origin.listFiles();
    	//Sort by number
    	Arrays.sort(files, new Comparator<File>(){
    		@Override
    		public int compare(File o1, File o2) {
		        int n1 = extractNumber(o1.getName());
		        int n2 = extractNumber(o2.getName());
		        return n1 - n2;
		    }
		            
		    private int extractNumber(String file_name) {
		        int indi = 0;
		        try {
		            int start = file_name.lastIndexOf("-") + 1;
		            int end = file_name.lastIndexOf('.');
		            String number = file_name.substring(start, end);
		            indi = Integer.parseInt(number);
		        } catch(Exception e) {
		            indi = 0; // if filename does not match the format default to 0
		        }
		        return indi;
		    } 
    	});
         
        for(File f : files) {
            System.out.println(f.getName());
        }

    	//End sort by number
    	int count = 0;
    	
    	DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
    	DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
    	
    	for (File file : files){	
    		System.out.println("Reading " + file.getAbsolutePath());
    		Document doc = dBuilder.parse(file);
	
			NodeList para = doc.getElementsByTagName("PARA");			
			for (int b = 0; b < para.getLength(); b++){
				Row row = sheet.createRow(rowNum++);;
				Node parag = para.item(b);
				if (parag.getNodeType() == Node.ELEMENT_NODE){
					Element paragrapf = (Element) parag;
					String requirement = paragrapf.getTextContent();
					//TITLE
					cell = row.createCell(TITLE);
					if ((coalesce(requirement, "Title")).equals("Title"))
						cell.setCellValue(requirement);
					else
						cell.setCellValue("");
					//CHAPTER
					cell = row.createCell(CHAPTER);
					if ((coalesce(requirement, "Chapter")).equals("Chapter"))
						cell.setCellValue(requirement);
					else
						cell.setCellValue("");
					//SUBCHAPTER
					cell = row.createCell(SUBCHAPTER);
					if ((coalesce(requirement, "Subchapter")).equals("Subchapter"))
						cell.setCellValue(requirement);
					else
						cell.setCellValue("");
					//ARTICLE
					cell = row.createCell(ARTICLE);
					if ((coalesce(requirement, "Article")).equals("Article"))
						cell.setCellValue(requirement);
					else
						cell.setCellValue("");
					//SECTION
					cell = row.createCell(SECTION);
					if ((coalesce(requirement, "§")).equals("§"))
						cell.setCellValue(requirement);
					else
						cell.setCellValue("");
						
					//Row row = sheet.createRow(rowNum++);;
					
					//cell = row.createCell(CHAPTER);
	    			//cell.setCellValue(chapter);
	    						    			
	    			cell = row.createCell(FILE_NAME);
	    			String file_name = (file.getName()).replaceAll("[^0-9]", "");
	    			cell.setCellValue(Integer.parseInt(file_name));
	    			
	    			cell = row.createCell(PARAGRAPH);
	    			try{
	    				cell.setCellValue(requirement);
	    			}
	    			catch (IllegalArgumentException e){
	    				cell.setCellValue("MAXIMUM LENGTH OF CELL EXCEEDED");
	    			}
				}
			}
	    	count++;
    	}
    	///////////////////////////////         ORDER       //////////////////////////////////////
		Scanner reader = new Scanner(System.in);  // Reading from System.in
		System.out.println("---------------------------------------------------");
		System.out.println("     FORMAT COLUMN A [FILE] IN ASCENDING ORDER     ");
		System.out.println("             Save the file and exit                ");
		System.out.println("---------------------------------------------------");
		//boolean confirmed = reader.nextBoolean();
		
    	///////////////////////////////         READ       ///////////////////////////////////////
  	
    	//TITLE
    	List<Integer> occurences_title = new ArrayList<Integer>();
		
		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++){
			if (!(((sheet.getRow(i)).getCell(TITLE)).getStringCellValue()).equals("")){
				occurences_title.add(i);
				System.out.println(" LINE: " + i );
			}
		}
		occurences_title.add(sheet.getPhysicalNumberOfRows());
		for (int j = 0; j < occurences_title.size() - 1; j++){
			System.out.println("FOUND TITLES AT: ");
			System.out.println("		" + occurences_title.get(j) + " TO " + occurences_title.get(j+1));
			String value = ((sheet.getRow(occurences_title.get(j))).getCell(TITLE)).getStringCellValue();
			System.out.println("		Copying: " + value);
			for (int index = occurences_title.get(j) + 1; index < occurences_title.get(j+1); index++){
				Cell cellCopy = (sheet.getRow(index)).getCell(TITLE);
				cellCopy.setCellValue(value);
				System.out.println("			UPDATE TITLE at " + index + " with VALUE of: "+ cellCopy.getStringCellValue());
			}
		}
		System.out.println("------------------  DELETE IN PROGRESS : TITLE  ------------------");
    	for (int k = occurences_title.size() - 1; k >= 0; k--){
    		removeRow(sheet, occurences_title.get(k));
    		System.out.println("Deleting row: " + occurences_title.get(k));
    	}
    	//CHAPTER
    	List<Integer> occurences_chapter = new ArrayList<Integer>();
		
		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++){
			if (!(((sheet.getRow(i)).getCell(CHAPTER)).getStringCellValue()).equals("")){
				occurences_chapter.add(i);
				System.out.println(" LINE: " + i );
			}
		}
		occurences_chapter.add(sheet.getPhysicalNumberOfRows());
		for (int j = 0; j < occurences_chapter.size() - 1; j++){
			System.out.println("FOUND TITLES AT: ");
			System.out.println("		" + occurences_chapter.get(j) + " TO " + occurences_chapter.get(j+1));
			String value = ((sheet.getRow(occurences_chapter.get(j))).getCell(CHAPTER)).getStringCellValue();
			System.out.println("		Copying: " + value);
			for (int index = occurences_chapter.get(j) + 1; index < occurences_chapter.get(j+1); index++){
				Cell cellCopy = (sheet.getRow(index)).getCell(CHAPTER);
				cellCopy.setCellValue(value);
				System.out.println("			UPDATE CHAPTER at " + index + " with VALUE of: "+ cellCopy.getStringCellValue());
			}
		}
		System.out.println("------------------  DELETE IN PROGRESS : CHAPTER  ------------------");
		for (int k = occurences_chapter.size() - 1; k >= 0; k--){
    		removeRow(sheet, occurences_chapter.get(k));
    		System.out.println("Deleting row: " + occurences_chapter.get(k));
    	}
		//SUBCHAPTER
    	List<Integer> occurences_subchapter = new ArrayList<Integer>();
		
		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++){
			if (!(((sheet.getRow(i)).getCell(SUBCHAPTER)).getStringCellValue()).equals("")){
				occurences_subchapter.add(i);
				System.out.println(" LINE: " + i );
			}
		}
		occurences_subchapter.add(sheet.getPhysicalNumberOfRows());
		for (int j = 0; j < occurences_subchapter.size() - 1; j++){
			System.out.println("FOUND TITLES AT: ");
			System.out.println("		" + occurences_subchapter.get(j) + " TO " + occurences_subchapter.get(j+1));
			String value = ((sheet.getRow(occurences_subchapter.get(j))).getCell(SUBCHAPTER)).getStringCellValue();
			System.out.println("		Copying: " + value);
			for (int index = occurences_subchapter.get(j) + 1; index < occurences_subchapter.get(j+1); index++){
				Cell cellCopy = (sheet.getRow(index)).getCell(SUBCHAPTER);
				cellCopy.setCellValue(value);
				System.out.println("			UPDATE SUBCHAPTER at " + index + " with VALUE of: "+ cellCopy.getStringCellValue());
			}
		}
		System.out.println("------------------  DELETE IN PROGRESS : SUBCHAPTER  ------------------");
		for (int k = occurences_subchapter.size() - 1; k >= 0; k--){
    		removeRow(sheet, occurences_subchapter.get(k));
    		System.out.println("Deleting row: " + occurences_subchapter.get(k));
    	}
    	//ARTICLE
    	List<Integer> occurences_article = new ArrayList<Integer>();
		
		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++){
			if (!(((sheet.getRow(i)).getCell(ARTICLE)).getStringCellValue()).equals("")){
				occurences_article.add(i);
				System.out.println(" LINE: " + i );
			}
		}
		occurences_article.add(sheet.getPhysicalNumberOfRows());
		for (int j = 0; j < occurences_article.size() - 1; j++){
			System.out.println("FOUND TITLES AT: ");
			System.out.println("		" + occurences_article.get(j) + " TO " + occurences_article.get(j+1));
			String value = ((sheet.getRow(occurences_article.get(j))).getCell(ARTICLE)).getStringCellValue();
			System.out.println("		Copying: " + value);
			for (int index = occurences_article.get(j) + 1; index < occurences_article.get(j+1); index++){
				Cell cellCopy = (sheet.getRow(index)).getCell(ARTICLE);
				cellCopy.setCellValue(value);
				System.out.println("			UPDATE SUBCHAPTER at " + index + " with VALUE of: "+ cellCopy.getStringCellValue());
			}
		}
		System.out.println("------------------  DELETE IN PROGRESS : ARTICLE  ------------------");
		for (int k = occurences_article.size() - 1; k >= 0; k--){
    		removeRow(sheet, occurences_article.get(k));
    		System.out.println("Deleting row: " + occurences_article.get(k));
    	}
		//SECTION
    	List<Integer> occurences_section = new ArrayList<Integer>();
		
		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++){
			if (!(((sheet.getRow(i)).getCell(SECTION)).getStringCellValue()).equals("")){
				occurences_section.add(i);
				System.out.println(" LINE: " + i );
			}
		}
		occurences_section.add(sheet.getPhysicalNumberOfRows());
		for (int j = 0; j < occurences_section.size() - 1; j++){
			System.out.println("FOUND TITLES AT: ");
			System.out.println("		" + occurences_section.get(j) + " TO " + occurences_section.get(j+1));
			String value = ((sheet.getRow(occurences_section.get(j))).getCell(SECTION)).getStringCellValue();
			System.out.println("		Copying: " + value);
			for (int index = occurences_section.get(j) + 1; index < occurences_section.get(j+1); index++){
				Cell cellCopy = (sheet.getRow(index)).getCell(SECTION);
				cellCopy.setCellValue(value);
				System.out.println("			UPDATE SECTION at " + index + " with VALUE of: "+ cellCopy.getStringCellValue());
			}
		}
		System.out.println("------------------  DELETE IN PROGRESS : SECTION  ------------------");
    	for (int k = occurences_section.size() - 1; k >= 0; k--){
    		removeRow(sheet, occurences_section.get(k));
    		System.out.println("Deleting row: " + occurences_section.get(k));
    	}
    	
       	System.out.println("TOTAL: " + count);
       	
        FileOutputStream fileOut = new FileOutputStream(xlsx);
		workbook.write(fileOut);
        workbook.close();
        fileOut.close();
    }
    
    private static void removeRow (Sheet sheet, int index){
    	int last = sheet.getLastRowNum();
    	if (index >= 0 && index < last)
    		sheet.shiftRows(index + 1, last, -1);
    	if (index == last){
    		Row removee = sheet.getRow(index);
    		if (removee != null)
    			sheet.removeRow(removee);
    	}
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
	
	//Not used because of .close() exception
	public static void copy(String xlsx) throws Exception{
		XSSFWorkbook output = new XSSFWorkbook(xlsx);
		Sheet opt = output.getSheetAt(0);
		List<Integer> occurences = new ArrayList<Integer>();
		
		for (int i = 1; i < opt.getPhysicalNumberOfRows(); i++){
			if (!(((opt.getRow(i)).getCell(TITLE)).getStringCellValue()).equals("")){
				occurences.add(i);
				System.out.println(" LINE: " + i );
			}
		}
		for (int j = 0; j < occurences.size() - 1; j++){
			System.out.println("FOUND TITLES AT: ");
			System.out.println("		" + occurences.get(j) + " TO " + occurences.get(j+1));
			String value = ((opt.getRow(occurences.get(j))).getCell(TITLE)).getStringCellValue();
			System.out.println("		Copying: " + value);
			for (int index = occurences.get(j) + 1; index < occurences.get(j+1); index++){
				Cell cell = (opt.getRow(index)).getCell(TITLE);
				cell.setCellValue(value);
				System.out.println("			UPDATE at " + index + " with VALUE of: "+ cell.getStringCellValue());
			}
		}

		FileOutputStream fileOut = new FileOutputStream(xlsx);
        output.write(fileOut);
        output.close();
        fileOut.close();
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
		//private final static int TITLE = 1; 		//HEADING, Regex excluding alpha characters
		cell = row.createCell(TITLE);
        cell.setCellValue("Title");
        cell.setCellStyle(style);
        
        //LEVEL 2 : CHAPTER	
		//private final static int CHAPTER = 2; 		//HEADING, Regex excluding alpha characters
		cell = row.createCell(CHAPTER);
        cell.setCellValue("Chapter");
        cell.setCellStyle(style);
		
		//LEVEL 3 : SUBCHAPTER
		//private final static int HAS_SUBCHAPTER = 0;
		cell = row.createCell(SUBCHAPTER);
        cell.setCellValue("Subchapter");
        cell.setCellStyle(style);
        
        //LEVEL 3.5 : ARTICLE
        cell = row.createCell(ARTICLE);
        cell.setCellValue("Article");
        cell.setCellStyle(style);

		//LEVEL 4 : SECTION
		//private final static int HAS_SUBCHAPTER = 0;
		cell = row.createCell(SECTION);
        cell.setCellValue("Section");
        cell.setCellStyle(style);
        
		//LEVEL 5 : PARAGRAPH
		//private final static int PARAGRAPH = 0;			//PARA
		cell = row.createCell(PARAGRAPH);
        cell.setCellValue("Requirement");
        cell.setCellStyle(style);
		
		//NEXT STEPS
		//private final static int CITATION = 0;
	    cell = row.createCell(CITATION);
        cell.setCellValue("Citation");
        cell.setCellStyle(style);
    }
}
