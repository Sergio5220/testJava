
package com.test.monitoring;


import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFCellUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.xmlbeans.*;



public class MonitoringFolder {
    
    public static Properties prop = new Properties();

	public static final String CONFIG_FILE = "./config.properties";
	public static final String LOCK_FILE = "lock_file";
        public static final String SALIDA_XLS = "salidaXls";
        public static final String PATH = "path";
	public static final String FILE_NAME = "fileName";
	public static final String EXT = "extension";


	private String sourcePath;
	private File originalFile;
	private String salidaXls;
        private String lockFile;
        private String path;
        private String extension;
	

    
    
    public static void main(String[] args) {
        File lock = null;
        MonitoringFolder m = new MonitoringFolder();
        try{//carga configuraciones
        	prop.load(new FileReader(new File(CONFIG_FILE)));
        	m.sourcePath = prop.getProperty("sourcePath");
        	m.salidaXls = prop.getProperty(SALIDA_XLS);
        	m.originalFile = new File(prop.getProperty(PATH)+prop.getProperty(FILE_NAME));
        	m.lockFile = prop.getProperty(LOCK_FILE);
        	m.path = prop.getProperty(PATH);
        	m.extension=prop.getProperty(EXT);
            List<String> dataFinal = new ArrayList<String>();

        	
       try {//Proceso de ejecucion
           lock = new File(m.lockFile);
           //Si no hay procesos ejecutandose
           if(!lock.exists()){
                lock.createNewFile();
                 File f=new File(m.path+m.salidaXls+"001"+m.extension);
                 m.buscarHojas();
		if(lock!=null){
                    lock.delete();
		}
		}else{
		}
		}catch (NullPointerException e){
                    e.printStackTrace();
		}catch (FileNotFoundException e) {
                    e.printStackTrace();
		}catch (IOException e) {
                    e.printStackTrace();
		}
		}catch(Exception e) {
                    e.printStackTrace();
		}
    }
    
    public void buscarHojas() throws IOException{
        String sheetName ="";
        String[] archivos = null;
        String ext = "";
        File f = new File(sourcePath);
        File[] contenido = f.listFiles();
        if (f.isDirectory()){
            //recorre el listado de archivos en el directorio
            for (File contenido1 : contenido) {
                //System.out.println(FilenameUtils.getExtension(contenido1.toString()));
                ext = FilenameUtils.getExtension(contenido1.toString());
                //valida que sea un archivo xls
                    if (ext.equals("xls")){
                        //consolida el archivo
                    	FileInputStream fis = null; 
                    	HSSFWorkbook wb = null; 
                    try {
                    	fis = new FileInputStream(contenido1.toString());
                        wb = new HSSFWorkbook(fis); 
                        ArrayList<String> data = new ArrayList();
                          consolidar("\\"+contenido1); 
                    } catch (Exception ex) {
                        Logger.getLogger(MonitoringFolder.class.getName()).log(Level.SEVERE, null, ex);
                    } finally {
                    	if(fis!=null) {
                    		fis.close();                    		
                    	}
                    	
                    }
                    

                    }else{
                    
                    }
                }   
            }
        }
    
    	public void consolidar(String excelPath) throws IOException {
    		/*HSSFWorkbook workbook = null; 
    		File file = new File(path+salidaXls+extension);
    		 FileOutputStream fileOut = new FileOutputStream(file);
    		 File file2 = new File(excelPath);
    		 FileInputStream fileIn = new FileInputStream(file2);
             int fCell = 0;
             int lCell = 0;
             int fRow = 0;
             int lRow = 0;
    		 if(!file.exists()) {
    			 try {
    		            workbook = (HSSFWorkbook)WorkbookFactory.create(file);
    		            HSSFWorkbook workbookIn = new HSSFWorkbook(fileIn);
    		        } catch (InvalidFormatException e) {
    		            e.printStackTrace();
    		        }
    			    int sheets = workbookIn.getNumberOfSheets();
    		        HSSFSheet sheet = workbook.createSheet("Sample sheet2");
    		    }*/
    		 
    		
    		System.out.println(excelPath);
            BufferedInputStream bis = new BufferedInputStream(new FileInputStream(excelPath));
            BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(path+salidaXls+extension, true));
            HSSFWorkbook workbook = new HSSFWorkbook(bis);
            HSSFWorkbook myWorkBook = new HSSFWorkbook();
            HSSFSheet sheet = null;
            HSSFRow row = null;
            HSSFCell cell = null;
            HSSFSheet mySheet = null;
            HSSFRow myRow = null;
            HSSFCell myCell = null;
            int sheets = workbook.getNumberOfSheets();
            int fCell = 0;
            int lCell = 0;
            int fRow = 0;
            int lRow = 0;
    		
            for (int iSheet = 0; iSheet < sheets; iSheet++) {
                sheet = workbook.getSheetAt(iSheet);
                if (sheet != null) {
                    mySheet = myWorkBook.createSheet(sheet.getSheetName()+iSheet);
                    fRow = sheet.getFirstRowNum();
                    lRow = sheet.getLastRowNum();
                    for (int iRow = fRow; iRow <= lRow; iRow++) {
                        row = sheet.getRow(iRow);
                        myRow = mySheet.createRow(iRow);
                        if (row != null) {
                            fCell = row.getFirstCellNum();
                            lCell = row.getLastCellNum();
                            for (int iCell = fCell; iCell < lCell; iCell++) {
                                cell = row.getCell(iCell);
                                myCell = myRow.createCell(iCell);
                                if (cell != null) {
                                    myCell.setCellType(cell.getCellType());
                                    switch (cell.getCellType()) {
                                    case HSSFCell.CELL_TYPE_BLANK:
                                        myCell.setCellValue("");
                                        break;

                                    case HSSFCell.CELL_TYPE_BOOLEAN:
                                        myCell.setCellValue(cell.getBooleanCellValue());
                                        break;

                                    case HSSFCell.CELL_TYPE_ERROR:
                                        myCell.setCellErrorValue(cell.getErrorCellValue());
                                        break;

                                    case HSSFCell.CELL_TYPE_FORMULA:
                                        myCell.setCellFormula(cell.getCellFormula());
                                        break;

                                    case HSSFCell.CELL_TYPE_NUMERIC:
                                        myCell.setCellValue(cell.getNumericCellValue());
                                        break;

                                    case HSSFCell.CELL_TYPE_STRING:
                                        myCell.setCellValue(cell.getStringCellValue());
                                        break;
                                    default:
                                        myCell.setCellFormula(cell.getCellFormula());
                                    }
                                }
                            }
                        }
                    }
                }
            }
            bis.close();        
            myWorkBook.write(bos);
            //bos.close();
      
    		
    	}
    	
    	
}
