
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


public class LoadFromExcel {
	
	private String absoluteProjectPath;
	//private String absoluteFilePath;
	private String nameFile = "client.xls";
	
	public String writeAbsoluteFilePath(){
		
		File file = new File(nameFile);		
		absoluteProjectPath = file.getAbsoluteFile().getParentFile().getAbsolutePath().toLowerCase();		
		//System.out.println(absoluteProjectPath);
		return absoluteProjectPath;
	}
	
	//"E:\Dropbox\QA_Ashkelon\Programming\Alex_Java\Load_from_file\src\client.xls"
	//tc0-regular <script> VS.Player.show("bottom-right",300,270,"HASHID",{"auto-play":true,"playback-delay":1,"extrab":1,"gubed":true});</script>
	public ArrayList<String> readFromExel(String filePath){
		Workbook wb = null;
		ArrayList<String> arlist = new ArrayList<String>();
		try {			
			wb = Workbook.getWorkbook(new File(absoluteProjectPath + "\\" + nameFile));
		} catch (Exception e) {			
			e.printStackTrace();
		}      
        String line = "";                        
        for (int sheetNo = 0; sheetNo < wb.getNumberOfSheets(); sheetNo++) {

            Sheet sheet = wb.getSheet(sheetNo);
            int tcNum = 0;
            String skobki = "";
            
            if(sheet.getName().trim() != null){
                int columns = sheet.getColumns();
                int rows = sheet.getRows();
                String data;
                String nameColplusData;
                
                for (int row = 1; row < rows; tcNum++, row++) {

                    for (int col = 1; col < columns; col++) {
                        data = sheet.getCell(col, row).getContents();
                        
                        if (data.equals("*")) nameColplusData = "";
                        
                        else {
                        	nameColplusData = sheet.getCell(col, 0).getContents();
                        
                        	if (col == 3)        skobki = "{";
                        	else if (col == 15)  skobki = "}";
                        	else                 skobki = ";";
                        	nameColplusData = nameColplusData +":"+ data;
                        	line =   line + nameColplusData + skobki;
                        }
                    }
                  // System.out.println(line);    
                   arlist.add("tc" + tcNum+ "-" + "regular " + "<script> "+ "VS.Player.show(" + line + ");</script>");
                   line = "";
                }
              /* Iterator iter = arlist.iterator();
        		while (iter.hasNext()){
        			String stroka = (String)iter.next();
        			System.out.println(stroka.toString());
        		}*/
            }
        }
        return arlist;
	}   

	public static void createFile(String path){
		File f = new File(path);
		
		try {
			
				boolean isCreated = f.createNewFile();
				if(isCreated)
					System.out.println("Created !!!");
				else 
					System.out.println("Not Created :((( ");
				
		} catch (IOException e) {			
			e.printStackTrace();
		}		
	}
	
	public static void writeToFile (String filePath, ArrayList<String> list){
		
		FileWriter fw;
		
		try {
			fw = new FileWriter(filePath);
			BufferedWriter bw = new BufferedWriter(fw);			
			
			Iterator iter = list.iterator();
    		while (iter.hasNext()){
    			String stroka = (String)iter.next();
    			System.out.println(stroka.toString());
    			
    			bw.write(stroka);
    			bw.flush();
    			bw.newLine();
    			
    		}
						
			bw.close();
		} catch (IOException e) {	
			e.printStackTrace();
		}
	}
	 
}