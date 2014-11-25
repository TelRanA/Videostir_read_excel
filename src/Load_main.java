import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import jxl.Workbook;



public class Load_main {
	
	public static void main(String[] args) {
		
		LoadFromExcel w = new LoadFromExcel();
		String path = "";
		path = w.writeAbsoluteFilePath();
		String pathCreatedFile = w.writeAbsoluteFilePath() +"\\"+ "input.txt";
		
		ArrayList<String> list = w.readFromExel(path);
		LoadFromExcel.createFile(path + "\\" + "input.txt");
				
		LoadFromExcel.writeToFile(pathCreatedFile, list);
	}
	

		
}
