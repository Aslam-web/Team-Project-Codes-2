/**
 * 
 * @author : Aslam
 * @Dependancies : ExcelHandler.java
 * @Details : This is the main file where the project is run. 
 * 				It sends the file that has the data to ExcelHandler.java
 * 
 */

package projects.excel;

import java.io.IOException;

public class Main {

	private static String fileLocation = "data/data.xls";

	/**
	 * creates an object of excel handler and calls the sendReports method by
	 * passing the location of the file
	 */

	public static void main(String[] args) throws IOException {

		ExcelHandler handler = new ExcelHandler();
		handler.sendReports(fileLocation);
	}
}