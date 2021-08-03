/**
 * 
 * @Author : Aslam
 * @Dependancies : MailSender.java
 * @Details : This class handles all the operations with excel sheet. 
 * 				(i.e reading , writing)
 * 
 */

package projects.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelHandler {

	private List<String> headers = new ArrayList<>();

	/**
	 * Accepts the file where the data about all the students is stored. Then it
	 * handles everythiing else by calling methods within this class as well as from
	 * {@link MailSender} class
	 */
	public void sendReports(String filename) throws IOException {
		Object[][] data = read(filename);

		// print data
		print(data);

		// Get the headers (i.e name, roll, etc) which are in the first row (i=0)
		for (int column = 0; column < data[0].length; column++) {
			Object header = data[0][column]; // returns the header as object
			headers.add((String) header); // casting the object to String and adding it to the headers list
		}

		// Iterate through all the student array starting form second row (i=1)
		for (int i = 1; i < data.length; i++) {

			// Pass each student array to the write method. It returns an array 
			// of length two, containing the parents' email and excel file location.
			String[] detailsForSending = write(data[i]);

			// pass the details to the send method of MailSender
			MailSender mailSender = new MailSender();
			/** 
			* 0th index -> parent's email address
			* 1st index -> excel file of the student
			*/
			mailSender.sendMail(detailsForSending[0], detailsForSending[1]);

		}

	}

	/**
	 * Reads the file passed to this method and returns a two dimentional array. The
	 * 2d array will contain all the data
	 */
	private Object[][] read(String filename) throws IOException {

		// Get the workbook
		Workbook workbook = new HSSFWorkbook(new FileInputStream(filename));

		// Get the first sheet(that is where the data is)
		Sheet sheet = workbook.getSheetAt(0);

		// No.of rows
		int rows = sheet.getPhysicalNumberOfRows();

		// No.of columns
		int columns = sheet.getRow(0).getPhysicalNumberOfCells();

		System.out.printf("Rows: %s, Columns: %s\n", rows, columns);
		System.out.println("---------------------------------------");

		/*
		 * Declare 2 dimensional String array and initialize it with the total no.of
		 * rows and columns Each row represents a student. And each column in that row
		 * represents that particular students' details The details can be number or
		 * string, so we declare the 2d array as Object array Object array can hold both
		 * String and numbers
		 */
		Object[][] data = new Object[rows][columns];

		for (int i = 0; i < rows; i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getLastCellNum(); j++) {

				try {
					// If the value is a String, it will be added to the array
					data[i][j] = row.getCell(j).getStringCellValue();

				} catch (IllegalStateException ise) {
					// if it is not a string(i.e number) then it will throw a ISE exception
					// And hence it will return a value as a double. we can cast it to (int)
					// Check what happens if we dont cast :)
					data[i][j] = (int) row.getCell(j).getNumericCellValue();
				}
			}
		}

		// close workbook
		workbook.close();
		return data;
	}

	/**
	 * This method accepts a single student details as an array. It creates a new
	 * excel file based on the student Then it returns that file's location and the
	 * parent email of the student
	 */
	private String[] write(Object[] student) throws IOException {

		// Create a workbook (excel file)
		Workbook workbook = new HSSFWorkbook();

		// Output file name
		String destination = "data/" + student[0] + "'s Record.xls";
		FileOutputStream output = new FileOutputStream(destination);

		// Create sheet (excel sheet)
		Sheet sheet = workbook.createSheet("Record");

		// Add all the details of a single student to the newly created sheet
		Row row;
		for (int i = 0; i < student.length; i++) {
			row = sheet.createRow(i); // create a row
			row.createCell(0).setCellValue(headers.get(i)); // set header on 1st column
			row.createCell(1).setCellValue(student[i].toString()); // set value on 2nd column
		}

		// Auto sizes column. see what happens if you dont declare this
//		sheet.autoSizeColumn(0);
//		sheet.autoSizeColumn(1);

		// TODO :
		// Find how to align the cell values to the right/left.
		// Try to make the headers as bold

		// Writes to the file specified
		workbook.write(output);
		workbook.close();
		System.out.printf("File created for student: %s at location : %s\n", student[0], destination);

		String parentEmail = (String) student[student.length - 1];
		return new String[] {parentEmail, destination};
	}

	// just for printing the data of 2d array so we can see in the console
	private void print(Object[][] data) {

		System.out.println("Displaying data :-");

		for (Object[] row : data) {
			for (int col = 0; col < row.length; col++) {
				if (col == 0)
					System.out.print(row[col] + "\t\t");
				else
					System.out.print(row[col] + "\t");
			}
			System.out.println();
		}

		System.out.println("---------------------------------------");
	}
}