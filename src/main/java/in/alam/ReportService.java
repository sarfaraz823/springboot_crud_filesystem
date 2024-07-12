package in.alam;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.stereotype.Service;

@Service
public class ReportService {

	String filePath = "D:\\MyDownload\\StudentRecord.xls";

	public String saveDataToFile(Student student) throws Exception {

		FileInputStream fileinputStream = new FileInputStream(new File(filePath));

		HSSFWorkbook workbook = new HSSFWorkbook(fileinputStream);
		HSSFSheet sheet = workbook.getSheetAt(0);

		int dataRowIndex = sheet.getLastRowNum() + 1;

		HSSFRow dataRow = sheet.createRow(dataRowIndex);
		dataRow.createCell(0).setCellValue(student.getStudentId());
		dataRow.createCell(1).setCellValue(student.getFirstName());
		dataRow.createCell(2).setCellValue(student.getLastName());
		dataRow.createCell(3).setCellValue(student.getAge());

		FileOutputStream outputStream = new FileOutputStream(filePath);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();

		return "File Saved Successfully....";

	}

	public List<Student> getAllStudent() throws IOException {

		List<Student> list = new ArrayList<>();
		FileInputStream fileinputStream = new FileInputStream(new File(filePath));

		HSSFWorkbook workbook = new HSSFWorkbook(fileinputStream);
		HSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			if (row.getRowNum() == 0) {
				continue;
			}

			if (sheet.getPhysicalNumberOfRows() == -1) {
				break;
			}

			Student obj = new Student();

			Cell cell1 = row.getCell(0);
			Cell cell2 = row.getCell(1);
			Cell cell3 = row.getCell(2);
			Cell cell4 = row.getCell(3);

			obj.setStudentId((long) cell1.getNumericCellValue());
			obj.setFirstName(cell2.getStringCellValue());
			obj.setLastName(cell3.getStringCellValue());
			obj.setAge((long) cell4.getNumericCellValue());
			list.add(obj);

		}
		return list;
	}

	public String updateStudent(Student student) throws IOException {

		FileInputStream fileinputStream = new FileInputStream(new File(filePath));

		HSSFWorkbook workbook = new HSSFWorkbook(fileinputStream);
		HSSFSheet sheet = workbook.getSheetAt(0);
		long studentId = (long) student.getStudentId();

		Row row = sheet.getRow((int) studentId);

		Cell cell1 = row.getCell(0);
		cell1.setCellValue(student.getStudentId());
		Cell cell2 = row.getCell(1);
		cell2.setCellValue(student.getFirstName());
		Cell cell3 = row.getCell(2);
		cell3.setCellValue(student.getLastName());
		Cell cell4 = row.getCell(3);
		cell4.setCellValue(student.getAge());

		fileinputStream.close();

		FileOutputStream fileOutputStream = new FileOutputStream(new File(filePath));
		workbook.write(fileOutputStream);
		workbook.close();
		fileOutputStream.close();

		return "Record is updated successfully..";
	}

	public String deleteStudent(long studentId) throws IOException {
		FileInputStream fileinputStream = new FileInputStream(new File(filePath));
		HSSFWorkbook workbook = new HSSFWorkbook(fileinputStream);
		HSSFSheet sheet = workbook.getSheetAt(0);
		HSSFRow row = sheet.getRow((int) studentId);
		sheet.removeRow(row);
		fileinputStream.close();
		FileOutputStream outFile = new FileOutputStream(new File(filePath));
		workbook.write(outFile);
		outFile.close();
		workbook.close();

		return "Record Deleted Succssfully...";
	}
}
