package in.alam;


import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ReportRestController {
	
	@Autowired
	private ReportService reportService;
	
	@GetMapping("/getAllStudent")
	public List<Student> getAllStudent() throws Exception{
		return reportService.getAllStudent();
	}
	
	@PostMapping("/saveDataToFile")
	public String saveDataToFile(@RequestBody Student student) throws Exception{
		return reportService.saveDataToFile(student);
	}
	
	@PutMapping("/updateStudent")
	public String updateDataToFile(@RequestBody Student student) throws Exception{
		return reportService.updateStudent(student);
	}
	
	@DeleteMapping("/deleteStudent/{studentId}")
	public String deleteStudent(@PathVariable("studentId") Long studentId) throws Exception{
		return reportService.deleteStudent(studentId);
	}
}
