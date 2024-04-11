package runner;

import org.testng.annotations.Test;

import pageValidation.windowHandel;

public class Excelfile_write_Runner {
  @Test
  public void runnerexcelfilereadwrite() throws Exception {
	  windowHandel handel =new windowHandel();
	  handel.excelfilewrite("Book1", 1, 2, "windows", "hi deepak");
	 
  }
}
