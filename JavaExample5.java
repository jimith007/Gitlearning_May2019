package step.com.test.javaexamples;

public class JavaExample5 {
	
	public void calculateSalary(String designation, int experience){
		int finalSalary=0;
		
		finalSalary = calcPerfAllow(designation) + calcLeaveAllow();
		
		System.out.println("Salary is " + finalSalary + " and the experience is " + experience);
	}
	
	public int calcPerfAllow(String designation){
		int perfAllow = 0;
		
		if(designation.contentEquals("Manager")){
			perfAllow = 1000;
		}else{
			perfAllow = 5000;
		}
		
		return perfAllow;
	}
	
	public int calcLeaveAllow(){
		int leaveAllow = 10;
		
		return leaveAllow;
	}
	
	public static void main(String args[]){
		JavaExample5 salary = new JavaExample5();
		
		salary.calculateSalary("Manager", 4);
		salary.calculateSalary("TestEngineer", 2);
	}
}