package excel;

import java.io.IOException;

public class Excelmain {
	public static void main(String[] args) throws IOException {
		ExcelRead er = new ExcelRead();
	
		for (int i = 0; i < er.rowsize(); i++) {  

			for (int j = 0; j < 3; j++) {  			
				String value = er.readData(i, j);
				System.out.println(value);
			}
		}
	}
}
