
import java.io.IOException;

public class Demo {

	public static void main(String[] args) throws IOException {
		
		String filePath = "/Volumes/Seagate Slim Drive/futures/01620500000880/01620500000880_2014-12-11.xls";

		System.out.println(ZJZK.getFXD(filePath));
		System.out.println(ZJZK.getZJBZJ(filePath));
		System.out.println(ZJZK.getBZJZY(filePath));
	
	}

}
