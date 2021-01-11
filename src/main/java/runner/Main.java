package runner;

import data.Data;
import helper.Const;

public class Main {

	public static void main(String[] abc) throws Exception {
		Data.setExcelFile(Const.Login_Prod_File, Const.Path_Prod_Data, "Sheet1");
		System.out.println(Data.getCellQuanity());
	}
}
