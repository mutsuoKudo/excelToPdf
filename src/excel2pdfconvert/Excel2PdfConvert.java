package excel2pdfconvert;

import java.io.File;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

public class Excel2PdfConvert {

	 public static void main(String[] args) {

		 File file = new File("C:\\Users\\user\\Desktop\\pdf_test");
		 File files[] = file.listFiles();

//		 pdf_testフォルダに入っているファイル分繰り返す
		 for (int i=0; i<files.length; i++) {

//			 確認用
			 System.out.println("ファイル" + (i+1) + "→" + files[i]);
//			 System.out.println("C:\\Users\\user\\Desktop\\pdf_test\\" + files[i]);
//			 System.out.println("C:\\Users\\user\\Desktop\\pdf_test\\" + files[i] + ".pdf");


//			ファイルの拡張子を取り除く
			 String basename = files[i].getName();
			 String text = basename.substring(0,basename.lastIndexOf('.'));

//			 確認用
			 System.out.println(basename);
//			 System.out.println("C:\\Users\\user\\Desktop\\pdf_test\\" + text + ".xlsx");
//			 System.out.println("C:\\Users\\user\\Desktop\\pdf_test\\" + text + ".pdf");



	        DefaultOfficeManagerConfiguration config = new DefaultOfficeManagerConfiguration();
	        //とりあえずサンプルなのでベタに設定
	        config.setOfficeHome("C:\\Program Files (x86)\\OpenOffice 4");
	        config.setPortNumber(8100);
	        //本来はタイムアウトの設定とかもしたほうが良い

	        OfficeManager officeManager = config.buildOfficeManager();
	        officeManager.start();

	        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
//	        ↓テスト用
//	        converter.convert(new File("C:\\Users\\user\\Desktop\\pdf_test\\" + text + ".xlsx"), new File("C:\\Users\\user\\Desktop\\pdf_test\\" + text + ".pdf"));
//			EXCELファイルは【\\192.168.71.4\Public\TEMP\髙橋さんへ\エンジニア作業報告書\\excel】
	        converter.convert(new File("C:\\Users\\user\\Desktop\\pdf_test\\" + text + ".xlsx"), new File("C:\\xampp\\htdocs\\hokushin_util\\ajax\\report_list\\report\\" + text + ".pdf"));
//	        converter.convert(new File("C:\\Users\\user\\Desktop\\pdf_test\\" + basename), new File("C:\\xampp\\htdocs\\hokushin_util\\ajax\\report_list\\report\\" + text + ".pdf"));

	        officeManager.stop();
	    }
	 }

}
