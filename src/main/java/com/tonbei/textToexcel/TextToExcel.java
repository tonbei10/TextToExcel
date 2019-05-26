package com.tonbei.textToexcel;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.CodeSource;
import java.security.ProtectionDomain;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TextToExcel {

	public static void main(String[] args) {

		try {
			File parepath = getApplicationPath(TextToExcel.class).getParent().toFile();
			File path = new File(parepath, "Output.xlsx");

			Workbook wb = new XSSFWorkbook();
			FileOutputStream fileOut = new FileOutputStream(path);

			for(String arg : args) {

				System.out.println(arg);

				File txt = new File(arg);
				BufferedReader reader = new BufferedReader(new FileReader(txt));

				int num = Integer.parseInt(reader.readLine());
				double x[] = new double[num];
				double y[] = new double[num];

				int i = 0;
				String line;
				while ((line = reader.readLine()) != null) {
					line = line.replace("  ", " ").replace("  ", " ");
					String[] sp = (line.startsWith(" ") ? line.substring(1) : line).split(" ", 2);
					x[i] = Double.parseDouble(sp[0]);
					y[i] = Double.parseDouble(sp[1]);
					i++;
					System.out.println(txt.getName() + " : " + i);
				}
				reader.close();

/*
				BufferedWriter writer = new BufferedWriter(new FileWriter(new File(parepath, "test.txt")));
				for(i = 0; i < num; i++) {
					writer.write(x[i] + " " + y[i]);
					writer.newLine();
				}
				writer.close();
*/

				String safeName = WorkbookUtil.createSafeSheetName(txt.getName());
				Sheet sheet1 = wb.createSheet(safeName);

				Row row0 = sheet1.createRow(0);
				row0.createCell(0).setCellValue("X");
				row0.createCell(1).setCellValue("Y");
				for(i = 0; i < num; i++) {
					Row rowi = sheet1.createRow(i + 1);
					rowi.createCell(0).setCellValue(x[i]);
					rowi.createCell(1).setCellValue(y[i]);
				}
			}

			wb.write(fileOut);
			wb.close();
			fileOut.close();

		}catch (Exception e) {
			e.printStackTrace();
		}finally {

		}
	}

	public static Path getApplicationPath(Class<?> cls) throws URISyntaxException {
		ProtectionDomain pd = cls.getProtectionDomain();
		CodeSource cs = pd.getCodeSource();
		URL location = cs.getLocation();
		URI uri = location.toURI();
		Path path = Paths.get(uri);
		return path;
	}

}