package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.WritableWorkbook;

public class Util {
	// 婧愭枃浠�
	private static Workbook sourceWorkbook = null;
	// 鐩爣鏂囦欢
	private static WritableWorkbook targetWorkbook = null;

	public static Workbook getSourceWorkbook() {
		InputStream inputStream = null;
		try {
			inputStream = new FileInputStream(new File("D:/source.xls"));
			sourceWorkbook = Workbook.getWorkbook(inputStream);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return sourceWorkbook;
	}

	public static WritableWorkbook getTargetWorkbook() {
		OutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream(new File("D:/target.xls"));
			targetWorkbook = Workbook.createWorkbook(outputStream);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return targetWorkbook;
	}

	public static List<String> getSetLittle(Workbook sourceWorkbook) {
		Set<String> setLittle = new HashSet<>();
		for (int i = 1; i < 4; i++) {
			Sheet sheet = sourceWorkbook.getSheet(i);
			for (Cell cell : sheet.getColumn(7)) {
				setLittle.add(cell.getContents());
			}
		}

		String string[] = setLittle.toArray(new String[] {});
		Arrays.sort(string);
		List<String> listLittle = new ArrayList<String>(Arrays.asList(string));
		listLittle.remove(listLittle.size() - 1);
		return listLittle;
	}

	public static List<String> getSetMiddle(Workbook sourceWorkbook) {
		Set<String> setMiddle = new HashSet<>();
		for (int i = 1; i < 4; i++) {
			Sheet sheet = sourceWorkbook.getSheet(i);
			for (Cell cell : sheet.getColumn(5)) {
				setMiddle.add(cell.getContents());
			}
		}

		String string[] = setMiddle.toArray(new String[] {});
		Arrays.sort(string);
		List<String> listMiddle = new ArrayList<String>(Arrays.asList(string));
		listMiddle.remove(listMiddle.size() - 1);
		return listMiddle;
	}

	public static List<String> getSetBig(Workbook sourceWorkbook) {
		Set<String> setBig = new HashSet<>();
		for (int i = 1; i < 4; i++) {
			Sheet sheet = sourceWorkbook.getSheet(i);
			for (Cell cell : sheet.getColumn(3)) {
				setBig.add(cell.getContents());
			}
		}

		String string[] = setBig.toArray(new String[] {});
		Arrays.sort(string);
		List<String> listBig = new ArrayList<String>(Arrays.asList(string));
		listBig.remove(listBig.size() - 1);
		return listBig;
	}
}
