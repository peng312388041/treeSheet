//姝ょ増鏈凡缁� 瀹炵幇寤虹珛sheet锛屽苟瀛樺叆鏁版嵁
package excel;

import java.io.IOException;
import java.util.Collections;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Main {

	public static void main(String[] args) throws BiffException, IOException,
			RowsExceededException, WriteException {

		long begin = System.currentTimeMillis();
		System.out.println();
		// 婧恇ook
		Workbook sourceWorkbook = Util.getSourceWorkbook();
		// 鐩爣book
		WritableWorkbook targetWorkbook = Util.getTargetWorkbook();
		// 鎵�鏈夊皬鍒嗙被id

		SheetNode root = new SheetNode();

		SheetTreeManage.initAllNodes(sourceWorkbook);
		SheetTreeManage.initTree(root, sourceWorkbook);

		for (SheetNode node : SheetTreeManage.allNodes) {
			SheetTreeManage.fillIndexSheet(node, targetWorkbook);
			// System.out.println(node.getId());
		}
		SheetTreeManage.setLink(targetWorkbook, SheetTreeManage.allNodes);

		targetWorkbook.write();
		targetWorkbook.close();
		long end = System.currentTimeMillis();
		System.out.println("鍏辫�楁椂" + (end - begin) / 1000 + 1 + "绉�");
	}
}
