package excel;

import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import jxl.Cell;
import jxl.Hyperlink;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableHyperlink;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class SheetTreeManage {
	public static List<SheetNode> allNodes = new ArrayList();

	public static int order = 1;

	public static void initAllNodes(Workbook sourceWorkbook) {
		for (String string : Util.getSetBig(sourceWorkbook)) { // 涓垎绫昏〃
			SheetNode node = new SheetNode();
			node.setId(string);
			node.setLevel(2);
			node.setLeaf(false);
			allNodes.add(node);
		}
		for (String string : Util.getSetMiddle(sourceWorkbook)) { // 灏忓垎绫昏〃
			SheetNode node = new SheetNode();
			node.setId(string);
			node.setLevel(3);
			node.setLeaf(false);
			allNodes.add(node);
		}
		for (String string : Util.getSetLittle(sourceWorkbook)) { // 鍟嗗搧琛�
			SheetNode node = new SheetNode();
			node.setId(string);
			node.setLevel(4);
			node.setLeaf(true);
			allNodes.add(node);
		}
	}

	public static void initTree(SheetNode node, Workbook sourceWorkbook) {

		// 瀵箂ource鐨勬墍鏈塻heet杩涜閬嶅巻
		for (int i = 1; i < 4; i++) {
			Sheet sourceSheet = sourceWorkbook.getSheet(i);
			// 瀵规瘡涓�琛岃繘琛岄亶鍘�
			for (int j = 1; j < sourceSheet.getRows(); j++) {
				Cell row[] = sourceSheet.getRow(j);
				String bigId = row[3].getContents();
				String middleId = row[5].getContents();
				String littleId = row[7].getContents();

				// 瀵瑰ぇ鍒嗙被琛ㄥ唴瀹硅繘琛屽垵濮嬪寲,鍗虫牴缁撶偣node
				// 濡傛灉褰撳墠璁板綍涓嶅瓨鍦ㄦ牴缁撶偣涓�
				SheetData datasBig = node.getData();
				if (!datasBig.exsit(bigId, datasBig.getAllCodes())) {
					IndexSheetDataItem data = new IndexSheetDataItem();

					String bigCode = row[3].getContents();
					String bigName = row[4].getContents();
					data.setName(bigName);
					data.setCode(bigCode);
					datasBig.getItems().add(data);
					node.setData(datasBig);
					node.setLevel(1);
					node.setId("0");
				}
				// 瀵逛腑鍒嗙被琛ㄨ繘琛屽垵濮嬪寲,鍒嗕袱閮ㄥ垎锛屼腑鍒嗙被琛ㄧ殑鍐呭锛屽拰鐖跺瓙鍏崇郴鐨勭‘绔�
				SheetNode middleNode = getNodeById(bigId, allNodes);

				// 濡傛灉褰撳墠鏉＄洰涓嶅瓨鍦�
				SheetData datasMiddle = middleNode.getData();
				if (!datasMiddle.exsit(middleId, datasMiddle.getAllCodes())) {
					IndexSheetDataItem data = new IndexSheetDataItem();
					String bigCode = row[5].getContents();
					String bigName = row[6].getContents();
					data.setName(bigName);
					data.setCode(bigCode);
					datasMiddle.getItems().add(data);
					middleNode.setData(datasMiddle);

					// 鑻ュ綋鍓嶄腑鍒嗙被琛ㄨ繕涓嶆槸澶у垎绫昏〃鐨勫瓙缁撶偣
					if (!exist(bigId, node.getChildList())) {
						node.getChildList().add(middleNode);
						middleNode.setParent(node);
					}
				}

				// 瀵瑰皬鍒嗙被琛ㄨ繘琛屽垵濮嬪寲,鍒嗕袱閮ㄥ垎锛屽皬鍒嗙被琛ㄧ殑鍐呭锛屽拰鐖跺瓙鍏崇郴鐨勭‘绔�
				SheetNode littleNode = getNodeById(middleId, allNodes);
				SheetData datasLittle = littleNode.getData();

				// 濡傛灉褰撳墠鏉＄洰涓嶅瓨鍦�
				if (!datasLittle.exsit(littleId, datasLittle.getAllCodes())) {
					IndexSheetDataItem data = new IndexSheetDataItem();
					String bigCode = row[7].getContents();
					String bigName = row[8].getContents();
					data.setName(bigName);
					data.setCode(bigCode);
					datasLittle.getItems().add(data);
					littleNode.setData(datasLittle);

					// 鑻ュ綋鍓嶅皬鍒嗙被琛ㄨ繕涓嶆槸涓垎绫昏〃鐨勫瓙缁撶偣
					SheetNode parent = getNodeById(bigId, allNodes);
					if (!exist(middleId, parent.getChildList())) {
						parent.getChildList().add(littleNode);
						littleNode.setParent(parent);
					}

				}

				// 瀵瑰晢鍝佽〃杩涜鍒濆鍖�

				SheetNode goodsNode = getNodeById(littleId, allNodes);
				SheetData goodsData = goodsNode.getData();
				GoodsSheetDataItem data = new GoodsSheetDataItem();
				data.setA(row[0].getContents());
				data.setB(row[1].getContents());
				data.setC(row[2].getContents());
				data.setD(row[3].getContents());
				data.setE(row[4].getContents());
				data.setF(row[5].getContents());
				data.setG(row[6].getContents());
				data.setH(row[7].getContents());
				data.setI(row[8].getContents());
				data.setJ(row[9].getContents());
				data.setK(row[10].getContents());
				data.setL(row[11].getContents());
				data.setM(row[12].getContents());
				data.setN(row[13].getContents());
				data.setO(row[14].getContents());
				data.setP(row[15].getContents());
				data.setQ(row[16].getContents());
				goodsData.getItems().add(data);
				goodsNode.setData(goodsData);

				SheetNode parent = getNodeById(middleId, allNodes);
				if (!exist(littleId, parent.getChildList())) {
					goodsNode.setParent(parent);
					parent.getChildList().add(goodsNode);
				}

			}
		}

		// 内排序
		SheetTreeManage.allNodes.add(0, node);
		for (SheetNode nodeSort : allNodes) {
			if (!nodeSort.isLeaf()) {
				Collections.sort(nodeSort.getData().getItems(),
						new DataComparator());
			}
		}

		// SheetTreeManage.allNodes.add(0, node);
		for (SheetNode nodeSort : allNodes) {
			if (nodeSort.getChildList() != null)
				Collections
						.sort(nodeSort.getChildList(), new SheetComparator());
		}

		// 设置number名称
		for (SheetNode nodeSort : allNodes) {
			if (nodeSort.level > 1)
				nodeSort.setNumber(nodeSort.getParent().getChildList()
						.indexOf(nodeSort) + 1);

		}

		for (SheetNode nodeSort : allNodes) {

			if (nodeSort.level == 1) {
				nodeSort.setName("大分类");
				List<String> stringList = new ArrayList<>();
				stringList.add("NO");
				stringList.add("大分類コード");
				stringList.add("大分類名称");
				nodeSort.getData().setHeader(stringList);
			}
			if (nodeSort.level == 2) {
				nodeSort.setName("中分类(" + nodeSort.getNumber() + ")");
				List<String> stringList = new ArrayList<>();
				stringList.add("NO");
				stringList.add("中分類コード");
				stringList.add("中分類名称");
				nodeSort.getData().setHeader(stringList);
			}
			if (nodeSort.level == 3) {
				nodeSort.setName("小分类(" + nodeSort.getParent().getNumber()
						+ "-" + nodeSort.getNumber() + ")");
				List<String> stringList = new ArrayList<>();
				stringList.add("NO");
				stringList.add("小分類コード");
				stringList.add("小分類名称");
				nodeSort.getData().setHeader(stringList);
			}
			if (nodeSort.level == 4) {
				nodeSort.setName("商品类("
						+ nodeSort.getParent().getParent().getNumber() + "-"
						+ nodeSort.getParent().getNumber() + "-"
						+ nodeSort.getNumber() + ")");

				List<String> stringList = new ArrayList<>();
				stringList.add("商品コード");
				stringList.add("商品名");
				stringList.add("JANコード");
				stringList.add("大分類コード");
				stringList.add("大分類名称");
				stringList.add("中分類コード");
				stringList.add("中分類名称");
				stringList.add("小分類コード");
				stringList.add("小分類名称");
				stringList.add("最重要成分");
				stringList.add("成分名称");
				stringList.add("メーカー名");
				stringList.add("ブランドコード");
				stringList.add("ブランド名");
				stringList.add("ランキング");
				stringList.add("ご提案卸価格(税抜）");
				stringList.add("URL");
				nodeSort.getData().setHeader(stringList);
			}
		}
	}

	public static void fillIndexSheet(SheetNode node,
			WritableWorkbook targetWorkbook) throws RowsExceededException,
			WriteException, IOException {
		WritableSheet targetSheet = targetWorkbook.createSheet(node.getName(),
				order++);

		int temp = 0;
		for (String string : node.getData().getHeader()) {
			Label label = new Label(temp++, 1, string);
			targetSheet.addCell(label);

			for (int i = 0; i < node.getData().getItems().size(); i++) {
				if (!node.isLeaf()) {
					{
						IndexSheetDataItem data = (IndexSheetDataItem) node
								.getData().getItems().get(i);
						WritableCell cellNumber = new Label(0, i + 2, i + 1
								+ "");
						WritableCell cellCode = new Label(1, i + 2,
								data.getCode());
						WritableCell cellName = new Label(2, i + 2,
								data.getName());
						targetSheet.addCell(cellCode);
						targetSheet.addCell(cellNumber);
						targetSheet.addCell(cellName);
					}

				} else {

					GoodsSheetDataItem data = (GoodsSheetDataItem) node
							.getData().getItems().get(i);
					WritableCell cellA = new Label(0, i + 2, data.getA());
					WritableCell cellB = new Label(1, i + 2, data.getB());
					WritableCell cellC = new Label(2, i + 2, data.getC());
					WritableCell cellD = new Label(3, i + 2, data.getD());
					WritableCell cellE = new Label(4, i + 2, data.getE());
					WritableCell cellF = new Label(5, i + 2, data.getF());
					WritableCell cellG = new Label(6, i + 2, data.getG());
					WritableCell cellH = new Label(7, i + 2, data.getH());
					WritableCell cellI = new Label(8, i + 2, data.getI());
					WritableCell cellJ = new Label(9, i + 2, data.getJ());
					WritableCell cellK = new Label(10, i + 2, data.getK());
					WritableCell cellL = new Label(11, i + 2, data.getL());
					WritableCell cellM = new Label(12, i + 2, data.getM());
					WritableCell cellN = new Label(13, i + 2, data.getN());
					WritableCell cellO = new Label(14, i + 2, data.getO());
					WritableCell cellP = new Label(15, i + 2, data.getP());
					WritableCell cellQ = new Label(16, i + 2, data.getQ());
					targetSheet.addCell(cellA);
					targetSheet.addCell(cellB);
					targetSheet.addCell(cellC);
					targetSheet.addCell(cellD);
					targetSheet.addCell(cellE);
					targetSheet.addCell(cellF);
					targetSheet.addCell(cellG);
					targetSheet.addCell(cellH);
					targetSheet.addCell(cellI);
					targetSheet.addCell(cellJ);
					targetSheet.addCell(cellK);
					targetSheet.addCell(cellL);
					targetSheet.addCell(cellM);
					targetSheet.addCell(cellN);
					targetSheet.addCell(cellO);
					targetSheet.addCell(cellP);
					targetSheet.addCell(cellQ);
				}
			}

		}

	}

	public static void setLink(WritableWorkbook targetWorkbook,
			List<SheetNode> list) throws RowsExceededException, WriteException {
		for (SheetNode node : list) {

			if (node.getLevel() < 4) {
				int temp = 0;
				for (Object item : node.getData().getItems()) {
					IndexSheetDataItem iitem = (IndexSheetDataItem) item;
					String id = iitem.getCode();
					WritableSheet targetSheet = targetWorkbook
							.getSheet(getNodeById(id, allNodes).getName());
					WritableHyperlink link = new WritableHyperlink(2,
							2 + temp++, iitem.getName(), targetSheet, 0, 0);
					targetWorkbook.getSheet(node.getName()).addHyperlink(link);
				}
			}

			if (node.getLevel() > 1) {
				WritableHyperlink rootLink = new WritableHyperlink(0, 0, "大分类",
						targetWorkbook.getSheet("大分类"), 0, 0);
				targetWorkbook.getSheet(node.getName()).addHyperlink(rootLink);
			}

			if (node.getLevel() == 3) {
				WritableHyperlink rootLink = new WritableHyperlink(4, 0, "中分类",
						targetWorkbook.getSheet(node.getParent().getName()), 0,
						0);
				targetWorkbook.getSheet(node.getName()).addHyperlink(rootLink);
			}
			if (node.getLevel() == 4) {
				WritableHyperlink littleLink = new WritableHyperlink(6, 0,
						"小分类", targetWorkbook.getSheet(node.getParent()
								.getName()), 0, 0);
				targetWorkbook.getSheet(node.getName())
						.addHyperlink(littleLink);

				WritableHyperlink middleLink = new WritableHyperlink(4, 0,
						"中分类", targetWorkbook.getSheet(node.getParent()
								.getParent().getName()), 0, 0);
				targetWorkbook.getSheet(node.getName())
						.addHyperlink(middleLink);
			}
		}
	}

	public static SheetNode getNodeById(String id, List<SheetNode> nodeList) {
		for (SheetNode node : nodeList) {
			if (id.equals(node.getId()))
				return node;
		}
		return null;
	}

	public static boolean exist(String id, List<SheetNode> nodeList) {
		if (getNodeById(id, nodeList) == null)
			return false;
		return true;
	}
}
