package excel;

import java.util.ArrayList;
import java.util.List;

public class SheetData {
	private List items = new ArrayList(); // 内容条目
	private List<String> header = new ArrayList<>(); // 表头

	public List getItems() {
		return items;
	}

	public void setItems(List items) {
		this.items = items;
	}

	public List<String> getHeader() {
		return header;
	}

	public void setHeader(List<String> header) {
		this.header = header;
	}

	public boolean exsit(String code, List<String> codesList) {
		for (String string : codesList) {
			if (string.equals(code))
				return true;
		}
		return false;
	}

	public List<String> getAllCodes() {
		List<String> codeList = new ArrayList<>();
		for (Object item : items) {
			codeList.add(((IndexSheetDataItem) item).getCode());
		}
		return codeList;
	}
}
