package excel;

import java.util.Comparator;

public class DataComparator implements Comparator<IndexSheetDataItem> {

	@Override
	public int compare(IndexSheetDataItem arg0, IndexSheetDataItem arg1) {
		return arg0.getCode().compareTo(arg1.getCode());
	}

}
