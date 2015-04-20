package excel;

import java.util.Comparator;

public class SheetComparator implements Comparator<SheetNode> {

	@Override
	public int compare(SheetNode arg0, SheetNode arg1) {
		return arg0.getId().compareTo(arg1.getId());
	}

}
