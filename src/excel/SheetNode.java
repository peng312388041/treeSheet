package excel;

import java.util.ArrayList;
import java.util.List;

public class SheetNode {
	int number; // 鐖剁粨鐐圭殑绗嚑涓瀛愮粨鐐�
	String id; // 琛╥d
	String name; // 琛ㄥ悕
	SheetData data = new SheetData(); // 琛ㄦ暟鎹�
	List<SheetNode> childList = new ArrayList<SheetNode>(); // 鎵�鏈夊瓙缁撶偣
	SheetNode parent; // 鐖剁粨鐐�
	boolean leaf; // 鏄惁涓哄晢鍝佺粨鐐�
	int level; // 瀛樻斁鏄鍑犲眰

	public int getNumber() {
		return number;
	}

	public void setNumber(int number) {
		this.number = number;
	}

	public int getLevel() {
		return level;
	}

	public void setLevel(int level) {
		this.level = level;
	}

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public SheetData getData() {
		return data;
	}

	public void setData(SheetData data) {
		this.data = data;
	}

	public List<SheetNode> getChildList() {
		return childList;
	}

	public void setChildList(List<SheetNode> childList) {
		this.childList = childList;
	}

	public SheetNode getParent() {
		return parent;
	}

	public void setParent(SheetNode parent) {
		this.parent = parent;
	}

	public boolean isLeaf() {
		return leaf;
	}

	public void setLeaf(boolean leaf) {
		this.leaf = leaf;
	}
}
