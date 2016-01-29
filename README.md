# Android_Excel
在android中生成excel
##效果图
![这里写图片描述](http://img.blog.csdn.net/20160129095143462)
##初始化数据
首先我们要先造下测试数据，这里我把数据写死在一个常量类Const中，如下：

```
public class Const {
	public interface OrderInfo{
		public static final String[][] orderOne = new String[][] {{ "123", "九龙", "13294352311",
			"武汉市关山口" },{ "124", "咱家", "13294352312",
			"武汉市水果湖" },{ "125", "陈家", "13294352315",
			"武汉市华师" },{ "126", "李", "13294352316",
			"武汉市杨家湾" }};
	}
}
```
理论上这些数据是从后台读过来的。
本文模拟打印订单的信息，所以这里还需要一个订单Model类：

```
public class Order implements Serializable {

	public String id;

	public String restPhone;

	public String restName;

	public String receiverAddr;

	
	public Order(String id,String restPhone, String restName, String receiverAddr) {
		this.id = id;
		this.restPhone = restPhone;
		this.restName = restName;
		
		this.receiverAddr = receiverAddr;
	}
}
```

##存内存卡
接下来我们要判断一下内存卡是否存在，内存是否足够大。先获取指定目录下内存的大小：

```
	/** 获取SD可用容量 */
	private static long getAvailableStorage(Context context) {
		String root = context.getExternalFilesDir(null).getPath();
		StatFs statFs = new StatFs(root);
		long blockSize = statFs.getBlockSize();
		long availableBlocks = statFs.getAvailableBlocks();
		long availableSize = blockSize * availableBlocks;
		// Formatter.formatFileSize(context, availableSize);
		return availableSize;
	}
```
这里用到的路径是getExternalFilesDir，它指定的是SDCard/Android/data/你的应用的包名/files/ 目录这个目录，它用来放一些长时间保存的数据，当应用被卸载时，会同时会删除。类似这种情况的还有getExternalCacheDir方法，只是它一般用来存放临时文件。之后通过StatFs来计算出可用容量的大小。
接下来在写入excel前对内存进行判断，如下：

```
if(!Environment.getExternalStorageState().equals(Environment.MEDIA_MOUNTED)&&getAvailableStorage()>1000000) {
			Toast.makeText(context, "SD卡不可用", Toast.LENGTH_LONG).show();
			return;
		}
		File file;
		File dir = new File(context.getExternalFilesDir(null).getPath());
		file = new File(dir, fileName + ".xls");
		if (!dir.exists()) {
			dir.mkdirs();
		}
```
如果内存卡不存在或内存小于1M，不进行写入，然后创建相应的文件夹并起名字。接下来重点看下如何写入excel。
##生成写入excel

 - 导入相关包
这里需要导入jxl包，它主要就是用于处理excel的，这个包我会附在项目放在github中，后面会给出链接。
 - 生成excel工作表
 以下代码是在指定路径下生成excel表，此时只是一个空表。
 

```
WritableWorkbook wwb;
		OutputStream os = new FileOutputStream(file);
		wwb = Workbook.createWorkbook(os);
```

 - 添加sheet表
 熟悉excel操作的都知道excel可以新建很多个sheet表。以下代码生成第一个工作表，名字为“订单”：
 

```
WritableSheet sheet = wwb.createSheet("订单", 0);
```

 - 添加excel表头
 添加excel的表头，这里可以自定义表头的样式，先看代码：
 

```
String[] title = { "订单", "店名", "电话", "地址" };
Label label;
		for (int i = 0; i < title.length; i++) {
			// Label(x,y,z) 代表单元格的第x+1列，第y+1行, 内容z
						// 在Label对象的子对象中指明单元格的位置和内容
			label = new Label(i, 0, title[i], getHeader());
			// 将定义好的单元格添加到工作表中
			sheet.addCell(label);
		}
```
这里表头信息我写死了。表的一个单元格对应一个Label，如label(0,0,"a")代表第一行第一列所在的单元格信息为a。getHeader()是自定义的样式，它返回一个 WritableCellFormat 。来看看如何自定义样式：
 

```
	public static WritableCellFormat getHeader() {
		WritableFont font = new WritableFont(WritableFont.TIMES, 10,
				WritableFont.BOLD);// 定义字体
		try {
			font.setColour(Colour.BLUE);// 蓝色字体
		} catch (WriteException e1) {
			e1.printStackTrace();
		}
		WritableCellFormat format = new WritableCellFormat(font);
		try {
			format.setAlignment(jxl.format.Alignment.CENTRE);// 左右居中
			format.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);// 上下居中
			 format.setBorder(Border.ALL, BorderLineStyle.THIN,
			 Colour.BLACK);// 黑色边框
			 format.setBackground(Colour.YELLOW);// 黄色背景
		} catch (WriteException e) {
			e.printStackTrace();
		}
		return format;
	}
```
看上面代码就很清楚了，通过获得WritableFont 来自定义字体的一些样式，如颜色大小等，通过WritableCellFormat 来设置文本框的样式，可以设置边框底色等。具体的可以查api文档，这里只给出例子。

 - 添加excel内容。


```
for (int i = 0; i < exportOrder.size(); i++) {
			Order order = exportOrder.get(i);

			Label orderNum = new Label(0, i + 1, order.id);
			Label restaurant = new Label(1, i + 1, order.restName);
			Label nameLabel = new Label(2,i+1,order.restPhone);
			Label address = new Label(3, i + 1, order.receiverAddr);
			
			sheet.addCell(orderNum);
			sheet.addCell(restaurant);
			sheet.addCell(nameLabel);
			sheet.addCell(address);
			Toast.makeText(context, "写入成功", Toast.LENGTH_LONG).show();
			
		}
```
