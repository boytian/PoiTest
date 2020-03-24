import cn.itcast.utils.ExcelParse;

//百万数据解析
public class PoiTest03 {

	//文件放在了resources
	public static void main(String[] args) throws  Exception{
		new ExcelParse().parse("G:\\java\\ssm\\ssm项目\\Saas\\day12\\出货表 (2).xlsx");
	}
}
