package cn.itcast.handler;

import cn.itcast.domain.ContractProductVo;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * 行解析器
 */
public class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

	private ContractProductVo vo = null;

	/**
	 * 开始解析某一行的时候，自动进行调用
	 * 参数 ： 行索引
	 */
	public void startRow(int i) {
		if(i>=2) {
			vo = new ContractProductVo();
		}
	}

	/**
	 * 完成解析某一行的时候，自动进行调用
	 * 参数：行索引
	 * 目的：在解析完成某一行的时候，完成业务逻辑
	 */
	public void endRow(int i) {
		System.out.println("解析完成第"+i+"行数据："+vo);
	}

	/**
	 * 开始行中每一个单元格到时候，自动调用的方法
	 *      cellname : 单元格名称（A3,H23,B2）
	 *      cellvalue ：单元格数据
	 *
	 */
	public void cell(String cellname, String cellvalue, XSSFComment xssfComment) {

		if(vo != null) {
			cellname = cellname.substring(0,1);
			if("B".equals(cellname)) {
				vo.setCustomName(cellvalue);
			}else if("C".equals(cellname)) {
				vo.setContractNo(cellvalue);
			}else if("D".equals(cellname)) {
				vo.setProductNo(cellvalue);
			}
		}
	}
}
