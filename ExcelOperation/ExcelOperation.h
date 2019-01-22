#pragma once

#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorksheets.h"
#include "CWorkbook.h"
#include "CWorksheet.h"
#include "CRange.h"
#include "CPageSetup.h"
#include "CFont0.h"

//该类使用mfc类型库，需在项目中自行初始化ole
//当前支持2010,、2013版本excel
//若使用2007版本，安装pdf插件，且需修改代码版本检查部分
class ExcelOperation
{
public:
	ExcelOperation(void);
	~ExcelOperation(void);
	
	//************************************************************************
	// Method:		CheckVersion	检查excel版本，当前支持2010(14.0),2013(15.0)
	// Returns:		bool
	// Author:		byshi
	// Date:		2018-12-11	
	//************************************************************************
	bool CheckVersion();
	
	//************************************************************************
	// Method:		Close
	// Returns:		void
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	void Close();
	
	//************************************************************************
	// Method:		CreateTwoDimSafeArray	创建二维安全数组
	// Returns:		bool
	// Parameter:	COleSafeArray & safeArray	数组名
	// Parameter:	DWORD dimElements1	一维的元素数量
	// Parameter:	DWORD dimElements2	二维的元素数量
	// Parameter:	long iLBound1	一维的下边界，默认值:1
	// Parameter:	long iLBound2	二维的下边界，默认值:1
	// Author:		byshi
	// Date:		2018-12-10	
	//************************************************************************
	bool CreateTwoDimSafeArray(COleSafeArray &safeArray, DWORD dimElements1, 
		DWORD dimElements2, long iLBound1 = 1, long iLBound2 = 1);
	
	//************************************************************************
	// Method:		GetCellsValue	获取批量单元格数据
	// Returns:		bool
	// Parameter:	COleVariant startCell	起始单元格
	// Parameter:	COleVariant endCell		结束单元格
	// Parameter:	VARIANT & iData		返回获取到的数据值
	// Author:		byshi
	// Date:		2018-12-10	
	//************************************************************************
	bool GetCellsValue(COleVariant startCell,COleVariant endCell,
		VARIANT &iData);
	//************************************************************************
	// Method:		GetCellValue	获取单元格数据
	// Returns:		bool
	// Parameter:	COleVariant rowIndex		单元格行号,("2",或数字2)
	// Parameter:	COleVariant columnIndex		单元格列号,("C",或数字3)
	// Parameter:	VARIANT & data		返回获取到的数据值
	// Author:		byshi
	// Date:		2018-12-10	
	//************************************************************************
	bool GetCellValue(COleVariant rowIndex, COleVariant columnIndex,VARIANT &data);
	
	//************************************************************************
	// Method:		IsFileExist		判断文件或文件夹是否存在
	// Returns:		bool
	// Parameter:	const CString & fileName	文件路径名
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool IsFileExist(const CString &fileName);
	
	//************************************************************************
	// Method:		IsRegionEqual	判断单元格区域是否与二维数组区域大小相同
	// Returns:		bool
	// Parameter:	CRange & iRange		单元格区域
	// Parameter:	COleSafeArray & iTwoDimArray	二维数组
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool IsRegionEqual(CRange &iRange,COleSafeArray &iTwoDimArray);
	
	//************************************************************************
	// Method:		Open	打开文件
	// Returns:		bool
	// Parameter:	const CString & iFileName	文件路径名
	// Parameter:	bool autoCreate		打开模式，默认文件不存在不自动创建
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool Open(const CString &iFileName, bool autoCreate = false);
	
	//************************************************************************
	// Method:		OpenFromTemplate	根据模板创建新文件
	// Returns:		bool
	// Parameter:	const CString & iFileName	模板文件路径名	
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool OpenFromTemplate(const CString &iFileName);
	
	//************************************************************************
	// Method:		Save	保存文件，OpenFromTemplate创建的文件不能使用该方法保存
	// Returns:		void
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool Save();
	
	//************************************************************************
	// Method:		SaveAs	另存为
	// Returns:		bool
	// Parameter:	const CString & iFileName	文件路径名
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SaveAs(const CString &iFileName);
	
	//************************************************************************
	// Method:		SaveAsPDF	另存为PDF文件
	// Returns:		bool
	// Parameter:	const CString & iFileName	文件路径名
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SaveAsPDF(const CString &iFileName);
	
	//************************************************************************
	// Method:		SetCellsValue	批量写入多个单元格数据
	// Returns:		bool
	// Parameter:	COleVariant startCell	起始单元格(左上角单元格),例如:"B5"或"$B$5"
	// Parameter:	COleVariant endCell		起始单元格(左上角单元格),例如:"C7"或"$C$7"
	// Parameter:	COleSafeArray & iTwoDimArray	二维数组(要写入的数据)
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SetCellsValue(COleVariant startCell,COleVariant endCell,
		COleSafeArray &iTwoDimArray);
	
	//************************************************************************
	// Method:		SetCellValue	写入单个单元格数据
	// Returns:		bool
	// Parameter:	COleVariant rowIndex		单元格行号,例如:"8"或数字表示
	// Parameter:	COleVariant columnIndex		单元格列号,例如:"C"或数字表示
	// Parameter:	COleVariant data			要写入的数据
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SetCellValue(COleVariant rowIndex, COleVariant columnIndex,COleVariant data);
	
	//************************************************************************
	// Method:		SwitchWorksheet		切换sheet
	// Returns:		bool
	// Parameter:	const CString & sheetName	sheet名称
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SwitchWorksheet(const CString &sheetName);

	void Test1();

public://not use
	//此处暂时不用，若需要的功能该类未封装，可使用以下函数获取book、sheet自定义操作
	
	//************************************************************************
	// Method:		GetRange	获取CRange对象的指针
	// Returns:		CRange *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CRange *GetRange();
	
	//************************************************************************
	// Method:		GetWorkbook		获取CWorkbook对象的指针
	// Returns:		CWorkbook *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorkbook *GetWorkbook();
	
	//************************************************************************
	// Method:		GetWorkbooks	获取CWorkbooks对象的指针
	// Returns:		CWorkbooks *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorkbooks *GetWorkbooks();
	
	//************************************************************************
	// Method:		GetWorksheet	获取CWorksheet对象的指针
	// Returns:		CWorksheet *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorksheet *GetWorksheet();
	
	//************************************************************************
	// Method:		GetWorksheets	获取CWorksheets对象的指针
	// Returns:		CWorksheets *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorksheets *GetWorksheets();

private:
	CApplication	m_app;
	CRange			m_range;
	CWorkbook		m_workbook;
	CWorkbooks		m_workbooks;
	CWorksheet		m_worksheet;
	CWorksheets		m_worksheets;

	COleVariant		covTrue;
	COleVariant		covFalse;
	COleVariant		covOptional;
};
