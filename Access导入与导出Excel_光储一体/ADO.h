// ADO.h: interface for the ADO class.
//
//////////////////////////////////////////////////////////////////////

//功能：将ADO的相关操作封装成一个类

#if !defined(AFX_ADO_H__E57AA002_2E56_4088_96D8_AA98EE73DAA3__INCLUDED_)
#define AFX_ADO_H__E57AA002_2E56_4088_96D8_AA98EE73DAA3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class ADO  
{

public:
	ADO();
	virtual ~ADO();
	
	UINT GetRecordsetCount(_RecordsetPtr pRecordset);			//返回记录集个数
	void CloseConn();											//关闭连接
	void CloseRecordset();										//关闭记录集
	_RecordsetPtr& OpenRecordset(CString sql);					//打开记录集
	void OnInitADOConn();										//初始化COM环境
	_RecordsetPtr m_pRecordset;									//智能指针
	_ConnectionPtr m_pConnection;								//智能指针

};

#endif // !defined(AFX_ADO_H__E57AA002_2E56_4088_96D8_AA98EE73DAA3__INCLUDED_)
