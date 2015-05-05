// ADO.h: interface for the ADO class.
//
//////////////////////////////////////////////////////////////////////

//���ܣ���ADO����ز�����װ��һ����

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
	
	UINT GetRecordsetCount(_RecordsetPtr pRecordset);			//���ؼ�¼������
	void CloseConn();											//�ر�����
	void CloseRecordset();										//�رռ�¼��
	_RecordsetPtr& OpenRecordset(CString sql);					//�򿪼�¼��
	void OnInitADOConn();										//��ʼ��COM����
	_RecordsetPtr m_pRecordset;									//����ָ��
	_ConnectionPtr m_pConnection;								//����ָ��

};

#endif // !defined(AFX_ADO_H__E57AA002_2E56_4088_96D8_AA98EE73DAA3__INCLUDED_)
