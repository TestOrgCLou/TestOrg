#ifndef DATA_STRUCT_H
#define DATA_STRUCT_H
#include <vector>

struct ST_DeviceType
{
	int	 iID;//���
	int	 iType;//�豸����
	char cDeviceName[64];//����
	char cDescribe[256];//����
};

typedef  std::vector<ST_DeviceType> DeviceTypeList;
#endif