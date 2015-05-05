#ifndef DATA_STRUCT_H
#define DATA_STRUCT_H
#include <vector>

struct ST_DeviceType
{
	int	 iID;//序号
	int	 iType;//设备类型
	char cDeviceName[64];//名字
	char cDescribe[256];//描述
};

typedef  std::vector<ST_DeviceType> DeviceTypeList;
#endif