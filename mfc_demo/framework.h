// header.h: 标准系统包含文件的包含文件，
// 或特定于项目的包含文件
//


#pragma once

// MFC 头文件 - 必须在其他头文件之前包含
#include <afxwin.h>         // MFC 核心组件和标准组件
#include <afxext.h>         // MFC 扩展
#include <afxdtctl.h>       // MFC 对 Internet Explorer 4 公共控件的支持
#ifndef _AFX_NO_OLE_SUPPORT
#include <afxole.h>         // MFC OLE 支持
#include <afxodlgs.h>       // MFC OLE 对话框支持
#endif // _AFX_NO_OLE_SUPPORT
#ifndef _AFX_NO_DB_SUPPORT
#include <afxdb.h>          // MFC ODBC 数据库支持
#endif // _AFX_NO_DB_SUPPORT

#include "targetver.h"
#define WIN32_LEAN_AND_MEAN             // 从 Windows 头文件中排除极少使用的内容
// C 运行时头文件
#include <stdlib.h>
#include <malloc.h>
#include <memory.h>
#include <tchar.h>
