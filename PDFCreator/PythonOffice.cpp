//#include "PythonOffice.h"
//#include<QDebug>
//
//namespace PythonOffice
//{
//
//	static PyObject* pythonModule = NULL;
//	static PyObject* func_ppt = NULL;
//	static PyObject* func_excel = NULL;
//	static PyObject* func_word = NULL;
//
//	bool ppt_to_pdf(QString sourceFile, QString destinationFile, QString& errorString)
//	{
//		if (!PythonOffice::init(errorString))
//		{
//			return false;
//		}
//
//		/*PyObject* func_ppt = PyObject_GetAttrString(PythonOffice::pythonModule, "PPT_To_PDF");
//		if (!func_ppt || !PyCallable_Check(func_ppt))
//		{
//			func_ppt = NULL;
//			errorString = QString("核心转化函数加载出错");
//
//			return false;
//		}*/
//
//		// 运行函数
//		PyObject* return_value = PyObject_CallFunction(PythonOffice::func_ppt, "ss", sourceFile.toStdString().c_str(), destinationFile.toStdString().c_str());
//		if (!return_value)
//		{
//			errorString = QString("函数执行出错");
//
//			return false;
//		}
//
//		char error_string[256];
//		int a = 0;
//
//		PyArg_ParseTuple(return_value, "i,s", &a, error_string);
//		errorString = QString::fromUtf8(error_string);
//
//		if (return_value)
//		{
//			Py_DECREF(return_value);
//		}
//		if (func_ppt)
//		{
//			Py_DECREF(func_ppt);
//		}
//		return a != 0 ? true : false;
//	}
//
//
//	bool word_to_pdf(QString sourceFile, QString destinationFile, QString& errorString)
//	{
//		if (!PythonOffice::init(errorString))
//		{
//			return false;
//		}
//
//		/*PyObject* func_word = PyObject_GetAttrString(PythonOffice::pythonModule, "Word_To_PDF");
//		if (!func_word || !PyCallable_Check(func_word))
//		{
//			func_word = NULL;
//			errorString = QString("核心转化函数加载出错");
//
//			return false;
//		}*/
//
//		// 运行函数
//		PyObject* return_value = PyObject_CallFunction(PythonOffice::func_word, "ss", sourceFile.toStdString().c_str(), destinationFile.toStdString().c_str());
//		if (!return_value)
//		{
//			errorString = QString("函数执行出错");
//			if (func_word)
//			{
//				Py_DECREF(func_word);
//			}
//			return false;
//		}
//
//		char error_string[256];
//		int a = 0;
//
//		PyArg_ParseTuple(return_value, "is", &a, error_string);
//		errorString = QString::fromUtf8(error_string);
//
//		if (return_value)
//		{
//			Py_DECREF(return_value);
//		}
//		if (func_word)
//		{
//			Py_DECREF(func_word);
//		}
//
//		return a != 0 ? true : false;
//	}
//
//	bool excel_to_pdf(QString sourceFile, QString destinationFile, QString& errorString)
//	{
//		if (!PythonOffice::init(errorString))
//		{
//			return false;
//		}
//
//		/*PyObject* func_excel = PyObject_GetAttrString(PythonOffice::pythonModule, "Excel_To_PDF");
//		if (!func_excel || !PyCallable_Check(func_excel))
//		{
//			func_excel = NULL;
//			errorString = QString("核心转化函数加载出错");
//			return false;
//		}*/
//
//		// 运行函数
//		PyObject* return_value = PyObject_CallFunction(PythonOffice::func_excel, "ss", sourceFile.toStdString().c_str(), destinationFile.toStdString().c_str());
//		if (!return_value)
//		{
//			errorString = QString("函数执行出错");
//			return false;
//		}
//
//		char error_string[256];
//		int a = 0;
//
//		PyArg_ParseTuple(return_value, "i,s", &a, error_string);
//		errorString = QString::fromUtf8(error_string);
//
//		if (return_value)
//		{
//			Py_DECREF(return_value);
//		}
//
//		return a != 0 ? true : false;
//	}
//
//
//	bool init(QString& errorString)
//	{
//		//初始化python模块
//		if (!Py_IsInitialized())
//		{
//			Py_Initialize();
//			if (!Py_IsInitialized())
//			{
//				errorString = QString("python模块初始化出错");
//				return false;
//			}
//		}
//
//		//导入python模块
//		PyRun_SimpleString("import win32com.client");
//		PyRun_SimpleString("sys.path.append('./')");
//
//		
//		if (!PythonOffice::pythonModule)
//		{
//			PythonOffice::pythonModule = PyImport_ImportModule("office_translator");
//			if (!PythonOffice::pythonModule)
//			{
//				errorString = QString("python模块加载出错");
//				return false;
//			}
//		}
//
//		if (!PythonOffice::func_ppt)
//		{
//			PythonOffice::func_ppt = PyObject_GetAttrString(PythonOffice::pythonModule, "PPT_To_PDF");
//			if (!func_ppt || !PyCallable_Check(func_ppt))
//			{
//				PythonOffice::func_ppt = NULL;
//				errorString = QString("核心转化函数加载出错");
//
//				return false;
//			}
//		}
//
//		if (!PythonOffice::func_word)
//		{
//			PythonOffice::func_word = PyObject_GetAttrString(PythonOffice::pythonModule, "Word_To_PDF");
//			if (!func_word || !PyCallable_Check(func_word))
//			{
//				PythonOffice::func_word = NULL;
//				errorString = QString("核心转化函数加载出错");
//
//				return false;
//			}
//		}
//
//		if (!PythonOffice::func_excel)
//		{
//			PythonOffice::func_excel = PyObject_GetAttrString(PythonOffice::pythonModule, "Excel_To_PDF");
//			if (!func_excel || !PyCallable_Check(func_excel))
//			{
//				PythonOffice::func_excel = NULL;
//				errorString = QString("核心转化函数加载出错");
//				return false;
//			}
//		}
//
//		return true;
//	}
//
//	void free()
//	{
//		if (PythonOffice::pythonModule)
//		{
//			Py_DECREF(PythonOffice::pythonModule);
//		}
//		if (PythonOffice::func_ppt)
//		{
//			Py_DECREF(PythonOffice::func_ppt);
//		}
//		if (PythonOffice::func_excel)
//		{
//			Py_DECREF(func_excel);
//		}
//		if (PythonOffice::func_word)
//		{
//			Py_DECREF(PythonOffice::func_word);
//		}
//		Py_Finalize();
//	}
//}