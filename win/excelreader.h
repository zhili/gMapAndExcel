#pragma once

// #include <standard library headers>
#include <string>
#include <list>

// #include <other library headers>
#include <ole2.h>

namespace msexcel {

	class excelreader {
	 /**
     * description
     */
    class ExcelException : public std::exception
    {
    public:
        ExcelException(const std::string& s = "") : msg(s) {}
        virtual ~ExcelException() throw() {}

    public:
        virtual const char* what() const throw() { return msg.c_str(); }
        virtual const std::string What() const throw() { return msg; }

    private:
        std::string msg;

    };

	public:
		excelreader(const std::wstring& filename = L"", const std::wstring& sheet = L"", bool visible = false);
		virtual ~excelreader();

	public:
		int rowCount();
		void SetData(const std::wstring& cell,
						 const std::wstring& data);     // set a range
		void SetData(int nRow, int nBeg,
						 const std::list<std::wstring>& data); // set a line
		void SetData(const std::wstring& colTag, int nBeg,
						 const std::list<std::wstring>& data); // set a column
		void SaveAsFile(const std::wstring& filename);
		double excelreader::GetDataAsDouble(int row, int col);
		std::string excelreader::GetDataAsString(int row, int col);
	private:
		void CreateNewWorkbook();
		void OpenWorkbook(const std::wstring& filename, const std::wstring& sheet);
		void GetActiveSheet();
		void GetSheet(const std::wstring& sheetname);
		std::wstring GetColumnName(int nIndex);
		bool CheckFilename(const std::wstring& filename);
		void InstanceExcelWorkbook(int visible);

		static void excelreader::AutoWrap(int autoType, VARIANT *pvResult,
										IDispatch *pDisp, LPOLESTR ptName,
										int cArgs...);

	private:
		IDispatch* pXlApp;
		IDispatch* pXlBooks;
		IDispatch* pXlBook;
		IDispatch* pXlSheet;

	};
}