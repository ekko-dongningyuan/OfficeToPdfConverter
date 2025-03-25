#pragma once

namespace OfficeHelper
{
	bool ConvertWordToPDF(const std::wstring& wordFilePath, const std::wstring& pdfFilePath);
	bool ConvertExcelToPDF(const std::wstring& excelFilePath, const std::wstring& pdfFilePath);
}
