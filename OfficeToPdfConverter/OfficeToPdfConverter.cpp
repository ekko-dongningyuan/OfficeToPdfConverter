#include "pch.h"
#include <string>

#include "OfficeToPdfConverter.h"
#include "CWordApplication.h"    
#include "CWordDocument.h"    
#include "CWordDocuments.h"    
#include "CExcelApplication.h"   
#include "CExcelWorkbook.h"   
#include "CExcelWorkbooks.h"   

// Word to PDF Conversion Function
bool OfficeHelper::ConvertWordToPDF(const std::wstring& wordFilePath, const std::wstring& pdfFilePath)
{
	CoInitialize(NULL);
	bool bResult = false;

	try
	{
		// Create Word application
		CWordApplication wordApp;
		LPDISPATCH lpDisp = NULL;

		// Create the Word application
		if (!wordApp.CreateDispatch(_T("Word.Application")))
		{
			AfxMessageBox(_T("Failed to start Word application"));
			CoUninitialize();
			return false;
		}

		// Make Word invisible
		wordApp.put_Visible(FALSE);
		wordApp.put_DisplayAlerts(FALSE);

		// Get Documents collection
		CWordDocuments documents;
		lpDisp = wordApp.get_Documents();
		documents.AttachDispatch(lpDisp);

		// Open the Word document
		CWordDocument doc;
		VARIANT vtFileName, vtReadOnly, vtMissing;
		vtFileName.vt = VT_BSTR;
		vtFileName.bstrVal = CString(wordFilePath.c_str()).AllocSysString();
		vtReadOnly.vt = VT_BOOL;
		vtReadOnly.boolVal = TRUE;
		vtMissing.vt = VT_ERROR;
		vtMissing.scode = DISP_E_PARAMNOTFOUND;

		lpDisp = documents.Open(&vtFileName, &vtMissing, &vtReadOnly
			, &vtMissing
			, &vtMissing
			, &vtMissing
			, &vtMissing
			, &vtMissing
			, &vtMissing
			, &vtMissing
			, &vtMissing
			, &vtMissing
			, &vtMissing
			, &vtMissing
			, &vtMissing
			, &vtMissing
		);
		doc.AttachDispatch(lpDisp);

		// Save as PDF
		VARIANT vtSaveFormat, vtPDFName;
		vtPDFName.vt = VT_BSTR;
		vtPDFName.bstrVal = CString(pdfFilePath.c_str()).AllocSysString();
		vtSaveFormat.vt = VT_I4;  // Changed from VT_INT to VT_I4
		vtSaveFormat.lVal = 17;   // Changed from intVal to lVal for VT_I4
		doc.SaveAs2(&vtPDFName, &vtSaveFormat,
			&vtMissing, &vtMissing, &vtMissing, &vtMissing,
			&vtMissing, &vtMissing, &vtMissing, &vtMissing,
			&vtMissing, &vtMissing, &vtMissing, &vtMissing,
			&vtMissing, &vtMissing, &vtMissing
		);

		// Close document
		VARIANT vtSaveChanges;
		vtSaveChanges.vt = VT_BOOL;
		vtSaveChanges.boolVal = FALSE;
		doc.Close(&vtSaveChanges
			, &vtMissing
			, &vtMissing
		);

		// Quit Word
		wordApp.Quit(&vtMissing
			, &vtMissing
			, &vtMissing
		);

		// Clean up
		doc.ReleaseDispatch();
		documents.ReleaseDispatch();
		wordApp.ReleaseDispatch();

		SysFreeString(vtFileName.bstrVal);
		SysFreeString(vtPDFName.bstrVal);

		bResult = true;
	}
	catch (COleException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		AfxMessageBox(szError);
		e->Delete();
	}
	catch (...)
	{
		AfxMessageBox(_T("Unknown exception occurred during Word to PDF conversion"));
	}

	CoUninitialize();
	return bResult;
}

// Excel to PDF Conversion Function
bool OfficeHelper::ConvertExcelToPDF(const std::wstring& excelFilePath, const std::wstring& pdfFilePath)
{
	CoInitialize(NULL);
	bool bResult = false;

	try
	{
		// Create Excel application
		CExcelApplication excelApp;
		LPDISPATCH lpDisp = NULL;

		if (!excelApp.CreateDispatch(_T("Excel.Application")))
		{
			AfxMessageBox(_T("Failed to start Excel application"));
			CoUninitialize();
			return false;
		}

		// Make Excel invisible
		excelApp.put_Visible(FALSE);
		excelApp.put_DisplayAlerts(FALSE);

		// Open the workbook
		VARIANT vtFileName, vtReadOnly, vtMissing;
		vtFileName.vt = VT_BSTR;
		vtFileName.bstrVal = CString(excelFilePath.c_str()).AllocSysString();
		vtReadOnly.vt = VT_BOOL;
		vtReadOnly.boolVal = TRUE;
		vtMissing.vt = VT_ERROR;
		vtMissing.scode = DISP_E_PARAMNOTFOUND;

		// Get Workbooks collection
		CExcelWorkbook workbook;
		lpDisp = excelApp.get_Workbooks();
		CExcelWorkbooks workbooks;
		workbooks.AttachDispatch(lpDisp);

		// Open workbook
		lpDisp = workbooks.Open(CString(excelFilePath.c_str()), vtMissing, vtReadOnly
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
		);
		workbook.AttachDispatch(lpDisp);

		// Save as PDF
		VARIANT vtPDFName, vtFormat, vtQuality;
		vtPDFName.vt = VT_BSTR;
		vtPDFName.bstrVal = CString(pdfFilePath.c_str()).AllocSysString();
		vtFormat.vt = VT_INT;
		vtFormat.intVal = 0; // xlTypePDF
		vtQuality.vt = VT_INT;
		vtQuality.intVal = 0; // xlQualityStandard

		// Export to PDF - equivalent to calling SaveAs with PDF format
		// Note: Depending on your Excel version, you might need the ExportAsFixedFormat method
		workbook.ExportAsFixedFormat(0, vtPDFName, vtQuality
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
			, vtMissing
		);

		// Close workbook
		VARIANT vtSaveChanges;
		vtSaveChanges.vt = VT_BOOL;
		vtSaveChanges.boolVal = FALSE;
		workbook.Close(vtSaveChanges
			, vtMissing
			, vtMissing
		);

		// Quit Excel
		excelApp.Quit();

		// Clean up
		workbook.ReleaseDispatch();
		workbooks.ReleaseDispatch();
		excelApp.ReleaseDispatch();

		SysFreeString(vtFileName.bstrVal);
		SysFreeString(vtPDFName.bstrVal);

		bResult = true;
	}
	catch (COleException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		AfxMessageBox(szError);
		e->Delete();
	}
	catch (...)
	{
		AfxMessageBox(_T("Unknown exception occurred during Excel to PDF conversion"));
	}

	CoUninitialize();
	return bResult;
}