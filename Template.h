#pragma once
#include <OpenXLSX.hpp>
#include <string>
using namespace std;
using namespace OpenXLSX;
class Template
{
	
protected:
	
	XLWorksheet dcnvs, wks;
	int posF, PosT;
	string sku;
	
public:

	void hijo(int posF, int posT, XLWorksheet wks, XLWorksheet dcnvs, string sku);
	void padre(int posF, int posT, XLWorksheet wks, XLWorksheet dcnvs, string sku, string tituloPadre);
};

