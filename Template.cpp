#include "Template.h";

void Template::hijo (int posF, int posT, XLWorksheet wks, XLWorksheet dcnvs, string sku) {
	
	wks.cell("A" + std::to_string(posT)).value() = "furnitureanddecor";
	wks.cell("B" + std::to_string(posT)).value() = dcnvs.cell("H" + to_string(posF)).value();
	wks.cell("C" + std::to_string(posT)).value() = "D'CANVAS";
	wks.cell("D" + std::to_string(posT)).value() = dcnvs.cell("I" + to_string(posF)).value();
	wks.cell("E" + std::to_string(posT)).value() = "EAN";
	wks.cell("F" + std::to_string(posT)).value() = dcnvs.cell("K" + to_string(posF)).value();
	wks.cell("G" + std::to_string(posT)).value() = "D'CANVAS";
	wks.cell("H" + std::to_string(posT)).value() = 9757416011;
	wks.cell("I" + std::to_string(posT)).value() = dcnvs.cell("J" + to_string(posF)).value();
	wks.cell("J" + std::to_string(posT)).value() = 100;
	wks.cell("K" + std::to_string(posT)).value() = 100;
	wks.cell("L" + std::to_string(posT)).value() = dcnvs.cell("M" + to_string(posF)).value();
	wks.cell("M" + std::to_string(posT)).value() = dcnvs.cell("M" + to_string(posF)).value();
	wks.cell("N" + std::to_string(posT)).value() = dcnvs.cell("N" + to_string(posF)).value();
	wks.cell("O" + std::to_string(posT)).value() = dcnvs.cell("O" + to_string(posF)).value();
	wks.cell("P" + std::to_string(posT)).value() = dcnvs.cell("P" + to_string(posF)).value();
	wks.cell("Q" + std::to_string(posT)).value() = dcnvs.cell("Q" + to_string(posF)).value();
	wks.cell("R" + std::to_string(posT)).value() = dcnvs.cell("R" + to_string(posF)).value();
	wks.cell("S" + std::to_string(posT)).value() = dcnvs.cell("S" + to_string(posF)).value();
	wks.cell("T" + std::to_string(posT)).value() = dcnvs.cell("T" + to_string(posF)).value();
	wks.cell("U" + std::to_string(posT)).value() = dcnvs.cell("U" + to_string(posF)).value();
	wks.cell("V" + std::to_string(posT)).value() = dcnvs.cell("AE6").value();
	wks.cell("W" + std::to_string(posT)).value() = sku;
	wks.cell("X" + std::to_string(posT)).value() = dcnvs.cell("AE1").value();
	wks.cell("Y" + std::to_string(posT)).value() = "Size";
	wks.cell("AA" + std::to_string(posT)).value() = dcnvs.cell("AE2").value();
	wks.cell("AE" + std::to_string(posT)).value() = dcnvs.cell("AF2").value();
	wks.cell("AF" + std::to_string(posT)).value() = dcnvs.cell("AG2").value();
	wks.cell("AG" + std::to_string(posT)).value() = dcnvs.cell("AH2").value();
	wks.cell("BC" + std::to_string(posT)).value() = dcnvs.cell("AI" + to_string(posF)).value();
	wks.cell("BQ" + std::to_string(posT)).value() = dcnvs.cell("AI" + to_string(posF)).value();
	wks.cell("BZ" + std::to_string(posT)).value() = dcnvs.cell("AD" + to_string(posF)).value();
	wks.cell("CB" + std::to_string(posT)).value() = dcnvs.cell("D" + to_string(posF)).value();
	wks.cell("DM" + std::to_string(posT)).value() = dcnvs.cell("F" + to_string(posF)).value();
	wks.cell("DO" + std::to_string(posT)).value() = dcnvs.cell("E" + to_string(posF)).value();
	wks.cell("DU" + std::to_string(posT)).value() = dcnvs.cell("E" + to_string(posF)).value();
	wks.cell("DV" + std::to_string(posT)).value() = 3;
	wks.cell("DW" + std::to_string(posT)).value() = dcnvs.cell("F" + to_string(posF)).value();
	wks.cell("DZ" + std::to_string(posT)).value() = dcnvs.cell("AE3").value();
	wks.cell("DZ" + std::to_string(posT)).value() = dcnvs.cell("AE3").value();
	wks.cell("EA" + std::to_string(posT)).value() = dcnvs.cell("AE3").value();
	wks.cell("EG" + std::to_string(posT)).value() = dcnvs.cell("AE3").value();
	wks.cell("EQ" + std::to_string(posT)).value() = dcnvs.cell("AE4").value();
	wks.cell("EW" + std::to_string(posT)).value() = "Tela Canvas";
	wks.cell("EX" + std::to_string(posT)).value() = "Madera";
	wks.cell("FE" + std::to_string(posT)).value() = dcnvs.cell("AE5").value();
}
void Template::padre (int posF, int posT, XLWorksheet wks, XLWorksheet dcnvs, string sku, string tituloPadre) {
	wks.cell("A" + std::to_string(posT)).value() = "furnitureanddecor";//tipo de producto
	wks.cell("B" + std::to_string(posT)).value() = sku;                //sku
	wks.cell("C" + std::to_string(posT)).value() = "D'CANVAS";         //brand
	wks.cell("F" + std::to_string(posT)).value() = tituloPadre;        //nombre del producto
	wks.cell("G" + std::to_string(posT)).value() = "D'CANVAS";         //Fabricante
	wks.cell("H" + std::to_string(posT)).value() = 9757416011;         //
	wks.cell("L" + std::to_string(posT)).value() = dcnvs.cell("M" + to_string(posF)).value();
	wks.cell("M" + std::to_string(posT)).value() = dcnvs.cell("M" + to_string(posF)).value();
	wks.cell("N" + std::to_string(posT)).value() = dcnvs.cell("N" + to_string(posF)).value();
	wks.cell("O" + std::to_string(posT)).value() = dcnvs.cell("O" + to_string(posF)).value();
	wks.cell("P" + std::to_string(posT)).value() = dcnvs.cell("P" + to_string(posF)).value();
	wks.cell("Q" + std::to_string(posT)).value() = dcnvs.cell("Q" + to_string(posF)).value();
	wks.cell("R" + std::to_string(posT)).value() = dcnvs.cell("R" + to_string(posF)).value();
	wks.cell("S" + std::to_string(posT)).value() = dcnvs.cell("S" + to_string(posF)).value();
	wks.cell("T" + std::to_string(posT)).value() = dcnvs.cell("T" + to_string(posF)).value();
	wks.cell("U" + std::to_string(posT)).value() = dcnvs.cell("U" + to_string(posF)).value();
	wks.cell("V" + std::to_string(posT)).value() = "Padre";
	wks.cell("Y" + std::to_string(posT)).value() = "Size";
	wks.cell("AA" + std::to_string(posT)).value() = dcnvs.cell("AE2").value();
}