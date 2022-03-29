#pragma once

#include <OpenXLSX.hpp>;
#include <msclr\marshal_cppstd.h>;
#include <OpenXLSX.hpp>;
#include "Template.h";

namespace Project1 {

	using namespace OpenXLSX;
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace msclr::interop;
	using namespace std;

	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			tmp = new Template();
			//
			//TODO: agregar código de constructor aquí
			//
		}
        Template *tmp ;

	
	protected:
		/// <summary>
		/// Limpiar los recursos que se estén usando.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^ button1;
	private: System::Windows::Forms::TextBox^ txtBoxNamePlantilla;
	private: System::Windows::Forms::TextBox^ txtBoxBaseDatos;
	private: System::Windows::Forms::ProgressBar^ progressBar1;
	private: System::Windows::Forms::NumericUpDown^ numericUpDownDesde;
	private: System::Windows::Forms::NumericUpDown^ numericUpDownHasta;


	private: System::Windows::Forms::Label^ label1;
	private: System::Windows::Forms::Label^ label2;
	protected:

	private:
		
		/// <summary>
		/// Variable del diseñador necesaria.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido de este método con el editor de código.
		/// </summary>
		void InitializeComponent(void)
		{
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(MyForm::typeid));
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->txtBoxNamePlantilla = (gcnew System::Windows::Forms::TextBox());
			this->txtBoxBaseDatos = (gcnew System::Windows::Forms::TextBox());
			this->progressBar1 = (gcnew System::Windows::Forms::ProgressBar());
			this->numericUpDownDesde = (gcnew System::Windows::Forms::NumericUpDown());
			this->numericUpDownHasta = (gcnew System::Windows::Forms::NumericUpDown());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->label2 = (gcnew System::Windows::Forms::Label());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericUpDownDesde))->BeginInit();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericUpDownHasta))->BeginInit();
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->BackColor = System::Drawing::SystemColors::Window;
			this->button1->Font = (gcnew System::Drawing::Font(L"Bahnschrift", 18, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button1->ForeColor = System::Drawing::Color::DarkGreen;
			this->button1->Location = System::Drawing::Point(22, 365);
			this->button1->Margin = System::Windows::Forms::Padding(4, 3, 4, 3);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(178, 49);
			this->button1->TabIndex = 0;
			this->button1->Text = L"Crear Excel";
			this->button1->UseVisualStyleBackColor = false;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// txtBoxNamePlantilla
			// 
			this->txtBoxNamePlantilla->Location = System::Drawing::Point(20, 104);
			this->txtBoxNamePlantilla->Margin = System::Windows::Forms::Padding(4, 3, 4, 3);
			this->txtBoxNamePlantilla->Name = L"txtBoxNamePlantilla";
			this->txtBoxNamePlantilla->Size = System::Drawing::Size(435, 23);
			this->txtBoxNamePlantilla->TabIndex = 1;
			this->txtBoxNamePlantilla->Text = L"Escribe el nombre de la platilla a generar";
			// 
			// txtBoxBaseDatos
			// 
			this->txtBoxBaseDatos->Location = System::Drawing::Point(22, 47);
			this->txtBoxBaseDatos->Margin = System::Windows::Forms::Padding(4, 3, 4, 3);
			this->txtBoxBaseDatos->Name = L"txtBoxBaseDatos";
			this->txtBoxBaseDatos->Size = System::Drawing::Size(433, 23);
			this->txtBoxBaseDatos->TabIndex = 2;
			this->txtBoxBaseDatos->Text = L"Escribe el nombre de tu archivo Base";
			// 
			// progressBar1
			// 
			this->progressBar1->ForeColor = System::Drawing::Color::Lime;
			this->progressBar1->Location = System::Drawing::Point(22, 454);
			this->progressBar1->Margin = System::Windows::Forms::Padding(4, 3, 4, 3);
			this->progressBar1->Name = L"progressBar1";
			this->progressBar1->Size = System::Drawing::Size(609, 29);
			this->progressBar1->Step = 1;
			this->progressBar1->Style = System::Windows::Forms::ProgressBarStyle::Continuous;
			this->progressBar1->TabIndex = 3;
			this->progressBar1->Visible = false;
			// 
			// numericUpDownDesde
			// 
			this->numericUpDownDesde->Location = System::Drawing::Point(343, 182);
			this->numericUpDownDesde->Margin = System::Windows::Forms::Padding(4, 3, 4, 3);
			this->numericUpDownDesde->Maximum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1048576, 0, 0, 0 });
			this->numericUpDownDesde->Name = L"numericUpDownDesde";
			this->numericUpDownDesde->Size = System::Drawing::Size(100, 23);
			this->numericUpDownDesde->TabIndex = 4;
			this->numericUpDownDesde->ValueChanged += gcnew System::EventHandler(this, &MyForm::numericUpDown1_ValueChanged);
			// 
			// numericUpDownHasta
			// 
			this->numericUpDownHasta->Location = System::Drawing::Point(343, 241);
			this->numericUpDownHasta->Margin = System::Windows::Forms::Padding(4, 3, 4, 3);
			this->numericUpDownHasta->Maximum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1048576, 0, 0, 0 });
			this->numericUpDownHasta->Name = L"numericUpDownHasta";
			this->numericUpDownHasta->Size = System::Drawing::Size(100, 23);
			this->numericUpDownHasta->TabIndex = 5;
			this->numericUpDownHasta->ValueChanged += gcnew System::EventHandler(this, &MyForm::numericUpDownHasta_ValueChanged);
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(62, 182);
			this->label1->Margin = System::Windows::Forms::Padding(4, 0, 4, 0);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(268, 16);
			this->label1->TabIndex = 6;
			this->label1->Text = L"Desde donde se va a iterar";
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Location = System::Drawing::Point(62, 250);
			this->label2->Margin = System::Windows::Forms::Padding(4, 0, 4, 0);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(268, 16);
			this->label2->TabIndex = 7;
			this->label2->Text = L"Hasta donde se va a iterar";
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(10, 16);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(0)), static_cast<System::Int32>(static_cast<System::Byte>(64)),
				static_cast<System::Int32>(static_cast<System::Byte>(64)));
			this->ClientSize = System::Drawing::Size(650, 504);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->numericUpDownHasta);
			this->Controls->Add(this->numericUpDownDesde);
			this->Controls->Add(this->progressBar1);
			this->Controls->Add(this->txtBoxBaseDatos);
			this->Controls->Add(this->txtBoxNamePlantilla);
			this->Controls->Add(this->button1);
			this->Font = (gcnew System::Drawing::Font(L"Lucida Console", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->ForeColor = System::Drawing::Color::LightYellow;
			this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedToolWindow;
			this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			this->Location = System::Drawing::Point(500, 500);
			this->Margin = System::Windows::Forms::Padding(4, 3, 4, 3);
			this->Name = L"MyForm";
			this->Text = L"Generador de plantilla";
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericUpDownDesde))->EndInit();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericUpDownHasta))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void MyForm_Load(System::Object^ sender, System::EventArgs^ e) {
	}


	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
		
		String^ name = txtBoxNamePlantilla->Text;
		string stdname = marshal_as<std::string>(name);
		String^ nameBase = txtBoxBaseDatos->Text;
		string stdnameBase = marshal_as<std::string>(nameBase);
		if (stdname == "excel.xlsx") {
			
			button1->Enabled = false;
			progressBar1->Visible = true;
			Decimal inicioD, finalD;
			inicioD = numericUpDownDesde->Value;
			finalD = numericUpDownHasta->Value;
			int inicio = System::Decimal::ToUInt16(inicioD);
			int final = System::Decimal::ToUInt16(finalD);
			XLDocument doc, doc2;//manejadores docs .xlsx
			doc.create(stdname);//Crear patilla llamado Template1.xlsx
			doc2.open(stdnameBase);//abrir archivo contenedor de la info
			XLWorksheet wks = doc.workbook().worksheet("Sheet1");//manejador de hoja
			XLWorksheet dcnvs = doc2.workbook().worksheet("SKU");//manejador de hoja
			int id, idant = 0;
			bool padreFlag = false;
			string sku;
			int rango = final - inicio;
			int paso = rango / 100;
			int pasoAnt = paso;
			

			for (int k = inicio, l = k, checkBar = 0; k <= final; k++, l++,checkBar++) {
				id = dcnvs.cell("B" +  to_string(k)).value();
				if (k > inicio) {
					idant = dcnvs.cell("B" + to_string(k - 1)).value();
					string cat = dcnvs.cell("C" + to_string(k)).value();
					sku.append(cat);
					sku.append(to_string(id));
					if (padreFlag) {
						padreFlag = false;
						k--;
					}
				}
				if (id == idant) {
					tmp->hijo(k, l + 1 - inicio, wks, dcnvs, sku);
					sku.erase();
				}
				else {
					padreFlag = true;
					string tituloPadre = dcnvs.cell("K" + to_string(k)).value();
					size_t found = tituloPadre.find_last_of(" ");
					size_t fin = tituloPadre.length();
					tituloPadre.erase(found, fin);
					tmp->padre(k, (l + 1 - inicio), wks, dcnvs, sku, tituloPadre);
					sku.erase();
				}
				
				if(checkBar==pasoAnt)
				{
					progressBar1->Increment(1);
					pasoAnt += paso;
				}

				
			}
			
			doc.save();
			doc.close();
			progressBar1->Visible = false;
			MessageBox::Show("Listo cerrar programa");
		}
		else {
			MessageBox::Show("No se pudo crear el excel.xlsx pruebe otra vez");

		}
	}
	private: System::Void numericUpDown1_ValueChanged(System::Object^ sender, System::EventArgs^ e) {
		if (numericUpDownDesde->Value < 0) {
			MessageBox::Show("El valor no puede ser menor que cero");
		}
	}
	private: System::Void numericUpDownHasta_ValueChanged(System::Object^ sender, System::EventArgs^ e) {
		if (numericUpDownDesde->Value < 0) {
			MessageBox::Show("El valor no puede ser menor que cero");
		}
	}
};
}
