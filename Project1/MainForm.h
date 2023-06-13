#pragma once

namespace Project1 {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// MainForm
	/// </summary>
	public ref class MainForm : public System::Windows::Forms::Form
	{
	public:
		MainForm(void)
		{
			InitializeComponent();
		}

	protected:
		~MainForm()
		{
			if (components)
			{
				delete components;
			}
		}

	protected:

	private: System::Collections::Generic::List<DataGridViewCell^>^ lgruzDataCells;
	private: System::Collections::Generic::List<DataGridViewCell^>^ sgruzDataCells;
	private: System::Collections::Generic::List<DataGridViewCell^>^ tgruzDataCells;
	private: System::Collections::Generic::List<DataGridViewCell^>^ busDataCells;
	private: System::Collections::Generic::List<DataGridViewCell^>^ carDataCells;
	private: DataGridViewCell^ lGruzAverage;
	private: DataGridViewCell^ sGruzAverage;
	private: DataGridViewCell^ tGruzAverage;
	private: DataGridViewCell^ busAverage;
	private: DataGridViewCell^ carAverage;
	private: DataGridViewCell^ lGruzPercents;
	private: DataGridViewCell^ sGruzPercents;
	private: DataGridViewCell^ tGruzPercents;
	private: DataGridViewCell^ busPercents;
	private: DataGridViewCell^ carPercents;
	private: DataGridViewCell^ sumData;
	private: DataGridViewCell^ percentsSumData;

	private: System::Windows::Forms::MenuStrip^ menuStrip1;
	private: System::Windows::Forms::ToolStripMenuItem^ saveToolStripMenuItem;
	private: System::Windows::Forms::DataGridView^ grid;

	private:
		System::ComponentModel::Container^ components;

#pragma region Windows Form Designer generated code
		void InitializeComponent(void)
		{
			this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
			this->saveToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->grid = (gcnew System::Windows::Forms::DataGridView());
			this->menuStrip1->SuspendLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->grid))->BeginInit();
			this->SuspendLayout();
			// 
			// menuStrip1
			// 
			this->menuStrip1->ImageScalingSize = System::Drawing::Size(20, 20);
			this->menuStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) { this->saveToolStripMenuItem });
			this->menuStrip1->Location = System::Drawing::Point(0, 0);
			this->menuStrip1->Name = L"menuStrip1";
			this->menuStrip1->Size = System::Drawing::Size(862, 28);
			this->menuStrip1->TabIndex = 1;
			this->menuStrip1->Text = L"menuStrip1";
			// 
			// saveToolStripMenuItem
			// 
			this->saveToolStripMenuItem->Name = L"saveToolStripMenuItem";
			this->saveToolStripMenuItem->Size = System::Drawing::Size(97, 24);
			this->saveToolStripMenuItem->Text = L"Сохранить";
			this->saveToolStripMenuItem->Click += gcnew System::EventHandler(this, &MainForm::saveToolStripMenuItem_Click);
			// 
			// grid
			// 
			this->grid->AllowUserToAddRows = false;
			this->grid->AllowUserToDeleteRows = false;
			this->grid->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->grid->Dock = System::Windows::Forms::DockStyle::Fill;
			this->grid->Location = System::Drawing::Point(0, 28);
			this->grid->Name = L"grid";
			this->grid->RowHeadersWidth = 51;
			this->grid->RowTemplate->Height = 24;
			this->grid->Size = System::Drawing::Size(862, 461);
			this->grid->TabIndex = 2;
			this->grid->CellEndEdit += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &MainForm::grid_CellEndEdit);
			// 
			// MainForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(8, 16);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(862, 489);
			this->Controls->Add(this->grid);
			this->Controls->Add(this->menuStrip1);
			this->MainMenuStrip = this->menuStrip1;
			this->Name = L"MainForm";
			this->Text = L"MainForm";
			this->Load += gcnew System::EventHandler(this, &MainForm::MainForm_Load);
			this->menuStrip1->ResumeLayout(false);
			this->menuStrip1->PerformLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->grid))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion


	private: System::Void MainForm_Load(System::Object^ sender, System::EventArgs^ e);

	private: void CreateGridData();

	private: void LoadDataFromTemplate(DataGridView^);

	private: void ProcessCell(System::Windows::Forms::DataGridView^ grid, int row, int column, String^ rawCellText);

	private: System::Void saveToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e);

	private: String^ GetCSVData(DataGridView^ grid);

	private: System::Void grid_CellEndEdit(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e);

	private: bool ValidateCellText(System::String^ text, System::Windows::Forms::DataGridViewCell^ cell);

	private: void RecalculateData();
};
}
