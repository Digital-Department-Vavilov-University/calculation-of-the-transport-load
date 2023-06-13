#include "MainForm.h"
#include <time.h>

#define CSV_DELIMETER ";"
#define CSV_DELIMETER_CHR ';'

using namespace System::Text;

System::Void Project1::MainForm::MainForm_Load(System::Object^ sender, System::EventArgs^ e)
{
    this->lgruzDataCells = gcnew System::Collections::Generic::List<DataGridViewCell^>();
    this->sgruzDataCells = gcnew System::Collections::Generic::List<DataGridViewCell^>();
    this->tgruzDataCells = gcnew System::Collections::Generic::List<DataGridViewCell^>();
    this->busDataCells = gcnew System::Collections::Generic::List<DataGridViewCell^>();
    this->carDataCells = gcnew System::Collections::Generic::List<DataGridViewCell^>();

    CreateGridData();

    return System::Void();
}

void Project1::MainForm::CreateGridData()
{
    LoadDataFromTemplate(this->grid);
}

System::Void Project1::MainForm::saveToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
    SaveFileDialog^ saveFileDialog = gcnew SaveFileDialog();
    saveFileDialog->Filter = "(csv file(*.csv)| *.csv";
    try
    {
        if (saveFileDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK)
        {
            String^ path = saveFileDialog->FileName;

            String^ collectedData = GetCSVData(this->grid);

            System::IO::File::WriteAllText(path, collectedData, System::Text::Encoding::UTF8);

            MessageBox::Show("Таблица успешно сохранена по пути\n" + path);
        }
        else {
            return System::Void();
        }
    }
    catch (Exception^ ex) {
        MessageBox::Show(this, "Не удалось сохранить. Возможно файл открыт в другой программе.", "Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }

    return System::Void();
}

System::String^ Project1::MainForm::GetCSVData(DataGridView^ grid)
{
    System::Text::StringBuilder^ sb = gcnew System::Text::StringBuilder();

    sb->Append(grid->TopLeftHeaderCell->Value->ToString());
    sb->Append(CSV_DELIMETER);
    for (int i = 0; i < grid->Columns->Count; i++) {
        System::Object^ header = grid->Columns[i]->HeaderCell->Value;
        if (header == nullptr) header = "";
        sb->Append(header->ToString());
        sb->Append(CSV_DELIMETER);
    }
    sb->Append(CSV_DELIMETER);
    sb->Append("\n");

    for (int i = 0; i < grid->Rows->Count; i++) {
        System::Object^ header = grid->Rows[i]->HeaderCell->Value;
        if (header == nullptr) header = "";
        sb->Append(header->ToString());
        sb->Append(CSV_DELIMETER);

        DataGridViewRow^ row = grid->Rows[i];
        for (int j = 0; j < row->Cells->Count; j++) {
            System::Object^ val = row->Cells[j]->Value;
            if (val == nullptr) val = "";
            sb->Append(row->Cells[j]->Value);
            sb->Append(CSV_DELIMETER);
        }
        sb->Append("\n");
    }

    return sb->ToString();
}

System::Void Project1::MainForm::grid_CellEndEdit(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e)
{
    DataGridViewCell^ cell = this->grid->CurrentCell;
    String^ text = cell->FormattedValue->ToString();

    if (!ValidateCellText(text, cell)) return;

    RecalculateData();

    return System::Void();
}

bool Project1::MainForm::ValidateCellText(String^ text, DataGridViewCell^ cell)
{
    if (text->Length == 0) {
        cell->Value = 0;
        MessageBox::Show("Введите натуральное число");
        return false;
    }

    for (int i = 0; i < text->Length; i++) {
        if (!Char::IsDigit(text[i])) {
            cell->Value = 0;
            MessageBox::Show("Введите натуральное число");
            return false;
        }
    }

    if (text->Length > 15) {
        cell->Value = text->Substring(0, 15);
        MessageBox::Show("Слишком большое значение ячейки");
        return false;
    }

    return true;
}

void Project1::MainForm::RecalculateData()
{
    clock_t tStart = clock();

    long long lgruzsum = 0;
    for (int i = 0; i < this->lgruzDataCells->Count; i++) {
        System::Object^ data = this->lgruzDataCells[i]->Value;
        if (data != nullptr) {
            lgruzsum += Int64::Parse(data->ToString());
        }
    }
    double lgruzav = lgruzsum / (double)this->lgruzDataCells->Count;
    this->lGruzAverage->Value = System::Math::Round(lgruzav, 2);

    long long sgruzsum = 0;
    for (int i = 0; i < this->sgruzDataCells->Count; i++) {
        System::Object^ data = this->sgruzDataCells[i]->Value;
        if (data != nullptr) {
            sgruzsum += Int64::Parse(data->ToString());
        }
    }
    double sgruzav = sgruzsum / (double)this->sgruzDataCells->Count;
    this->sGruzAverage->Value = System::Math::Round(sgruzav, 2);

    long long tgruzsum = 0;
    for (int i = 0; i < this->tgruzDataCells->Count; i++) {
        System::Object^ data = this->tgruzDataCells[i]->Value;
        if (data != nullptr) {
            tgruzsum += Int64::Parse(data->ToString());
        }
    }
    double tgruzav = tgruzsum / (double)this->tgruzDataCells->Count;
    this->tGruzAverage->Value = System::Math::Round(tgruzav, 2);

    long long bussum = 0;
    for (int i = 0; i < this->busDataCells->Count; i++) {
        System::Object^ data = this->busDataCells[i]->Value;
        if (data != nullptr) {
            bussum += Int64::Parse(data->ToString());
        }
    }
    double busav = bussum / (double)this->busDataCells->Count;
    this->busAverage->Value = System::Math::Round(busav, 2);

    long long carsum = 0;
    for (int i = 0; i < this->carDataCells->Count; i++) {
        System::Object^ data = this->carDataCells[i]->Value;
        if (data != nullptr) {
            carsum += Int64::Parse(data->ToString());
        }
    }
    double carav = carsum / (double)this->carDataCells->Count;
    this->carAverage->Value = System::Math::Round(carav, 2);

    double avsum = lgruzav + sgruzav + tgruzav + busav + carav;
    this->sumData->Value = System::Math::Round(avsum, 2);

    double lgruzperc = (lgruzav / avsum) * 100;
    this->lGruzPercents->Value = System::Math::Round(lgruzperc, 2);
    double sgruzperc = (sgruzav / avsum) * 100;
    this->sGruzPercents->Value = System::Math::Round(sgruzperc, 2);
    double tgruzperc = (tgruzav / avsum) * 100;
    this->tGruzPercents->Value = System::Math::Round(tgruzperc, 2);
    double busperc = (busav / avsum) * 100;
    this->busPercents->Value = System::Math::Round(busperc, 2);
    double carperc = (carav / avsum) * 100;
    this->carPercents->Value = System::Math::Round(carperc, 2);

    double percentSum = lgruzperc + sgruzperc + tgruzperc + busperc + carperc;
    this->percentsSumData->Value = System::Math::Round(percentSum, 2);
}

void Project1::MainForm::LoadDataFromTemplate(System::Windows::Forms::DataGridView^ grid)
{
    array<String^>^ rows = System::IO::File::ReadAllLines("template.csv", System::Text::Encoding::UTF8);
    array<String^>^ row0 = rows[0]->Split(CSV_DELIMETER_CHR);
    grid->TopLeftHeaderCell->Value = row0[0];
    for (int i = 1; i < row0->Length; i++) {
        grid->Columns->Add(row0[i], row0[i]);
        grid->Columns[i - 1]->SortMode = DataGridViewColumnSortMode::NotSortable;
    }
    grid->Rows->Add(rows->Length - 1);
    for (int i = 1; i < rows->Length; i++) {
        array<String^>^ cells = rows[i]->Split(CSV_DELIMETER_CHR);

        grid->Rows[i - 1]->HeaderCell->Value = cells[0];
        grid->Rows[i - 1]->HeaderCell->Style->BackColor = System::Drawing::Color::LightGray;

        for (int j = 1; j < cells->Length; j++) {
            ProcessCell(grid, i, j, cells[j]);
        }
    }
}

void Project1::MainForm::ProcessCell(System::Windows::Forms::DataGridView^ grid, int row, int column, String^ rawCellText)
{
    RegularExpressions::Regex^ tagRegex = gcnew RegularExpressions::Regex("<\\w*>");

    DataGridViewCell^ cell = grid->Rows[row - 1]->Cells[column - 1];
    String^ textData = rawCellText;
    textData = tagRegex->Replace(textData, String::Empty);
    cell->Value = textData;
    if (!rawCellText->Contains("<EDIT>"))
    {
        cell->ReadOnly = true;
        cell->Style->BackColor = System::Drawing::Color::LightGray;
    }

    if (rawCellText->Contains("<LGRUZDATA>")) {
        this->lgruzDataCells->Add(cell);
    }
    if (rawCellText->Contains("<SGRUZDATA>")) {
        this->sgruzDataCells->Add(cell);
    }
    if (rawCellText->Contains("<TGRUZDATA>")) {
        this->tgruzDataCells->Add(cell);
    }
    if (rawCellText->Contains("<BUSDATA>")) {
        this->busDataCells->Add(cell);
    }
    if (rawCellText->Contains("<CARDATA>")) {
        this->carDataCells->Add(cell);
    }

    if (rawCellText->Contains("<LGRUZAV>")) {
        this->lGruzAverage = cell;
    }
    if (rawCellText->Contains("<SGRUZAV>")) {
        this->sGruzAverage = cell;
    }
    if (rawCellText->Contains("<TGRUZAV>")) {
        this->tGruzAverage = cell;
    }
    if (rawCellText->Contains("<BUSAV>")) {
        this->busAverage = cell;
    }
    if (rawCellText->Contains("<CARAV>")) {
        this->carAverage = cell;
    }

    if (rawCellText->Contains("<LGRUZPERC>")) {
        this->lGruzPercents = cell;
    }
    if (rawCellText->Contains("<SGRUZPERC>")) {
        this->sGruzPercents = cell;
    }
    if (rawCellText->Contains("<TGRUZPERC>")) {
        this->tGruzPercents = cell;
    }
    if (rawCellText->Contains("<BUSPERC>")) {
        this->busPercents = cell;
    }
    if (rawCellText->Contains("<CARPERC>")) {
        this->carPercents = cell;
    }

    if (rawCellText->Contains("<AVSUM>")) {
        this->sumData = cell;
    }
    if (rawCellText->Contains("<PERCSUM>")) {
        this->percentsSumData = cell;
    }
}
