using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Diagnostics;

namespace Schedule_parser
{
    public partial class Form1 : Form
    {
        static List<string> professors = new List<string>();
        public List<Data> dataList = new List<Data>();
        public List<string> numbers = new List<string> { "1", "2", "3", "4", "5", "6" };
        public List<DateTime> dateList = new List<DateTime>();
        public static List<DataEntry> entryList = new List<DataEntry>();
        int lastRow, lastCol;
        public string group;

        public Form1()
        {
            InitializeComponent();
        }

        //Чтение exc
        public void ReadExcelFile()
        {
            //Regex pattern1 = new Regex(@".\s[a-zA-Z0-9]{2}\W[a-zA-Z0-9]{2}\W[a-zA-Z0-9]{2}.\W\s..\s[a-zA-Z0-9]{2}\W[a-zA-Z0-9]{2}\W[a-zA-Z0-9]{2}.\W.");

            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл Excel|*.XLSX;*.XLS";
            openDialog.ShowDialog();

            var ExcelApp = new Excel.Application();
            //Книга.
            var WorkBookExcel = ExcelApp.Workbooks.Open(openDialog.FileName);
            //Таблица.
            var WorkSheetExcel = (Excel.Worksheet)WorkBookExcel.Sheets[1];

            var lastCell = WorkSheetExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            string[,] list = new string[lastCell.Row, lastCell.Column];

            for (int i = 0; i < (int)lastCell.Row; i++)
            {
                for (int j = 0; j < (int)lastCell.Column; j++)
                {
                    list[i, j] = WorkSheetExcel.Cells[i + 1, j + 1].Text.ToString();//считал текст ячейки в строку
                    lastCol = j;

                    //Определяет какой группе принадлежит прочитанный файл расписания (находит значение переменной "group")
                    if (j != 0 && list[i, j - 1] == "№ пары")
                    {
                        group = list[i, j];
                    }
                    ///////////
                }
                lastRow = i;
            }

            //Закрытие сессии Excel
            WorkBookExcel.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ExcelApp.Quit(); // вышел из Excel
            GC.Collect(); // убрал за собой
            GC.WaitForPendingFinalizers();

            Process[] processes = Process.GetProcessesByName("Excel");
            foreach (Process p in processes)
            {
                p.Kill();
            }

            //Функции выполняются после того, как файл Excel был прочитан:
            //Создает список из ячеек файла таблицы расписания, которые содержат информацию о занятиях (номер пары, дата, предмет, преподаватель)
            addToDataList(list);
            //Создает список из объектов класса DataEntry, которые представляют собой записи о занятиях, разделенные по полям класса (group, date, subject, professor)
            dataList.ForEach(GetEntry);
            dateList.Sort();
            SortEntryListByDate(entryList);
        }

        public void multiSelect()
        {

        }

        public void SortEntryListByDate(List<DataEntry> entryList)
        {
            
        }

        public void CreateExcelFile(int count)
        {

        }

        //Все строки excel файла добавляются в список
        public void addToDataList(string[,] l)
        {
            for (int i = 0; i < lastRow; i++)
                for (int j = 0; j < lastCol; j++)
                {
                    bool contains = numbers.Contains(l[i, j], StringComparer.OrdinalIgnoreCase);
                    if (contains)
                    {
                        var arr = Regex.Split(l[i, j + 1], "\n");
                        var halls = Regex.Split(l[i, j + 3], "\n");
                        int k = 0;
                        Data data;
                        DataEntry entry;

                        foreach (string info in arr)
                        {
                            if (info.Contains("-лаб.раб.") || info.Contains("-теория") || info.Contains("-практика") || info.Contains("-ДИФ.ЗАЧЕТ") || info.Contains("-ЗАЧЕТ"))
                            {
                                //subject
                                string subject = info.Substring(0, info.IndexOf(";"));
                                string content = info.Remove(0, info.IndexOf(";") + 2);
                                //

                                //professor
                                string[] substrings = Regex.Split(info, ": ");
                                string professor = substrings[1];
                                //Удалить из поля professor лишние символы
                                if (professor.Contains(";"))
                                     professor = professor.Replace(";", "");
                                //Добавить профессора в combobox список профессоров (если уже не был добавлен)
                                if (!professorsComboBox.Items.Contains(professor)) professorsComboBox.Items.Add(professor);

                                char[] chars = { '|', ':' };
                                content = Regex.Replace(content, professor, "");
                                content = ReplaceSeparators(content, '-', chars, '#');
                                content = content.Remove(content.LastIndexOf("#"), 1);

                                string[] splitContent = Regex.Split(content, "#");
                                foreach (string str in splitContent)
                                {
                                    string[] damnedSplit = Regex.Split(str, "-");
                                    //   entry = new DataEntry(l[i, j], damnedSplit[0], subject, professor, group, "xxx", damnedSplit[1]);
                                    k = 0;
                                    data = new Data(l[i, j], damnedSplit[0] + subject + "; " + damnedSplit[1] + ": " + professor, halls[k]);
                                    if (halls.Length > 1) k++;

                                    //Добавить профессора в combobox список профессоров (если уже не был добавлен)
                                    if (data.data != "")
                                    {
                                        //Удалить из поля professor лишние символы
                                        if (data.professor.Contains(";"))
                                            data.professor = data.professor.Replace(";", "");
                                        if (!professorsComboBox.Items.Contains(data.professor)) professorsComboBox.Items.Add(data.professor);
                                    }
                                    dataList.Add(data);
                                }
                            }
                            else
                            {
                                k = 0;
                                data = new Data(l[i, j], info, halls[k]);
                                if (halls.Length > 1) k++;

                                //Добавить профессора в combobox список профессоров (если уже не был добавлен)
                                if (data.data != "")
                                {
                                    //Удалить из поля professor лишние символы
                                    if (data.professor.Contains(";"))
                                        data.professor = data.professor.Replace(";", "");
                                    if (!professorsComboBox.Items.Contains(data.professor)) professorsComboBox.Items.Add(data.professor);
                                }
                                dataList.Add(data);
                            }
                        }
                    }
                }
        }

        //Преобразовать строки с датами формата "с хх.хх.ххг. по хх.хх.ххг.", "хх,хх,хх.хх.ххг." в отдельные записи
        public void GetEntry(Data data)
        {
            if (data.data != "")
            {
                DateTime date1, date2;
                string[] splitDate = Regex.Split(data.date, ";");
                for (int i = 0; i < splitDate.Length; i++)
                {
                    if (Regex.IsMatch(splitDate[i], @"[с]\W\d\d.\d\d.\d\d[г].\W[п][о]\W\d\d.\d\d.\d\d[г]."))
                    {
                        var index1 = splitDate[i].IndexOf("с");
                        var index2 = splitDate[i].IndexOf("по");
                        date1 = DateTime.ParseExact(splitDate[i].Substring(index1 + 2, 8), "dd.MM.yy", CultureInfo.InvariantCulture);
                        date2 = DateTime.ParseExact(splitDate[i].Substring(index2 + 3, 8), "dd.MM.yy", CultureInfo.InvariantCulture);

                        int shutdownCounter = 0;
                        while (date1 != date2)
                        {
                            DataEntry entry = new DataEntry(data.number, date1.ToString("dd.MM.yy"), data.subject, data.professor, group, data.lectureHall, data.type);
                            RemoveSpareSimbolsInDate(entry);
                            entryList.Add(entry);

                            //Заполнение списка дат
                            if (!dateList.Contains(DateTime.ParseExact(entry.date, "dd.MM.yy", CultureInfo.InvariantCulture)))
                                dateList.Add(DateTime.ParseExact(entry.date, "dd.MM.yy", CultureInfo.InvariantCulture));
                            //Заполнение списка дат

                            date1 = date1.AddDays(7);
                            if (date1 == date2)
                            {
                                entry = new DataEntry(data.number, date1.ToString("dd.MM.yy"), data.subject, data.professor, group, data.lectureHall, data.type);
                                RemoveSpareSimbolsInDate(entry);
                                entryList.Add(entry);

                                //Заполнение списка дат
                                if (!dateList.Contains(DateTime.ParseExact(entry.date, "dd.MM.yy", CultureInfo.InvariantCulture)))
                                    dateList.Add(DateTime.ParseExact(entry.date, "dd.MM.yy", CultureInfo.InvariantCulture));
                                //Заполнение списка дат
                            }
                            if (shutdownCounter == 25) break;
                            shutdownCounter++;
                        }
                    }
                    else if (splitDate[i].Contains(","))
                    {
                        var index = Regex.Match(splitDate[i], @"\d\d[.]\d\d[г]");
                        string monthAndYear = splitDate[i].Substring(index.Index, 5);
                        string[] splitStr = Regex.Split(splitDate[i], ",");

                        for (int j = 0; j < splitStr.Length - 1; j++)
                        {
                            splitStr[j] = splitStr[j] + "." + monthAndYear;
                        }
                        for (int k = 0; k < splitStr.Length; k++)
                        {
                            DataEntry entry = new DataEntry(data.number, splitStr[k], data.subject, data.professor, group, data.lectureHall, data.type);
                            RemoveSpareSimbolsInDate(entry);
                            entryList.Add(entry);

                            //Заполнение списка дат
                            if (!dateList.Contains(DateTime.ParseExact(entry.date, "dd.MM.yy", CultureInfo.InvariantCulture)))
                                dateList.Add(DateTime.ParseExact(entry.date, "dd.MM.yy", CultureInfo.InvariantCulture));
                            //Заполнение списка дат
                        }
                    }
                    else if (Regex.IsMatch(splitDate[i], @"\d\d[.]\d\d[.]\d\d[г]."))
                    {
                        DataEntry entry = new DataEntry(data.number, splitDate[i], data.subject, data.professor, group, data.lectureHall, data.type);
                        RemoveSpareSimbolsInDate(entry);
                        entryList.Add(entry);

                        //Заполнение списка дат
                        if (!dateList.Contains(DateTime.ParseExact(entry.date, "dd.MM.yy", CultureInfo.InvariantCulture)))
                            dateList.Add(DateTime.ParseExact(entry.date, "dd.MM.yy", CultureInfo.InvariantCulture));
                        //Заполнение списка дат
                    }
                }
            }
        }

        public void RemoveSpareSimbolsInDate(DataEntry entry)
        {
            if (entry.date.Contains("г. ")) entry.date = entry.date.Replace("г. ", "");
            if (entry.date.Contains("г.")) entry.date = entry.date.Replace("г.", "");
            if (entry.date.Contains("г")) entry.date = entry.date.Replace("г", "");
            if (entry.date.Contains(" ")) entry.date = entry.date.Replace(" ", "");
        }

        //Заменяет разделители
        public static string ReplaceSeparators(string input, char startingChar, char[] oldChar, char newChar)
        {
            List<int> indexes = new List<int>();
            char[] chArr = input.ToCharArray();
            for (int i = 0; i < chArr.Length; i++)
            {
                if (chArr[i] == startingChar)
                {
                    int j = i;
                    while (chArr[j] != ';' && chArr[j] != ':')
                        j++;
                    indexes.Add(j);
                }
            }

            for (int i = 0; i < chArr.Length; i++)
            {
                if (indexes.Contains(i)) chArr[i] = newChar;
            }
            string s = new string(chArr);
            return s;
        }

        //Print entries from entryList to textBox
        public void PrintEntryList(List<DataEntry> entryList)
        {
            string professor = professorsComboBox.Text;
            string str = "";
            for (int i = 0; i < entryList.Count; i++)
            {
                if (entryList[i].professor.Contains(professor))
                {
                    str = str + entryList[i].number + ") " + entryList[i].date + " " + entryList[i].subject + " " + entryList[i].professor + " " + entryList[i].grp + " " + entryList[i].lectureHall + " " + entryList[i].type + Environment.NewLine;
                }
            }
            richTextBox1.Text = str;
        }

        //
        void print(string str)
        {
            richTextBox1.Text += str + Environment.NewLine;
        }

        //Open file
        private void OpenFile_button_Click(object sender, EventArgs e)
        {
            ReadExcelFile();
        }

        //Print to textBox
        private void button1_Click(object sender, EventArgs e)
        {
            PrintEntryList(entryList);
        }
    }

    public class Data
    {
        public string data;

        public string number;
        public string date;
        public string subject;
        public string professor;
        public string lectureHall;
        public string type;

        public Data(string numberOfLecture, string info, string hall)
        {
            number = numberOfLecture;
            data = info;
            lectureHall = hall;
            if (data != "") Split(data);
        }

        public void Split(string info)
        {
            string pat4 = @"[г]..[а-яА-Я]{3}";
            string temp = "";

            //professor
            string[] substrings = Regex.Split(info, ": ");
            professor = substrings[1].Trim();

            info = substrings[0];

                //date
                substrings = Regex.Split(info, "[г]..[а-яА-Я]{3}");
                date = substrings[0] + "г.";
                //date = substrings[0];

                //subject
                temp = info.Remove(0, date.Length);
                subject = Regex.Split(temp, "; ")[0];
                type = Regex.Split(temp, "; ")[1];
                //subject = GetSubjectShortname(subject);
        }

        public string GetSubjectShortname(string subject)
        {
            char[] temp = subject.ToCharArray();
            string subj = "";
            foreach (char ch in temp)
            {
                if (Char.IsUpper(ch)) subj += ch; 
            }
            return subj;
        }
    }

    public class DataEntry
    {
        public string grp;

        public string number;
        public string date;
        public string subject;
        public string type;
        public string professor;
        public string lectureHall;

        public DataEntry(string num, string Date, string Subject, string Professor, string group, string hall, string Type)
        {
            //if (Date.Contains("г.")) Date.Replace("г.", "");

            grp = group;
            lectureHall = hall;
            number = num;
            date = Date;
            subject = Subject;
            type = Type;
            professor = Professor;
        }
    }
}



