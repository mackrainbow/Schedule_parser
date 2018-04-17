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
        public List<DataEntry> profesorsLectures = new List<DataEntry>();
        public static List<DataEntry> entryList = new List<DataEntry>();
        public SortedDictionary<string, List<DataEntry>> lections = new SortedDictionary<string, List<DataEntry>>(new SortComparer());
        public int mon, tue, wed, thur, fri, sat = 0;
        public int hallRowIndex, hallColIndex;
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
            openDialog.Multiselect = true;
            openDialog.ShowDialog();
            //string fileName = Path.GetFileName(openDialog.FileName);

            foreach (String file in openDialog.FileNames)
            {
                multiSelect(file);
            }
        }

        public void multiSelect(string fileName)
        {
            var ExcelApp = new Excel.Application();
            //Книга.
            var WorkBookExcel = ExcelApp.Workbooks.Open(fileName);
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
                    if (list[i, j] == "ауд.")
                    {
                        hallRowIndex = i;
                        hallColIndex = j;
                    }

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
            dataList.Clear();
            GetProfessorsLectures(entryList);
            // CreateExcelFile(entryList[1]);
        }

        public void GetProfessorsLectures(List<DataEntry> entryList)
        {
            foreach (DataEntry entry in entryList)
            {
                if (entry.professor == professorsComboBox.Text)
                {
                    if (entry.date == "29.01.18")
                    {

                    }
                    profesorsLectures.Add(entry);
                }

            }

            foreach (DataEntry entry in profesorsLectures)
            {
                if (lections.ContainsKey(entry.date))
                {
                    lections[entry.date].Add(entry);
                }
                else
                {
                    var lectionDay = new List<DataEntry>();
                    lectionDay.Add(entry);
                    lections.Add(entry.date, lectionDay);
                }
            }      
        }

        private void professorsComboBox_TextChanged(object sender, EventArgs e)
        {
            GetProfessorsLectures(entryList);
            SaveFileButton.Enabled = true;
        }

        private void SaveFileButton_Click(object sender, EventArgs e)
        {
            CreateExcelFile(entryList[1]);
        }

        public void CreateExcelFile(DataEntry entry)
        {
            var info = entry;

            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            object misValue = System.Reflection.Missing.Value;

            var xlWorkBook = xlApp.Workbooks.Add(misValue);
            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            var cells = xlWorkSheet.Cells;

            cells[3, 1] = "Понедельник";
            cells[11, 1] = "Вторник";
            cells[19, 1] = "Среда";
            cells[30, 1] = "Четверг";
            cells[41, 1] = "Пятница";
            cells[52, 1] = "Суббота";

            //Понедельник
            var range = xlWorkSheet.Range[cells[3, 1], cells[10, 1]];
            range.Merge();
            range.Cells.Orientation = Excel.XlOrientation.xlUpward;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            //Вторник
            range = xlWorkSheet.Range[cells[11, 1], cells[18, 1]];
            range.Merge();
            range.Cells.Orientation = Excel.XlOrientation.xlUpward;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            //Среда
            range = xlWorkSheet.Range[cells[19, 1], cells[26, 1]];
            range.Merge();
            range.Cells.Orientation = Excel.XlOrientation.xlUpward;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            //Четверг
            range = xlWorkSheet.Range[cells[27, 1], cells[34, 1]];
            range.Merge();
            range.Cells.Orientation = Excel.XlOrientation.xlUpward;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            //Пятница
            range = xlWorkSheet.Range[cells[35, 1], cells[42, 1]];
            range.Merge();
            range.Cells.Orientation = Excel.XlOrientation.xlUpward;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            //Суббота
            range = xlWorkSheet.Range[cells[43, 1], cells[50, 1]];
            range.Merge();
            range.Cells.Orientation = Excel.XlOrientation.xlUpward;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            foreach (string key in lections.Keys)
            {
                switch (DateTime.ParseExact(key, "dd.MM.yy", CultureInfo.InvariantCulture).DayOfWeek)
                {
                    case DayOfWeek.Monday:
                        {
                            mon++;
                            break;
                        }
                    case DayOfWeek.Tuesday:
                        {
                            tue++;
                            break;
                        }
                    case DayOfWeek.Wednesday:
                        {
                            wed++;
                            break;
                        }
                    case DayOfWeek.Thursday:
                        {
                            thur++;
                            break;
                        }
                    case DayOfWeek.Friday:
                        {
                            fri++;
                            break;
                        }
                    case DayOfWeek.Saturday:
                        {
                            sat++;
                            break;
                        }
                }
                                //CreateBlockTemplate(x, y, xlWorkSheet, cells, key);
                                //FillBlock(x, y, xlWorkSheet, cells, key, lections[key]);
            }

            int x = 3;
            int y = 2;
            int a1 = 3;
            int a2 = 3;
            int a3 = 3;
            int a4 = 3;
            int a5 = 3;
            int a6 = 3;

            foreach (string key in lections.Keys)
            {
                switch (DateTime.ParseExact(key, "dd.MM.yy", CultureInfo.InvariantCulture).DayOfWeek)
                {
                    case DayOfWeek.Monday:
                        {
                            CreateBlockTemplate(a1, y, xlWorkSheet, cells, key);
                            FillBlock(a1, y, xlWorkSheet, cells, key, lections[key]);
                            a1 += 8;
                            break;
                        }
                    case DayOfWeek.Tuesday:
                        {
                            CreateBlockTemplate(a2, y + 8, xlWorkSheet, cells, key);
                            FillBlock(a2, y + 8, xlWorkSheet, cells, key, lections[key]);
                            a2 += 8;
                            break;
                        }
                    case DayOfWeek.Wednesday:
                        {
                            CreateBlockTemplate(a3, y + 16, xlWorkSheet, cells, key);
                            FillBlock(a3, y + 16, xlWorkSheet, cells, key, lections[key]);
                            a3 += 8;
                            break;
                        }
                    case DayOfWeek.Thursday:
                        {
                            CreateBlockTemplate(a4, y + 24, xlWorkSheet, cells, key);
                            FillBlock(a4, y + 24, xlWorkSheet, cells, key, lections[key]);
                            a4 += 8;
                            break;
                        }
                    case DayOfWeek.Friday:
                        {
                            CreateBlockTemplate(a5, y + 32, xlWorkSheet, cells, key);
                            FillBlock(a5, y + 32, xlWorkSheet, cells, key, lections[key]);
                            a5 += 8;
                            break;
                        }
                    case DayOfWeek.Saturday:
                        {
                            CreateBlockTemplate(a6, y + 40, xlWorkSheet, cells, key);
                            FillBlock(a6, y + 40, xlWorkSheet, cells, key, lections[key]);
                            a6 += 8;
                            break;
                        }
                }
            }

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Файл Excel|*.XLSX;*.XLS";
            saveDialog.RestoreDirectory = true;


          //  xlWorkBook.SaveAs("d:\\Professor's_Schedule.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }

        public void FillBlock(int y, int x, Excel.Worksheet xlWorkSheet, Excel.Range cells, string date, List<DataEntry> entryList)
        {
            for (int i = 1; i <= 6; i++)
            {
                //№ пары
                cells[x + 2 + i, y - 1] = i;
                DataEntry entry = GetEntry(entryList, i);
                if (entry != null)
                {
                    if (entry.date == "29.01.18")
                    {

                    }
                    if (cells[x + 2 + i, y].Value2 == null)
                    {
                        cells[x + 2 + i, y] = entry.grp;
                        cells[x + 2 + i, y + 1] = entry.subject;
                        cells[x + 2 + i, y + 2] = entry.lectureHall;
                        cells[x + 2 + i, y + 3] = entry.type;
                    }
                    else
                    {
                        cells[x + 2 + i, y - 1] = "! " + cells[x + 2 + i, y - 1].Text.ToString();
                        cells[x + 2 + i, y] = cells[x + 2 + i, y].Text.ToString() + " " + entry.grp;

                        var range = xlWorkSheet.Range[cells[x + 2 + i, y], cells[x + 2 + i, y + 3]];
                        range.Font.Color = ColorTranslator.ToOle(Color.Red);
                    }
                }
            }
        }

        public void CreateBlockTemplate(int y, int x, Excel.Worksheet xlWorkSheet, Excel.Range cells, string date)
        {
            cells[x + 1, y] = date; //date
            cells[x + 2, y - 1] = "№ пары";
            cells[x + 2, y] = "Группа";
            cells[x + 2, y + 1] = "Предмет";
            cells[x + 2, y + 2] = "Кабинет";
            cells[x + 2, y + 3] = "Вид занятия";
            var range = xlWorkSheet.Range[cells[x + 1, y - 1], cells[x + 1, y + 3]];
            range.Merge();
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        public DataEntry GetEntry(List<DataEntry> entryList, int value)
        {
            foreach (DataEntry entry in entryList)
            {
                if (entry.number == value.ToString())
                    return entry;
            }
            return null;
        }

        /* public bool Contains(List<DataEntry> entryList, int value)
         {
             foreach (DataEntry entry in entryList)
             {
                 if (entry.number == value.ToString())
                     return true;
             }
             return false;
         } */



        //Все строки excel файла добавляются в список
        public void addToDataList(string[,] l)
        {
            for (int i = 0; i < lastRow; i++)
                for (int j = 0; j < lastCol; j++)
                {
                    if (l[i, j] == "№ пары") group = l[i, j + 1];
                    bool contains = numbers.Contains(l[i, j], StringComparer.OrdinalIgnoreCase);
                    if (contains)
                    {
                        var arr = Regex.Split(l[i, j + 1], "\n");
                        var halls = Regex.Split(l[i, hallColIndex], "\n");
                        int k = 0;
                        int n = 0;
                        Data data;

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
                                 //   k = 0;
                                    data = new Data(l[i, j], damnedSplit[0] + subject + "; " + damnedSplit[1] + ": " + professor, halls[n], group);
                                    if (halls.Length > 1) n++;

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
                              //  k = 0;
                                data = new Data(l[i, j], info, halls[k], group);
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
                bool check = true;
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
                            DataEntry entry = new DataEntry(data.number, date1.ToString("dd.MM.yy"), data.subject, data.professor, data.group, data.lectureHall, data.type);
                            RemoveSpareSimbolsInDate(entry);
                            entryList.Add(entry);

                            //Заполнение списка дат
                            if (!dateList.Contains(DateTime.ParseExact(entry.date, "dd.MM.yy", CultureInfo.InvariantCulture)))
                                dateList.Add(DateTime.ParseExact(entry.date, "dd.MM.yy", CultureInfo.InvariantCulture));
                            //Заполнение списка дат

                            date1 = date1.AddDays(7);
                            if (date1 == date2)
                            {
                                entry = new DataEntry(data.number, date1.ToString("dd.MM.yy"), data.subject, data.professor, data.group, data.lectureHall, data.type);
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
                            DataEntry entry = new DataEntry(data.number, splitStr[k], data.subject, data.professor, data.group, data.lectureHall, data.type);
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
                        DataEntry entry = new DataEntry(data.number, splitDate[i], data.subject, data.professor, data.group, data.lectureHall, data.type);
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

        //Open file
        private void OpenFile_button_Click(object sender, EventArgs e)
        {
            profesorsLectures.Clear();
            professors.Clear();
            dataList.Clear();
            profesorsLectures.Clear();
            dateList.Clear();
            entryList.Clear();
            lections.Clear();


            ReadExcelFile();
        }
    }

    public class Data
    {
        public string data;

        public string group;
        public string number;
        public string date;
        public string subject;
        public string professor;
        public string lectureHall;
        public string type;

        public Data(string numberOfLecture, string info, string hall, string grp)
        {
            group = grp; 
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
            substrings = Regex.Split(info, "[г]...[а-яА-Я]{3}");
            date = substrings[0] + "г.";
            //date = substrings[0];

            //subject
            temp = info.Remove(0, date.Length);
            subject = Regex.Split(temp, "; ")[0];
            type = Regex.Split(temp, "; ")[1];
            subject = GetSubjectShortname(subject);
        }

        public string GetSubjectShortname(string subject)
        {
            char[] temp = subject.ToCharArray();
            string subj = "";
            for (int i = 0; i < temp.Length; i++)
            {
                if (i != 0 && temp[i - 1] == ' ') subj += Char.ToUpper(temp[i]);
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

    public class SortComparer : IComparer<string>
    {
        public int Compare (string x, string y)
        {
            DateTime xDate = DateTime.ParseExact(x, "dd.MM.yy", CultureInfo.InvariantCulture);
            DateTime yDate = DateTime.ParseExact(y, "dd.MM.yy", CultureInfo.InvariantCulture);

            return xDate.CompareTo(yDate);
        }
    }
}



