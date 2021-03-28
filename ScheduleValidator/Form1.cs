using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

using ADOX;

namespace ScheduleValidator
{
    struct StudyGroup {
        public string GroupName;
        public int ID1;
        public int ID2;
    };
    struct MyGroup
    {
        public string ObjectName;
        public int ID;
    };
    struct ScheduleRecord
    {
        public int ID;
        public string Subject;
        public int Subgroup;
        public string Teacher;
        public string Room;
        public string Type;
        public int WeekNumber;
        public int DayOfWeek;
        public int LessonNumber;
    };
    public partial class Form1 : Form
    {
        private string databasePath = string.Empty;
        private string DSN = string.Empty;
        private const string MdbMask = "Microsoft Access DB (*.mdb)|*.mdb";
        private const string XlsxMask = "Microsoft Excel (*.xlsx, *.xls)|*.xlsx;*.xls";
        public Form1()
        {
            this.MinimumSize = new Size(800, 600);
            InitializeComponent();
            this.setDatabasePath("C:\\Users\\Victoria\\Documents\\test.mdb");
            //this.initDB();
            this.databaseOpened(true);
            // this.importFromXLS("C:\\Users\\Victoria\\Documents\\Осень-2019-2020.xlsx");
            // this.importFromXLS("C:\\Users\\Victoria\\Documents\\Осень-2018-2019.xls");
            // this.importFromXLS("C:\\Users\\Victoria\\Documents\\Весна-2019-2020.xlsx");
            // this.importFromXLS("C:\\Users\\Victoria\\Documents\\Весна-2018-2019.xls");
            // this.importFromXLS("C:\\Users\\Victoria\\Documents\\Осень-2020-2021.xlsx");
        }

        private void setDatabasePath(string path)
        {
            this.databasePath = path;
            if (path != string.Empty)
            {
                this.DSN = "Provider=Microsoft.Jet.OLEDB.4.0;";
                this.DSN += "Data Source=" + this.databasePath + ";Jet OLEDB:Engine Type=5";
            } else
            {
                this.DSN = string.Empty;
            }
        }

        private void databaseOpened(bool isOpened)
        {
            errors.Text = "";
            if (isOpened)
            {
                this.textDatabasePath.Text = this.databasePath;
                this.importFromXLSToolStripMenuItem.Enabled = true;
                this.closeDatabaseToolStripMenuItem.Enabled = true;
                this.checkForErrorsToolStripMenuItem.Enabled = true;
            } else {
                this.textDatabasePath.Text = string.Empty;
                this.importFromXLSToolStripMenuItem.Enabled = false;
                this.closeDatabaseToolStripMenuItem.Enabled = false;
                this.checkForErrorsToolStripMenuItem.Enabled = false;
            }
        }

        private void initDB()
        {
            if (File.Exists(this.databasePath))
            {
                File.Delete(this.databasePath);
            }
            this.databaseOpened(true);
            Catalog cat = new ADOX.Catalog();
            cat.Create(this.DSN);
            ADOX.Table group = new ADOX.Table();
            ADOX.Table room = new ADOX.Table();
            ADOX.Table subject = new ADOX.Table();
            ADOX.Table teacher = new ADOX.Table();
            ADOX.Table schedule = new ADOX.Table();

            //Create the table and it's fields. 
            group.Name = "Group";
            group.Columns.Append("Group_ID", ADOX.DataTypeEnum.adInteger);
            group.Keys.Append("PrimaryKey", KeyTypeEnum.adKeyPrimary, "Group_ID");
            group.Columns["Group_ID"].ParentCatalog = cat;
            group.Columns["Group_ID"].Properties["AutoIncrement"].Value = true;
            group.Columns.Append("Name");  // полное название группы ИС-О-20/2     
            group.Columns.Append("Year", ADOX.DataTypeEnum.adInteger);  // курс (год поступления): 20
            group.Columns.Append("Speciality");  // специализация: ИС
            group.Columns.Append("Iteration", ADOX.DataTypeEnum.adInteger);   // если несколько групп одной специализации: /2
            group.Columns.Append("Subgroup", ADOX.DataTypeEnum.adInteger);  // подгруппа по иностранному
            cat.Tables.Append(group);

            room.Name = "Room";
            room.Columns.Append("Room_ID", ADOX.DataTypeEnum.adInteger);
            room.Keys.Append("PrimaryKey", KeyTypeEnum.adKeyPrimary, "Room_ID");
            room.Columns["Room_ID"].ParentCatalog = cat;
            room.Columns["Room_ID"].Properties["AutoIncrement"].Value = true;
            room.Columns.Append("Name");  // название аудитории 202б
            cat.Tables.Append(room);

            subject.Name = "Subject";
            subject.Columns.Append("Subject_ID", ADOX.DataTypeEnum.adInteger);
            subject.Keys.Append("PrimaryKey", KeyTypeEnum.adKeyPrimary, "Subject_ID");
            subject.Columns["Subject_ID"].ParentCatalog = cat;
            subject.Columns["Subject_ID"].Properties["AutoIncrement"].Value = true;
            subject.Columns.Append("Name");  // name of subject
            cat.Tables.Append(subject);

            teacher.Name = "Teacher";
            teacher.Columns.Append("Teacher_ID", ADOX.DataTypeEnum.adInteger);
            teacher.Keys.Append("PrimaryKey", KeyTypeEnum.adKeyPrimary, "Teacher_ID");
            teacher.Columns["Teacher_ID"].ParentCatalog = cat;
            teacher.Columns["Teacher_ID"].Properties["AutoIncrement"].Value = true;
            teacher.Columns.Append("Name");  // teacher's name
            cat.Tables.Append(teacher);

            schedule.Name = "Schedule";
            schedule.Columns.Append("Schedule_ID", ADOX.DataTypeEnum.adInteger);
            schedule.Keys.Append("PrimaryKey", KeyTypeEnum.adKeyPrimary, "Schedule_ID");
            schedule.Columns["Schedule_ID"].ParentCatalog = cat;
            schedule.Columns["Schedule_ID"].Properties["AutoIncrement"].Value = true;
            schedule.Columns.Append("WeekNumber", ADOX.DataTypeEnum.adInteger);  // номер недели 1 или 2
            schedule.Columns.Append("DayOfWeek", ADOX.DataTypeEnum.adInteger);   // день недели
            schedule.Columns.Append("LessonNumber", ADOX.DataTypeEnum.adInteger);// номер пары 
            schedule.Columns.Append("LessonType");  // тип урока: лек сем лаб
            schedule.Columns.Append("Group_ID", ADOX.DataTypeEnum.adInteger);    // id подгруппы 
            schedule.Columns.Append("Subject_ID", ADOX.DataTypeEnum.adInteger);  // id предмета
            schedule.Columns.Append("Room_ID", ADOX.DataTypeEnum.adInteger);     // id кабинета
            schedule.Columns.Append("Teacher_ID", ADOX.DataTypeEnum.adInteger);  // id преподавателя
            schedule.Columns["LessonType"].Attributes = ADOX.ColumnAttributesEnum.adColNullable;
            // schedule.Columns["LessonType"].Properties["Default"].Value = null;
            schedule.Columns["Room_ID"].Attributes = ADOX.ColumnAttributesEnum.adColNullable;
            schedule.Columns["Teacher_ID"].Attributes = ADOX.ColumnAttributesEnum.adColNullable;
            schedule.Keys.Append("ForeignKey_Group_ID", ADOX.KeyTypeEnum.adKeyForeign, "Group_ID", "Group", "Group_ID");
            schedule.Keys.Append("ForeignKey_Subject_ID", ADOX.KeyTypeEnum.adKeyForeign, "Subject_ID", "Subject", "Subject_ID");
            schedule.Keys.Append("ForeignKey_Room_ID", ADOX.KeyTypeEnum.adKeyForeign, "Room_ID", "Room", "Room_ID");
            schedule.Keys.Append("ForeignKey_Teacher_ID", ADOX.KeyTypeEnum.adKeyForeign, "Teacher_ID", "Teacher", "Teacher_ID");

            cat.Tables.Append(schedule);

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(group);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(room);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(subject);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(teacher);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(schedule);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cat.Tables);
            //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cat.ActiveConnection);

            var con = cat.ActiveConnection;
            //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cat.ActiveConnection);
            if (con != null)
                con.Close();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cat);
        }

        private void newDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // TODO create new Database
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                // Stream myStream;
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                saveFileDialog.Filter = MdbMask;
                saveFileDialog.FilterIndex = 2;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.OverwritePrompt = true;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {

                    this.setDatabasePath(saveFileDialog.FileName);
                    this.initDB();
                }
            }

        }

        private void openDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileDialog.Filter = MdbMask;
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    this.setDatabasePath(openFileDialog.FileName);
                    this.databaseOpened(true);

                }
            }
        }

        private void closeDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.databasePath = string.Empty;
            this.databaseOpened(false);
        }

        private OleDbConnection returnXlsConnection(string fileName)
        {
            string extension = Path.GetExtension(fileName);
            Console.WriteLine(fileName);
            if (extension == ".xlsx")
            {
                return new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0;");
            }
            else
            {
                return new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + "; Jet OLEDB:Engine Type=5;Extended Properties=\"Excel 8.0;\"");
            }
        }

        private string NormalizeName(string teacher)
        {
            string name = teacher;
            Match match = Regex.Match(name, "([А-ЯЁ])[,\\.]?\\s*([А-ЯЁ])[,\\.]?\\s*([А-ЯЁ][а-яё]+)");
            if (match.Success == true)
            {
                name = match.Groups[1].Value + '.' + match.Groups[2].Value + '.' + match.Groups[3].Value;
            }
            match = Regex.Match(name, "([А-ЯЁ][а-яё]+)\\s+([А-ЯЁ])[,\\.]?\\s*([А-ЯЁ])[,\\.]?\\s*");
            if (match.Success == true)
            {
                name = match.Groups[2].Value + '.' + match.Groups[3].Value + '.' + match.Groups[1].Value;
            }
            return name;
        }

        private string NormalizeType(string Type)
        {
            if (Type == null || Type == string.Empty)
            {
                return Type;
            }
            if (Type[0] == 'c' || Type[0] == 'с')
            {
                return "сем";
            }
            if (Type.Length > 1 && Type[1] == 'а')
            {
                return "лаб";
            }
            // Console.WriteLine(Type + " -> лек");
            return "лек";
        }

        private void importFromXLS(string filename)
        {
            string sheetName;
            string message = "Не поддерживаемый формат XLS/XLSX!";
            string caption = "Импорт данных из XLS";

            DataSet ds = new DataSet();
            using (OleDbConnection con = this.returnXlsConnection(filename))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        cmd.Connection = con;
                        try
                        {
                            con.Open();
                        }
                        catch (System.Data.OleDb.OleDbException) {
                            MessageBox.Show(message, caption);
                            return;
                        }
                        finally
                        {

                        }
                        DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        for (int i = 0; i < dtExcelSchema.Rows.Count; i++)
                        {
                            sheetName = dtExcelSchema.Rows[i]["TABLE_NAME"].ToString();
                            if (sheetName.Contains("Область_печати") || sheetName.Contains("Print_Area"))
                            {
                                continue;
                            }
                            Console.WriteLine(sheetName);
                            DataTable dt = new DataTable(sheetName);
                            cmd.Connection = con;
                            cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                            oda.SelectCommand = cmd;
                            Console.WriteLine(oda.ToString());
                            try
                            {
                                oda.Fill(dt);
                                dt.TableName = sheetName;
                                ds.Tables.Add(dt);
                            }
                            catch (System.Data.OleDb.OleDbException) { }
                            finally
                            {

                            }


                        }
                    }
                }
            }
            if (ds != null && ds.Tables[0] != null)
            {
                OleDbConnection conn = new OleDbConnection(this.DSN);
                conn.Open();
                var group_cache = new Dictionary<string, StudyGroup>() { };
                var teacher_cache = new Dictionary<string, MyGroup>() { };
                var subject_cache = new Dictionary<string, MyGroup>() { };
                var room_cache = new Dictionary<string, MyGroup>() { };
                foreach (DataTable table in ds.Tables)
                {
                    var group_index = new Dictionary<int, StudyGroup>() { };
                    Console.WriteLine("Table name: " + table.TableName);
                    int current_week = -2;
                    int first_lesson_column = -2;  // -2 -> undefined
                    int last_lesson_column = -2;
                    int table_header_row = -2;
                    int last_lesson = 6;
                    int current_weekday = 0;  // sunday
                    foreach (DataRow row in table.Rows)
                    {
                        int current_row = table.Rows.IndexOf(row);
                        int current_lesson = -2;
                        if (current_week > 0)
                        {
                            string value = row[first_lesson_column - 1].ToString();
                            Match match = Regex.Match(value, "(\\d{1})");
                            if (match.Success == true)
                            {
                                current_lesson = Int32.Parse(match.Groups[1].Value);
                                last_lesson = current_lesson;
                                if (current_lesson == 1)
                                {
                                    current_weekday++;
                                    Console.WriteLine("Day of week: " + ((DayOfWeek)current_weekday).ToString());
                                }
                                Console.WriteLine(current_lesson);
                                for (int i = 0; i <= last_lesson_column - first_lesson_column; i++)
                                {
                                    var lesson = row[first_lesson_column + i].ToString().Trim();
                                    if (lesson == string.Empty || lesson == ",")
                                    {
                                        // Console.WriteLine("empty lesson");
                                        continue;
                                    }
                                    string[] lessons = Regex.Split(lesson, "\\-\\-+[\\r\\n]*");
                                    var ParsedLessons = new Dictionary<int, ScheduleRecord>() { };
                                    int j = 0;
                                    foreach(string sublesson in lessons)
                                    {
                                        // 2п. Объектно-ориентированнное программирование  лаб 303    Г.А.Жаркова  
                                        match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-ЯA-Z\\d][А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+?)[\\s\\r\\n]+(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|л)[\\s\\r\\n]+([Ааод\\d][\\s\\d/\\.а-я]+?)[\\s\\r\\n]*([А-Яа-я]+[\\.,]?\\s?[\\r\\n]*[А-Я]\\.?\\.?[А-Яа-яЁё\\-]+)[\\s\\r\\n]*$");
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-ЯA-Z\\d][А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+?)[\\s\\r\\n]+(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|л)[\\s\\r\\n]+([Ааод\\d][\\s\\d/\\.а-я]+?)[\\s\\r\\n]+([А-Яа-яЁё\\-]+\\s?[А-Я]\\.?[А-Я]\\.\\.?)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord() {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Type = match.Groups[4].ToString().Trim(),
                                                Room = match.Groups[5].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                Teacher = match.Groups[6].ToString().Trim(),
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+?)[\\s\\r\\n]+([А-Я]?\\.?[А-Я]?\\.?[\\.\\s]?[А-Я][а-яЁё\\-]{3,})\\.?[\\s\\r\\n]+(сем|Сем|лаб|Лаб|лек|сем\\.|cем|сем\\s+сем|с|с\\.|лаб\\.|лек\\.|л\\.|л(?!аб))[\\s\\r\\n]*([АаоПд\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Type = match.Groups[4].ToString().Trim(),
                                                Room = match.Groups[5].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                Teacher = match.Groups[6].ToString().Trim(),
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+?)[\\s\\r\\n]+([А-Яа-яЁё\\-]+\\s{0,2}[А-Я]\\.?[А-Я]\\.?\\.?)[\\s\\r\\n]*(сем|Сем|лаб|Лаб|лек|сем\\.|сем\\s+сем|с|с\\.|лаб\\.|лек\\.|л\\.|л)[\\s\\r\\n]*([АаоПд\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Teacher = match.Groups[4].ToString().Trim(),
                                                Type = match.Groups[5].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                Room = match.Groups[6].ToString().Trim(),
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }

                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-ЯA-Z][А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+?)[\\s\\r\\n]+(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|л)[\\s\\r\\n]+([А-Я]\\.?[А-Я]\\.\\.?[А-Яа-яЁё\\-]+)[\\s\\r\\n]+([Аа\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Teacher = match.Groups[4].ToString().Trim(),
                                                Type = match.Groups[5].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                Room = match.Groups[6].ToString().Trim(),
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+?)[\\s\\r\\n]+(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|л)[\\s\\r\\n]+([А-Яа-яЁё\\-]+\\s?[А-Я]\\.?[А-Я]\\.\\.?)[\\s\\r\\n]+([Аа\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Type = match.Groups[4].ToString().Trim(),
                                                Teacher = match.Groups[5].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                Room = match.Groups[6].ToString().Trim(),
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+)[\\s\\r\\n]+([А-Я]\\.?[А-Я]\\.\\.?[А-Яа-яЁё\\-]+)[\\s\\r\\n]+([Аа\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]+(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|л)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Type = match.Groups[4].ToString().Trim(),
                                                Teacher = match.Groups[5].ToString().Trim(),
                                                Room = match.Groups[6].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+)[\\s\\r\\n]+([Аа\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]+(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|л)[\\s\\r\\n]+([А-Я]\\.?[А-Я]\\.\\.?[А-Яа-яЁё\\-]+)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Teacher = match.Groups[4].ToString().Trim(),
                                                Room = match.Groups[5].ToString().Trim(),
                                                Type = match.Groups[6].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+)[\\s\\r\\n]+([А-Я]\\.?[А-Я]\\.\\.?[А-Яа-яЁё\\-]+)[\\s\\r\\n]+([Аа\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Room = match.Groups[4].ToString().Trim(),
                                                Type = match.Groups[5].ToString().Trim(),
                                                Teacher = match.Groups[6].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+)[\\s\\r\\n]+([А-Яа-яЁё\\-]+\\s?[А-Я]\\.?[А-Я]\\.\\.?)[\\s\\r\\n]+([Аа\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Type = "лаб",
                                                Teacher = match.Groups[4].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                Room = match.Groups[5].ToString().Trim(),
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "((\\d)\\s{0,2}п\\.?\\s*)?([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+)[\\s\\r\\n]+([Аа\\d][\\s\\d/\\.а-я]+?)[\\s\\r\\n]+([А-Я]\\.?[А-Я]\\.\\.?[А-Яа-яЁё\\-]+)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Type = "лаб",
                                                Teacher = match.Groups[4].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                Room = match.Groups[5].ToString().Trim(),
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+)[\\s\\r\\n]+([А-Я]\\.?[А-Я]\\.\\.?[А-Яа-яЁё\\-]+)[\\s\\r\\n]+(\\d)\\s?п\\.?\\s?(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|сем \\.|л)\\s+([Аа\\d][\\s\\d/\\.а-я]+?),[\\s\\r\\n]+(\\d)\\s?п\\.?\\s?(сем|лаб|Лаб|лек|сем\\.|с|лаб\\.|лек\\.|л\\.|сем \\.|л)\\s+([Аа\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]*");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[3].ToString().Trim(),
                                                Type = "лаб",
                                                Room = match.Groups[4].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                Teacher = match.Groups[5].ToString().Trim(),
                                                LessonNumber = current_lesson
                                            };
                                            if (match.Groups[2].ToString() != string.Empty)
                                            {
                                                schedule_record.Subgroup = Int32.Parse(match.Groups[2].ToString().Trim());
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        /// 
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+)[\\s\\r\\n]+([А-Я]\\.?[А-Я]\\.\\.?[А-Яа-яЁё\\-]+)[\\s\\r\\n]+(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|сем \\.|л),[\\s\\r\\n]+(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|сем \\.|л)\\s+([Аа\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]*");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[1].ToString().Trim(),
                                                Teacher = match.Groups[2].ToString().Trim(),
                                                LessonNumber = Int32.Parse(match.Groups[3].ToString().Trim()),
                                                Type = match.Groups[4].ToString().Trim(),
                                                Room = match.Groups[5].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[1].ToString().Trim(),
                                                Teacher = match.Groups[2].ToString().Trim(),
                                                LessonNumber = Int32.Parse(match.Groups[6].ToString().Trim()),
                                                Type = match.Groups[7].ToString().Trim(),
                                                Room = match.Groups[8].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        ///
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+)[\\s\\r\\n]+(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|сем \\.|л)\\s+([АаодП\\d][\\s\\d/\\.а-я]+)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[1].ToString().Trim(),
                                                Teacher = match.Groups[2].ToString().Trim(),
                                                LessonNumber = current_lesson,
                                                Type = match.Groups[3].ToString().Trim(),
                                                Room = match.Groups[5].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[1].ToString().Trim(),
                                                Teacher = match.Groups[2].ToString().Trim(),
                                                LessonNumber = current_lesson + 1,
                                                Type = match.Groups[4].ToString().Trim(),
                                                Room = match.Groups[5].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+)[\\s\\r\\n]+([А-Я]\\.?[А-Я]\\.\\.?[А-Яа-яЁё\\-]+)[\\s\\r\\n]+(сем|лаб|Лаб|лек|сем\\.|с|с\\.|лаб\\.|лек\\.|л\\.|сем \\.|л)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[1].ToString().Trim(),
                                                Type = match.Groups[2].ToString().Trim(),
                                                Room = match.Groups[3].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                LessonNumber = current_lesson
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            match = Regex.Match(sublesson, "([А-Яа-яЁёA-Za-z\\d:/\\(\\)\\-\\.,\\s]+)[\\s\\r\\n]*$");
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[1].ToString().Trim(),
                                                Teacher = match.Groups[2].ToString().Trim(),
                                                Type = match.Groups[3].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                LessonNumber = current_lesson
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                        if (match.Success != true)
                                        {
                                            Console.WriteLine("Failed to parse: " + sublesson);
                                        }
                                        else
                                        {
                                            var schedule_record = new ScheduleRecord()
                                            {
                                                Subject = match.Groups[1].ToString().Trim(),
                                                WeekNumber = current_week,
                                                DayOfWeek = current_weekday,
                                                LessonNumber = current_lesson
                                            };
                                            ParsedLessons[j++] = schedule_record;
                                            continue;
                                        }
                                    }
                                    if (ParsedLessons.Values.Count == 1 && ParsedLessons.Values.First().Subgroup == 0)
                                    {
                                        ScheduleRecord record = ParsedLessons[0];
                                        record.Subgroup = 2;
                                        ParsedLessons[1] = record;
                                        record.Subgroup = 1;
                                        ParsedLessons[0] = record;
                                    }
                                    foreach (ScheduleRecord record in ParsedLessons.Values)
                                    {
                                        int teacher_id = 0;
                                        int room_id = 0;
                                        int subject_id = 0;
                                        OleDbCommand cmd;
                                        if (record.Teacher != string.Empty && record.Teacher != null)
                                        {
                                            string teacher = this.NormalizeName(record.Teacher);
                                            if (teacher_cache.ContainsKey(teacher) != true)
                                            {
                                                MyGroup teacher_group = new MyGroup()
                                                {
                                                    ObjectName = teacher,
                                                    ID = -1
                                                };
                                                cmd = new OleDbCommand();
                                                cmd.CommandType = CommandType.Text;
                                                cmd.CommandText = "INSERT INTO [Teacher] ([Name]) VALUES (?);";
                                                cmd.Parameters.Add("@Name", OleDbType.VarChar).Value = teacher;
                                                cmd.Connection = conn;
                                                cmd.ExecuteNonQuery();
                                                cmd.CommandText = "SELECT @@Identity";
                                                teacher_group.ID = (int)cmd.ExecuteScalar();
                                                teacher_cache[teacher] = teacher_group;
                                            }
                                            teacher_id = teacher_cache[teacher].ID;
                                        }
                                        if (record.Room != string.Empty && record.Room != null)
                                        {
                                            if (room_cache.ContainsKey(record.Room) != true)
                                            {
                                                MyGroup room_group = new MyGroup()
                                                {
                                                    ObjectName = record.Room,
                                                    ID = -1
                                                };
                                                cmd = new OleDbCommand();
                                                cmd.CommandType = CommandType.Text;
                                                cmd.CommandText = "INSERT INTO [Room] ([Name]) VALUES (?);";
                                                cmd.Parameters.Add("@Name", OleDbType.VarChar).Value = record.Room;
                                                cmd.Connection = conn;
                                                cmd.ExecuteNonQuery();
                                                cmd.CommandText = "SELECT @@Identity";
                                                room_group.ID = (int)cmd.ExecuteScalar();
                                                room_cache[record.Room] = room_group;
                                            }
                                            room_id = room_cache[record.Room].ID;
                                        }
                                        if (record.Subject != string.Empty && record.Subject != null)
                                        {
                                            if (subject_cache.ContainsKey(record.Subject) != true)
                                            {
                                                MyGroup subject_group = new MyGroup()
                                                {
                                                    ObjectName = record.Subject,
                                                    ID = -1
                                                };
                                                cmd = new OleDbCommand();
                                                cmd.CommandType = CommandType.Text;
                                                cmd.CommandText = "INSERT INTO [Subject] ([Name]) VALUES (?);";
                                                cmd.Parameters.Add("@Name", OleDbType.VarChar).Value = record.Subject;
                                                cmd.Connection = conn;
                                                cmd.ExecuteNonQuery();
                                                cmd.CommandText = "SELECT @@Identity";
                                                subject_group.ID = (int)cmd.ExecuteScalar();
                                                subject_cache[record.Subject] = subject_group;
                                            }
                                            subject_id = subject_cache[record.Subject].ID;
                                        }
                                        cmd = new OleDbCommand();
                                        cmd.CommandType = CommandType.Text;
                                        cmd.CommandText = "INSERT INTO [Schedule] ([WeekNumber], [DayOfWeek], [LessonNumber], [LessonType], [Group_ID], [Subject_ID], [Room_ID], [Teacher_ID]) VALUES (?, ?, ?, ?, ?, ?, ?, ?);";
                                        cmd.Parameters.Add("@WeekNumber", OleDbType.Integer).Value = record.WeekNumber;
                                        cmd.Parameters.Add("@DayOfWeek", OleDbType.Integer).Value = record.DayOfWeek;
                                        cmd.Parameters.Add("@LessonNumber", OleDbType.Integer).Value = record.LessonNumber;
                                        string norm_type = this.NormalizeType(record.Type);
                                        if (norm_type == string.Empty || norm_type == null)
                                        {
                                            cmd.Parameters.Add("@LessonType", OleDbType.VarChar).Value = DBNull.Value;
                                        } else
                                        {
                                            cmd.Parameters.Add("@LessonType", OleDbType.VarChar).Value = norm_type;
                                        }
                                        if (record.Subgroup == 1)
                                        {
                                            cmd.Parameters.Add("@Group_ID", OleDbType.Integer).Value = group_index[i].ID1;
                                        } else
                                        {
                                            cmd.Parameters.Add("@Group_ID", OleDbType.Integer).Value = group_index[i].ID2;
                                        }
                                        cmd.Parameters.Add("@Subject_ID", OleDbType.Integer).Value = subject_id;
                                        if (room_id == 0)
                                        {
                                            cmd.Parameters.Add("@Room_ID", OleDbType.VarChar).Value = DBNull.Value;
                                        }
                                        else
                                        {
                                            cmd.Parameters.Add("@Room_ID", OleDbType.VarChar).Value = room_id;
                                        }
                                        if (teacher_id == 0)
                                        {
                                            cmd.Parameters.Add("@Teacher_ID", OleDbType.VarChar).Value = DBNull.Value;
                                        }
                                        else
                                        {
                                            cmd.Parameters.Add("@Teacher_ID", OleDbType.VarChar).Value = teacher_id;
                                        }
                                        cmd.Connection = conn;
                                        cmd.ExecuteNonQuery();
                                        Console.WriteLine(record.Subgroup + "п, " + record.Subject + ", " + record.Teacher + ", " + record.Room + ", " + record.Type);
                                    }
                                }
                            }
                            else
                            {
                                if (last_lesson < 7)
                                {
                                    continue;
                                }
                                break;
                            }
                            continue;
                        }
                        foreach (DataColumn column in table.Columns)
                        {
                            string value = row[column].ToString();
                            if (table_header_row == -2 || current_row == table_header_row)  // first of all - find table header
                            {
                                Match match = Regex.Match(value, "([А-Я]{2,7})\\s*-\\s*([А-Я])\\s*-\\s*(\\d{2,4})\\s*/\\s*(\\d{1})");
                                if (match.Success == true)
                                {
                                    table_header_row = current_row;
                                    string speciality = match.Groups[1].Value;
                                    string form = match.Groups[2].Value;
                                    string year = match.Groups[3].Value;
                                    string iteration = match.Groups[4].Value;
                                    string groupName = speciality + '-' + form +  '-' + year + '/' + iteration;
                                    Console.WriteLine(groupName);
                                    int current_column = table.Columns.IndexOf(column);
                                    if (first_lesson_column == -2)
                                    {
                                        first_lesson_column = current_column;
                                        for(int i = current_column - 1; i >= 0; i--)
                                        {
                                            match = Regex.Match(row[i].ToString(), "(\\d)");
                                            if (match.Success == true)
                                            {
                                                current_week = Int32.Parse(match.Groups[1].Value);
                                                //Console.WriteLine(current_week + " current week");
                                            }
                                        }
                                    }
                                    if (group_cache.ContainsKey(groupName) != true) {
                                        StudyGroup study_group = new StudyGroup() {
                                            GroupName = groupName,
                                            ID1 = -1,
                                            ID2 = -1
                                        };
                                        OleDbCommand cmd = new OleDbCommand();
                                        cmd.CommandType = CommandType.Text;
                                        cmd.CommandText = "INSERT INTO [Group] ([Name],[Year],[Speciality],[Iteration],[Subgroup]) VALUES (?,?,?,?,?);";
                                        cmd.Parameters.Add("@Name", OleDbType.VarChar).Value = groupName;
                                        cmd.Parameters.Add("@Year", OleDbType.Integer).Value = Int32.Parse(year);
                                        cmd.Parameters.Add("@Speciality", OleDbType.VarChar).Value = speciality;
                                        cmd.Parameters.Add("@Iteration", OleDbType.Integer).Value = Int32.Parse(iteration);
                                        cmd.Parameters.Add("@Subgroup", OleDbType.Integer).Value = 1;
                                        cmd.Connection = conn;
                                        cmd.ExecuteNonQuery();
                                        cmd.CommandText = "SELECT @@Identity";
                                        study_group.ID1 = (int)cmd.ExecuteScalar();
                                        cmd = new OleDbCommand();
                                        cmd.CommandType = CommandType.Text;
                                        cmd.CommandText = "INSERT INTO [Group] ([Name],[Year],[Speciality],[Iteration],[Subgroup]) VALUES (?,?,?,?,?);";
                                        cmd.Parameters.Add("@Name", OleDbType.VarChar).Value = groupName;
                                        cmd.Parameters.Add("@Year", OleDbType.Integer).Value = Int32.Parse(year);
                                        cmd.Parameters.Add("@Speciality", OleDbType.VarChar).Value = speciality;
                                        cmd.Parameters.Add("@Iteration", OleDbType.Integer).Value = Int32.Parse(iteration);
                                        cmd.Parameters.Add("@Subgroup", OleDbType.Integer).Value = 2;
                                        cmd.Connection = conn;
                                        cmd.ExecuteNonQuery();
                                        cmd.CommandText = "SELECT @@Identity";
                                        study_group.ID2 = (int)cmd.ExecuteScalar();
                                        group_cache[groupName] = study_group;
                                    }
                                    group_index[current_column - first_lesson_column] = group_cache[groupName];
                                    last_lesson_column = current_column;
                                }

                            }
                            
                        }
                    }
                }
                conn.Close();
            }
        }

        private void importFromXLSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string message = "Перезаписать базу данных расписания?";
            string caption = "Импорт данных из XLS";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result;

            result = MessageBox.Show(message, caption, buttons);
            if (result != System.Windows.Forms.DialogResult.Yes)
            {
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileDialog.Filter = XlsxMask;
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    this.initDB();
                    this.importFromXLS(openFileDialog.FileName);
                }
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // TODO close DB
            this.Close();
        }

        private void checkForErrorsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(this.DSN);
            conn.Open();

            var group_cache = new Dictionary<int, string>() { };

            OleDbCommand command = new OleDbCommand("SELECT [Group_ID], [Name] FROM [Group];", conn);

            OleDbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                int Group_ID = reader.GetInt32(0);
                string Name = reader.GetValue(1).ToString();
                group_cache[Group_ID] = Name;
            };

            command = new OleDbCommand("SELECT Schedule.Schedule_ID, Schedule.WeekNumber, Schedule.DayOfWeek, Schedule.LessonNumber, Schedule.LessonType, Schedule.Group_ID, Group.Name, Room.Name, Subject.Name, Teacher.Name FROM Teacher INNER JOIN(Subject INNER JOIN(Room INNER JOIN([Group] INNER JOIN Schedule ON Group.Group_ID = Schedule.Group_ID) ON Room.Room_ID = Schedule.Room_ID) ON Subject.Subject_ID = Schedule.Subject_ID) ON Teacher.Teacher_ID = Schedule.Teacher_ID;", conn);

            reader = command.ExecuteReader();

            var single_group = new Dictionary<string, ScheduleRecord>() { };
            var single_teacher = new Dictionary<string, ScheduleRecord>() { };
            var single_room = new Dictionary<string, ScheduleRecord>() { };
            var single_lab = new Dictionary<string, ScheduleRecord>() { };
            string found_errors = string.Empty;
            errors.Text = "Поиск ошибок...";

            while (reader.Read())
            {
                ScheduleRecord record = new ScheduleRecord()
                {
                    ID = reader.GetInt32(0),
                    WeekNumber = reader.GetInt32(1),
                    DayOfWeek = reader.GetInt32(2),
                    LessonNumber = reader.GetInt32(3),
                    Type = reader.GetString(4),
                    Subgroup = reader.GetInt32(5),
                    Room = reader.GetValue(7).ToString(),
                    Teacher = reader.GetValue(9).ToString(),
                };
                // 1.1. Никакая группа не может находится в одном кабинете с двумя и более преподавателями на одной паре,
                // в одно и то же время; 
                // 1.2. Никакая группа не может находится в разных кабинетах с одним преподавателем на одной паре,
                // в одно и то же время;
                // Group_Name+WeekNumber+DayOfWeek+LessonNumber
                string key1 = record.Subgroup + "+" + record.WeekNumber + "+" + record.DayOfWeek + "+" + record.LessonNumber;
                if (single_group.ContainsKey(key1) == true) {
                    found_errors += "1. Record duplicates: \r\n" +
                        record.ID + ". group: " + group_cache[record.Subgroup] + ", week: " + record.WeekNumber + ", day of week: " + ((DayOfWeek)record.DayOfWeek).ToString() + ", lesson number: " + record.LessonNumber + ", teacher: " + record.Teacher + ", room: " + record.Room + "\r\n" +
                        single_group[key1].ID + ". group: " + group_cache[single_group[key1].Subgroup] + ", week: " + single_group[key1].WeekNumber + ", day of week: " + ((DayOfWeek)single_group[key1].DayOfWeek).ToString() + ", lesson number: " + single_group[key1].LessonNumber + ", teacher: " + single_group[key1].Teacher + ", room: " + single_group[key1].Room + "\r\n";
                } else
                {
                    single_group[key1] = record;
                }

                // 2.1. Никакой преподаватель не может сразу присутствовать в разных кабинетах у одной группы;
                // 2.2. Никакой преподаватель не может сразу присутствовать в разных кабинетах у нескольких групп;
                // Teacher_ID+WeekNumber+DayOfWeek+LessonNumber
                string key2 = record.Teacher + "+" + record.WeekNumber + "+" + record.DayOfWeek + "+" + record.LessonNumber;
                if (single_teacher.ContainsKey(key2) == true)
                {
                    if (record.Room != single_teacher[key2].Room)
                    {
                        found_errors += "2. Record duplicates: \r\n" +
                            record.ID + ". group: " + group_cache[record.Subgroup] + ", week: " + record.WeekNumber + ", day of week: " + ((DayOfWeek)record.DayOfWeek).ToString() + ", lesson number: " + record.LessonNumber + ", teacher: " + record.Teacher + ", room: " + record.Room + "\r\n" +
                            single_teacher[key2].ID + ". group: " + group_cache[single_teacher[key2].Subgroup] + ", week: " + single_teacher[key2].WeekNumber + ", day of week: " + ((DayOfWeek)single_teacher[key2].DayOfWeek).ToString() + ", lesson number: " + single_teacher[key2].LessonNumber + ", teacher: " + single_teacher[key2].Teacher + ", room: " + single_teacher[key2].Room + "\r\n";
                    }
                }
                else
                {
                    single_teacher[key2] = record;
                }

                // ignore without room_id or teacher_id
                // WeekNumber+DayOfWeek+LessonNumber+Room_ID=Teacher_ID
                if (record.Room != null && record.Room != string.Empty && record.Teacher != null && record.Teacher != string.Empty)
                {
                    string key3 = record.Room + "+" + record.WeekNumber + "+" + record.DayOfWeek + "+" + record.LessonNumber;
                    if (single_room.ContainsKey(key3) == true)
                    {
                        if (record.Teacher != single_room[key3].Teacher)  // record.Group_Name != group_cache[Group_ID]
                        {
                            found_errors += "3. Record duplicates: \r\n" +
                                record.ID + ". group: " + group_cache[record.Subgroup] + ", week: " + record.WeekNumber + ", day of week: " + ((DayOfWeek)record.DayOfWeek).ToString() + ", lesson number: " + record.LessonNumber + ", teacher: " + record.Teacher + ", room: " + record.Room + "\r\n" +
                                single_room[key3].ID + ". group: " + group_cache[single_room[key3].Subgroup] + ", week: " + single_room[key3].WeekNumber + ", day of week: " + ((DayOfWeek)single_room[key3].DayOfWeek).ToString() + ", lesson number: " + single_room[key3].LessonNumber + ", teacher: " + single_room[key3].Teacher + ", room: " + single_room[key3].Room + "\r\n";
                        }
                    }
                    else
                    {
                        single_room[key3] = record;
                    }
                }

                // WeekNumber+DayOfWeek+LessonNumber+Room_ID
                if (record.Room != null && record.Room != string.Empty)
                {
                    string key4 = record.Room + "+" + record.WeekNumber + "+" + record.DayOfWeek + "+" + record.LessonNumber;
                    if (single_lab.ContainsKey(key4) == true && (single_lab[key4].Type == "лаб" || record.Type == "лаб") && (group_cache[single_lab[key4].Subgroup] != group_cache[record.Subgroup]))
                    {
                        found_errors += "4. Record duplicates: \r\n" +
                            record.ID + ". group: " + group_cache[record.Subgroup] + ", week: " + record.WeekNumber + ", day of week: " + ((DayOfWeek)record.DayOfWeek).ToString() + ", lesson number: " + record.LessonNumber + ", teacher: " + record.Teacher + ", room: " + record.Room + "\r\n" +
                            single_lab[key4].ID + ". group: " + group_cache[single_lab[key4].Subgroup] + ", week: " + single_lab[key4].WeekNumber + ", day of week: " + ((DayOfWeek)single_lab[key4].DayOfWeek).ToString() + ", lesson number: " + single_lab[key4].LessonNumber + ", teacher: " + single_lab[key4].Teacher + ", room: " + single_lab[key4].Room + "\r\n";
                    } else
                    {
                        single_lab[key4] = record;
                    }
                }

            }
            if (found_errors != string.Empty)
            {
                errors.Text = found_errors;
            } else
            {
                errors.Text = "Ошибки не найдены";
            }
            reader.Close();
            conn.Close();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // TODO show about window
        }
    }
}
