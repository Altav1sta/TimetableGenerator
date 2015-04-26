using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace TimeTableGenerating
{
    public class Generator
    {
        private int count;
        private double[] w;
        public string connectionString;
        public Lesson[, ,] timetable;

        public struct Lesson
        {
            public readonly int id;
            public readonly bool isLect;
            public readonly int tId;
            public readonly int room;

            public Lesson(int id, bool isLection, int teacher, int room)
            {
                this.id = id;
                isLect = isLection;
                tId = teacher;
                this.room = room;
            }

            public string toString(string conStr)
            {
                OleDbConnection con = new OleDbConnection(conStr);
                con.Open();
                OleDbCommand cmd = new OleDbCommand("SELECT Lesson FROM Lessons WHERE Lessons.ID = " + id, con);
                string title = Convert.ToString(cmd.ExecuteScalar());
                cmd = new OleDbCommand("SELECT TName FROM Teachers WHERE Teachers.ID = " + tId, con);
                string teacher = Convert.ToString(cmd.ExecuteScalar());
                string lection;
                if (isLect) lection = ", лекция, ";
                else lection = ", практика, ";
                con.Close();

                return title + lection + teacher + ", " + room.ToString();
            }
        }

        public Generator(int count, double[] w, string conStr)
        {
            this.count = count;
            this.w = new double[w.Length];
            for (int i = 0; i < w.Length; i++)
                this.w[i] = w[i];
            connectionString = conStr;
        }
        

        public void generateTimeTable()
        {
            OleDbConnection con = new OleDbConnection(connectionString);
            con.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT count(*) FROM Lessons", con);
            int[] lessonsID = new int[Convert.ToInt32(cmd.ExecuteScalar())];

            cmd = new OleDbCommand("SELECT ID FROM Lessons", con);
            OleDbDataReader odr = cmd.ExecuteReader();
            for (int i = 0; i < lessonsID.Length; i++)
            {
                odr.Read();                
                lessonsID[i] = odr.GetInt32(0);
            }
            
            con.Close();

            lessonsID = sortLessonsByDegreeOfFreedom(lessonsID);


            timetable = new Lesson[count, 5, 5];
            for (int i = 0; i < lessonsID.Length; i++)
            {
                con.Open();
                cmd = new OleDbCommand("SELECT * FROM Lessons WHERE Lessons.ID = " + lessonsID[i], con);
                odr = cmd.ExecuteReader();
                odr.Read();
                bool isLect = odr.GetBoolean(2);
                int group = odr.GetInt32(3);
                int tID = odr.GetInt32(4);


                // Вытащим массив всех подходящих аудиторий                
                cmd = new OleDbCommand("SELECT Room FROM Rooms WHERE (Capacity >= " +
                                                        "(SELECT Population FROM Groups WHERE Number = " + group + ")" +
                                                    ") AND (Projector = " +
                                                        "(SELECT Projector FROM Lessons WHERE ID = " + lessonsID[i] + ")" +
                                                    ") AND (Laboratory = " +
                                                        "(SELECT Laboratory FROM Lessons WHERE ID = " + lessonsID[i] + ")" +
                                                    ") AND (Computers = " +
                                                        "(SELECT Computers FROM Lessons WHERE ID = " + lessonsID[i] + ")" +
                                                    ") AND (Gym = " +
                                                        "(SELECT Gym FROM Lessons WHERE ID = " + lessonsID[i] + "))", con);
                OleDbDataReader reader = cmd.ExecuteReader();
                List<int> rooms = new List<int>();
                while (reader.Read())
                {
                    rooms.Add(reader.GetInt32(0));
                }


                // Заменим фактический номер группы на ее порядковый номер, как индекс массива
                int groupIndex = 0;
                cmd = new OleDbCommand("SELECT Number FROM Groups", con);
                reader = cmd.ExecuteReader();
                for (int num = 0; num < count; num++)
                {
                    reader.Read();
                    if (group == reader.GetInt32(0))
                    {
                        groupIndex = num;
                        break;
                    }
                }

                con.Close();

                                
                
                double max = -1.0;
                int r_max = -1;
                int day_max = -1;
                int pos_max = -1;
                foreach (int r in rooms)
                {
                    // Подберем нужное время
                    double R_max = -10.0;
                    int d = -1;
                    int p = -1;
                    double R;
                    for (int day = 0; day < 5; day++)
                    {
                        for (int pos = 0; pos < 5; pos++)
                        {
                            if (timetable[groupIndex, day, pos].id != 0) continue;
                            bool flag = false;
                            for (int t = 0; t < count; t++)
                            {
                                if ((timetable[t, day, pos].id != 0) && (t != groupIndex) && ((timetable[t, day, pos].room == r) || (timetable[t, day, pos].tId == tID)))
                                {
                                    flag = true;
                                    break;
                                }
                            }
                            if (!flag)
                            {
                                R = 0;
                                R += w[0] * measureOfQualityInTimeSpace(groupIndex, tID, 0, day, pos);
                                R += w[1] * measureOfQualityInTimeSpace(groupIndex, tID, 1, day, pos);
                                R += w[2] * measureOfQualityInTimeSpace(groupIndex, tID, 2, day, pos);
                                R += w[3] * measureOfQualityInTimeSpace(groupIndex, tID, 3, day, pos);
                                R += w[4] * measureOfQualityInTimeSpace(groupIndex, tID, 4, day, pos);

                                if (R > R_max)
                                {
                                    R_max = R;
                                    d = day;
                                    p = pos;
                                }
                            }
                        }
                    }

                    if (R_max + w[5] * measureOfQualityInRoomSpace(group, r) > max)
                    {               
                        max = R_max + w[5] * measureOfQualityInRoomSpace(group, r);
                        r_max = r;
                        day_max = d;
                        pos_max = p;
                    }
                }
                

                timetable[groupIndex, day_max, pos_max] = new Lesson(lessonsID[i], isLect, tID, r_max);
                
            }

        }
        
        public double getMeasureOfQuality()
        {
            // Getting indexes of groups
            int[] groups = new int[count];
            OleDbConnection con = new OleDbConnection(connectionString);
            con.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT Number FROM Groups", con);
            OleDbDataReader reader = cmd.ExecuteReader();
            int counter = 0;
            while (reader.Read())
            {
                groups[counter] = reader.GetInt32(0);
                counter++;
            }

            // Пересчитываем качество расписания
            double R_total = 0;
            for (int i = 0; i < count; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    for (int k = 0; k < 5; k++)
                    {
                        if (timetable[i, j, k].id == 0) continue;

                        Lesson tmp = timetable[i, j, k];
                        timetable[i, j, k] = new Lesson();
                        R_total += w[0] * measureOfQualityInTimeSpace(i, tmp.tId, 0, j, k);
                        R_total += w[1] * measureOfQualityInTimeSpace(i, tmp.tId, 1, j, k);
                        R_total += w[2] * measureOfQualityInTimeSpace(i, tmp.tId, 2, j, k);
                        R_total += w[3] * measureOfQualityInTimeSpace(i, tmp.tId, 3, j, k);
                        R_total += w[4] * measureOfQualityInTimeSpace(i, tmp.tId, 4, j, k);
                        R_total += w[5] * measureOfQualityInRoomSpace(groups[i], tmp.room);
                        timetable[i, j, k] = tmp;
                    }
                }
            }

            return R_total;
        }

        private double measureOfQualityInRoomSpace(int group, int room)
        {
            OleDbConnection con = new OleDbConnection(connectionString);
            con.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT Capacity FROM Rooms WHERE Room = " + room, con);
            int capacity = Convert.ToInt32(cmd.ExecuteScalar());
            cmd = new OleDbCommand("SELECT Population FROM Groups WHERE Number = " + group, con);
            int population = Convert.ToInt32(cmd.ExecuteScalar());
            con.Close();

            return 1.0 - (capacity - population) / (double)capacity;
        }

        private double measureOfQualityInTimeSpace(int group, int tID, int numberOfMeasure, int day, int lessonPosition)
        {                   
            switch (numberOfMeasure)
            {
                // Появление окна в расписании группы студентов
                case 0:
                    // Ситуация, когда одно окно в 3 занятия дробится на 2 окна по 1 занятию
                    if ((lessonPosition == 2) && (timetable[group, day, 0].id != 0) && (timetable[group, day, 1].id == 0) && (timetable[group, day, 3].id == 0) && (timetable[group, day, 4].id != 0))
                    {
                        return 0;
                    }
                    // Остальные ситуации
                    else
                    {
                        int first = -10, last = -10;
                        for (int i = 0; i < 5; i++)
                        {
                            if (timetable[group, day, i].id != 0) 
                            {
                                last = i;
                                if (first == -10) first = i;
                            }
                        }

                        if ((lessonPosition < first - 1) || (lessonPosition > last + 1))
                            return 0;
                        else
                            return 1.0;
                    }

                // Появление окна в расписании преподавателя
                case 1:
                    // Сначала узнаем расписание преподавателя на день
                    Lesson[] tmpArr = new Lesson[5];
                    for (int i = 0; i < count; i++)
                    {
                        for (int j = 0; j < 5; j++)
                        {
                            if ((timetable[i, day, j].id != 0) && (timetable[i, day, j].tId == tID))
                            {
                                tmpArr[j] = timetable[i, day, j];
                            }
                        }
                    }
                    // Ситуация, когда одно окно в 3 занятия дробится на 2 окна по 1 занятию
                    if ((lessonPosition == 2) && (tmpArr[0].id != 0) && (tmpArr[1].id == 0) && (tmpArr[3].id == 0) && (tmpArr[4].id != 0))
                    {
                        return 0;
                    }
                    // Остальные ситуации
                    else
                    {
                        int first = -10, last = -10;
                        for (int i = 0; i < 5; i++)
                        {
                            if (tmpArr[i].id != 0)
                            {
                                last = i;
                                if (first == -10) first = i;
                            }
                        }

                        if ((lessonPosition == first - 2) || (lessonPosition == last + 2))
                            return 0;
                        else
                            return 1.0;
                    }

                // Исчезновения окна в расписании группы студентов
                case 2:
                    if ((lessonPosition == 0) || (lessonPosition == 4))
                    {
                        return 0;
                    }
                    else
                    {
                        if ((timetable[group, day, lessonPosition - 1].id != 0) && (timetable[group, day, lessonPosition + 1].id != 0))
                            return 1.0;
                        else
                            return 0;
                    }

                // Исчезновение окна в расписании преподавателя
                case 3:
                    // Сначала узнаем расписание преподавателя на день
                    tmpArr = new Lesson[5];
                    for (int i = 0; i < count; i++)
                    {
                        for (int j = 0; j < 5; j++)
                        {
                            if ((timetable[i, day, j].id != 0) && (timetable[i, day, j].tId == tID))
                            {
                                tmpArr[j] = timetable[i, day, j];
                            }
                        }
                    }
                    if ((lessonPosition == 0) || (lessonPosition == 4))
                    {
                        return 0;
                    }
                    else
                    {
                        if ((tmpArr[lessonPosition - 1].id != 0) && (tmpArr[lessonPosition + 1].id != 0))
                            return 1.0;
                        else
                            return 0;
                    }

                // Загруженность расписания группы студентов в этот день
                case 4:
                    int sum = 1;
                    for (int p = 0; p < 5; p++)
                    {
                        if ((p != lessonPosition) && (timetable[group, day, p].id != 0))
                            sum++;
                    }
                    return (Math.Sqrt(Math.Exp(5)) + 1 - Math.Sqrt(Math.Exp(sum))) / Math.Sqrt(Math.Exp(5));

                default:
                    return 0;               
            }
        }

        private int[] sortLessonsByDegreeOfFreedom(int[] lessonsID)
        {
            double[] degreeOfFreedom = new double[lessonsID.Length];
            int[] sortedID = new int[lessonsID.Length];
            OleDbConnection con = new OleDbConnection(connectionString);
            con.Open();

            // Вычисляем степени свободы
            for (int i = 0; i < lessonsID.Length; i++)
            {                                                                  
                OleDbCommand cmd = new OleDbCommand("SELECT count(*) FROM Rooms WHERE (Capacity >= " + 
                                                        "(SELECT Population FROM Groups WHERE Number = " +
                                                            "(SELECT Group FROM Lessons WHERE ID = " + lessonsID[i] + ")" +
                                                        ")" +
                                                    ") AND (Projector = " + 
                                                        "(SELECT Projector FROM Lessons WHERE ID = " + lessonsID[i] + ")" + 
                                                    ") AND (Laboratory = " + 
                                                        "(SELECT Laboratory FROM Lessons WHERE ID = " + lessonsID[i] + ")" +
                                                    ") AND (Computers = " +
                                                        "(SELECT Computers FROM Lessons WHERE ID = " + lessonsID[i] + ")" +
                                                    ") AND (Gym = " +
                                                        "(SELECT Gym FROM Lessons WHERE ID = " + lessonsID[i] + "))", con);
                int a = Convert.ToInt32(cmd.ExecuteScalar());
                if (a == 0)
                {
                    System.Windows.Forms.MessageBox.Show("Для предмета с индексом " + lessonsID[i].ToString() + " невозможно найти аудиторию!");
                    break;
                }
                
                cmd = new OleDbCommand("SELECT count(*) FROM Lessons WHERE Lessons.Group = (SELECT Group FROM Lessons WHERE ID = " + lessonsID[i] + ")", con);
                int g = Convert.ToInt32(cmd.ExecuteScalar());
                
                cmd = new OleDbCommand("SELECT count(*) FROM Lessons WHERE Teacher = (SELECT Teacher FROM Lessons WHERE ID = " + lessonsID[i] + ")", con);
                int p = Convert.ToInt32(cmd.ExecuteScalar());
                
                degreeOfFreedom[i] = a / (double)(g * p);
            }

            con.Close();
            
            
            // Сортируем массив индексов занятий по возрастанию степени свободы
            for (int i = 0; i < lessonsID.Length; i++)
            {
                double minValue = degreeOfFreedom[0];
                int ind = 0;
                for (int j = 0; j < lessonsID.Length; j++)
                {
                    if ((minValue == 0) || (degreeOfFreedom[j] != 0) && (degreeOfFreedom[j] < minValue))
                    {
                        minValue = degreeOfFreedom[j];
                        ind = j;
                    }
                }
                degreeOfFreedom[ind] = 0;
                sortedID[i] = lessonsID[ind];
            }

            return sortedID;
        }

    }
}
