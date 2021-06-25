// Copyright © 2021 Gorokhov Vladislav Dmitrievich. All rights reserved.
// Contacts: <gorohov2017vladislav@yandex.ru>
// License: https://www.gnu.org/licenses/old-licenses/gpl-2.0.html

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MySqlConnector;
using Wsav_alg;
using Color = System.Drawing.Color;

namespace Algorythm
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView1.Select();
            Opacity = 0;
            Timer timer = new Timer();
            timer.Tick += new EventHandler((sender, e) => 
            { 
                if ((Opacity+=0.05d) >= 0.92) timer.Stop(); 
            });
            timer.Interval = 50;
            timer.Start();
        }

        private long countColumns;
        private string idColumnName;
        
        private string host;
        private int port;
        private string database;
        private string tablename;
        private string username;
        private string password;

        private void textBox_host_TextChanged(object sender, EventArgs e)
        {
            host = textBox_host.Text;
        }
        
        private void textBox_port_TextChanged(object sender, EventArgs e)
        {
            port = Int32.Parse(textBox_port.Text);
        }
        
        private void textBox_database_TextChanged(object sender, EventArgs e)
        {
            database = textBox_database.Text;
        }
        
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            tablename = textBox3.Text;
        }
        
        private void textBox_username_TextChanged(object sender, EventArgs e)
        {
            username = textBox_username.Text;
        }
        
        private void textBox_password_TextChanged(object sender, EventArgs e)
        {
            password = textBox_password.Text;
        }

        private void LoadData()
        {
            
            string connectString = "Server=" + host + ";Database=" + database + ";port=" + port + ";User ID=" + username + ";password=" + password;
            
            MySqlConnection myConnection = new MySqlConnection(connectString);
            
            string query = "SELECT COUNT(*) AS NUMBEROFCOLUMNS FROM INFORMATION_SCHEMA.COLUMNS WHERE table_name  = '" + tablename + "';";
            
            myConnection.Open();
            MySqlCommand command = new MySqlCommand(query, myConnection);
            countColumns =  (long) command.ExecuteScalar();
            myConnection.Close();

            DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
            col.Name = "str_num";
            col.HeaderText = "№ стратегии";
            dataGridView1.Columns.Add(col);
            for (int i = 1; i < countColumns; ++i)
            {
                col = new DataGridViewTextBoxColumn();
                col.Name = $"nature{i}";
                col.HeaderText = $"П{i}";
                dataGridView1.Columns.Add(col);
            }

            query = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE table_name='" + tablename + "' AND COLUMN_NAME LIKE '%id%';";
            myConnection.Open();
            command = new MySqlCommand(query, myConnection);
            idColumnName = (string) command.ExecuteScalar();
            myConnection.Close();
            
            query = "SELECT * FROM " + tablename + " ORDER BY " + idColumnName;
            myConnection.Open();
            command = new MySqlCommand(query, myConnection);
            
            MySqlDataReader reader = command.ExecuteReader();
            List<string[]> data = new List<string[]>();
            while (reader.Read())
            {
                data.Add(new string[countColumns]);
                for (int i = 0; i < countColumns; ++i)
                {
                    data[data.Count - 1][i] = reader[i].ToString();
                }
            }
            reader.Close();
            
            myConnection.Close();

            foreach (string[] str in data)
            {
                dataGridView1.Rows.Add(str);
            }
            
            dataGridView1.AutoResizeColumns();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (host != null && port != 0 && database != null && tablename != null && username != null && password != null)
            {
                try
                {
                    LoadData();
                    button1.Enabled = false;
                    button2.Enabled = true; 
                }
                catch (MySqlException)
                {
                    MessageBox.Show("Проверьте корректность введённых данных и повторите попытку.", "Отказано в доступе!");
                }
            }
            dataGridView1.Select();
        }
         
        public static bool Formula4346(double Sav_Sow, double Sav_s, double Sav_i, double W_Sosav, double W_s, double W_i) {
            return (((Sav_Sow - Sav_s)*W_i - (W_s - W_Sosav)*Sav_i) < (W_Sosav*Sav_Sow - W_s*Sav_s));
        }
        
        public static double Find_alpha_WSav(double Sav_Sow, double Sav_s, double W_Sosav, double W_s) {
            return (Math.Round((Sav_Sow - Sav_s)/((Sav_Sow - Sav_s) + (W_s - W_Sosav)), 3));
        }
        
        public static void WSaviONends(double[] Sav_i, double[] W_i, ref double[,] WSavi_ends) {
            for (int i = 0; i < W_i.Length; ++i) {
                if (Sav_i[i] != 0) {
                    WSavi_ends[i, 0] = -Sav_i[i];
                }
                else {
                    WSavi_ends[i, 0] = 0;
                }
                WSavi_ends[i, 1] = W_i[i];
            }
        }
        
        public static bool Cross(double x1, double y1, double x2, double y2, double x3, double y3, double x4, double y4, ref double[] crossDot) {
            double n;
            if (y2 - y1 != 0) {  
                double q = - (x2 - x1) / (y2 - y1);
                double sn = (x3 - x4) + (y3 - y4) * q; 
                if (sn == 0) {
                    return false; 
                }
                double fn = (x3 - x1) + (y3 - y1) * q; 
                n = fn / sn;
            }
            else {
                if (y3 - y4 == 0) { 
                    return false;
                }
                n = (y3 - y1) / (y3 - y4);   
            }
            crossDot[0] = Math.Round(x3 + (x4 - x3) * n, 12);  
            crossDot[1] = y3 + (y4 - y3) * n;  
            return true;
        }
        
        public static double Wsav_i(int i, double alpha, double[] W_i, double[] Sav_i) {
            return Math.Round((alpha*W_i[i] - (1 - alpha)*Sav_i[i]), 2);
        }
        
        public static int Oft(List<int> list, int count_strat) { 
            int[] count = new int[count_strat]; 
            foreach (var element in list) {
                count[element - 1]++;
            }
            int max = -1, i_max = -1;
            for (int i = 0; i < count.Length; ++i) {
                if (count[i] > max) {
                    max = count[i];
                    i_max = i;
                }
            }
            return (i_max + 1);
        }
        
        public void FindWSavOptimal(bool intersection, int[] WOstrategies, int[] SavOstrategies , int[] WSavIntersection, int count_conditions, double[,] matrix, double Sav_s, double W_s, double[] Sav_i, double[] W_i) {
            bool equals = false; 
            if (intersection) {

                string str;
                str = "{";
                for (int i = 0; i < SavOstrategies.Length; ++i)
                {
                    str += $"A{SavOstrategies[i] + 1}";
                    if (i != (SavOstrategies.Length - 1))
                    {
                        str += ", ";
                    }
                }
                str += "}, при \u03B1 = 0";
                listBox3.Items.Add(str);
                str = "{";
                for (int i = 0; i < WSavIntersection.Length; ++i) {
                    str += $"A{WSavIntersection[i]+1}";
                    if (i != (WSavIntersection.Length - 1)) {
                        str += ", ";
                    }
                }
                str += "}, при \u03B1 принадлежащем интервалу (0, 1)";
                listBox3.Items.Add(str);
                str = "{";
                for (int i = 0; i < WOstrategies.Length; ++i)
                {
                    str += $"A{WOstrategies[i] + 1}";
                    if (i != (WOstrategies.Length - 1))
                    {
                        str += ", ";
                    }
                }
                str += "}, при \u03B1 = 1";
                listBox3.Items.Add(str);

                equals = SavOstrategies.OrderBy(item => item).SequenceEqual(WSavIntersection.OrderBy(item => item));
                equals = equals && WOstrategies.OrderBy(item => item).SequenceEqual(WSavIntersection.OrderBy(item => item));
                if (equals) {
                    foreach (var element in WSavIntersection) {
                        str = $"Стратегия A{element + 1} является доминантой.";
                        listBox3.Items.Add(str);
                    }
                }
            }
            else {

                double[,] SavO_matrix = new double[SavOstrategies.Length, count_conditions];
                for (int i = 0; i < SavOstrategies.Length; ++i) {
                    for (int j = 0; j < count_conditions; ++j) {
                        SavO_matrix[i, j] = matrix[SavOstrategies[i], j];
                    }
                }
                double SavO_WO_sum = int.MinValue;
                for (int i = 0; i < W_i.Length; ++i) {
                    for (int j = 0; j < SavOstrategies.Length; ++j) {
                        if ((i == SavOstrategies[j]) && (W_i[i] > SavO_WO_sum)) {
                            SavO_WO_sum = W_i[i];
                        }
                    }
                }
                List<int> SavO_WO_strategies_list = new List<int>();
                for (int i = 0; i < W_i.Length; ++i) {
                    if (W_i[i] == SavO_WO_sum) {
                        SavO_WO_strategies_list.Add(i);
                    }
                }
                int[] SavO_WO_strategies = SavO_WO_strategies_list.ToArray(); 

                double[,] WO_matrix = new double[WOstrategies.Length, count_conditions];
                for (int i = 0; i < WOstrategies.Length; ++i) {
                    for (int j = 0; j < count_conditions; ++j) {
                        WO_matrix[i, j] = matrix[WOstrategies[i], j];
                    }
                }
                double WO_SavO_sum = double.MaxValue;
                for (int i = 0; i < Sav_i.Length; ++i) {
                    for (int j = 0; j < WOstrategies.Length; ++j) {
                        if ((i == WOstrategies[j]) && (Sav_i[i] < WO_SavO_sum)) {
                            WO_SavO_sum = Sav_i[i];
                        }
                    }
                }
                List<int> WO_SavO_strategies_list = new List<int>();
                for (int i = 0; i < Sav_i.Length; ++i) {
                    if (Sav_i[i] == WO_SavO_sum) {
                        WO_SavO_strategies_list.Add(i);
                    }
                }
                int[] WO_SavO_strategies = WO_SavO_strategies_list.ToArray(); 

                List<int> WSavUnion_list = new List<int>();
                WSavUnion_list.AddRange(WOstrategies);
                WSavUnion_list.AddRange(SavOstrategies);
                int[] WSavUnion = WSavUnion_list.Distinct().ToArray(); 

                List<int> NoWorSavStrategiesList = new List<int>();
                bool f;
                for (int i = 0; i < matrix.GetLength(0); ++i) {
                    f = false;
                    for (int j = 0; j < WSavUnion.Length; ++j) {
                        if (i == WSavUnion[j]) {
                            f = true;
                            break;
                        }
                    }
                    if (!f) { 
                        NoWorSavStrategiesList.Add(i);
                    }
                }
                int[] NoWorSavStrategies = NoWorSavStrategiesList.ToArray();

                bool true_for_everyone = true;
                for (int i = 0; i < NoWorSavStrategies.Length; ++i) {
                    if (!Formula4346(WO_SavO_sum, Sav_s, Sav_i[NoWorSavStrategies[i]], SavO_WO_sum, W_s, W_i[NoWorSavStrategies[i]])) {
                        true_for_everyone = false;
                        break;
                    }
                }
                string str;
                if ((!true_for_everyone) || (NoWorSavStrategies.Length == 0)) { 
                    str = "О структуре множества стратегий, оптимальных по критерию Вальда-Сэвиджа, ничего определенного сказать нельзя.";
                    listBox3.Items.Add(str);
                }
                else {
                    double alpha_WSav = Find_alpha_WSav(WO_SavO_sum, Sav_s, SavO_WO_sum, W_s);

                    str = "{";
                    for (int i = 0; i < SavOstrategies.Length; ++i)
                    {
                        str += $"A{SavOstrategies[i] + 1}";
                        if (i != (SavOstrategies.Length - 1))
                        {
                            str += ", ";
                        }
                    }
                    str += "}, при \u03B1 = 0";
                    listBox3.Items.Add(str);
                    str = "{";
                    for (int i = 0; i < SavO_WO_strategies.Length; ++i) {
                        str += $"A{SavO_WO_strategies[i] + 1}";
                        if (i != (SavO_WO_strategies.Length - 1)) {
                            str += ", ";
                        }
                    }
                    str += $"}}, при \u03B1 принадлежащем интервалу (0, {alpha_WSav})";
                    listBox3.Items.Add(str);
                    List<int> WOSavO_SavOWO_Union_list = new List<int>();
                    WOSavO_SavOWO_Union_list.AddRange(WO_SavO_strategies);
                    WOSavO_SavOWO_Union_list.AddRange(SavO_WO_strategies);
                    int[] WOSavO_SavOWO_Union = WOSavO_SavOWO_Union_list.Distinct().ToArray();
                    str = "{";
                    for (int i = 0; i < WOSavO_SavOWO_Union.Length; ++i) {
                        str += $"A{WOSavO_SavOWO_Union[i] + 1}";
                        if (i != (WOSavO_SavOWO_Union.Length - 1)) {
                            str += ", ";
                        }
                    }
                    str += $"}}, при \u03B1 = {alpha_WSav}";
                    listBox3.Items.Add(str);
                    str = "{";
                    for (int i = 0; i < WO_SavO_strategies.Length; ++i) {
                        str += $"A{WO_SavO_strategies[i] + 1}";
                        if (i != (WO_SavO_strategies.Length - 1)) {
                            str += ", ";
                        }
                    }
                    str += $"}}, при \u03B1 принадлежащем интервалу ({alpha_WSav}, 1)";
                    listBox3.Items.Add(str);
                    str = "{";
                    for (int i = 0; i < WOstrategies.Length; ++i)
                    {
                        str += $"A{WOstrategies[i] + 1}";
                        if (i != (WOstrategies.Length - 1))
                        {
                            str += ", ";
                        }
                    }
                    str += "}, при \u03B1 = 1";
                    listBox3.Items.Add(str);
                }
            }

            {
                double[,] WSavi_ends = new double[W_i.Length, 2];
                WSaviONends(Sav_i, W_i, ref WSavi_ends); 

                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                column.Name = "WSav_i";
                column.HeaderText = "(WSav)i(\u03B1)";
                dataGridView4.Columns.Add(column);
                column = new DataGridViewTextBoxColumn();
                column.Name = "0";
                column.HeaderText = "\u03B1 = 0";
                dataGridView4.Columns.Add(column);
                column = new DataGridViewTextBoxColumn();
                column.Name = "1";
                column.HeaderText = "\u03B1 = 1";
                dataGridView4.Columns.Add(column);

                for (int i = 0; i < WSavi_ends.GetLength(0); i++)
                {
                    dataGridView4.Rows.Add();
                    dataGridView4[0, i].Value = $"A{i + 1}";
                    for (int j = 1; j < WSavi_ends.GetLength(1) + 1; j++)
                    {
                        dataGridView4[j, i].Value = WSavi_ends[i, j - 1];
                    }
                }
                
                dataGridView4.AutoResizeColumns();

                double[] dot = new double[2]; 
                double[,] intersected = new double[W_i.Length, W_i.Length]; 
                List<double> intersections_list = new List<double>();
                for (int i = 0; i < W_i.Length; ++i) {
                    for (int j = (i + 1); j < W_i.Length; ++j) {
                        if (Cross(0, WSavi_ends[i, 0], 1, WSavi_ends[i, 1], 0, WSavi_ends[j, 0], 1, WSavi_ends[j, 1], ref dot)) {
                            if ((dot[0] > 0) && (dot[0] < 1)) { 
                                intersected[i, j] = dot[0];
                                intersections_list.Add(dot[0]);
                            }
                        }
                    }
                }
                intersections_list.Add(0); 
                intersections_list.Add(1);
                double[] intersections = intersections_list.Distinct().OrderBy(item => item).ToArray(); 
                
                column = new DataGridViewTextBoxColumn();
                column.Name = "WSav_num";
                column.HeaderText = "№№ i отрезков (WSav)i(\u03B1)";
                dataGridView5.Columns.Add(column);
                for (int i = 1; i <= W_i.Length; i++)
                {
                    column = new DataGridViewTextBoxColumn();
                    column.Name = $"{i}";
                    column.HeaderText = $"{i}";
                    dataGridView5.Columns.Add(column);
                }

                for (int i = 0; i < intersected.GetLength(0); i++)
                {
                    dataGridView5.Rows.Add();
                    dataGridView5[0, i].Value = $"{i + 1}";
                    for (int j = 1; j < intersected.GetLength(1) + 1; j++)
                    {
                        var data = Math.Round(intersected[i, j - 1], 6);
                        dataGridView5[j, i].Value = data;
                    }
                }
                
                dataGridView5.AutoResizeColumns();

                double[,] WSaviINeachAlpha = new double[intersections.Length, W_i.Length]; 
                for (int i = 0; i < W_i.Length; ++i) {
                    WSaviINeachAlpha[0, i] = WSavi_ends[i, 0];
                }
                for (int i = 1; i < intersections.Length - 1; ++i) {
                    for (int j = 0; j < W_i.Length; ++j) {
                        WSaviINeachAlpha[i, j] = Wsav_i(j, intersections[i], W_i, Sav_i);
                    }
                }
                for (int i = 0; i < W_i.Length; ++i) {
                    WSaviINeachAlpha[(intersections.Length - 1), i] = WSavi_ends[i, 1];
                }

                column = new DataGridViewTextBoxColumn();
                column.Name = "alpha";
                column.HeaderText = "Значения выигрыш-показателя α";
                dataGridView6.Columns.Add(column);
                for (int i = 1; i <= W_i.Length; i++)
                {
                    column = new DataGridViewTextBoxColumn();
                    column.Name = $"WSav_{i}";
                    column.HeaderText = $"(WSav){i}(α)";
                    dataGridView6.Columns.Add(column);
                }
                for (int i = 0; i < WSaviINeachAlpha.GetLength(0); i++)
                {
                    dataGridView6.Rows.Add();
                    dataGridView6[0, i].Value = Math.Round(intersections[i], 6);
                    for (int j = 1; j < WSaviINeachAlpha.GetLength(1) + 1; j++)
                    {
                        dataGridView6[j, i].Value = WSaviINeachAlpha[i, j - 1];
                    }
                }
                
                dataGridView6.AutoResizeColumns();

                List<int>[,] PlacesINeachAlpha = new List<int>[intersections.Length, W_i.Length]; 
                for (int i = 0; i < PlacesINeachAlpha.GetLength(0); ++i) {
                    for (int j = 0; j < PlacesINeachAlpha.GetLength(1); ++j) {
                        PlacesINeachAlpha[i, j] = new List<int>();
                    }
                }
                double max; 
                int ind_max, count_max;
                int place;
                for (int i = 0; i < intersections.Length; ++i) { 
                    place = 1;
                    for (int l = 0; l < W_i.Length; ++l) { 
                        if (place <= W_i.Length) { 
                            max = double.MinValue;
                            ind_max = -1;
                            count_max = 1;
                            for (int j = 0; j < W_i.Length; ++j) {
                                if ((WSaviINeachAlpha[i, j] > max) && (!PlacesINeachAlpha[i,j].Any())) {
                                    max = WSaviINeachAlpha[i, j];
                                    ind_max = j;
                                }
                            }
                            for (int j = 0; j < W_i.Length; ++j) { 
                                if ((WSaviINeachAlpha[i, j] == max) && (j != ind_max)) {
                                    count_max++;
                                }
                            }
                            for (int j = 0; j < W_i.Length; ++j) { 
                                if (WSaviINeachAlpha[i, j] == max) {
                                    for (int k = place; k < (place + count_max); ++k) {
                                        PlacesINeachAlpha[i, j].Add(k);
                                    }
                                }
                            }
                            place += count_max;
                        }
                    }
                }

                int[,] PlacesINeachInterval = new int[(PlacesINeachAlpha.GetLength(0) - 1), PlacesINeachAlpha.GetLength(1)]; 
                for (int i = 0; i < (PlacesINeachAlpha.GetLength(0) - 1); ++i) {
                    for (int j = 0; j < PlacesINeachAlpha.GetLength(1); ++j) {
                        List<int> placesUnion = new List<int>();
                        placesUnion.AddRange(PlacesINeachAlpha[i, j]);
                        placesUnion.AddRange(PlacesINeachAlpha[i + 1, j]);
                        PlacesINeachInterval[i, j] = Oft(placesUnion, PlacesINeachAlpha.GetLength(1));
                    }
                }

                column = new DataGridViewTextBoxColumn();
                column.Name = "alpha";
                column.HeaderText = "Значения выигрыш-показателя α";
                dataGridView8.Columns.Add(column);
                for (int i = 1; i <= W_i.Length; i++)
                {
                    column = new DataGridViewTextBoxColumn();
                    column.Name = $"str_{i}";
                    column.HeaderText = $"A{i}";
                    dataGridView8.Columns.Add(column);
                }
                
                int h = 0;
                int[] PlacesINeachA;
                string places;
                for (int i = 0; i < PlacesINeachAlpha.GetLength(0) + PlacesINeachInterval.GetLength(0) - 1; i+=2) 
                {
                    dataGridView8.Rows.Add();
                    dataGridView8.Rows.Add();
                    dataGridView8[0, i].Value = Math.Round(intersections[i - h], 6);
                    dataGridView8[0, i + 1].Value = $"{Math.Round(intersections[i - h], 6)} < α < {Math.Round(intersections[i - h + 1], 6)}";
                    for (int j = 1; j < PlacesINeachAlpha.GetLength(1) + 1; j++)
                    {
                        PlacesINeachA = PlacesINeachAlpha[i - h, j - 1].ToArray();
                        places = Convert.ToString(PlacesINeachA[0]);
                        for (int k = 1; k < PlacesINeachA.Length; k++)
                        {
                            places += ("/" + Convert.ToString(PlacesINeachA[k]));
                        }
                        dataGridView8[j, i].Value = places;
                        dataGridView8[j, i + 1].Value = PlacesINeachInterval[i - h, j - 1];
                    }
                    h++;
                }
                dataGridView8.Rows.Add();
                dataGridView8[0, PlacesINeachAlpha.GetLength(0) + PlacesINeachInterval.GetLength(0) - 1].Value = Math.Round(intersections[intersections.Length - 1], 6);
                for (int i = 1; i < PlacesINeachAlpha.GetLength(1) + 1; i++)
                {
                    PlacesINeachA = PlacesINeachAlpha[PlacesINeachAlpha.GetLength(0) - 1, i - 1].ToArray();
                    places = Convert.ToString(PlacesINeachA[0]);
                    for (int k = 1; k < PlacesINeachA.Length; k++)
                    {
                        places += ("/" + Convert.ToString(PlacesINeachA[k]));
                    }
                    dataGridView8[i, PlacesINeachAlpha.GetLength(0) + PlacesINeachInterval.GetLength(0) - 1].Value = places;
                }
                
                dataGridView8.AutoResizeColumns();
                
                List<string>[,] orders = new List<string>[PlacesINeachAlpha.GetLength(0) + PlacesINeachInterval.GetLength(0), W_i.Length]; 
                for (int i = 0; i < orders.GetLength(0); ++i) {
                    for (int j = 0; j < orders.GetLength(1); ++j) {
                        orders[i, j] = new List<string>();
                    }
                }
                int p = 0;
                for (int i = 0; i < orders.GetLength(0); i += 2) { 
                    for (int j = 0; j < PlacesINeachAlpha.GetLength(1); ++j) {
                        for (int place_ = 1; place_ <= W_i.Length; place_++) { 
                            if (PlacesINeachAlpha[i - p, j].Contains(place_)) { 
                                orders[i, (place_ - 1)].Add($"A{j + 1}");
                            }
                        }
                    }
                    p++;
                }
                p = 1; 
                for (int i = 1; i < orders.GetLength(0) - 1; i += 2) { 
                    for (int j = 0; j < PlacesINeachInterval.GetLength(1); ++j) {
                        for (int place_ = 1; place_ <= W_i.Length; place_++) {
                            if (PlacesINeachInterval[i - p, j] == place_) {
                                orders[i, (place_ - 1)].Add($"A{j + 1}");
                            }
                        }
                    }
                    p++;
                }

                column = new DataGridViewTextBoxColumn();
                column.Name = "alpha";
                column.HeaderText = "Значения выигрыш-показателя α";
                dataGridView7.Columns.Add(column);
                for (int i = 1; i <= W_i.Length; i++)
                {
                    column = new DataGridViewTextBoxColumn();
                    column.Name = $"place_{i}";
                    column.HeaderText = $"{i}";
                    dataGridView7.Columns.Add(column);
                }
                
                p = 0;
                int m = 0;
                string strats;
                for (int i = 0; i < orders.GetLength(0); ++i)
                {
                    dataGridView7.Rows.Add();
                    if (i % 2 == 0) { 
                        dataGridView7[0, i].Value = $"\u03B1 = {Math.Round(intersections[p], 6)}";
                        p++;
                    }
                    else { 
                        dataGridView7[0, i].Value = $"{Math.Round(intersections[m], 6)} < \u03B1 < {Math.Round(intersections[m + 1], 6)}";
                        m++;
                    }
                    for (int j = 1; j < orders.GetLength(1) + 1; ++j) {
                        string[] array = orders[i, j - 1].ToArray();
                        strats = array[0];
                        for (int l = 1; l < array.Length; ++l) { 
                            strats += ("/" + array[l]);
                        }
                        dataGridView7[j, i].Value = strats;
                    }
                }
                
                dataGridView7.AutoResizeColumns();
                
            }
            
        }

        private string destinationPath;
        
        private void WriteExcelFile()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Значения выигрыш-показателя α", typeof(string));
            for (int i = 1; i < dataGridView7.Columns.Count; i++)
            {
                table.Columns.Add($"{i}", typeof(string));
            }
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                table.Rows.Add();
            }
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView7.Columns.Count; j++)
                {
                    table.Rows[i][j] = dataGridView7[j, i].Value;
                }
            }

            MessageBox.Show("Выберите путь сохранения и введите название создаваемого файла", "Сохранение результата");
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                destinationPath = saveFileDialog1.FileName;
            }

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(destinationPath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Приоритетные очереди" };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<string> columns = new List<string>();
                foreach (DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                   Row newRow = new Row();
                   foreach (string col in columns)
                   {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                   }

                   sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            double[,] payoff_matrix = new double[dataGridView1.RowCount, dataGridView1.ColumnCount - 1]; 
            for (int i = 0; i < payoff_matrix.GetLength(0); i++) 
            {
                for (int j = 1; j < payoff_matrix.GetLength(1) + 1; j++)
                {
                    payoff_matrix[i, j - 1] = Convert.ToDouble(dataGridView1[j, i].Value); 
                }
            }
            
            DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
            column.Name = "str_num";
            column.HeaderText = "№ стратегии";
            dataGridView2.Columns.Add(column);
            for (int i = 1; i < countColumns; ++i) 
            {
                column = new DataGridViewTextBoxColumn();
                column.Name = $"nature{i}";
                column.HeaderText = $"П{i}";
                dataGridView2.Columns.Add(column);
            }
            
            for (int i = 0; i < payoff_matrix.GetLength(0); i++) 
            {
                dataGridView2.Rows.Add();
                for (int j = 1; j < payoff_matrix.GetLength(1) + 1; j++)
                {
                    dataGridView2[0, i].Value = i + 1;
                    dataGridView2[j, i].Value = payoff_matrix[i, j - 1];
                }
            }

            double Ws;
            int[] WOptimalStrategies; 
            double[] Wi; 
            WSav_algorythm.FindWoptimal(payoff_matrix, out WOptimalStrategies, out Ws, out Wi);

            column = new DataGridViewTextBoxColumn();
            column.Name = "Wi";
            column.HeaderText = "Wi";
            dataGridView2.Columns.Add(column);
            
            for (int i = 0; i < Wi.GetLength(0); i++)
            {
                dataGridView2[dataGridView2.ColumnCount - 1, i].Value = Wi[i];
            }

            dataGridView2.AutoResizeColumns();
            
            textBox1.Text = Convert.ToString(Ws);

            List<string> WOptimalStrategiesList = new List<string>();
            for (int i = 0; i < WOptimalStrategies.Length; i++)
            {
                WOptimalStrategiesList.Add($"A{WOptimalStrategies[i]+1}");
            }
            listBox1.DataSource = WOptimalStrategiesList;

            double[,] risk_matrix;
            WSav_algorythm.CreateRiskMatrix(payoff_matrix, out risk_matrix);
            
            column = new DataGridViewTextBoxColumn();
            column.Name = "str_num";
            column.HeaderText = "№ стратегии";
            dataGridView3.Columns.Add(column);
            for (int i = 1; i < countColumns; ++i) 
            {
                column = new DataGridViewTextBoxColumn();
                column.Name = $"nature{i}";
                column.HeaderText = $"П{i}";
                dataGridView3.Columns.Add(column);
            }
            
            for (int i = 0; i < risk_matrix.GetLength(0); i++)
            {
                dataGridView3.Rows.Add();
                for (int j = 1; j < risk_matrix.GetLength(1) + 1; j++)
                {
                    dataGridView3[0, i].Value = i + 1;
                    dataGridView3[j, i].Value = risk_matrix[i, j - 1];
                }
            }

            double Savs;
            int[] SavOptimalStrategies;
            double[] Savi; 
            WSav_algorythm.FindSavOptimal(risk_matrix, out SavOptimalStrategies, out Savs, out Savi);
            
            column = new DataGridViewTextBoxColumn();
            column.Name = "Savi";
            column.HeaderText = "Savi";
            dataGridView3.Columns.Add(column);
            
            for (int i = 0; i < Savi.GetLength(0); i++)
            {
                dataGridView3[dataGridView3.ColumnCount - 1, i].Value = Savi[i];
            }
            
            dataGridView3.AutoResizeColumns();
            
            textBox2.Text = Convert.ToString(Savs);

            List<string> SavOptimalStrategiesList = new List<string>();
            for (int i = 0; i < SavOptimalStrategies.Length; i++)
            {
                SavOptimalStrategiesList.Add($"A{SavOptimalStrategies[i]+1}");
            }
            listBox2.DataSource = SavOptimalStrategiesList;

            int[] WSavIntersection; 
            bool WOSOintersection = WSav_algorythm.IsIntersectionNotEmpty(WOptimalStrategies, SavOptimalStrategies, out WSavIntersection);

            FindWSavOptimal(WOSOintersection, WOptimalStrategies, SavOptimalStrategies, WSavIntersection,
                payoff_matrix.GetLength(1), payoff_matrix, Savs, Ws, Savi, Wi);

            button2.Enabled = false;
            dataGridView7.Select();

            if (checkBox1.Checked)
            {
                WriteExcelFile();
            }
        }

        private void ClearAll_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
            dataGridView4.Rows.Clear();
            dataGridView4.Columns.Clear();
            dataGridView5.Rows.Clear();
            dataGridView5.Columns.Clear();
            dataGridView6.Rows.Clear();
            dataGridView6.Columns.Clear();
            dataGridView7.Rows.Clear();
            dataGridView7.Columns.Clear();
            dataGridView8.Rows.Clear();
            dataGridView8.Columns.Clear();
            listBox1.DataSource = null;
            listBox2.DataSource = null;
            listBox3.Items.Clear();
            textBox1.Text = null;
            textBox2.Text = null;
            button1.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button1.FlatAppearance.BorderSize = 0;
            button1.FlatAppearance.MouseOverBackColor = Color.Transparent;
            button2.FlatAppearance.BorderSize = 0;
            button2.FlatAppearance.MouseOverBackColor = Color.Transparent;
            ClearAll.FlatAppearance.BorderSize = 0;
            ClearAll.FlatAppearance.MouseOverBackColor = Color.Transparent;
            
            ToolTip toolTip_button1 = new ToolTip();
            toolTip_button1.SetToolTip(button1, "Подключить базу данных MySQL.\nДля корректного подключения необходимо, чтобы\nназвание столбца с номерами стратегий содержало слово id.");
            ToolTip toolTip_button2 = new ToolTip();
            toolTip_button2.SetToolTip(button2, "Сформировать приоритетную очередь");
            ToolTip toolTip_ClearAll = new ToolTip();
            toolTip_ClearAll.SetToolTip(ClearAll, "Очистить всё");
            ToolTip toolTip_dataGridView1 = new ToolTip();
            toolTip_dataGridView1.SetToolTip(dataGridView1, "Платёжная матрица, импортированная из базы данных.");
            
        }
        
    }
}   