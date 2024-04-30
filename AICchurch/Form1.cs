using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using System.Data.SqlClient;
using System.Xml.Linq;
using static AICchurch.Form1;
using System.Data; 

namespace AICchurch
{
    public partial class Form1 : Form
    {

        private readonly checkUser _user;
        private readonly int _userId;

        private bool isDragging = false;
        private int mouseX;
        private int mouseY;

        private List<int> selectedServicesIds = new List<int>();

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]

        private static extern IntPtr CreateRoundRectRgn

        (

        int nLeftRect,
        int nTopRect,
        int nRightRect,
        int nBottomRect,
        int nWidthEllipse,
        int nHeightEllipse
    );


        public Form1(checkUser user, int userId)
        {
            _user = user;
            _userId = userId; 

            InitializeComponent();

            
            dataGridView2.CellContentClick += dataGridView2_CellContentClick;
            dataGridView4.CellContentClick += dataGridView4_CellContentClick;


            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 25, 25));
            pnlNav.Height = btnKorust.Height;
            pnlNav.Top = btnKorust.Top;
            pnlNav.Left = btnKorust.Left;


            TaskPanel.MouseDown += Control_MouseDown;
            TaskPanel.MouseMove += Control_MouseMove;
            TaskPanel.MouseUp += Control_MouseUp;

            TaskPanel2.MouseDown += Control_MouseDown;
            TaskPanel2.MouseMove += Control_MouseMove;
            TaskPanel2.MouseUp += Control_MouseUp;

            lightPicture.Add(AddUser);
            lightPicture.Add(DelUser);
            lightPicture.Add(UpdateUser);
            lightPicture.Add(SearchUser);

            lightPicture.Add(AddPos);
            lightPicture.Add(PosUp);
            lightPicture.Add(DelPos);
            lightPicture.Add(PosScr);

            MouseTrigger();

        }

        private bool useServiceNames = false;

  


        private void Control_MouseDown(object sender, MouseEventArgs e)
        {
            isDragging = true;
            mouseX = e.X;
            mouseY = e.Y;
        }

        private void Control_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                int deltaX = e.X - mouseX;
                int deltaY = e.Y - mouseY;
                Left += deltaX;
                Top += deltaY;
            }
        }

        private void Control_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }


        private void Hide_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void close_Click(object sender, EventArgs e)
        {
            Application.Exit();

        }


        private void IsAdmin()
        {

            btnKorust.Visible = _user.IsAdmin;
            btnzamov.Visible = _user.IsAdmin;


            AddPos.Visible = _user.IsAdmin;
            PosUp.Visible = _user.IsAdmin;
            DelPos.Visible = _user.IsAdmin;
            PosScr.Visible = _user.IsAdmin;

            if (_user.IsAdmin)
            {
                pictureBox3.Visible = false;
            }
            else
            {
                pictureBox3.Visible = true;
            }

            btnmat.Visible = _user.IsAdmin;
            btnmat2.Visible = _user.IsAdmin;
            tabControl1.SelectedIndex = 5;

        }





        private List<PictureBox> lightPicture = new List<PictureBox>();

        private void MouseTrigger()


        {
            foreach (PictureBox pictureBox in lightPicture)
            {
                pictureBox.MouseEnter += MouseEnter; 
                pictureBox.MouseLeave += MouseLeave; 
            }
        }

        private void MouseEnter(object sender, EventArgs e)
        {
            PictureBox pictureBox = sender as PictureBox;
            if (pictureBox != null)
            {
                pictureBox.BackColor = Color.FromArgb(185, 162, 148);
            }
        }

        private void MouseLeave(object sender, EventArgs e)
        {
            PictureBox pictureBox = sender as PictureBox;
            if (pictureBox != null)
            {
                pictureBox.BackColor = Color.Transparent; 
            }
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            IsAdmin();
            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            string Sql = $"select Логін from [dbo].[Користувачі] where Логін = '{DataSend.text}'";
            SqlCommand scmd = new SqlCommand(Sql, con);
            con.Open();
            SqlDataReader sur = scmd.ExecuteReader();

            string status = _user.IsAdmin ? "Админ" : "Користувач";
            UserLogin.Text = $"{_user.Login} | {status}";

            con.Close();

            if (_user.IsAdmin)
            {
                btnKorust_Click(sender, e);
            }
            else
            {
                btnposlyg_Click(sender, e);
            }
        }


        private void btnKorust_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnKorust.Height;
            pnlNav.Top = btnKorust.Top;
            pnlNav.Left = btnKorust.Left;

            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();
            SqlCommand command = new SqlCommand($"SELECT * FROM [dbo].[Користувачі] ", con);

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            da.Fill(dt);


            dataGridView1.DataSource = dt;

            if (dataGridView1.Columns.Contains("Пароль"))
            {
                dataGridView1.Columns["Пароль"].Visible = false;
                dataGridView1.Columns["ID"].Visible = false;
            }

            tabControl1.SelectedIndex = 0;
        }


        private void btnposlyg_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnposlyg.Height;
            pnlNav.Top = btnposlyg.Top;
            btnposlyg.BackColor = Color.FromArgb(204, 199, 199);

            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();
            SqlCommand command = new SqlCommand($"SELECT * FROM [dbo].[Послуги] ", con);

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            da.Fill(dt);

            dataGridView2.DataSource = dt;

            if (dataGridView2.Columns.Contains("ID послуги"))
            {
                dataGridView2.Columns["ID послуги"].Visible = false;
               
            }

            bool buyColumnExists = dataGridView2.Columns.Cast<DataGridViewColumn>().Any(column => column.HeaderText == "Придбати");

            if (!buyColumnExists)
            {
                DataGridViewButtonColumn buyButtonColumn = new DataGridViewButtonColumn();
                buyButtonColumn.HeaderText = "Придбати";
                buyButtonColumn.Text = "в кошик";
                buyButtonColumn.UseColumnTextForButtonValue = true;
                dataGridView2.Columns.Add(buyButtonColumn);
            }

            dataGridView2.CellContentClick -= dataGridView2_CellContentClick;
            dataGridView2.CellContentClick += dataGridView2_CellContentClick;


            tabControl1.SelectedIndex = 1;
        }

        private void btnzamov_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnzamov.Height;
            pnlNav.Top = btnzamov.Top;
            btnzamov.BackColor = Color.FromArgb(204, 199, 199);
            tabControl1.SelectedIndex = 2;

            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();
                string sql = @"SELECT Замовлення.[ID Замовлення], Послуги.[Назва послуги], Користувачі.[Логін] AS [Логін Користувача], Замовлення.[Загальна сума], Замовлення.[Дата] 
             FROM Замовлення
             LEFT JOIN Послуги ON Замовлення.[ID Послуги] = Послуги.[ID послуги]
             LEFT JOIN Користувачі ON Замовлення.[ID Користувача] = Користувачі.[ID]";
                SqlCommand command = new SqlCommand(sql, con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView3.DataSource = dt;

                if (dataGridView3.Columns.Contains("ID Замовлення"))
                {
                    dataGridView3.Columns["ID Замовлення"].Visible = false;
                }

             
                FillComboBox(con);
                
                con.Close();
            }
        }

        private void FillComboBox(SqlConnection con)
        {

            comboPos.Items.Clear();

            string sqlPos = "SELECT DISTINCT [Назва послуги] FROM Послуги";
            SqlCommand commandPos = new SqlCommand(sqlPos, con);

            using (SqlDataReader reader = commandPos.ExecuteReader())
            {
                while (reader.Read())
                {
                    comboPos.Items.Add(reader["Назва послуги"]);
                }
            }

            comboUsers.Items.Clear();

            string sqlUsers = "SELECT DISTINCT [Логін] AS [Ім'я Користувача] FROM Користувачі";
            SqlCommand commandUsers = new SqlCommand(sqlUsers, con);

            using (SqlDataReader reader = commandUsers.ExecuteReader())
            {
                while (reader.Read())
                {
                    comboUsers.Items.Add(reader["Ім'я Користувача"]);
                }
            }

            combomat.Items.Clear();

            string sqlVMat = "SELECT DISTINCT [Назва Матеріалу] FROM [Витратні Матеріали]";
            SqlCommand commandVMat = new SqlCommand(sqlVMat, con);

            using (SqlDataReader reader = commandVMat.ExecuteReader())
            {
                while (reader.Read())
                {
                    combomat.Items.Add(reader["Назва Матеріалу"]);
                }
            }


        }

        private void btnmat_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnmat.Height;
            pnlNav.Top = btnmat.Top;
            btnmat.BackColor = Color.FromArgb(204, 199, 199);
            tabControl1.SelectedIndex = 3;


            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();
            SqlCommand command = new SqlCommand($"SELECT * FROM [dbo].[Витратні Матеріали] ", con);

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView6.DataSource = dt;

            if (dataGridView6.Columns.Contains("ID Матеріалу"))
            {
                dataGridView6.Columns["ID Матеріалу"].Visible = false;
            }

       
            FillComboBox(con);
        }

        private void btnmat2_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnmat2.Height;
            pnlNav.Top = btnmat2.Top;
            btnmat2.BackColor = Color.FromArgb(204, 199, 199);

            tabControl1.SelectedIndex = 4;

            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();
                string sql = @"SELECT vm.ID, vm.[ID Замовлення], m.[Назва Матеріалу], vm.[Кількість], vm.[Дата Використання]
                      FROM [dbo].[Витрачені матеріали] vm
                      LEFT JOIN [dbo].[Витратні Матеріали] m ON vm.[ID Матеріалу] = m.[ID Матеріалу]";
                SqlCommand command = new SqlCommand(sql, con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView7.DataSource = dt;

                if (dataGridView7.Columns.Contains("ID"))
                {
                    dataGridView7.Columns["ID"].Visible = false;
                    dataGridView7.Columns["ID Замовлення"].Visible = false;
                }

                con.Close();
            }
        }



        private void btnKorust_Leave(object sender, EventArgs e)
        {
            btnKorust.BackColor = Color.FromArgb(217, 217, 217);
        }

        private void btnposlyg_Leave(object sender, EventArgs e)
        {
            btnposlyg.BackColor = Color.FromArgb(217, 217, 217);
        }

        private void btnzamov_Leave(object sender, EventArgs e)
        {
            btnzamov.BackColor = Color.FromArgb(217, 217, 217);
        }

        private void btnmat_Leave(object sender, EventArgs e)
        {
            btnmat.BackColor = Color.FromArgb(217, 217, 217);
        }

        private void btnmat2_Leave(object sender, EventArgs e)
        {
            btnmat2.BackColor = Color.FromArgb(217, 217, 217);
        }
        private void cart_Leave(object sender, EventArgs e)
        {
            cart.BackColor = Color.FromArgb(217, 217, 217);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {

            AllBtn frm = new AllBtn();
            frm.AllConrol.SelectedTab = frm.tabPage1;

            frm.UpdateUser.Visible = false;

            frm.Show();

        }


        public class UpData
        {
            public string ID { get; set; }
            public string Логін { get; set; }
            public string Пошта { get; set; }

            public string IDпослуги { get; set; }
            public string ОписПослуги { get; set; }
            public string НазваПослуги { get; set; }
            public string ВартістьПослуги { get; set; }

            public string IDМатеріалу { get; set; }
            public string НазваМатеріалу { get; set; }
            public string ОписМатеріалу { get; set; }
            public string ОдиниціВиміру { get; set; }
            public string КількістьНаСкладі { get; set; }
            public string ЦінаЗаодиницю { get; set; }


            public string IDМат { get; set; }
            public string IDЗамовленняМат { get; set; }
            public string IDМатеріалуМат { get; set; }
            public string КількістьМат { get; set; }
            public string IDМатеріалуМат2 { get; set; }

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string column1Data = dataGridView1.SelectedRows[0].Cells["ID"].Value.ToString();
                string column2Data = dataGridView1.SelectedRows[0].Cells["Логін"].Value.ToString();
                string column3Data = dataGridView1.SelectedRows[0].Cells["Пошта"].Value.ToString();

                UpData upData = new UpData
                {
                    ID = column1Data,
                    Логін = column2Data,
                    Пошта = column3Data,

                };

                AllBtn frm = new AllBtn();
                frm.Addbtn.Visible = false;
                frm.label3.Text = "редагування";

                frm.SetData(upData);

                frm.Show();
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();

            SqlCommand command;

            if (!string.IsNullOrEmpty(txtID.Text))
            {
                command = new SqlCommand($"DELETE [dbo].[Користувачі] WHERE ID = @ID", con);
                command.Parameters.AddWithValue("@ID", int.Parse(txtID.Text));
                command.ExecuteNonQuery();
            }
            else
            {

                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите пользователя для удаления.");
                    con.Close();
                    return;
                }

              
                int selectedRowIndex = dataGridView1.SelectedRows[0].Index;
                int idToDelete = Convert.ToInt32(dataGridView1.Rows[selectedRowIndex].Cells["ID"].Value);
                command = new SqlCommand($"DELETE [dbo].[Користувачі] WHERE ID = @ID", con);
                command.Parameters.AddWithValue("@ID", idToDelete);
                command.ExecuteNonQuery();
            }

            SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM [dbo].[Користувачі]", con);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;

            txtID.Clear();

            con.Close();
            MessageBox.Show("Пользователь удален.");
        }



        private void btnSearch_Click(object sender, EventArgs e)
        {

            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();

            SqlCommand command;
            SqlDataAdapter da = new SqlDataAdapter();
            DataTable dt = new DataTable();



            if (string.IsNullOrWhiteSpace(txtID.Text) && string.IsNullOrWhiteSpace(txtlog.Text) && string.IsNullOrWhiteSpace(txtmail.Text))
            {
                command = new SqlCommand($"SELECT * FROM [dbo].[Користувачі]", con);
            }


            else if (!string.IsNullOrWhiteSpace(txtID.Text))
            {
                command = new SqlCommand($"SELECT * FROM [dbo].[Користувачі] WHERE ID = @ID ", con);
                command.Parameters.AddWithValue("ID", int.Parse(txtID.Text));
            }

            else if (!string.IsNullOrWhiteSpace(txtlog.Text))
            {
                command = new SqlCommand($"SELECT * FROM [dbo].[Користувачі] WHERE Логін = @Логін ", con);
                command.Parameters.AddWithValue("Логін",txtlog.Text);
            }

            else if (!string.IsNullOrWhiteSpace(txtmail.Text))
            {
                command = new SqlCommand($"SELECT * FROM [dbo].[Користувачі] WHERE Пошта = @Пошта ", con);
                command.Parameters.AddWithValue("Пошта",txtmail.Text);
            }

            else
            {
                MessageBox.Show("Введите хотя бы одно значение для поиска.");
                con.Close();
                return;
            }

            da.SelectCommand = command;
            da.Fill(dt);

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Записи не найдены");
            }
            else
            {
                dataGridView1.DataSource = dt;
            }

            con.Close();

        }

        private void PoslugAdd_Click(object sender, EventArgs e)
        {

            AllBtn frm = new AllBtn();
            frm.AllConrol.SelectedTab = frm.tabPage2;

            frm.PosUp.Visible = false;

            frm.Show();

            frm.label3.Text = "Додавання";

        }

        private void PosUpdate_Click(object sender, EventArgs e)
        {

            if (dataGridView2.SelectedRows.Count > 0)
            {
                string column1 = dataGridView2.SelectedRows[0].Cells["ID послуги"].Value.ToString();
                string column2 = dataGridView2.SelectedRows[0].Cells["Назва послуги"].Value.ToString();
                string column3 = dataGridView2.SelectedRows[0].Cells["Опис послуги"].Value.ToString();
                string column4 = dataGridView2.SelectedRows[0].Cells["Вартість послуги"].Value.ToString();

                UpData upData = new UpData
                {


                    IDпослуги = column1,
                    НазваПослуги = column2,
                    ОписПослуги = column3,
                    ВартістьПослуги = column4,

                };

                AllBtn frm = new AllBtn();
                frm.AllConrol.SelectedTab = frm.tabPage2;

                frm.AddPos.Visible = false;
                frm.label3.Text = "редагування";
                frm.SetData(upData);

                frm.Show();
    
            }

        }

        private void PosDel_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();

            SqlCommand command;

            if (!string.IsNullOrEmpty(IDPos.Text))
            {
                command = new SqlCommand($"DELETE [dbo].[Послуги] WHERE [ID послуги] = @IDпослуги", con);
                command.Parameters.AddWithValue("@IDпослуги", int.Parse(IDPos.Text));
                command.ExecuteNonQuery(); 
            }
            else if (dataGridView2.SelectedRows.Count > 0)
            {
                int selectedRowIndex = dataGridView2.SelectedRows[0].Index;
                int idToDelete = Convert.ToInt32(dataGridView2.Rows[selectedRowIndex].Cells["ID послуги"].Value);
                command = new SqlCommand($"DELETE [dbo].[Послуги] WHERE [ID послуги] = @IDпослуги", con);
                command.Parameters.AddWithValue("@IDпослуги", idToDelete);
                command.ExecuteNonQuery();
            }
           

            SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM [dbo].[Послуги]", con);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);
            dataGridView2.DataSource = dataTable;

            IDPos.Clear();

            con.Close();
            MessageBox.Show("Услуга удалена.");
        }

        private void PosSearch_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();

            SqlCommand command;
            SqlDataAdapter da = new SqlDataAdapter();
            DataTable dt = new DataTable();

            if (string.IsNullOrWhiteSpace(IDPos.Text))
            {
                command = new SqlCommand($"SELECT * FROM [dbo].[Послуги]", con);
            }
            else
            {
                command = new SqlCommand($"SELECT * FROM [dbo].[Послуги]  WHERE [Назва послуги] LIKE '%' + @searchTerm + '%'", con);
                command.Parameters.AddWithValue("@searchTerm", IDPos.Text);
            }

            da.SelectCommand = command;
            da.Fill(dt);

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Записи не найдены");
            }
            else
            {
                dataGridView2.DataSource = dt;
            }

            con.Close();
        }


        private void TestButton_Click(object sender, EventArgs e)
        {
            AllBtn frm = new AllBtn();
            frm.AllConrol.SelectedTab = frm.tabPage2;
            frm.Show();
        }

        private void ZamovAdd_Click(object sender, EventArgs e)
        {


            AllBtn frm = new AllBtn();
            frm.AllConrol.SelectedTab = frm.Zatmat;

            frm.PosUp.Visible = false;

            frm.Show();

        }

        private void ZamovSearch_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();

                string sqlQuery = "SELECT Замовлення.[ID Замовлення], Послуги.[Назва послуги], Користувачі.[Логін] AS [Логін Користувача], Замовлення.[Загальна сума], Замовлення.[Дата] " +
                                  "FROM Замовлення " +
                                  "LEFT JOIN Послуги ON Замовлення.[ID Послуги] = Послуги.[ID послуги] " +
                                  "LEFT JOIN Користувачі ON Замовлення.[ID Користувача] = Користувачі.[ID] " +
                                  "WHERE 1 = 1";

                List<SqlParameter> parameters = new List<SqlParameter>();

                if (!string.IsNullOrWhiteSpace(comboPos.Text))
                {
                    sqlQuery += " AND Послуги.[Назва послуги] LIKE @ServiceName";
                    parameters.Add(new SqlParameter("@ServiceName", "%" + comboPos.Text + "%"));
                }

                if (!string.IsNullOrWhiteSpace(comboUsers.Text))
                {
                    sqlQuery += " AND Користувачі.[Логін] LIKE @UserName";
                    parameters.Add(new SqlParameter("@UserName", "%" + comboUsers.Text + "%"));
                }

                if (dateTimePicker1.Value <= dateTimePicker2.Value)
                {
                    DateTime startDate = dateTimePicker1.Value.Date + new TimeSpan(00, 00, 00);
                    DateTime endDate = dateTimePicker2.Value.Date + new TimeSpan(23, 59, 59);

                    sqlQuery += " AND Замовлення.[Дата] BETWEEN @StartDate AND @EndDate";
                    parameters.Add(new SqlParameter("@StartDate", startDate));
                    parameters.Add(new SqlParameter("@EndDate", endDate));
                }

                using (SqlCommand command = new SqlCommand(sqlQuery, con))
                {
                    command.Parameters.AddRange(parameters.ToArray());

                    using (SqlDataAdapter da = new SqlDataAdapter(command))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        if (dt.Rows.Count == 0)
                        {
                            MessageBox.Show("Записи не найдены");
                        }
                        else
                        {
                            dataGridView3.DataSource = dt;
                        }
                    }
                }
            }
        }

        private void ZamovDel_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();

                SqlCommand command;

               
               if (dataGridView3.SelectedRows.Count > 0)
                {
                    int selectedRowIndex = dataGridView3.SelectedRows[0].Index;
                    int idToDelete = Convert.ToInt32(dataGridView3.Rows[selectedRowIndex].Cells["ID Замовлення"].Value);
                    command = new SqlCommand($"DELETE FROM [dbo].[Замовлення] WHERE [ID Замовлення] = @ID", con);
                    command.Parameters.AddWithValue("@ID", idToDelete);
                    command.ExecuteNonQuery(); 
                }

                SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM [dbo].[Замовлення]", con);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView3.DataSource = dataTable;

                MessageBox.Show("Удалено");

                con.Close();
            }
        }

        private void cart_Click(object sender, EventArgs e)
        {
            pnlNav.Height = cart.Height;
            pnlNav.Top = cart.Top;
            pnlNav.Left = cart.Left;

            cart.BackColor = Color.FromArgb(204, 199, 199);
            tabControl1.SelectedIndex = 6;

            if (dataGridView4.Columns.Contains("ID послуги"))
            {
                dataGridView4.Columns["ID послуги"].Visible = false;
            }
        }



        private bool deleteButtonAdded = false; 

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex >= 0 && dataGridView2.Columns[e.ColumnIndex] is DataGridViewButtonColumn)
            {
                int serviceId = Convert.ToInt32(dataGridView2.Rows[e.RowIndex].Cells["ID послуги"].Value);

                selectedServicesIds.Add(serviceId);

                label14.Text = string.Join(", ", selectedServicesIds);

                FindServicesInDatabase(selectedServicesIds);
                CalculateTotalCost();

                if (!deleteButtonAdded)
                {

                    NoTovar.Visible = false;

                    DataGridViewButtonColumn deleteButtonColumn = new DataGridViewButtonColumn();
                    deleteButtonColumn.HeaderText = "Видалити";
                    deleteButtonColumn.Text = "Видалити";
                    deleteButtonColumn.UseColumnTextForButtonValue = true;
                    dataGridView4.Columns.Add(deleteButtonColumn);
                    deleteButtonAdded = true; 
                }
            }
        }


        private void FindServicesInDatabase(List<int> serviceIds)
        {
            using (SqlConnection connection = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                connection.Open();

                string query = "SELECT * FROM [dbo].[Послуги] WHERE [ID послуги] IN (" + string.Join(",", serviceIds) + ")";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter(command);

                    DataTable dataTable = new DataTable();

                    adapter.Fill(dataTable);

                    dataGridView4.DataSource = dataTable;
                }
                connection.Close();
            }
        }

        private void CalculateTotalCost()
        {
            decimal totalCost = 0;

            foreach (DataGridViewRow row in dataGridView4.Rows)
            {
                if (row.Cells["Вартість послуги"].Value != null)
                {
                    totalCost += Convert.ToDecimal(row.Cells["Вартість послуги"].Value);
                }
            }

            sum.Text = $"Загальна сума: {totalCost}".ToString();
        }


        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex >= 0 && dataGridView4.Columns[e.ColumnIndex] is DataGridViewButtonColumn)
            {
                int serviceId = Convert.ToInt32(dataGridView4.Rows[e.RowIndex].Cells["ID послуги"].Value);

                selectedServicesIds.Remove(serviceId);

                dataGridView4.Rows.RemoveAt(e.RowIndex);

                CalculateTotalCost();
            }
        }

        private void FinalZamov_Click(object sender, EventArgs e)
        {
            if (selectedServicesIds.Count == 0)
            {
                MessageBox.Show("Выберите услуги для заказа.");
                return;
            }

            if (_userId == 0)
            {
                MessageBox.Show("Не удалось определить пользователя.");
                return;
            }

            using (SqlConnection connection = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                connection.Open();

                string insertQuery = "INSERT INTO [dbo].[Замовлення] ([ID Послуги], [ID Користувача], [Загальна сума], [Дата]) VALUES ";

                foreach (int serviceId in selectedServicesIds)
                {
                    string selectQuery = $"SELECT [Вартість послуги] FROM [dbo].[Послуги] WHERE [ID послуги] = {serviceId}";

                    using (SqlCommand command = new SqlCommand(selectQuery, connection))
                    {
                        object result = command.ExecuteScalar();

                        if (result != null && result != DBNull.Value)
                        {
                            decimal serviceCost = Convert.ToDecimal(result);

                            insertQuery += $"({serviceId}, {_userId}, {serviceCost}, GETDATE()), ";
                        }
                    }
                }


                insertQuery = insertQuery.TrimEnd(',', ' ');

                using (SqlCommand command = new SqlCommand(insertQuery, connection))
                {
                    int rowsAffected = command.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Заказ успешно добавлен.");
                    }
                    else
                    {
                        MessageBox.Show("Не удалось добавить заказ.");
                    }
                }
            }
        }

        private void Addmat_Click(object sender, EventArgs e)
        {



            AllBtn frm = new AllBtn();
            frm.AllConrol.SelectedTab = frm.Zatmat;

            frm.Upzatmat.Visible = false;

            frm.Show();
            frm.label3.Text = "Додавання";

        }

        private void Upmat_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string column1 = dataGridView6.SelectedRows[0].Cells["ID Матеріалу"].Value.ToString();
                string column2 = dataGridView6.SelectedRows[0].Cells["Назва Матеріалу"].Value.ToString();
                string column3 = dataGridView6.SelectedRows[0].Cells["Опис Матеріалу"].Value.ToString();
                string column4 = dataGridView6.SelectedRows[0].Cells["Одиниці Виміру"].Value.ToString();
                string column5 = dataGridView6.SelectedRows[0].Cells["Кількість На Складі"].Value.ToString();
                string column6 = dataGridView6.SelectedRows[0].Cells["Ціна за одиницю"].Value.ToString();

                UpData upData = new UpData
                {

                    IDМатеріалу = column1,
                    НазваМатеріалу = column2,
                    ОписМатеріалу = column3,
                    ОдиниціВиміру = column4,
                    КількістьНаСкладі = column5,
                    ЦінаЗаодиницю = column6,
                };



                AllBtn frm = new AllBtn();
                frm.AllConrol.SelectedTab = frm.Zatmat;

                frm.SetData(upData);

                frm.Addzatmat.Visible = false;
                frm.label3.Text = "редагування";
                frm.Show();

            }
        }

        private void Delmat_Click(object sender, EventArgs e)
        {


            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();

            SqlCommand command;

            if (dataGridView6.SelectedRows.Count > 0)
            {
                int selectedRowIndex = dataGridView6.SelectedRows[0].Index;
                int idToDelete = Convert.ToInt32(dataGridView6.Rows[selectedRowIndex].Cells["ID Матеріалу"].Value);
                command = new SqlCommand($"DELETE [dbo].[Витратні Матеріали] WHERE [ID Матеріалу] = @IDМатеріалу", con);
                command.Parameters.AddWithValue("@IDМатеріалу", idToDelete);
                command.ExecuteNonQuery(); 
            }


            SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM [dbo].[Витратні Матеріали]", con);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);
            dataGridView6.DataSource = dataTable;

            IDPos.Clear();

            con.Close();
            MessageBox.Show("Услуга удалена.");

        }

        private void Searchmat_Click(object sender, EventArgs e)
        {

            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();

            SqlCommand command;
            SqlDataAdapter da = new SqlDataAdapter();
            DataTable dt = new DataTable();



            if (string.IsNullOrWhiteSpace(combomat.Text))
            {
                command = new SqlCommand($"SELECT * FROM [dbo].[Витратні Матеріали]", con);
            }
            else
            {
                command = new SqlCommand($"SELECT * FROM [dbo].[Витратні Матеріали]  WHERE [Назва Матеріалу] LIKE '%' + @searchTerm + '%'", con);
                command.Parameters.AddWithValue("@searchTerm", combomat.Text);
            }



            da.SelectCommand = command;
            da.Fill(dt);

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Записи не найдены");
            }
            else
            {
                dataGridView6.DataSource = dt;
            }

            con.Close();

        }

        private void AddPotmat_Click(object sender, EventArgs e)
        {

            AllBtn frm = new AllBtn();
            frm.AllConrol.SelectedTab = frm.PotMat;

            frm.UpPotmat.Visible = false;
            frm.Show();
            frm.label3.Text = "Додавання";
        }


        private void UpPotmat_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string column1 = dataGridView7.SelectedRows[0].Cells["ID"].Value.ToString();
                string column2 = dataGridView7.SelectedRows[0].Cells["ID Замовлення"].Value.ToString();

                string column3;
                if (dataGridView7.Columns.Contains("ID Матеріалу"))
                    column3 = dataGridView7.SelectedRows[0].Cells["ID Матеріалу"].Value.ToString();
                else if (dataGridView7.Columns.Contains("Назва Матеріалу"))
                    column3 = dataGridView7.SelectedRows[0].Cells["Назва Матеріалу"].Value.ToString();
                else
                    column3 = ""; 

                string column4 = dataGridView7.SelectedRows[0].Cells["Кількість"].Value.ToString();


                UpData upData = new UpData
                {
                    IDМат = column1,
                    IDЗамовленняМат = column2,
                    IDМатеріалуМат = column3,
                    КількістьМат = column4,
                };

                AllBtn frm = new AllBtn();
                frm.AllConrol.SelectedTab = frm.PotMat;

                frm.AddPotmat.Visible = false;
                frm.label3.Text = "редагування";
                frm.SetData(upData);

                frm.Show();

            }
        }


        private void exituser(object sender, EventArgs e)
        {
            MessageBox.Show("Лив с позором");
        }

        private void ExitUser_Click(object sender, EventArgs e)
        {
            Login frm = new Login();
            frm.Show();
            this.Hide();
        }



        private void SearchPotmat_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();

                string sqlQuery = "SELECT vm.ID, vm.[ID Замовлення], m.[Назва Матеріалу], vm.[Кількість], vm.[Дата Використання] " +
                                  "FROM [dbo].[Витрачені матеріали] vm " +
                                  "LEFT JOIN [dbo].[Витратні Матеріали] m ON vm.[ID Матеріалу] = m.[ID Матеріалу] " +
                                  "WHERE 1 = 1";

                List<SqlParameter> parameters = new List<SqlParameter>();

                if (!string.IsNullOrWhiteSpace(IDVtrMat.Text))
                {
                    sqlQuery += " AND m.[Назва Матеріалу] LIKE @MaterialName";
                    parameters.Add(new SqlParameter("@MaterialName", "%" + IDVtrMat.Text + "%"));
                }

                if (dateTimePicker3.Value <= dateTimePicker4.Value)
                {
                    DateTime startDate = dateTimePicker3.Value.Date + new TimeSpan(0, 0, 0);
                    DateTime endDate = dateTimePicker4.Value.Date + new TimeSpan(23, 59, 59);

                    sqlQuery += " AND vm.[Дата Використання] BETWEEN @StartDate AND @EndDate";
                    parameters.Add(new SqlParameter("@StartDate", startDate));
                    parameters.Add(new SqlParameter("@EndDate", endDate));
                }

                using (SqlCommand command = new SqlCommand(sqlQuery, con))
                {
                    command.Parameters.AddRange(parameters.ToArray());

                    using (SqlDataAdapter da = new SqlDataAdapter(command))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        if (dt.Rows.Count == 0)
                        {
                            MessageBox.Show("Записи не найдены");
                        }
                        else
                        {
                            dataGridView7.DataSource = dt;
                        }
                    }
                }
            }
        }


        private void Export1_Click(object sender, EventArgs e)
        {
            ExportToWord();
        }

        private void ExportToWord()
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Add();

            Microsoft.Office.Interop.Word.Table table = wordDoc.Tables.Add(wordDoc.Range(), dataGridView3.Rows.Count + 1, dataGridView3.Columns.Count - 1);

            int wordColumnIndex = 1;
            foreach (DataGridViewColumn column in dataGridView3.Columns)
            {
                if (column.HeaderText != "ID Замовлення")
                {
                    table.Cell(1, wordColumnIndex).Range.Text = column.HeaderText;
                    wordColumnIndex++;
                }
            }
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                wordColumnIndex = 1;
                for (int j = 0; j < dataGridView3.Columns.Count; j++)
                {
                    if (dataGridView3.Columns[j].HeaderText != "ID Замовлення")
                    {
                        if (dataGridView3.Rows[i].Cells[j].Value != null)
                        {
                            if (dataGridView3.Columns[j].ValueType == typeof(DateTime)) 
                            {
                                DateTime dateValue = (DateTime)dataGridView3.Rows[i].Cells[j].Value;
                                table.Cell(i + 2, wordColumnIndex).Range.Text = dateValue.ToShortDateString();
                            }
                            else
                            {
                                table.Cell(i + 2, wordColumnIndex).Range.Text = dataGridView3.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                        else
                        {
                            table.Cell(i + 2, wordColumnIndex).Range.Text = "";
                        }
                        wordColumnIndex++;
                    }
                }
            }

            table.Borders.Enable = 1;

            wordApp.Visible = true;
        }

        private void Export2_Click(object sender, EventArgs e)
        {
            ExportToWord2();
        }

        private void ExportToWord2()
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Add();

            Microsoft.Office.Interop.Word.Table table = wordDoc.Tables.Add(wordDoc.Range(), dataGridView7.Rows.Count + 1, dataGridView7.Columns.Count - 2); 
           
            int wordColumnIndex = 1;
            foreach (DataGridViewColumn column in dataGridView7.Columns)
            {
                if (column.Index > 1 && column.Index < dataGridView7.Columns.Count) 
                {
                    table.Cell(1, wordColumnIndex).Range.Text = column.HeaderText;
                    wordColumnIndex++;
                }
            }

            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                wordColumnIndex = 1;
                for (int j = 0; j < dataGridView7.Columns.Count; j++)
                {
                    if (dataGridView7.Columns[j].Index > 1 && dataGridView7.Columns[j].Index < dataGridView7.Columns.Count) 
                    {
                        if (dataGridView7.Rows[i].Cells[j].Value != null)
                        {
                            if (dataGridView7.Columns[j].ValueType == typeof(DateTime) && j == dataGridView7.Columns.Count - 1) 
                            {
                                DateTime dateValue = (DateTime)dataGridView7.Rows[i].Cells[j].Value;
                                table.Cell(i + 2, wordColumnIndex).Range.Text = dateValue.ToShortDateString();
                            }
                            else
                            {
                                table.Cell(i + 2, wordColumnIndex).Range.Text = dataGridView7.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                        else
                        {
                            table.Cell(i + 2, wordColumnIndex).Range.Text = "";
                        }
                        wordColumnIndex++;
                    }
                }
            }

            table.Borders.Enable = 1;

            wordApp.Visible = true;
        }
    }
}
