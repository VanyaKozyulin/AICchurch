using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Cryptography;

namespace AICchurch
{
    public partial class register : Form
    {


        private bool isDragging = false;
        private int mouseX;
        private int mouseY;


        public register()
        {
            InitializeComponent();

            TaskPanel.MouseDown += TaskPanel_MouseDown;
            TaskPanel.MouseMove += TaskPanel_MouseMove;
            TaskPanel.MouseUp += TaskPanel_MouseUp;

            txtpassword.UseSystemPasswordChar = true;

        }

        private void TaskPanel_MouseDown(object sender, MouseEventArgs e)
        {
            isDragging = true;
            mouseX = e.X;
            mouseY = e.Y;
        }

        private void TaskPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                int deltaX = e.X - mouseX;
                int deltaY = e.Y - mouseY;
                Left += deltaX;
                Top += deltaY;
            }
        }

        private void TaskPanel_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }



        private void button1_Click(object sender, EventArgs e)
        {
            var login = txtUserName.Text;
            var email = txtemail.Text;
            var password = txtpassword.Text;

            bool hasError = false;

            if (password.Length < 6)
            {
                label7.Text = "Не менее 6 символів";
                hasError = true;
            }

            if (login.Length < 3)
            {
                label9.Text = "Не менее 3 символів";
                hasError = true;
            }

            if (!password.Any(char.IsUpper))
            {
                label7.Text = "Має бути хоча б одна велика літера";
                hasError = true;
            }

            if (login.Any(char.IsPunctuation))
            {
                label9.Text = "Не повинно бути спец символів";
                hasError = true;
            }

            if (!email.Contains("@"))
            {
                label8.Text = "Введіть коректну адресу";
                hasError = true;
            }

            if (hasError)
                return;

            var hashedPassword = HashMD5.hashPassword(password);

            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();

                SqlCommand checkLoginCmd = new SqlCommand("SELECT COUNT(*) FROM [dbo].[Користувачі] WHERE Логін = @Логін", con);
                checkLoginCmd.Parameters.AddWithValue("@Логін", login);
                int loginCount = (int)checkLoginCmd.ExecuteScalar();

                SqlCommand checkEmailCmd = new SqlCommand("SELECT COUNT(*) FROM [dbo].[Користувачі] WHERE Пошта = @Пошта", con);
                checkEmailCmd.Parameters.AddWithValue("@Пошта", email);
                int emailCount = (int)checkEmailCmd.ExecuteScalar();

            
                if (loginCount > 0)
                {
                    MessageBox.Show("Такий логін вже існує");
                    return;
                }

                if (emailCount > 0)
                {
                    MessageBox.Show("Така Пошта вже існує");
                    return;
                }

                SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Користувачі] (Логін, Пошта, Пароль, is_admin) VALUES (@Логін, @Пошта, @Пароль, 0)", con);

                command.Parameters.AddWithValue("@Логін", login);
                command.Parameters.AddWithValue("@Пошта", email);
                command.Parameters.AddWithValue("@Пароль", hashedPassword);

                command.ExecuteNonQuery();
                con.Close();

                PerformLogin(login, hashedPassword);
            }
        }

        private void PerformLogin(string login, string password)
        {
            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();

                SqlCommand command = new SqlCommand($"SELECT COUNT(*) FROM [dbo].[Користувачі] WHERE Логін = @Логін AND Пароль = '{password}'", con);
                command.Parameters.Add("@Логін", SqlDbType.VarChar).Value = login;
                int count = (int)command.ExecuteScalar();

                if (count > 0)
                {
                    SqlCommand userCommand = new SqlCommand($"SELECT * FROM [dbo].[Користувачі] WHERE Логін = @Логін", con);
                    userCommand.Parameters.Add("@Логін", SqlDbType.VarChar).Value = login;
                    SqlDataReader reader = userCommand.ExecuteReader();

                    if (reader.Read())
                    {
                        int userId = Convert.ToInt32(reader["ID"]);
                        string userName = reader["Логін"].ToString();
                        string email = reader["Пошта"] == DBNull.Value ? "" : reader["Пошта"].ToString();
                        bool isAdmin = reader["is_admin"] == DBNull.Value ? false : Convert.ToBoolean(reader["is_admin"]);

                        checkUser currentUser = new checkUser(userName, isAdmin);

              
                        DataSend.text = login;

                        Form1 frm = new Form1(currentUser, userId);

                        frm.Show();
                        this.Hide();
                    }

                    reader.Close();
                }
                else
                {
                    MessageBox.Show("Учетная запись не существует. Пожалуйста, проверьте логин и пароль.", "Ошибка входа", MessageBoxButtons.OK);
                }
            }
        }




        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                txtpassword.UseSystemPasswordChar = false;
            }
            else
            {
                txtpassword.UseSystemPasswordChar = true;
            }

        }

        private void label3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Login frm = new Login();
            frm.Show();
            this.Hide();
        }

    }
}
