using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace AICchurch
{
    public partial class Login : Form
    {

        private bool isDragging = false;
        private int mouseX;
        private int mouseY;

        public Login()
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
            var passwordUser = HashMD5.hashPassword(txtpassword.Text);

            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();

                SqlCommand command = new SqlCommand($"SELECT COUNT(*) FROM [dbo].[Користувачі] WHERE Логін = @Логін AND Пароль = @Пароль", con);
                command.Parameters.Add("@Логін", SqlDbType.VarChar).Value = txtUserName.Text;
                command.Parameters.Add("@Пароль", SqlDbType.VarChar).Value = passwordUser;
                int count = (int)command.ExecuteScalar();

                if (count > 0)
                {
                    SqlCommand userCommand = new SqlCommand($"SELECT * FROM [dbo].[Користувачі] WHERE Логін = @Логін", con);
                    userCommand.Parameters.Add("@Логін", SqlDbType.VarChar).Value = txtUserName.Text;
                    SqlDataReader reader = userCommand.ExecuteReader();

                    if (reader.Read())
                    {
                        int userId = Convert.ToInt32(reader["ID"]);
                        string userName = reader["Логін"].ToString();
                        string userEmail = reader["Пошта"] == DBNull.Value ? "" : reader["Пошта"].ToString();
                        string userPassword = reader["Пароль"] == DBNull.Value ? "" : reader["Пароль"].ToString();
                        bool isAdmin = reader["is_admin"] == DBNull.Value ? false : Convert.ToBoolean(reader["is_admin"]);

        
                        checkUser currentUser = new checkUser(userName, isAdmin);

                        DataSend.text = txtUserName.Text;

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


        private void label2_Click(object sender, EventArgs e)
        {
            txtpassword.Clear();
            txtUserName.Clear();
            txtUserName.Focus();
        }

        private void label3_Click(object sender, EventArgs e)
        {
            Application.Exit();
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


        private void button2_Click(object sender, EventArgs e)
        {

            register frm = new register();
            frm.Show();
            this.Hide();

        }
    }
}
