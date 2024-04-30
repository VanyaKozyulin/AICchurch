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
using static AICchurch.Form1;

namespace AICchurch
{
    public partial class AllBtn : Form
    {
        private string connectionString = @"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True";
        private int originalWidth;


        private bool isDragging = false;
        private int mouseX;
        private int mouseY;


        public AllBtn()
        {
            InitializeComponent();

            label3.MouseDown += TaskPanel_MouseDown;
            label3.MouseMove += TaskPanel_MouseMove;
            label3.MouseUp += TaskPanel_MouseUp;


            string sql = "SELECT [ID Матеріалу], [Назва Матеріалу] FROM [dbo].[Витратні Матеріали]";
            AddToComboBox(sql, ComboNameID, "ID Матеріалу", "Назва Матеріалу");

            ComboNameID.SelectedIndexChanged += ComboNameID_SelectedIndexChanged;

            ComboNameID.SelectedIndex = -1;
            ComboNameID.Text = "";
        }

        private void ComboNameID_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboNameID.SelectedIndex != -1)
            {
                string selectedID = ComboNameID.SelectedValue.ToString();
                IDmat.Text = selectedID;
            }
        }

        public void AddToComboBox(string sql, ComboBox comboBox, string valueMember, string displayMember)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(sql, connection);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();

                var list = new BindingSource();
                list.DataSource = reader;

                comboBox.DataSource = list;
                comboBox.ValueMember = valueMember;
                comboBox.DisplayMember = displayMember;

                reader.Close();
            }
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


        private void Addbtn_Click(object sender, EventArgs e)
        {
            var password = HashMD5.hashPassword(txtPass.Text);

            bool isAdmin = checkBoxAdmin.Checked;

            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();
                SqlCommand command = new SqlCommand($"INSERT INTO [dbo].[Користувачі] (Логін, Пошта, Пароль, is_admin) VALUES(@Логін, @Пошта, @Пароль, @IsAdmin)", con);

                command.Parameters.AddWithValue("@Логін", txtLogin.Text);
                command.Parameters.AddWithValue("@Пошта", txtemail.Text);
                command.Parameters.AddWithValue("@Пароль", password);
                command.Parameters.AddWithValue("@IsAdmin", isAdmin);

                command.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Користувач успішно доданий");
            }
        }

        public void SetData(UpData upData)
        {
            txtID.Text = upData.ID;
            txtLogin.Text = upData.Логін;
            txtemail.Text = upData.Пошта;


            IDPos.Text = upData.IDпослуги;
            NamePos.Text = upData.НазваПослуги;
            OpisPos.Text = upData.ОписПослуги;
            CostPos.Text = upData.ВартістьПослуги;

            IDzatmat.Text = upData.IDМатеріалу;
            Namezatmat.Text = upData.НазваМатеріалу;
            Opiszatmat.Text = upData.ОписМатеріалу;
            Unitszatmat.Text = upData.ОдиниціВиміру;
            Colozatmat.Text = upData.КількістьНаСкладі;
            Costzatmat.Text = upData.ЦінаЗаодиницю;

            IDPotmat.Text = upData.IDМат;
            IDzam.Text = upData.IDЗамовленняМат;
            IDmat.Text = upData.IDМатеріалуМат;
            ColoPotmat.Text = upData.КількістьМат;
        }



        private void UpdateUser_Click(object sender, EventArgs e)
        {
            bool isAdmin = checkBoxAdmin.Checked;

            var password = HashMD5.hashPassword(txtPass.Text);

            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();

                SqlCommand command = new SqlCommand($"UPDATE [dbo].[Користувачі] SET", con);

                if (!string.IsNullOrEmpty(txtLogin.Text))
                {
                    command.CommandText += " Логін=@Логін,";
                    command.Parameters.AddWithValue("@Логін", txtLogin.Text);
                }

                if (!string.IsNullOrEmpty(txtemail.Text))
                {
                    command.CommandText += " Пошта=@Пошта,";
                    command.Parameters.AddWithValue("@Пошта", txtemail.Text);
                }

                if (!string.IsNullOrEmpty(txtPass.Text))
                {
                    command.CommandText += " Пароль=@Пароль,";
                    command.Parameters.AddWithValue("@Пароль", password);
                }

                command.CommandText += " is_admin=@IsAdmin WHERE ID=@ID";
                command.Parameters.AddWithValue("@IsAdmin", isAdmin);
                command.Parameters.AddWithValue("@ID", int.Parse(txtID.Text));

                if (command.CommandText.EndsWith(","))
                {
                    command.CommandText = command.CommandText.Remove(command.CommandText.Length - 1);
                }

                command.ExecuteNonQuery();

                con.Close();

                MessageBox.Show("оновило!");
            }
        }


        private void AddPos_Click(object sender, EventArgs e)
        {

            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();
            SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Послуги] ([Назва послуги], [Опис послуги], [Вартість послуги]) VALUES(@НазваПослуги, @ОписПослуги, @ВартістьПослуги)", con);

            command.Parameters.AddWithValue("@НазваПослуги", NamePos.Text);
            command.Parameters.AddWithValue("@ОписПослуги", OpisPos.Text);
            command.Parameters.AddWithValue("@ВартістьПослуги", CostPos.Text);

            command.ExecuteNonQuery();

            con.Close();
            MessageBox.Show("Додано успішно");

        }



        private void PosUp_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();

            SqlCommand command = new SqlCommand($"UPDATE [dbo].[Послуги] SET", con);

            if (!string.IsNullOrEmpty(NamePos.Text))
            {
                command.CommandText += " [Назва послуги]=@НазваПослуги,";
                command.Parameters.AddWithValue("@НазваПослуги", NamePos.Text);
            }

            if (!string.IsNullOrEmpty(OpisPos.Text))
            {
                command.CommandText += " [Опис послуги]=@ОписПослуги,";
                command.Parameters.AddWithValue("@ОписПослуги", OpisPos.Text);
            }

            if (!string.IsNullOrEmpty(CostPos.Text))
            {
                command.CommandText += " [Вартість послуги]=@ВартістьПослуги,";
                command.Parameters.AddWithValue("@ВартістьПослуги", CostPos.Text);
            }

            if (command.CommandText.EndsWith(","))
            {
                command.CommandText = command.CommandText.Remove(command.CommandText.Length - 1);
            }

            command.CommandText += " WHERE [ID послуги]=@IDпослуги";
            command.Parameters.AddWithValue("@IDпослуги", int.Parse(IDPos.Text));

            command.ExecuteNonQuery();

            con.Close();
            MessageBox.Show("оновило!");
        }



        private void AllBtn_Load(object sender, EventArgs e)
        {
            originalWidth = OpisPos.Width;

            OpisPos.Enter += TextBox_Enter;
            OpisPos.Leave += TextBox_Leave;

        }

        private void TextBox_Enter(object sender, EventArgs e)
        {

            TextBox textBox = (TextBox)sender;
            textBox.Width += 50;
            textBox.Left -= 25; 
            textBox.Height += 50;

        }

        private void TextBox_Leave(object sender, EventArgs e)
        {

            TextBox textBox = (TextBox)sender;
            textBox.Width = originalWidth;
            textBox.Height -= 50;
            textBox.Left += 25;
        }

        private void txtPass_TextChanged(object sender, EventArgs e)
        {

        }

        private void Addzatmat_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();

            if (string.IsNullOrEmpty(Namezatmat.Text) || string.IsNullOrEmpty(Opiszatmat.Text) || string.IsNullOrEmpty(Unitszatmat.Text) || string.IsNullOrEmpty(Colozatmat.Text) || string.IsNullOrEmpty(Costzatmat.Text))
            {
                MessageBox.Show("Все поля должны быть заполнены!");
                con.Close();
                return;
            }

            if (!int.TryParse(Colozatmat.Text, out int colo) || !int.TryParse(Costzatmat.Text, out int cost))
            {
                MessageBox.Show("Кількість на складі та Ціна за одиницю повинні бути числом!");
                con.Close();
                return;
            }

            SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Витратні Матеріали] ([Назва Матеріалу], [Опис Матеріалу], [Одиниці Виміру], [Кількість На Складі], [Ціна за одиницю]) VALUES (@НазваМатеріалу, @ОписМатеріалу, @ОдиниціВиміру, @КількістьНаСкладі, @Ціназаодиницю)", con);

            command.Parameters.AddWithValue("@НазваМатеріалу", Namezatmat.Text);
            command.Parameters.AddWithValue("@ОписМатеріалу", Opiszatmat.Text);
            command.Parameters.AddWithValue("@ОдиниціВиміру", Unitszatmat.Text);
            command.Parameters.AddWithValue("@КількістьНаСкладі", colo);
            command.Parameters.AddWithValue("@Ціназаодиницю", cost);

            command.ExecuteNonQuery();

            con.Close();
            MessageBox.Show("Додано успішно");
        }


        private void Upzatmat_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True");
            con.Open();

            SqlCommand command = new SqlCommand($"UPDATE [dbo].[Витратні Матеріали] SET", con);

            if (!string.IsNullOrEmpty(Namezatmat.Text))
            {
                command.CommandText += " [Назва Матеріалу]=@НазваМатеріалу,";
                command.Parameters.AddWithValue("@НазваМатеріалу", Namezatmat.Text);
            }

            if (!string.IsNullOrEmpty(Opiszatmat.Text))
            {
                command.CommandText += " [Опис Матеріалу]=@ОписМатеріалу,";
                command.Parameters.AddWithValue("@ОписМатеріалу", Opiszatmat.Text);
            }

            if (!string.IsNullOrEmpty(Unitszatmat.Text))
            {
                command.CommandText += " [Одиниці Виміру]=@ОдиниціВиміру,";
                command.Parameters.AddWithValue("@ОдиниціВиміру", Unitszatmat.Text);
            }

            if (!string.IsNullOrEmpty(Colozatmat.Text))
            {
                if (int.TryParse(Colozatmat.Text, out int colo))
                {
                    command.CommandText += " [Кількість На Складі]=@КількістьНаСкладі,";
                    command.Parameters.AddWithValue("@КількістьНаСкладі", colo);
                }
                else
                {
                    MessageBox.Show("Кількість на складі повинна бути числом!");
                    con.Close();
                    return;
                }
            }

            if (!string.IsNullOrEmpty(Costzatmat.Text))
            {
                if (int.TryParse(Costzatmat.Text, out int cost))
                {
                    command.CommandText += " [Ціна за одиницю]=@Ціназаодиницю,";
                    command.Parameters.AddWithValue("@Ціназаодиницю", cost);
                }
                else
                {
                    MessageBox.Show("Ціна за одиницю повинна бути числом!");
                    con.Close();
                    return;
                }
            }

            if (command.CommandText.EndsWith(","))
            {
                command.CommandText = command.CommandText.Remove(command.CommandText.Length - 1);
            }

            command.CommandText += " WHERE [ID Матеріалу]=@IDМатеріалу";
            command.Parameters.AddWithValue("@IDМатеріалу", int.Parse(IDzatmat.Text));

            command.ExecuteNonQuery();

            con.Close();
            MessageBox.Show("Оновлено!");
        }


        private void AddPotmat_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();
                SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Витрачені матеріали] ([ID Замовлення], [ID Матеріалу], [Кількість], [Дата Використання]) VALUES (@IDЗамовлення, @IDМатеріалу, @Кількість, GETDATE())", con);

                command.Parameters.AddWithValue("@IDЗамовлення", IDzam.Text);
                command.Parameters.AddWithValue("@IDМатеріалу", IDmat.Text);
                command.Parameters.AddWithValue("@Кількість", ColoPotmat.Text);

                command.ExecuteNonQuery();

                UpdateMaterialQuantity(con, int.Parse(IDmat.Text), int.Parse(ColoPotmat.Text), true);

                con.Close();
                MessageBox.Show("Додано успішно");
            }
        }

        private void UpPotmat_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(@"Data Source=yaPC\MSSQLSERVER01; Initial Catalog=church; Integrated Security=True"))
            {
                con.Open();

                if (!int.TryParse(ColoPotmat.Text, out int colo))
                {
                    MessageBox.Show("Кількість повинна бути числом!");
                    con.Close();
                    return;
                }

                SqlCommand getQuantityCommand = new SqlCommand($"SELECT [Кількість] FROM [dbo].[Витрачені матеріали] WHERE [ID] = @ID", con);
                getQuantityCommand.Parameters.AddWithValue("@ID", int.Parse(IDPotmat.Text));
                int previousQuantity = Convert.ToInt32(getQuantityCommand.ExecuteScalar());

                SqlCommand command = new SqlCommand($"UPDATE [dbo].[Витрачені матеріали] SET", con);

                if (!string.IsNullOrEmpty(IDzam.Text))
                {
                    command.CommandText += " [ID Замовлення]=@IDЗамовлення,";
                    command.Parameters.AddWithValue("@IDЗамовлення", IDzam.Text);
                }

                if (!string.IsNullOrEmpty(IDmat.Text))
                {
                    command.CommandText += " [ID Матеріалу]=@IDМатеріалу,";
                    command.Parameters.AddWithValue("@IDМатеріалу", IDmat.Text);
                }

                if (!string.IsNullOrEmpty(ColoPotmat.Text))
                {
                    command.CommandText += " [Кількість]=@Кількість,";
                    command.Parameters.AddWithValue("@Кількість", colo);
                }

                if (command.CommandText.EndsWith(","))
                {
                    command.CommandText = command.CommandText.Remove(command.CommandText.Length - 1);
                }

                command.CommandText += " WHERE [ID] = @ID";
                command.Parameters.AddWithValue("@ID", int.Parse(IDPotmat.Text));

                command.ExecuteNonQuery();

                int newQuantity = colo;

                bool isIncreasing = newQuantity > previousQuantity;

                UpdateMaterialQuantity(con, int.Parse(IDmat.Text), Math.Abs(newQuantity - previousQuantity), isIncreasing);

                con.Close();
                MessageBox.Show("оновило!");
            }
        }

        private void UpdateMaterialQuantity(SqlConnection connection, int materialId, int quantity, bool subtract)
        {
            string operation = subtract ? "-" : "+";
            string query = $"UPDATE [dbo].[Витратні Матеріали] SET [Кількість На Складі] = [Кількість На Складі] {operation} @Quantity WHERE [ID Матеріалу] = @MaterialId";

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@MaterialId", materialId);
                command.Parameters.AddWithValue("@Quantity", quantity);
                command.ExecuteNonQuery();
            }
        }

        private void exit_Click(object sender, EventArgs e)
        {

            this.Hide();

        }
    }
}
