using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Drawing.Text;
using System.Data.OleDb;

namespace fms
{
    public partial class Form1 : Form
    {
        DataSet ds = new DataSet();
        SqlConnection cs = new SqlConnection("Data source = SIMBA;initial catalog=Sarah;Persist Security Info=True;User ID=sa;Password=pass");

        SqlDataAdapter da = new SqlDataAdapter();

        BindingSource bs = new BindingSource();
        int stu_id;
        String stu_fname, stu_lname, stu_name, stu_dept, stu_add, stu_img, stu_fac;
        public byte [] stu_pic;
        ImageConverter converter = new ImageConverter();

        public Form1()
        {
            InitializeComponent();
        }

        public void GetData()
        {
            stu_id = int.Parse(textBox1.Text);
            stu_fname = textBox2.Text;
            stu_lname = textBox3.Text;
            stu_name = stu_fname +" "+ stu_lname;
            stu_dept = textBox4.Text;
            stu_add = textBox5.Text;
            stu_img = PB.ImageLocation;
            stu_fac = textBox6.Text;
            Image img = PB.Image;
            stu_pic = (byte[])converter.ConvertTo(img, typeof(byte[]));

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            errorProvider1.Clear();
            if (textBox2.Text == "")
                errorProvider1.SetError(textBox2, "Please Provide the First Name!");
            else if (textBox3.Text == "")
                errorProvider1.SetError(textBox3, "Please Provide Book Last Name!");
            else if (textBox4.Text == "")
                errorProvider1.SetError(textBox4, "Please Provide Course!");
            else if (textBox5.Text == "")
                errorProvider1.SetError(textBox5, "Please Provide Department!");
            else if (textBox6.Text == "")
                errorProvider1.SetError(textBox6, "Please Provide Address!");
            else
            {
                try
                {
                    GetData();
                    string subPath = @"C:\Program Files\FilesOfFiles\";
                    try
                    {
                        bool exists = Directory.Exists(subPath);
                        if (!exists)
                        {
                            DirectoryInfo dir = Directory.CreateDirectory(subPath);
                            dir.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                            //DirectorySecurity dSecurity = dir.GetAccessControl();
                            //dSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), FileSystemRights.FullControl, InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit, PropagationFlags.NoPropagateInherit, AccessControlType.Allow));
                            //dir.SetAccessControl(dSecurity);
                        }
                    }
                    catch(Exception c)
                    {
                        MessageBox.Show(c.Message, "Could not create folder");
                    }
                    OleDbConnection myConnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0;" + @"Data source= C:\Program Files\FilesOfFiles\FileManagerDb.accdb");
                    OleDbCommand myCommand = new OleDbCommand();
                    myConnection.Open();
                    myCommand.Connection = myConnection;
                    myCommand.CommandText = "INSERT INTO MyTable (Stu_Id, Stu_Name, Stu_Dept, Stu_Fac, Stu_Add, Stu_Img) VALUES (@p1, @p2, @p3, @p4, @p5, @p6)";
                    myCommand.Parameters.AddWithValue("@p1", stu_id);
                    myCommand.Parameters.AddWithValue("@p2", stu_name);
                    myCommand.Parameters.AddWithValue("@p3", stu_dept);
                    myCommand.Parameters.AddWithValue("@p4", stu_fac);
                    myCommand.Parameters.AddWithValue("@p5", stu_add);
                    myCommand.Parameters.AddWithValue("@p6", stu_pic);
                    myCommand.ExecuteNonQuery();
                    myCommand.Connection.Close();
                    MessageBox.Show("New Record Inserted", "PoGWorld", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //da.InsertCommand = new SqlCommand("insert into Training values(@FirstName,@LastName,@Course,@Department,@Address)", cs);
                    //da.InsertCommand.Parameters.Add("@FirstName", SqlDbType.VarChar).Value = textBox2.Text;
                    //da.InsertCommand.Parameters.Add("@LastName", SqlDbType.VarChar).Value = textBox3.Text;
                    //da.InsertCommand.Parameters.Add("@Course", SqlDbType.VarChar).Value = textBox4.Text;
                    //da.InsertCommand.Parameters.Add("@Department", SqlDbType.VarChar).Value = textBox5.Text;
                    //da.InsertCommand.Parameters.Add("@Address", SqlDbType.VarChar).Value = textBox5.Text;

                    //cs.Open();
                    //da.InsertCommand.ExecuteNonQuery();
                    MessageBox.Show("Successfully Saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    myConnection.Close();
                }
                catch 
                {
                    MessageBox.Show("Data Exists!", "File Manager", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            OleDbConnection myConnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0;" + @"Data source= C:\Program Files\FilesOfFiles\FileManagerDb.accdb");
            OleDbCommand myCommand = new OleDbCommand();
            myConnection.Open();
            myCommand.Connection = myConnection;
            myCommand.CommandText = "select Stu_Id as 'MATRIC NO', Stu_Name as NAME, Stu_Dept as DEPARTMENT, Stu_Fac as FACULTY, Stu_Add as ADDRESS, Stu_Img as PASSPORT from MyTable order by Stu_Name";
            myCommand.ExecuteNonQuery();
            using (myConnection)
            {
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(myCommand.CommandText, myConnection))
                {
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    dgv.DataSource = ds.Tables[0];
                }
            }
            myCommand.Connection.Close();

        }

        private void First_Click(object sender, EventArgs e)
        {
            if (groupBox2.Visible) groupBox2.Hide();
        }

        private void Next_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "")
                errorProvider1.SetError(groupBox1, "Please Register yourself first!");
            else
            {
                label10.Text = textBox1.Text;
                int store = int.Parse(label10.Text);
                OleDbConnection myConnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0;" + @"Data source= C:\Program Files\FilesOfFiles\FileManagerDb.accdb");
                OleDbCommand myCommand = new OleDbCommand();
                myConnection.Open();
                myCommand.Connection = myConnection;
                myCommand.CommandText = "select count(*) from MyTable where Stu_Id = " + store;
                int stu_count = (int) myCommand.ExecuteScalar();
                if (stu_count == 0) MessageBox.Show("Please Save Before Proceeding", "File Manager", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                {
                    if (!groupBox2.Visible) groupBox2.Show();
                    if (textBox11.Text == "")
                        errorProvider1.SetError(textBox11, "Please Provide Course code!");
                    else if (textBox12.Text == "")
                        errorProvider1.SetError(textBox12, "Please Provide Course Title!");
                    else if (textBox10.Text == "")
                        errorProvider1.SetError(textBox10, "Please Provide Department");
                    else if (textBox9.Text == "")
                        errorProvider1.SetError(textBox9, "Please Provide Course Unit!");
                    else
                    {

                    }
                        //else MessageBox.Show("Box is visible");
                }
                    myCommand.Connection.Close();
            }
        }

        private void Previous_Click(object sender, EventArgs e)
        {
            bs.MovePrevious();
            dgvUpdate();
            ViewRecords();
        }

        private void Last_Click(object sender, EventArgs e)
        {
            bs.MoveLast();
            dgvUpdate();
            ViewRecords();
        }
        private void dgvUpdate()
        {
            dgv.ClearSelection();
            dgv.Rows[bs.Position].Selected = true;
        }
        private void ViewRecords()
        {
            label3.Text = "  ViewRecords  " + bs.Position + " Of " + (bs.Count - 1);

        }
        private void clear()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";

        }
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection myConnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0;" + @"Data source= C:\Program Files\FilesOfFiles\FileManagerDb.accdb");
                OleDbCommand myCommand = new OleDbCommand();
                myConnection.Open();
                myCommand.Connection = myConnection;
                try
                {
                    GetData();
                    myCommand.CommandText = "UPDATE MyTable SET Stu_Name = @p2, Stu_Dept = @p3, Stu_Fac = @p4, Stu_Add = @p5, Stu_Img = @p6 where Stu_Id = @p1";
                    myCommand.Parameters.AddWithValue("@p1", stu_id);
                    myCommand.Parameters.AddWithValue("@p2", stu_name);
                    myCommand.Parameters.AddWithValue("@p3", stu_dept);
                    myCommand.Parameters.AddWithValue("@p4", stu_fac);
                    myCommand.Parameters.AddWithValue("@p5", stu_add);
                    myCommand.Parameters.AddWithValue("@p6", stu_pic);
                    myCommand.ExecuteNonQuery();
                    MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch(Exception i)
                {
                    MessageBox.Show(i.Message);
                }
                clear();
            }
            
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr;
                dr = MessageBox.Show("Are U Sure?There is No Undo Once Records is Deleted ", "Confirmation Deletion", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {

                    da.DeleteCommand = new SqlCommand("Delete from Training where StudentID=@StudentID", cs);
                    da.DeleteCommand.Parameters.Add("@StudentID", SqlDbType.Int).Value = ds.Tables[0].Rows[bs.Position][0];

                    cs.Open();
                    da.DeleteCommand.ExecuteNonQuery();

                    cs.Close();
                    ds.Clear();
                    da.Fill(ds);
                    MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                {
                    MessageBox.Show("Deletion Canceled");
                }
            }


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        bool lr = true;
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (lr)
            {

                lblmoving.Location = new Point(lblmoving.Location.X + 5, lblmoving.Location.Y);

            }
            else
            {
                lblmoving.Location = new Point(lblmoving.Location.X - 5, lblmoving.Location.Y);
            }
            if (lblmoving.Location.X + lblmoving.Width >= this.Width)
            {
                lr = false;
            }
            if (lblmoving.Location.X <= 0)
            {
                lr = true;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public void btnBrowser_Click(object sender, EventArgs e)
        {

            OpenFileDialog OpenFile = new OpenFileDialog();
            try
            {
                OpenFile.FileName = "";
                OpenFile.Title = "Select a Passport-sized Photograph";
                OpenFile.Filter = "Image files: (*.jpg)|*.jpg|(*.jpeg)|*.jpeg|(*.png)|*.png|(*.Gif)|*.Gif|(*.bmp)|*.bmp| All Files (*.*)|*.*";
                DialogResult res = OpenFile.ShowDialog();
                if (res == DialogResult.OK)
                {
                    this.PB.Image = Image.FromFile(OpenFile.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "error!!!");
            }

        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            clear();
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow dgvr = this.dgv.Rows[e.RowIndex];
                textBox1.Text = dgvr.Cells["'MATRIC NO'"].Value.ToString();
                textBox2.Text = dgvr.Cells["NAME"].Value.ToString();
                textBox4.Text = dgvr.Cells["DEPARTMENT"].Value.ToString();
                textBox5.Text = dgvr.Cells["ADDRESS"].Value.ToString();
                textBox6.Text = dgvr.Cells["FACULTY"].Value.ToString();
                try 
                {
                    MemoryStream ms = new MemoryStream((byte[])dgvr.Cells["PASSPORT"].Value);
                    PB.Image = Image.FromStream(ms);
                    
                }
                catch
                {
                    PB.Image = Properties.Resources.user_96;
                }
             }
        }
        public void TextOnImage(String i)
        {
            PB.Paint += new PaintEventHandler((sender, e) =>
            {
                e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                SizeF textSize = e.Graphics.MeasureString(i, Font);
                PointF locationToDraw = new PointF();
                locationToDraw.X = (PB.Width / 2) - (textSize.Width / 2);
                locationToDraw.Y = (PB.Height / 2) - (textSize.Height / 2);
                e.Graphics.DrawString(i, Font, Brushes.Black, locationToDraw);
            });
        }
        
        private void BRemove_Click(object sender, EventArgs e)
        {
            //this.PB.Image = System.Drawing.Image.FromFile(Application.StartupPath.ToString() + "\\Image\\zawadi.jpg");
        }

        private void lblmoving_Click(object sender, EventArgs e)
        {

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            label5.Text = DateTime.Now.ToString();
        }

        private void PB_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
