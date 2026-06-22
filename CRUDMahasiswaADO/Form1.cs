using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CRUDMahasiswaADO
{
    public partial class Form1 : Form
    {
        SqlConnection conn;
        string connectionString = "Data Source=Nimra\\NIMRA;Initial Catalog=DBAkademikADO;Integrated Security=True";
        DAL dbLogic = new DAL();
        private BindingSource bindingSource = new BindingSource();
        private DataTable dtMahasiswa = new DataTable();

        public Form1()
        {
            InitializeComponent();
            conn = new SqlConnection(connectionString);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cmbJK.DataSource = new string[] { "L", "P" };

            dgvMahasiswa.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvMahasiswa.MultiSelect = false;
            dgvMahasiswa.ReadOnly = true;
            dgvMahasiswa.AllowUserToAddRows = false;
            dgvMahasiswa.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            bindingNavigator1.BindingSource = bindingSource;

            LoadData();
        }

        private void LoadData()
        {
            try
            {
                dtMahasiswa = dbLogic.GetMhs();

                bindingSource.DataSource = dtMahasiswa;
                dgvMahasiswa.DataSource = bindingSource;

                BindControls();
                HitungTotal();
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("Load Error: " + ex.Message);
            }
        }

        private void HitungTotal()
        {
            try
            {
                int total = dbLogic.CountMhs();
                lblTotal.Text = "Total Mahasiswa: " + total;
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("Gagal menghitung total: " + ex.Message);
            }
        }

        private void BindControls()
        {
            txtNIM.DataBindings.Clear();
            txtNama.DataBindings.Clear();
            cmbJK.DataBindings.Clear();
            dtpTanggalLahir.DataBindings.Clear();
            txtAlamat.DataBindings.Clear();
            txtKodeProdi.DataBindings.Clear();

            txtNIM.DataBindings.Add("Text", bindingSource, "NIM");
            txtNama.DataBindings.Add("Text", bindingSource, "Nama");
            cmbJK.DataBindings.Add("Text", bindingSource, "JenisKelamin");

            dtpTanggalLahir.DataBindings.Add(
                "Value",
                bindingSource,
                "TanggalLahir",
                true,
                DataSourceUpdateMode.OnPropertyChanged,
                DateTime.Now
            );

            txtAlamat.DataBindings.Add("Text", bindingSource, "Alamat");
            txtKodeProdi.DataBindings.Add("Text", bindingSource, "NamaProdi");
        }

        private void ClearForm()
        {
            txtNIM.Enabled = true;
            txtNIM.Clear();
            txtNama.Clear();
            cmbJK.SelectedIndex = -1;
            txtAlamat.Clear();
            txtKodeProdi.Clear();
            dtpTanggalLahir.Value = DateTime.Now;
            fotoMhs.Image = null;
            txtNIM.Focus();
        }

        private byte[] ConvertImageToBytes(PictureBox pb)
        {
            if (pb.Image == null)
                return null;

            using (MemoryStream ms = new MemoryStream())
            {
                pb.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                return ms.ToArray();
            }
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] imgBytes = ConvertImageToBytes(fotoMhs);

                dbLogic.InsertMhs(txtNIM.Text, txtNama.Text, txtAlamat.Text, cmbJK.Text, dtpTanggalLahir.Value.Date, txtKodeProdi.Text, imgBytes);

                MessageBox.Show("Data mahasiswa berhasil ditambahkan");
                ClearForm();
                LoadData();
            }
            catch (SqlException ex)
            {
                SimpanLog("Rollback Insert :" + ex.Message);
                MessageBox.Show("SQL Error :" + ex.Message);
            }
            catch (Exception ex)
            {
                SimpanLog("General Error :" + ex.Message);
                MessageBox.Show("General Error :" + ex.Message);
            }
        }

        private void btnConncet_Click(object sender, EventArgs e)
        {
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
                MessageBox.Show("Connected!");
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void dtpTanngalLahir_ValueChanged(object sender, EventArgs e)
        {
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] imgBytes = ConvertImageToBytes(fotoMhs);

                dbLogic.UpdateMhs(txtNIM.Text, txtNama.Text, txtAlamat.Text, cmbJK.Text, dtpTanggalLahir.Value.Date, txtKodeProdi.Text, imgBytes);

                MessageBox.Show("Data mahasiswa berhasil diubah");
                ClearForm();
                LoadData();
            }
            catch (SqlException ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("SQL Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("General Error: " + ex.Message);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtNIM.Text))
                {
                    MessageBox.Show("NIM is required!");
                    return;
                }

                DialogResult dg = MessageBox.Show(
                    "Yakin ingin menghapus data?",
                    "Konfirmasi",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (dg == DialogResult.Yes)
                {
                    dbLogic.DeleteMhs(txtNIM.Text);
                    MessageBox.Show("Data mahasiswa berhasil dihapus");
                    ClearForm();
                    LoadData();
                }
            }
            catch (SqlException ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("SQL Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("General Error: " + ex.Message);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void btnBackup_Click(object sender, EventArgs e)
        {
            try
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();

                string query = @"
        IF OBJECT_ID('dbo.Mahasiswa_Backup') IS NOT NULL
            DROP TABLE dbo.Mahasiswa_Backup;

        SELECT * INTO dbo.Mahasiswa_Backup
        FROM dbo.Mahasiswa;
        ";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Backup updated (old backup replaced)");
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("Backup gagal: " + ex.Message);
            }
            finally
            {
                if (conn != null && conn.State != ConnectionState.Closed)
                {
                    conn.Close();
                }
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            try
            {
                dbLogic.resetData();
                MessageBox.Show("Data berhasil direset");
                LoadData();
            }
            catch (SqlException ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("SQL Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("General Error: " + ex.Message);
            }
        }

        private void btnTestInjection_Click(object sender, EventArgs e)
        {
            try
            {
                dbLogic.testInject(txtNIM.Text);
                LoadData();
            }
            catch (SqlException ex)
            {
                if (ex.Message.Contains("safe"))
                {
                    SimpanLog(ex.Message);
                    MessageBox.Show("SQL Error : Unsafe UPDATE operation not allowed");
                }
                else
                {
                    SimpanLog(ex.Message);
                    MessageBox.Show("SQL Error :" + ex.Message);
                }
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("General Error :" + ex.Message);
            }
        }

        public void SimpanLog(string message)
        {
            dbLogic.InsertLog(message);
        }

        private void btnRekapData_Click(object sender, EventArgs e)
        {
            RekapMahasiswa fm3 = new RekapMahasiswa();
            fm3.Show();
            this.Hide();
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                fotoMhs.Image = Image.FromFile(ofd.FileName);
                fotoMhs.SizeMode = PictureBoxSizeMode.StretchImage;
            }
        }

        private void dgvMahasiswa_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataRow row = ((DataRowView)bindingSource[e.RowIndex]).Row;

                if (row["Foto"] != DBNull.Value)
                {
                    byte[] imgBytes = (byte[])row["Foto"];
                    using (MemoryStream ms = new MemoryStream(imgBytes))
                    {
                        fotoMhs.Image = Image.FromStream(ms);
                        fotoMhs.SizeMode = PictureBoxSizeMode.StretchImage;
                    }
                }
                else
                {
                    fotoMhs.Image = null;
                }

                txtNIM.Enabled = false;
            }
        }

        private void btnImpExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel Workbook| *. xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataReader.ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataReader.ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                            DataTable dt = result.Tables[0];
                            dgvMahasiswa.DataSource = dt;
                            dgvMahasiswa.Enabled = false;

                            btnImpDb.Enabled = true;
                            btnInsert.Enabled = false;
                            btnUpdate.Enabled = false;
                            btnDelete.Enabled = false;
                            btnLoad.Enabled = false;
                            btnReset.Enabled = false;
                            btnTestInjection.Enabled = false;

                        }
                    }
                }
            }
        }

        private void btnImpDb_Click(object sender, EventArgs e)

        {

            try

            {

                DataTable dt = (DataTable)dgvMahasiswa.DataSource;

                if (dt == null || dt.Rows.Count == 0)

                {

                    MessageBox.Show("Tidak ada data untuk diimport.");

                    return;

                }

                int sukses = 0;

                foreach (DataRow row in dt.Rows)

                {

                    string nim = row["NIM"].ToString().Trim();
                    if (nim.Length > 10) nim = nim.Substring(0, 10);

                    string nama = row["Nama"].ToString().Trim();

                    string jk = row["JenisKelamin"].ToString().Trim();

                    string alamat = row["Alamat"].ToString().Trim();

                    string kodeProdi = row["NamaProdi"].ToString().Trim();

                    string fotoPath = row.Table.Columns.Contains("FotoPath")

                                        ? row["FotoPath"].ToString().Trim()

                                        : string.Empty;

                    if (string.IsNullOrEmpty(nim) || string.IsNullOrEmpty(nama))

                        continue;

                    DateTime tglLahir;

                    if (!DateTime.TryParse(row["TanggalLahir"].ToString(), out tglLahir))

                        continue;

                    byte[] ConvertImageFromPath(string path)

                    {

                        if (string.IsNullOrWhiteSpace(path))

                            return null;

                        if (!File.Exists(path))

                            return null;

                        return File.ReadAllBytes(path);

                    }

                    byte[] fotoBytes = ConvertImageFromPath(fotoPath);

                    dbLogic.InsertMhs(nim, nama, alamat, jk, tglLahir, kodeProdi, fotoBytes);

                    sukses++;

                }

                MessageBox.Show("Data mahasiswa berhasil ditambahkan: " + sukses + " data");

                ClearForm();

                LoadData();

                dgvMahasiswa.Enabled = true;

                btnImpDb.Enabled = false;

                btnInsert.Enabled = true;

                btnUpdate.Enabled = true;

                btnDelete.Enabled = true;

                btnLoad.Enabled = true;

                btnReset.Enabled = true;

                btnTestInjection.Enabled = true;

            }

            catch (SqlException ex)

            {

                SimpanLog("Rollback Insert :" + ex.Message);

                MessageBox.Show("SQL Error :" + ex.Message);

            }

            catch (Exception ex)

            {

                SimpanLog("General Error :" + ex.Message);

                MessageBox.Show("General Error :" + ex.Message);

            }

        }
    }
}