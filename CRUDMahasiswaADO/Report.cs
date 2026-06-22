using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using CrystalDecisions.CrystalReports.Engine;

namespace CRUDMahasiswaADO
{
    public partial class Report : Form
    {
        string connectionString = "Data Source=Nimra\\NIMRA;Initial Catalog=DBAkademikADO;Integrated Security=True";

        SqlConnection conn;
        SqlDataAdapter da;
        DataTable dtMahasiswa;

        ListMahasiswa listMahasiswa = new ListMahasiswa();

        string prodi { get; set; }
        DateTime tglmasuk { get; set; }

        // Modified constructor
        public Report(string Prodi, DateTime TglMasuk)
        {
            InitializeComponent();

            prodi = Prodi;
            tglmasuk = TglMasuk;

            conn = new SqlConnection(connectionString);

            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();

                SqlCommand cmd = new SqlCommand("sp_Report", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@inProdi", prodi);
                cmd.Parameters.AddWithValue("@inTglMsuk", tglmasuk.Year);

                da = new SqlDataAdapter(cmd);
                dtMahasiswa = new DataTable();
                da.Fill(dtMahasiswa);

                conn.Close();

                // Feed data into Crystal Report
                foreach (DataRow row in dtMahasiswa.Rows)
                {
                    listMahasiswa.Add(new MahasiswaData
                    {
                        Nama = row["Nama"].ToString(),
                        JenisKelamin = row["JenisKelamin"].ToString(),
                        Alamat = row["Alamat"].ToString(),
                        NamaProdi = row["NamaProdi"].ToString(),
                        TanggalDaftar = Convert.ToDateTime(row["TanggalDaftar"])
                    });
                }

                ReportMahasiswa report = new ReportMahasiswa();
                report.SetDataSource(listMahasiswa);
                crystalReportViewer1.ReportSource = report;
                crystalReportViewer1.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Gagal load data: " + ex.Message);
            }
        }
    }
}