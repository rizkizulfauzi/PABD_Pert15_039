using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace CRUDMahasiswaADO
{
    public partial class report : Form
    {
        DAL dbLogic = new DAL();

        static string connectionString =
            "Data Source=NJOELL18\\MRZULFAUZI;Initial Catalog=DBAkademikADO;User ID=sa;Password=Mrzxxx18";

        SqlConnection conn =
            new SqlConnection(connectionString);

        SqlDataAdapter da;
        DataTable dtMahasiswa;
        DataTable dtProdi;

        public report()
        {
            InitializeComponent();
        }

        public report(string prodi, DateTime tglMasuk)
        {
            InitializeComponent();

            try
            {
                dtMahasiswa =
                    dbLogic.getDataRekap(
                        prodi,
                        tglMasuk);

                dataGridView1.DataSource =
                    dtMahasiswa;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void report_Load(
            object sender,
            EventArgs e)
        {
            dtpTanggalMasuk.Format =
                DateTimePickerFormat.Custom;

            dtpTanggalMasuk.CustomFormat =
                "yyyy";

            dtpTanggalMasuk.ShowUpDown =
                true;

            cmbProdi.DropDownStyle =
                ComboBoxStyle.DropDownList;

            btnCetak.Enabled = false;

            try
            {
                if (conn.State ==
                    ConnectionState.Closed)
                {
                    conn.Open();
                }

                SqlCommand cmd =
                    new SqlCommand(
                        "SELECT NamaProdi FROM ProgramStudi",
                        conn);

                da = new SqlDataAdapter(cmd);

                dtProdi =
                    new DataTable();

                da.Fill(dtProdi);

                cmbProdi.DataSource =
                    dtProdi;

                cmbProdi.DisplayMember =
                    "NamaProdi";

                cmbProdi.ValueMember =
                    "NamaProdi";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void button1_Click(
            object sender,
            EventArgs e)
        {
            try
            {
                if (conn.State ==
                    ConnectionState.Closed)
                {
                    conn.Open();
                }

                SqlCommand cmd =
                    new SqlCommand(
                        "sp_Report",
                        conn);

                cmd.CommandType =
                    CommandType.StoredProcedure;

                cmd.Parameters.Add(
                    "@inProdi",
                    SqlDbType.VarChar,
                    50).Value =
                    cmbProdi.SelectedValue;

                cmd.Parameters.Add(
                    "@inTglMasuk",
                    SqlDbType.VarChar,
                    4).Value =
                    dtpTanggalMasuk
                    .Value.Year
                    .ToString();

                da =
                    new SqlDataAdapter(cmd);

                dtMahasiswa =
                    new DataTable();

                da.Fill(dtMahasiswa);

                dataGridView1.DataSource =
                    dtMahasiswa;

                btnCetak.Enabled =
                    dtMahasiswa.Rows.Count > 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void btnCetak_Click(
            object sender,
            EventArgs e)
        {
            Report2 frm =
                new Report2(
                    cmbProdi.SelectedValue
                        .ToString(),
                    dtpTanggalMasuk.Value);

            frm.Show();
            this.Hide();
        }

        private void simpanLog(
            string pesan)
        {
            string path =
                Path.Combine(
                    Application.StartupPath,
                    "errorlog.txt");

            File.AppendAllText(
                path,
                DateTime.Now +
                " - " +
                pesan +
                Environment.NewLine);
        }
    }
}