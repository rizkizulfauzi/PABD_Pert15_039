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

namespace CRUDMahasiswaADO
{
    public partial class Report2: Form
    {
        static string connectionString = "Data Source=NJOELL18\\MRZULFAUZI;Initial Catalog=DBAkademikADO;User ID=sa;Password=Mrzxxx18";


        SqlConnection conn = new SqlConnection(connectionString);
        SqlDataAdapter da;
        DataTable dtMahasiswa;

        CrystalReport1 CrystalReport1 = new CrystalReport1();

        string prodi { get; set; }
        DateTime tglmasuk { get; set; }

        public Report2(string Prodi, DateTime TglMasuk)
        {
            InitializeComponent();

            prodi = Prodi;
            tglmasuk = TglMasuk;

            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }

                SqlCommand cmd = new SqlCommand("sp_Report", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@inProdi", prodi);
                cmd.Parameters.AddWithValue("@inTglMasuk", tglmasuk.Year);

                da = new SqlDataAdapter(cmd);
                dtMahasiswa = new DataTable();
                da.Fill(dtMahasiswa);

                conn.Close();

                CrystalReport1.SetDataSource(dtMahasiswa);
                crystalReportViewer1.ReportSource = CrystalReport1;
                crystalReportViewer1.Refresh();
            }
            catch (Exception ex)
            {
                //simpanLog(ex.Message);
                MessageBox.Show("Gagal load data: " + ex.Message);
            }
        }

        private void Report2_Load(object sender, EventArgs e)
        {

        }
    }
}
