using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.IO;

namespace CRUDMahasiswaADO
{
    internal class DAL
    {
        // Fungsi untuk mengambil IP Address lokal secara dinamis (Keperluan Deployment)
        public static string GetLocalIPAddress()
        {
            string localIP = string.Empty;
            try
            {
                var host = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        localIP = ip.ToString();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error mengambil IP lokal: " + ex.Message);
            }
            return localIP;
        }

        // Fungsi Connection String yang fleksibel menggunakan IP dinamis
        public static string GetConnectionString()
        {
            // Sesuaikan User ID (sa) dan Password SQL Server kamu masing-masing
            return "Data Source=NJOELL18\\MRZULFAUZI;Initial Catalog=DBAkademikADO;User ID=sa;Password=Mrzxxx18";
        }

        // Inisialisasi objek SQL
        private SqlConnection conn = new SqlConnection(GetConnectionString());
        private SqlDataAdapter da;
        private DataTable dtMahasiswa;

        // 1. Method Menghitung Jumlah Mahasiswa (Output Parameter)
        public int CountMhs()
        {
            // Menggunakan 'using' menjamin koneksi otomatis tertutup meskipun error
            using (SqlConnection conn = new SqlConnection(GetConnectionString()))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("sp_CountMahasiswa", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlParameter outputParam = new SqlParameter("@pCount", SqlDbType.Int) { Direction = ParameterDirection.Output };
                    cmd.Parameters.Add(outputParam);
                    cmd.ExecuteNonQuery();
                    return Convert.ToInt32(outputParam.Value);
                }
            }
        }

        // 2. Method Mengambil Semua Data Mahasiswa (Termasuk BLOB Foto)
        public DataTable GetMhs()
        {
            if (conn.State == ConnectionState.Closed) conn.Open();
            SqlCommand cmd = new SqlCommand("sp_GetMahasiswa", conn);
            cmd.CommandType = CommandType.StoredProcedure;

            da = new SqlDataAdapter(cmd);
            dtMahasiswa = new DataTable();
            da.Fill(dtMahasiswa);
            conn.Close();
            return dtMahasiswa;
        }

        // 3. Method Insert Data Mahasiswa dengan Fitur Transaksi & BLOB Foto
        public void InsertMhs(string nim, string nama, string alamat, string jeniskelamin, DateTime tanggallahir, string kodeProdi, byte[] foto)
        {
            if (conn.State == ConnectionState.Closed) conn.Open();
            SqlTransaction trans = conn.BeginTransaction();
            try
            {
                SqlCommand command = new SqlCommand("sp_InsertMahasiswa", conn, trans);
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@PNIM", nim);
                command.Parameters.AddWithValue("@pNama", nama);
                command.Parameters.AddWithValue("@pAlamat", alamat);
                command.Parameters.AddWithValue("@pJenisKelamin", jeniskelamin);
                command.Parameters.AddWithValue("@pTanggalLahir", tanggallahir);
                command.Parameters.AddWithValue("@pKodeProdi", kodeProdi);
                command.Parameters.AddWithValue("@pFoto", (object)foto ?? DBNull.Value); // Menangani jika foto kosong

                command.ExecuteNonQuery();
                trans.Commit();
            }
            catch (Exception)
            {
                trans.Rollback();
                throw;
            }
            finally
            {
                conn.Close();
            }
        }

        // 4. Method Update Data Mahasiswa (Termasuk Update BLOB Foto)
        public void UpdateMhs(string nim, string nama, string alamat, string jeniskelamin, DateTime tanggallahir, string kodeProdi, byte[] foto)
        {
            if (conn.State == ConnectionState.Closed) conn.Open();
            try
            {
                SqlCommand command = new SqlCommand("sp_UpdateMahasiswa", conn);
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@PNIM", nim);
                command.Parameters.AddWithValue("@pNama", nama);
                command.Parameters.AddWithValue("@pAlamat", alamat);
                command.Parameters.AddWithValue("@pJenisKelamin", jeniskelamin);
                command.Parameters.AddWithValue("@pTanggalLahir", tanggallahir);
                command.Parameters.AddWithValue("@pKodeProdi", kodeProdi);
                command.Parameters.AddWithValue("@pFoto", (object)foto ?? DBNull.Value);

                command.ExecuteNonQuery();
            }
            finally
            {
                conn.Close();
            }
        }

        // 5. Method Delete Data Mahasiswa
        public void DeleteMhs(string nim)
        {
            if (conn.State == ConnectionState.Closed) conn.Open();
            try
            {
                SqlCommand cmd = new SqlCommand("sp_DeleteMahasiswa", conn);
                cmd.CommandType = CommandType.StoredProcedure;

                // wajib tambahkan parameter @NIM
                cmd.Parameters.AddWithValue("@NIM", nim);

                cmd.ExecuteNonQuery();
            }
            finally
            {
                conn.Close();
            }
            conn.Close();
        }
    


        // 6. Method Ambil Data untuk Chart Dashboard (Semua Data)
        public DataTable getAllDataChart()
        {
            if (conn.State == ConnectionState.Closed) conn.Open();
            SqlCommand cmd = new SqlCommand("sp_DashBoard", conn);
            cmd.CommandType = CommandType.StoredProcedure;

            da = new SqlDataAdapter(cmd);
            dtMahasiswa = new DataTable();
            da.Fill(dtMahasiswa);
            conn.Close();
            return dtMahasiswa;
        }

        // 7. Method Ambil Data untuk Chart Dashboard Filter per Tahun
        public DataTable getDataChartByTahun(DateTime thMasuk)
        {
            if (conn.State == ConnectionState.Closed) conn.Open();
            SqlCommand cmd = new SqlCommand("sp_DashBoardByTahun", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@inTglMsuk", thMasuk.Year.ToString());

            da = new SqlDataAdapter(cmd);
            dtMahasiswa = new DataTable();
            da.Fill(dtMahasiswa);
            conn.Close();
            return dtMahasiswa;
        }

        // 8. Method Log Message
        public void InsertLog(string message)
        {
            if (conn.State == ConnectionState.Closed) conn.Open();
            try
            {
                SqlCommand cmd = new SqlCommand("sp_LogMessage", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@psn", message);
                cmd.ExecuteNonQuery();
            }
            finally
            {
                conn.Close();
            }
        }
        public void resetData()
        {
            if (conn.State == ConnectionState.Closed) conn.Open();
            try
            {
                SqlCommand cmd = new SqlCommand("sp_ResetMahasiswa", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.ExecuteNonQuery();
            }
            finally
            {
                conn.Close();
            }
        }

        public void testInject(string nim)
        {
            if (conn.State == ConnectionState.Closed) conn.Open();
            try
            {
                // contoh raw query untuk simulasi injection
                string query = $"UPDATE Mahasiswa SET Nama = 'Injected' WHERE NIM = '{nim}'";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
            }
            catch
            {
                throw new Exception("Unsafe UPDATE operation not allowed");
            }
            finally { conn.Close(); }
        }

        public DataTable GetMahasiswaByNIM(string nim)
        {
            using (SqlConnection conn = new SqlConnection(GetConnectionString()))
            {
                SqlCommand cmd = new SqlCommand("sp_GetMahasiswaByNIM", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pNIM", nim); // sesuaikan dengan nama di SP

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                return dt;
            }
        }



        public DataTable getDataRekap(string prodi, DateTime tglMasuk)
        {
            using (SqlConnection conn =
                new SqlConnection(GetConnectionString()))
            {
                conn.Open();

                SqlCommand cmd =
                    new SqlCommand("sp_Report", conn);

                cmd.CommandType =
                    CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@Prodi", prodi);
                cmd.Parameters.AddWithValue("@TglMasuk", tglMasuk);

                SqlDataAdapter da =
                    new SqlDataAdapter(cmd);

                DataTable dt =
                    new DataTable();

                da.Fill(dt);

                return dt;
            }
        }

    }
}