using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;

namespace Bases_Persona
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnIngresar_Click(object sender, EventArgs e)
        {
            Ingresar();
        }
        private void Ingresar()
        {
            SqlConnection cn = new SqlConnection(@"server=localhost\SQLEXPRESS;database=Registros;integrated security=true");
            cn.Open();
            SqlCommand cm = new SqlCommand("SELECT Usuario, Contraseña FROM dbo.DatosUsuario WHERE Usuario = @Usuario AND Contraseña = @Contraseña", cn);
            cm.Parameters.AddWithValue("@Usuario", textBox1.Text);
            cm.Parameters.AddWithValue("@Contraseña", textBox2.Text);
            SqlDataReader dr = cm.ExecuteReader();
            if (dr.Read())
            {
                this.Hide();
                Datos uno = new Datos();
                uno.Show();
            }
        }
    }
}
