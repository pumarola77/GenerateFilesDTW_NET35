using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using Npgsql;


namespace GenerateFilesDTW_NET35
{
    public partial class WindowMain : Form
    {

        private String pathFile;
        private StreamWriter sw;
        private Boolean lb_connectSAP = false;
        private Boolean lb_connectXSINFO = false;

        public WindowMain()
        {
            InitializeComponent();
        }

        private Boolean TryConnectSAP(int option)
        {
            try
            {

                // Realitzar la connexio si esta informat el servidor i la base de dades

                if ((!String.IsNullOrEmpty(this.serverSAP.Text)) &&
                     (!String.IsNullOrEmpty(this.nameBDSAP.Text)))
                {

                    SqlConnection connection = new SqlConnection
                    {
                        ConnectionString = "SERVER=" + this.serverSAP.Text + "; Initial Catalog =" + this.nameBDSAP.Text + "; Integrated Security=true"
                    };
                    connection.Open();

                    if (option == 1)
                    {
                        textBox1.AppendText("Connection is established (SAPB1)");
                        textBox1.AppendText(Environment.NewLine);
                    }

                    connection.Close();
                    return true;
                }
                else
                {
                    textBox1.AppendText("Connection is not established (SAPB1)");
                    textBox1.AppendText(Environment.NewLine);

                    return false;
                }

            }
            catch (Exception ex)
            {
                textBox1.AppendText("Can not open connection" + ex.Message);
                textBox1.AppendText(Environment.NewLine);

                return false;
            }
        }

        private Boolean TryConnectXSINFO(int option)
        {
            try
            {
                if ((!String.IsNullOrEmpty(this.serverXSNIFO.Text)) &&
                    (!String.IsNullOrEmpty(this.nameBDXSINFO.Text)) &&
                    (!String.IsNullOrEmpty(this.textPort.Text)) &&
                    (!String.IsNullOrEmpty(this.textUserID.Text)) &&
                    (!String.IsNullOrEmpty(this.textPassword.Text))
                    )
                {
                    String SqlConnection = "Server=" + this.serverXSNIFO.Text + ";Port=" + this.textPort.Text + ";User id=" + this.textUserID.Text + ";Password=" + this.textPassword.Text + ";DataBase=" + this.nameBDXSINFO.Text;
                    NpgsqlConnection connection = new NpgsqlConnection(SqlConnection);

                    connection.Open();

                    if (option == 1)
                    {
                        textBox1.AppendText("Connection is Established XSINFO");
                        textBox1.AppendText(Environment.NewLine);
                    }

                    connection.Close();
                    return true;
                }
                else
                {
                    textBox1.AppendText("Connection is not established (XSINFO)");
                    textBox1.AppendText(Environment.NewLine);

                    return false;
                }
            }
            catch (Exception ex)
            {
                textBox1.AppendText("Can not open connection" + ex.Message);
                textBox1.AppendText(Environment.NewLine);
                return false;
            }
        }

        private void GenerateFiles_Click(object sender, EventArgs e)
        {
            BtCancel.Enabled = false;
            GenerateFiles.Enabled = false;

            // Provar les connexions de la base de dades
            if (!String.IsNullOrEmpty(this.serverSAP.Text))
            {
                lb_connectSAP = TryConnectSAP(0);
            }

            if (!String.IsNullOrEmpty(this.serverXSNIFO.Text))
            {
                lb_connectXSINFO = TryConnectXSINFO(0);
            }

            // Llegeix path on estan ubicats els fitxers xml.
            String pathXML = PathSQL.Text;
            String FileXml;
            String sPath;

            if ((!String.IsNullOrEmpty(PathSQL.Text)) &&
                 (!String.IsNullOrEmpty(PathTXT.Text)) &&
                 (Directory.Exists(PathSQL.Text)) &&
                 (Directory.Exists(PathTXT.Text)) &&
                 ((lb_connectSAP) || (lb_connectXSINFO))
               )
            {

                // Comprova si hi han fitxers xml 
                // al directori arrel
                DirectoryInfo di = new DirectoryInfo(pathXML);
                foreach (var fi in di.GetFiles("*.xml"))
                {
                    sPath = pathXML + "//" + fi.ToString();
                    CreateReaderFiles(sPath);
                }

                // Comprova si hi han carpetes i comprova existencia 
                // de fitxers xml

                DirectoryInfo di2;
                String[] dirs = Directory.GetDirectories(pathXML, "*", SearchOption.AllDirectories);
                foreach (string dir in dirs)
                {
                    // Lectura de cada directori buscant fitxers xml
                    di2 = new DirectoryInfo(dir.ToString());

                    foreach (var fi in di2.GetFiles("*.xml"))
                    {
                        FileXml = fi.ToString();
                        sPath = dir.ToString() + "//" + FileXml;

                        textBox1.AppendText("Reader File " + fi.ToString());
                        textBox1.AppendText(Environment.NewLine);

                        CreateReaderFiles(sPath);
                    }
                }

                BtCancel.Enabled = true;
                GenerateFiles.Enabled = true;

                textBox1.AppendText("The process has finished. Check files.");
                textBox1.AppendText(Environment.NewLine);
            }
            else
            {

                if (!Directory.Exists(PathSQL.Text))
                {
                    textBox1.AppendText(PathSQL.Text + " Is not valid a directory");
                    textBox1.AppendText(Environment.NewLine);
                }

                if (!Directory.Exists(PathTXT.Text))
                {
                    textBox1.AppendText(PathTXT.Text + " Is not valid a directory");
                    textBox1.AppendText(Environment.NewLine);
                }

                if (lb_connectSAP == false || lb_connectXSINFO == false)
                {
                    textBox1.AppendText("Connection is not established");
                    textBox1.AppendText(Environment.NewLine);
                }

                BtCancel.Enabled = true;
                GenerateFiles.Enabled = true;
            }
        }


        private void CreateReaderFiles(String sPath)
        {
            XmlTextReader reader = new XmlTextReader(sPath);
            String tag = null;
            Int32 contador = 0;

            try
            {
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:

                            if (reader.Name.Equals("table"))
                            {
                                if (reader.AttributeCount == 3)
                                {

                                    textBox1.AppendText("Generate File Table " + reader.GetAttribute(0));
                                    textBox1.AppendText(Environment.NewLine);
                                    CreateFileFolder(reader.GetAttribute(0), reader.GetAttribute(1), reader.GetAttribute(2));
                                }
                                else
                                {
                                    throw new InvalidOperationException("Incorrect Values in Tag <Table>");
                                }
                            }

                            if (reader.Name.Equals("first-row"))
                            {
                                tag = reader.Name;
                                contador = 1;
                            }

                            if (reader.Name.Equals("second-row"))
                            {
                                tag = reader.Name;
                                contador = 1;
                            }

                            if (reader.Name.Equals("sql"))
                            {
                                tag = reader.Name;                                
                            }

                            break;

                        case XmlNodeType.Text:
                            if ((tag == "first-row") || (tag == "second-row"))
                            {
                                if (contador > 1)
                                {
                                    sw.Write('\t');
                                    contador = contador + 1;
                                }
                                else
                                {
                                    contador = contador + 1;
                                }
                                                                    
                                sw.Write(reader.Value);
                                //sw.Write('\t');                               
                            }

                            if (tag == "sql")
                            {
                                if (lb_connectSAP == true)
                                {
                                    ExecuteSQLLoadFileSQLServer(reader.Value);
                                }

                                if (lb_connectXSINFO == true)
                                {
                                    ExecuteSQLLoadFilePostgres(reader.Value);
                                }
                            }

                            break;

                        case XmlNodeType.EndElement:
                            if (reader.Name.Equals("first-row") || reader.Name.Equals("second-row"))
                            {
                                sw.Write('\r');
                            }
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                textBox1.AppendText("Message : " + e.Message);
                textBox1.AppendText(Environment.NewLine);
            }
            finally
            {
                if (reader != null)
                {
                    // Close read XML
                    reader.Close();
                    // Close File
                    sw.Close();

                }
            }
        }

        private void ConnectSAP_Click(object sender, EventArgs e)
        {
            try
            {
                TryConnectSAP(1);
            }
            catch (Exception ex)
            {
                textBox1.AppendText("Can not open connection " + ex.Message);
                textBox1.AppendText(Environment.NewLine);
            }
        }

        private void ConnectXSINFO_Click(object sender, EventArgs e)
        {
            try
            {
                TryConnectXSINFO(1);
            }
            catch (Exception ex)
            {
                textBox1.AppendText("Can not open connection " + ex.Message);
                textBox1.AppendText(Environment.NewLine);
            }
        }

        private void BtCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void FindPath_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                PathSQL.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void DestPath_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                PathTXT.Text = folderBrowserDialog1.SelectedPath;
            }
        }


        /**
        * Create File and Folder
        **/
        private void CreateFileFolder(String nameTable, String FileTable, String FolderTable)
        {
            String path;

            if (Directory.Exists(PathTXT.Text))
            {
                path = PathTXT.Text + "//" + FolderTable;
                // Create Directory
                if (!Directory.Exists(path))
                {
                    DirectoryInfo di = Directory.CreateDirectory(path);
                }

                // Create File                
                pathFile = PathTXT.Text + "//" + FolderTable + "//" + FileTable;
                if (File.Exists(pathFile))
                {
                    File.Delete(pathFile);
                }

                sw = File.CreateText(pathFile);
            }
        }

        private void ExecuteSQLLoadFilePostgres(String sql)
        {
            String SqlConnection = "Server=" + this.serverXSNIFO.Text + ";Port=" + this.textPort.Text + ";User id=" + this.textUserID.Text + ";Password=" + this.textPassword.Text + ";DataBase=" + this.nameBDXSINFO.Text + ";Pooling=false;CommandTimeout=100";
            NpgsqlConnection connection = new NpgsqlConnection(SqlConnection);
            connection.Open();

            NpgsqlCommand cmd = new NpgsqlCommand(sql, connection);
            NpgsqlDataReader dataReader = cmd.ExecuteReader();

            while (dataReader.HasRows)
            {
                while (dataReader.Read())
                {
                    Object[] values = new Object[dataReader.FieldCount];
                    int fieldcount = dataReader.GetValues(values);

                    for (int i = 0; i < fieldcount; i++)
                    {
                        if ( i > 0 )
                            sw.Write('\t');

                        sw.Write(dataReader[i]);
                        //sw.Write('\t');
                    }
                    sw.Write('\r');
                }
                dataReader.NextResult();
            }
        }


        /**
        * Reader SQL and Load the registers
        * SQL SERVER         
        **/
        private void ExecuteSQLLoadFileSQLServer(String sql)
        {
            SqlConnection connection = new SqlConnection
            {
                ConnectionString = "SERVER=" + this.serverXSNIFO.Text + "; Initial Catalog =" + this.nameBDXSINFO.Text + "; Integrated Security=true"
            };
            connection.Open();
            SqlCommand cmd = new SqlCommand(sql, connection);
            SqlDataReader reader = cmd.ExecuteReader();

            while (reader.HasRows)
            {
                while (reader.Read())
                {
                    Object[] values = new Object[reader.FieldCount];
                    int fieldcount = reader.GetValues(values);

                    for (int i = 0; i < fieldcount; i++)
                    {
                        if (i > 0)
                            sw.Write('\t');

                        sw.Write(reader[i]);
                        //sw.Write('\t');
                    }
                    sw.Write('\r');
                }
                reader.NextResult();
            }
        }
    }
}
