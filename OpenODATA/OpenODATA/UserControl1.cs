using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Xml;
using System.IO;
using System.Net;
using Microsoft.Office.Tools;

using Microsoft.Office.Tools.Ribbon;
using Microsoft.Data.Edm;
using System.Net.Http;
using Microsoft.Data.OData;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;



namespace OpenODATA
{
    public partial class UserControl1 : UserControl
    {

        ODataHelper odatahelper = new ODataHelper();
        ODataHelper.TableInformation[] tablas;
        int endRow;
        int endCol;
        string[][] data;
        List<List<string>> listaDeDatosCrudos = new List<List<string>>();



        public UserControl1()
        {
            InitializeComponent();
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {
            string path = @"data.txt";
            if (File.Exists(path))
            {
                String[] text = System.IO.File.ReadAllLines(path);
                loadCombo(text);
            }
                


        }
        public void loadCombo(String[] text)
        {
            //Interaction.MsgBox("csacas");
            comboBox2.Items.Clear();
            foreach (var row in text)
            {
                // Interaction.MsgBox(row);
                String[] row2 = row.Split(new String[] { "_&"}, StringSplitOptions.None);
                comboBox2.Items.Add(row2[0]);
                
            }
            comboBox2.Refresh();


        }

        private void button3_Click(object sender, EventArgs e)
        {
            var url = Interaction.InputBox("Conectar al servidor", "OpenOData", "url");
            //UrlP = url;
            //http://services.odata.org/V4/OData/OData.svc/"
            //http://services.odata.org/V4/Northwind/Northwind.svc/
            if (url == "url" || url == "" || url == " ")
            {
                Interaction.MsgBox("Direccion invalida");
            }
            else
            {
            try
                {


                    var request = WebRequest.CreateHttp(url + "$metadata");
                    WebResponse response = request.GetResponse();
                    Stream datastream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(datastream);

                    string responseFromServer = reader.ReadToEnd();

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(new StringReader(responseFromServer));
                    reader.Close();
                    response.Close();


               XmlNodeList node = xmlDoc.GetElementsByTagName("Schema");
                    string id = node[0].InnerText;
                    
                    string path = @"data.txt";

                if (!File.Exists(path))
                {
                    var f = File.Create(path);
                    f.Close();
                }

                    
                     String[] text = System.IO.File.ReadAllLines(path);


                        if (!text.Contains(url)){
                            TextWriter tw = new StreamWriter(path, true);
                            tw.WriteLine(url + "_&" + id);
                            tw.Close();
                        
                        }
                    
                        loadCombo(System.IO.File.ReadAllLines(path));
                    Interaction.MsgBox("Conexion exitosa");



                }
                catch
                {
                    Interaction.MsgBox("Error al conectar");
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Interaction.MsgBox(comboBox2.SelectedItem);

            odatahelper.metadata = null;
            odatahelper.tables = null;

            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            activeSheet.Cells.ClearContents();

            comboBox1.Text = "";

            string url = comboBox2.SelectedItem.ToString();

            odatahelper.odataEndpointURL = new Uri(url);
            odatahelper.odataMetadataURL = new Uri(url + "$metadata");
            

            odatahelper.GetTables();

            MessageBox.Show("Conectado...");


            tablas = odatahelper.tables;
            comboBox1.Items.Clear();
            foreach (ODataHelper.TableInformation tabla in tablas)
            {
                //RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                string item = tabla.tableName;
                comboBox1.Items.Add(item);
            }
        }

        public List<ODataHelper.Change> buscarCambios(List<List<string>> listaDeDatosActualizados, List<List<string>> listaDeDatosCrudos)
        {

            int index = saberLlaveDeTabla();
            List<ODataHelper.Change> listaDeCambios = new List<ODataHelper.Change>();

            foreach (List<string> dato_actualizado in listaDeDatosActualizados)
            {
                //MessageBox.Show(dato_actualizado[index]);
                List<List<string>> listaExistentes = listaDeDatosCrudos.FindAll(x => x.ElementAt(index) == dato_actualizado[index]);
                if (listaExistentes.Count > 0)
                {
                    foreach (List<string> trabajando in listaExistentes)
                    {
                        if (trabajando.SequenceEqual(dato_actualizado))
                        {
                            //MessageBox.Show("No cambio: " + dato_actualizado[1]);
                        }
                        else
                        {
                            //MessageBox.Show("Cambio: " + dato_actualizado[1]);
                            listaDeCambios.Add(creacionDeCambio("UPDATE", dato_actualizado[index], dato_actualizado.ToArray()));
                        }

                        listaDeDatosCrudos.Remove(trabajando);
                    }
                }
                else
                {
                    //MessageBox.Show("Insertar: "+ dato_actualizado[1]);
                    listaDeCambios.Add(creacionDeCambio("INSERT", dato_actualizado[index], dato_actualizado.ToArray()));
                }
            }

            foreach (List<string> eliminar in listaDeDatosCrudos)
            {
                //MessageBox.Show("Eliminar: " + eliminar[1]);
                listaDeCambios.Add(creacionDeCambio("DELETE", eliminar[index], eliminar.ToArray()));
            }

            return listaDeCambios;
        }

        public ODataHelper.Change creacionDeCambio(string operacion, string id, string[] data)
        {
            ODataHelper.Change change = new ODataHelper.Change();
            change.operation = operacion;
            change.id = id;
            change.data = data;

            return change;
        }
        public int saberLlaveDeTabla()
        {
            string[] header = odatahelper.GetHeaders(comboBox1.Text);
            string saberKey = "";
            foreach (ODataHelper.TableInformation tabla in tablas)
            {
                if (tabla.tableName == comboBox1.Text)
                {
                    saberKey = tabla.key;
                }
            }

            int index = Array.IndexOf(header, saberKey);

            return index;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listaDeDatosCrudos.Clear();
            data = odatahelper.ReadDataFromODataWithHeaders(comboBox1.Text, "");

            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            activeSheet.Cells.ClearContents();

            endRow = data.GetLength(0);
            endCol = data[0].GetLength(0);

            for (int i = 1; i <= endRow; i++)
            {
                List<string> datos_crudos = new List<string>();
                for (int y = 1; y <= endCol; y++)
                {
                    var data_ = data[i - 1][y - 1];
                    
                    ((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[i, y]).Value2 = data_;
                   
                    datos_crudos.Add(data_);
                }
                if (i != 1)
                {
                    listaDeDatosCrudos.Add(datos_crudos);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);

            int row = 2;
            int col = 1;

            List<List<string>> listaDeDatosActualizados = new List<List<string>>();
            List<string> llavesPrimarias = new List<string>();
            int index = saberLlaveDeTabla();

            while (((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row, col]).Value2 != null)
            {
                List<string> datos_act = new List<string>();

                for (int y = 1; y <= endCol; y++)
                {
                    var data_ = ((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row, col]).Value2;
                    if (data_ is double)
                    {
                        data_ = ((double)((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row, col]).Value2).ToString();
                    }

                    if (data_ == null)
                    {
                        datos_act.Add("");
                    }
                    else
                    {
                        datos_act.Add(data_);
                    }

                    if (index == y - 1)
                    {
                        llavesPrimarias.Add(data_);
                    }
                    col++;
                }
                row++;
                col = 1;
                listaDeDatosActualizados.Add(datos_act);

            }

            if (llavesPrimarias.Count.ToString() == llavesPrimarias.Distinct().Count().ToString())
            {
                //buscarCambios(listaDeDatosActualizados, listaDeDatosCrudos);
                odatahelper.UpdateDataToOData(buscarCambios(listaDeDatosActualizados, listaDeDatosCrudos).ToArray(), comboBox1.Text, odatahelper.GetHeaders(comboBox1.Text));

                //MessageBox.Show("Tenemos el string actualizado...");
                //button2_Click(null, null);
            }
            else
            {
                MessageBox.Show("No se puede actualizar por que hay llaves repetidas.");
            }

        }
    }
}
