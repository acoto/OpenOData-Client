using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Net;
using Microsoft.Office.Tools;

namespace OpenODATA
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
          /*  var url = Interaction.InputBox("Conectar al servidor", "OpenOData", "url");
            //UrlP = url;
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


                    foreach (XmlNode node in xmlDoc.GetElementsByTagName("EntitySet"))
                    {
                        Console.WriteLine(node.Attributes["Name"].Value);
                    }


                    


               }
                catch
                {
                    Interaction.MsgBox("Error al conectar");
                }
            }*/

            if (Globals.ThisAddIn.TaskPane.Visible){
                Globals.ThisAddIn.TaskPane.Visible = false;
            }
            else
            {
                Globals.ThisAddIn.TaskPane.Visible = true;
            }






        }
    }
}
