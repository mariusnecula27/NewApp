using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace NewApp
{
    public class xmlFile
    {
        public XmlElement nodePrincipal;
        public XmlElement nodePrincipal2;
        public XmlDocument doc;


        public xmlFile()
        {
            doc = new XmlDocument();

            //(1) the xml declaration is recommended, but not mandatory
            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);


            XmlElement element1 = doc.CreateElement(string.Empty, "AUTOSAR", string.Empty);
            doc.AppendChild(element1);

            XmlElement element2 = doc.CreateElement(string.Empty, "AR-PACKAGES", string.Empty);
            element1.AppendChild(element2);

            XmlElement element3 = doc.CreateElement(string.Empty, "AR-PACKAGES", string.Empty);
            element2.AppendChild(element3);

            XmlElement element4 = doc.CreateElement(string.Empty, "SHORT-NAME", string.Empty);
            XmlText text2 = doc.CreateTextNode("RootP_NetworkDesc");
            element4.AppendChild(text2);
            element3.AppendChild(element4);

            XmlElement element5 = doc.CreateElement(string.Empty, "AR-PACKAGES", string.Empty);
            element3.AppendChild(element5);

            nodePrincipal = doc.CreateElement(string.Empty, "AR-PACKAGES", string.Empty);
            element5.AppendChild(nodePrincipal);

            nodePrincipal2 = doc.CreateElement(string.Empty, "AR-PACKAGES", string.Empty);
            element5.AppendChild(nodePrincipal2);

            //doc.Save("D:\\newFile.xml");
        }

        public void WriteSignalNode(List<string> lineElements, XmlElement nodePrincipal, XmlDocument doc)
        {
            XmlElement signal = doc.CreateElement(string.Empty, "I-SIGNAL", string.Empty);
            nodePrincipal.AppendChild(signal);

            XmlElement element1 = doc.CreateElement(string.Empty, "SHORT-NAME", string.Empty);
            XmlText text1 = doc.CreateTextNode(lineElements[0]);
            element1.AppendChild(text1);
            signal.AppendChild(element1);

            XmlElement element2 = doc.CreateElement(string.Empty, "DATA-TYPE-POLICY", string.Empty);
            XmlText text2 = doc.CreateTextNode(lineElements[1]);
            element2.AppendChild(text2);
            signal.AppendChild(element2);

            XmlElement element3 = doc.CreateElement(string.Empty, "LENGHT", string.Empty);
            XmlText text3 = doc.CreateTextNode(lineElements[2]);
            element3.AppendChild(text3);
            signal.AppendChild(element3);

            XmlElement element4 = doc.CreateElement(string.Empty, "IMPLEMENTATION-DATA-TYPE-REF", string.Empty);
            XmlText text4 = doc.CreateTextNode(lineElements[3]);
            element4.AppendChild(text4);
            signal.AppendChild(element4);

            XmlElement element5 = doc.CreateElement(string.Empty, "SYSTEM-SIGNAL-REF", string.Empty);
            XmlText text5 = doc.CreateTextNode(lineElements[4]);
            element5.AppendChild(text5);
            signal.AppendChild(element5);

        }

        public void WriteECU(List<String> lineElements, XmlElement nodeEnter, XmlDocument doc)
        {
            XmlElement ecusignal = doc.CreateElement(string.Empty, "ECUINSTANCES", string.Empty);
            nodeEnter.AppendChild(ecusignal);

            XmlElement element1 = doc.CreateElement(string.Empty, "SHORT-NAME", string.Empty);
            XmlText text1 = doc.CreateTextNode(lineElements[0]);
            element1.AppendChild(text1);
            ecusignal.AppendChild(element1);

            XmlElement element2 = doc.CreateElement(string.Empty, "GW-TIME-BASE", string.Empty);
            XmlText text2 = doc.CreateTextNode(lineElements[1]);
            element2.AppendChild(text2);
            ecusignal.AppendChild(element2);

            XmlElement element3 = doc.CreateElement(string.Empty, "TX-TIME-BASE", string.Empty);
            XmlText text3 = doc.CreateTextNode(lineElements[2]);
            element3.AppendChild(text3);
            ecusignal.AppendChild(element3);

            XmlElement element4 = doc.CreateElement(string.Empty, "RX-TIME-BASE", string.Empty);
            XmlText text4 = doc.CreateTextNode(lineElements[3]);
            element4.AppendChild(text4);
            ecusignal.AppendChild(element4);

            XmlElement element5 = doc.CreateElement(string.Empty, "CYCLIC-TRANSMISSION", string.Empty);
            XmlText text5 = doc.CreateTextNode(lineElements[4]);
            element5.AppendChild(text5);
            ecusignal.AppendChild(element5);

            XmlElement element6 = doc.CreateElement(string.Empty, "SLEEP-MODE", string.Empty);
            XmlText text6 = doc.CreateTextNode(lineElements[5]);
            element6.AppendChild(text6);
            ecusignal.AppendChild(element6);

            XmlElement element7 = doc.CreateElement(string.Empty, "SUPPORTED_WAKE-UP", string.Empty);
            XmlText text7 = doc.CreateTextNode(lineElements[6]);
            element7.AppendChild(text7);
            ecusignal.AppendChild(element7);







        }

    }
}
