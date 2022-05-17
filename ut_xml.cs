using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace CUT_M
{
    public class ut_xml
    {
        public static string _err;
        public static string ValueXML(string File, string Value)
        {
            string LaValeur = "";
            XmlTextReader test = null;
            try
            {
                test = new XmlTextReader(File);
                test.WhitespaceHandling = WhitespaceHandling.None;
                while (test.Read())
                {
                    if (test.LocalName == Value)
                    {
                        LaValeur = test.ReadString();
                        test.Read();
                    }
                }
            }
            catch (Exception e)
            {
                _err = e.ToString();
            }
            finally
            {
                if (test != null) test.Close();
            }

            return LaValeur;
        }
    }
}
