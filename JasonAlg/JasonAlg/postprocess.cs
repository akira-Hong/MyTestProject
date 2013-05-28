using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Text;

namespace JasonAlg
{
    public class postprocess
    {

        public static String ProcessPost(string parameters)
        {
            string xmlstring = "";
            int retry = 3;
            //  do
            // {
            xmlstring = ProcessPostAct(parameters);
            retry--;
            //  } while (customer.ErrorCode != (int)TrackbackSendResult.Success && retry>=0);
            return xmlstring;
        }

        public static  String ProcessPostAct(string parameters)
        {
            HttpWebRequest webRequest = (HttpWebRequest)HttpWebRequest.Create("http://www.computerlangs.com/pleaseallowme.aspx");
            webRequest.Method = "POST";
            webRequest.Timeout = 15000;
            webRequest.ContentType = "application/x-www-form-urlencoded";//"multipart/form-data";
            string resultXmlString = "";
            string message = "";

            try
            {
                UTF8Encoding encoding = new UTF8Encoding();
                byte[] postdata = encoding.GetBytes(parameters);
                webRequest.ContentLength = postdata.Length;
                Stream writer = webRequest.GetRequestStream();
                writer.Write(postdata, 0, postdata.Length);
                writer.Close();
                using (HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse())
                {
                    using (StreamReader streamReader = new StreamReader(webResponse.GetResponseStream(), Encoding.Default, true))
                    {
                        resultXmlString = streamReader.ReadToEnd();
                        //streamReader.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                resultXmlString = "";
                Console.WriteLine("ProcessPost: Error " + message);
            }
            return resultXmlString;
        }
    }
}
