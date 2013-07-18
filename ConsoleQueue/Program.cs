using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Messaging;
using System.Web;
using System.Net;
using System.Xml;
using System.Runtime.Serialization.Json;
using Microsoft.Office.Server.Powerpoint.Web.Services;
using System.ServiceProcess;

namespace ConsoleQueue
{
    class Program
    {
        static string _rootUrl = "http://app2013";
        static int _maxTried = 50;
        static int _intervalTried = 5;

        static void ConvertWord(string sourcePath, string targetPath)
        {
            string wopiUrl = string.Format("{0}/oh/wopi/files/@/wFileId?wFileId={1}",
                _rootUrl, HttpUtility.UrlEncode(sourcePath));
            string wordUrl = string.Format("{0}/wv/WordViewer/request.pdf?WOPIsrc={1}&access_token=1&type=printpdf",
                _rootUrl, HttpUtility.UrlEncode(wopiUrl));

            int tryCount = 0;
            while (tryCount < _maxTried)
            {
                Console.WriteLine("Time: " + tryCount);
                HttpWebRequest request = WebRequest.Create(wordUrl) as HttpWebRequest;
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                string ctype = response.ContentType.ToLower();
                Console.WriteLine(ctype);
                if (ctype.Contains("text/xml"))
                {
                    // Xml
                    XmlDocument xmldoc = new XmlDocument();
                    xmldoc.Load(response.GetResponseStream());
                    XmlNode statusNode = xmldoc.SelectSingleNode("/docdata/status");
                    if (statusNode != null && statusNode.InnerText == "InProgress")
                    {
                        Console.WriteLine("InProgress");
                        tryCount++;
                        System.Threading.Thread.Sleep(_intervalTried * 1000);
                    }
                    else
                    {
                        Console.WriteLine(xmldoc.OuterXml);
                        break;
                    }
                }
                else if (ctype.Contains("application/pdf"))
                {
                    // PDF
                    using (FileStream fs = new FileStream(targetPath, FileMode.Create))
                    {
                        response.GetResponseStream().CopyTo(fs);
                    }
                    break;
                }
            }
        }

        static void ConvertPPT(string sourcePath, string targetPath)
        {
            string wopiUrl = string.Format("{0}/oh/wopi/files/@/wFileId?wFileId={1}",
                _rootUrl, HttpUtility.UrlEncode(sourcePath));
            string contentStr = string.Format("{{\"presentationId\":\"WOPIsrc={0}&access_token=1\"}}",
                HttpUtility.UrlEncode(wopiUrl));
            byte[] content = Encoding.UTF8.GetBytes(contentStr);
            string restUrl = string.Format("{0}/p/ppt/view.svc/jsonAnonymous/Print", _rootUrl);
            Console.WriteLine(restUrl);
            Console.WriteLine(contentStr);

            int tryCount = 0;
            while (tryCount < _maxTried)
            {
                Console.WriteLine("Time: " + tryCount);
                HttpWebRequest req = WebRequest.Create(restUrl) as HttpWebRequest;
                req.Method = "POST";
                req.ContentType = "application/json; charset=utf-8";

                req.ContentLength = content.Length;
                var reqStream = req.GetRequestStream();
                reqStream.Write(content, 0, content.Length);
                reqStream.Close();

                // Response
                var resStream = req.GetResponse().GetResponseStream();
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(ServiceResult),
                    new List<Type> { typeof(ServiceError), typeof(PptViewingService.PrintResult) });
                var res = ser.ReadObject(resStream) as ServiceResult;
                if (res.Result == null)
                {
                    // InProgress
                    if ((int)res.Error.Code == 59)
                    {
                        Console.WriteLine("InProgress");
                        tryCount++;
                        System.Threading.Thread.Sleep(_intervalTried * 1000);
                    }
                    else
                    {
                        Console.WriteLine(res.Error.Message);
                        break;
                    }
                }
                else
                {
                    // Done
                    string downUrl = (res.Result as PptViewingService.PrintResult).PrintUrl;
                    downUrl = string.Format("{0}/p{1}", _rootUrl, downUrl.TrimStart('.'));
                    WebClient wc = new WebClient();
                    wc.DownloadFile(downUrl, targetPath);
                    break;
                }
            }
        }

        static void Main(string[] args)
        {
            //ConvertWord(@"\\app2013\docs\d.doc", @"\\app2013\ConvertResult\d.pdf");
            ConvertPPT(@"\\app2013\docs\y.ppt", @"\\app2013\ConvertResult\y.pdf");
            //string s = "http%3A%2F%2Fapp2013%2Foh%2Fwopi%2Ffiles%2F%40%2FwFileId%3FwFileId%3D%255C%255Capp2013%255Cdocs%255Cc%252Edoc";
            //Console.WriteLine(System.Web.HttpUtility.UrlDecode(s));
            return;
            MessageQueue queue = new MessageQueue(@".\private$\ConversionJob");
            //string xml = @"<Job Source='\\xxxx\123\456\abc\中文\哈哈哈哈.docx' Target='xxxxxxxxxxxx' />";
            //queue.Send(xml, "title abc");

            queue.MessageReadPropertyFilter.ArrivedTime = true;
            //while (true)
            //{
                try
                {
                    var msg = queue.Peek(new TimeSpan(0, 0, 5));
                    msg.Formatter = new XmlMessageFormatter(new Type[] { typeof(string) });
                    Console.WriteLine(msg.ArrivedTime);
                    Console.WriteLine(msg.Body);
                    XmlDocument xmldoc = new XmlDocument();
                    xmldoc.LoadXml(Convert.ToString(msg.Body));
                    Console.WriteLine(xmldoc.DocumentElement.Attributes["Source"].Value);
                }
                catch (MessageQueueException ex)
                {
                    if (ex.MessageQueueErrorCode == MessageQueueErrorCode.IOTimeout)
                    {
                        Console.WriteLine("Timeout.");
                    }
                }
            //}

            //string url = @"c:\docs\a.docx";
            //string url = @"\\app2013\docs\a.docx";
            //FileInfo fi = new FileInfo(url);
            //Console.WriteLine(fi.Length);
            //byte[] contents = File.ReadAllBytes(url);
            //Console.WriteLine(contents.Length);
        }
    }
}
