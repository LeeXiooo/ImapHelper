using ImapHelper.Entity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Security;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;

namespace ImapHelper
{
    public class ImapClient : IDisposable
    {
        TcpClient tcpclient = new TcpClient();
        SslStream sslstream;
        StreamWriter sw;
        StreamReader reader;

        #region public

        /// <summary>
        /// 连接
        /// </summary>
        /// <returns></returns>
        public void Connection(string serveraddress, int port)
        {
            try
            {
                tcpclient.Connect(serveraddress, port);
                sslstream = new SslStream(tcpclient.GetStream());
                sslstream.AuthenticateAsClient(serveraddress);
                bool flag = sslstream.IsAuthenticated;
                if (!flag)
                {
                    throw new Exception("sslstream IsAuthenticated return false");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 登录
        /// </summary>
        /// <returns></returns>
        public void login(string username, string password)
        {
            sw = new StreamWriter(sslstream);
            // Assigned reader to stream
            reader = new StreamReader(sslstream);
            sw.WriteLine("a01 LOGIN " + username + " " + password);
            sw.Flush();
            try
            {
                string strTemp = string.Empty;
                while ((strTemp = reader.ReadLine()) != null)
                {
                    if (strTemp.IndexOf("OK LOGIN completed", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        break;
                    }
                    else if (strTemp.IndexOf("NO LOGIN failed", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        throw new Exception("NO LOGIN failed");
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 查看文件夹
        /// </summary>
        /// <returns></returns>
        public string FoldersList()
        {
            sw.WriteLine("a02 list \"\" *");
            sw.Flush();
            string str = string.Empty;
            try
            {
                string strTemp = string.Empty;
                while ((strTemp = reader.ReadLine()) != null)
                {
                    if (strTemp.IndexOf("OK LIST completed", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        break;
                    }
                    str += strTemp + "\r";
                }
            }
            catch (Exception ex)
            {
                return "error:" + ex.Message;
            }
            return str;
        }

        /// <summary>
        /// 选择文件夹
        /// </summary>
        /// <param name="folder"></param>
        public bool SelectFolder(string folder)
        {
            sw.WriteLine("a03 select " + folder);
            sw.Flush();
            try
            {
                string strTemp = string.Empty;
                while ((strTemp = reader.ReadLine()) != null)
                {
                    if (strTemp.IndexOf("NO SELECT failed", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return false;
                    }
                    if (strTemp.IndexOf("SELECT completed", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            return false;
        }

        /// <summary>
        /// 获取邮件id
        /// </summary>
        /// <param name="dateTime"></param>
        public string[] Search(DateTime dateTime, bool isReverse = true)
        {
            string strDateTime = ToSearchDt(dateTime);
            sw.WriteLine("a04 SEARCH SINCE " + strDateTime);
            sw.Flush();
            string str = string.Empty;
            try
            {
                string strTemp = string.Empty;
                while ((strTemp = reader.ReadLine()) != null)
                {
                    if (strTemp.IndexOf("SEARCH", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        str = Regex.Replace(strTemp, "SEARCH", "", RegexOptions.IgnoreCase)
                            .Replace("*", "").Trim();
                    }
                    break;
                }
            }
            catch (Exception ex)
            {
                ;
            }
            if (string.IsNullOrEmpty(str)) return null;
            string[] arr = Regex.Split(str, @"\s+");
            if (isReverse)
            {
                arr = arr.Reverse().ToArray();
            }
            return arr;
        }

        /// <summary>
        /// 获取邮件头
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public Header GetHeader(int ID)
        {
            var strHeader = fetch(ID, "body[header]");
            var header = HeaderParse(strHeader);
            return header;
        }

        /// <summary>
        /// 获取邮件主体
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public Body GetBody(int ID)
        {
            var strBody = fetch(ID, "body[1]");
            List<string> list = new List<string>(strBody.Split(new string[] { "\r" }, StringSplitOptions.RemoveEmptyEntries));
            var contentTypeList = list.Where(d => d.StartsWith("Content-Type:", StringComparison.OrdinalIgnoreCase)).ToList();
            if (contentTypeList.Count() <= 1)
            {
                var body = BodyParse(strBody);
                return body;
            }
            else
            {
                strBody = fetch(ID, "body[1.1]");
                var body = BodyParse(strBody);
                return body;
            }
        }

        /// <summary>
        /// 获取附件
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public List<Attachment> GetAttachments(int ID)
        {
            var body = fetch(ID, "BODY");
            List<string> bodyList = new List<string>(body.Split(new string[] { ")" }, StringSplitOptions.RemoveEmptyEntries));
            bodyList = bodyList.Where(q => q.Contains("name")).ToList();
            int index = 2;
            List<Attachment> list = new List<Attachment>();
            foreach (var entity in bodyList)
            {
                var name = entity.Split('}')[1].Replace("\r", "");
                var attachment = fetch(ID, "body[" + index + "]", true);
                list.Add(new Attachment
                {
                    Name = name,
                    bs64 = attachment,
                });
                index++;
            }
            //while (true)
            //{
            //    var attachment = fetch(ID, "body[" + index + "]", true);
            //    if (string.IsNullOrEmpty(attachment)) break;
            //    try
            //    {
            //        byte[] bytes = Convert.FromBase64String(attachment);
            //    }
            //    catch (Exception ex)
            //    {
            //        index++;
            //        continue;
            //    }
            //    list.Add(new Attachment
            //    {
            //        bs64 = attachment,
            //    });
            //    index++;
            //}
            return list;
        }

        #endregion public

        #region private

        /// <summary>
        /// 获取邮件
        /// </summary>
        /// <param name="id"></param>
        /// <param name="data"></param>
        private string fetch(int id, string data, bool line = false)
        {
            //fetch 2918 body[header]
            sw.WriteLine("b01 fetch " + id + " " + data);
            sw.Flush();
            string str = string.Empty;
            try
            {
                Regex reg1 = new Regex(@"body\[[1-9][0-9]{0,1}\]");
                bool ismatch = reg1.IsMatch(data);
                if (ismatch && data != "body[1]")
                {
                    //100k
                    var clen = 1024 * 100;
                    var read = new Char[clen];
                    var count = reader.Read(read, 0, clen);
                    while (count > 0)
                    {
                        var strTemp = new string(read, 0, count);
                        str += strTemp;
                        if (strTemp.IndexOf("OK FETCH completed", StringComparison.OrdinalIgnoreCase) != -1)
                            break;
                        count = reader.Read(read, 0, clen);
                    }
                    var arr = str.Split('\r');
                    string[] b = new string[arr.Length - 3];
                    Array.Copy(arr, 1, b, 0, arr.Length - 3);
                    str = string.Join("\r", b).Replace("\r\n", "").Replace(")", "");
                }
                else
                {
                    string strTemp = string.Empty;
                    while ((strTemp = reader.ReadLine()) != null)
                    {
                        if (strTemp.IndexOf("OK SEARCH completed", StringComparison.OrdinalIgnoreCase) != -1 || strTemp.ToLower().IndexOf(data, StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            continue;
                        }
                        if (strTemp.IndexOf("OK FETCH completed", StringComparison.OrdinalIgnoreCase) != -1 || strTemp.IndexOf("BAD invalid command or parameters", StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            break;
                        }
                        if (line)
                        {
                            str += strTemp;
                        }
                        else
                        {
                            str += strTemp + "\r";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ;
            }
            return str;
        }

        /// <summary>
        /// 解析邮件头
        /// </summary>
        /// <param name="strHeader"></param>
        /// <returns></returns>
        private Header HeaderParse(string strHeader)
        {
            List<string> list = new List<string>(strHeader.Split(new string[] { "\r" }, StringSplitOptions.RemoveEmptyEntries));
            Header hearer = new Header();
            //获取From
            var from = GetFieldValue(list, "From:");
            var fromName = from.Split(' ')[0].Trim();
            hearer.FromName = MineToStr(fromName);
            try
            {
                hearer.FromAddress = from.Split(' ')[1].Replace("<", "").Replace(">", "").Trim();
            }
            catch
            {
                var index = list.FindIndex(q => q.StartsWith("From:", StringComparison.OrdinalIgnoreCase));
                hearer.FromAddress = list[index+1].Replace("<", "").Replace(">", "").Trim();
            }
            //获取To
            var to = GetFieldValue(list, "To:");
            hearer.To = to;
            //MessageID
            var messageID = GetFieldValue(list, "Message-ID:");
            hearer.MessageID = messageID.Replace("<", "").Replace(">", "").Trim();
            //Subject
            var subject = GetFieldValue(list, "Subject:");
            hearer.Subject = MineToStr(subject);
            var subjects = list.Where(q => q.StartsWith(" =?", StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (var sub in subjects)
            {
                hearer.Subject += MineToStr(sub);
            }
            //Date
            var date = GetFieldValue(list, "Date:");
            hearer.Date = DateTime.Parse(date);
            return hearer;
        }

        private string GetFieldValue(List<string> list, string fieldname)
        {
            var fv = list.FirstOrDefault(q => q.StartsWith(fieldname, StringComparison.OrdinalIgnoreCase));
            fv = Regex.Replace(fv, fieldname, "", RegexOptions.IgnoreCase).Trim();
            return fv;
        }

        /// <summary>
        /// 解析邮件主体
        /// </summary>
        /// <param name="strBody"></param>
        /// <returns></returns>
        private Body BodyParse(string strBody)
        {
            List<string> list = new List<string>(strBody.Split(new string[] { "\r" }, StringSplitOptions.RemoveEmptyEntries));
            Body body = new Body();
            //是否有附件
            var hasPart = list.Any(d => d.StartsWith("------=", StringComparison.OrdinalIgnoreCase));
            body.HasPart = hasPart;
            //类型 编码
            var contentTypeList = list.Where(d => d.StartsWith("Content-Type:", StringComparison.OrdinalIgnoreCase)).ToList();
            if (contentTypeList.Count == 1)
            {
                var contentType = contentTypeList.FirstOrDefault();
                if (contentType != null)
                {
                    contentType = Regex.Replace(contentType, "Content-Type:", "", RegexOptions.IgnoreCase).Trim();
                    body.ContentType = contentType.Split(';')[0];
                    body.Charset = contentType.Split(';')[1].Split('=')[1].Replace("\"", "").ToLower().Trim();
                }
            }
            else
            {
                var index = list.FindLastIndex(d => d.StartsWith("------=", StringComparison.OrdinalIgnoreCase));
                //多个取最后一个
                foreach (var type in contentTypeList)
                {
                    var contentType = type.Replace("Content-Type:", "").Trim();
                    body.ContentType = contentType.Split(';')[0];
                    body.Charset = contentType.Split(';')[1].Split('=')[1].Replace("\"","").ToLower().Trim();
                }
            }
            //内容主体
            var contentList = list.Where(d => !d.StartsWith("------=", StringComparison.OrdinalIgnoreCase)
                                       && !d.StartsWith("Content-Type:", StringComparison.OrdinalIgnoreCase)
                                       && !d.StartsWith("Content-Transfer-Encoding:", StringComparison.OrdinalIgnoreCase)
                                       && !d.StartsWith(")", StringComparison.OrdinalIgnoreCase)).ToList();
            var strContent = string.Join("\r", contentList.ToArray()).Replace("\r", "").Replace(")", "");
            if (strContent.IndexOf("<html", StringComparison.OrdinalIgnoreCase) != -1)
            {
                body.Content = strContent;
                body.ContentType = "text/html";
            }
            else
            {
                if (contentTypeList.Count > 1)
                {
                    strContent = strContent.Split('=')[2] + "=";
                }
                byte[] bytes = Convert.FromBase64String(strContent);
                body.Content = Encoding.GetEncoding(body.Charset).GetString(bytes);
            }
            return body;
        }

        /// <summary>
        /// MIME 编码 转 字符串
        /// </summary>
        /// <param name="mine"></param>
        /// <returns></returns>
        private string MineToStr(string mine)
        {
            if (mine.IndexOf("?", StringComparison.OrdinalIgnoreCase) != -1)
            {
                var coding = mine.Split('?')[1];
                var bs64 = mine.Split('?')[3];
                try
                {
                    byte[] bytes = Convert.FromBase64String(bs64);
                    var str = Encoding.GetEncoding(coding).GetString(bytes);
                    return str;
                }
                catch (Exception ex)
                {
                    return bs64;
                }
            }
            else
            {
                return mine;
            }
        }

        /// <summary>
        /// 转成Search命令的日期格式
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        private string ToSearchDt(DateTime dateTime)
        {
            string[] arr = dateTime.Date.ToString("yyyy-MM-dd").Split('-');
            string month = "Jan";
            if (arr[1] == "02")
            {
                month = "Feb";
            }
            else if (arr[1] == "03")
            {
                month = "Mar";
            }
            else if (arr[1] == "04")
            {
                month = "Apr";
            }
            else if (arr[1] == "05")
            {
                month = "May";
            }
            else if (arr[1] == "06")
            {
                month = "Jun";
            }
            else if (arr[1] == "07")
            {
                month = "Jul";
            }
            else if (arr[1] == "08")
            {
                month = "Aug";
            }
            else if (arr[1] == "09")
            {
                month = "Sep";
            }
            else if (arr[1] == "10")
            {
                month = "Oct";
            }
            else if (arr[1] == "11")
            {
                month = "Nov";
            }
            else if (arr[1] == "12")
            {
                month = "Dec";
            }
            return arr[2] + '-' + month + '-' + arr[0];
        }

        #endregion private

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            reader?.Close();
            sw?.Close();
            sslstream?.Close();
            tcpclient.Close();
        }
    }
}
