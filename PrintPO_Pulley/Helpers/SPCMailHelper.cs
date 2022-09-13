using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.Helpers
{
    public class SPCMailHelper
    {
        private static SPCMailHelper instance;
        public static SPCMailHelper Instance
        {
            get {
                if (instance == null)
                    instance = new SPCMailHelper();
                return instance;
            }
            set
            {
                instance = value;
            }
        }

        public int PrepareSPCMail(string fromEmail, string toAddress, string subject, string body, string mailsCc)
        {
            //16-11-2020: cập nhật TLS1.2 cho mail
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            //16-11-2020

            //2019-11-28 Start lock
            //using (System.Net.Mail.MailMessage myMail = new System.Net.Mail.MailMessage())
            //{
            //    myMail.From = new MailAddress(fromEmail);
            //    myMail.To.Add(toAddress);
            //    myMail.Subject = subject;
            //    myMail.IsBodyHtml = true;
            //    myMail.Body = body;
            //    myMail.IsBodyHtml = true;
            //    try
            //    {
            //        using (System.Net.Mail.SmtpClient s = new System.Net.Mail.SmtpClient("mailnew.spclt.com.vn"))
            //        {
            //            //s.DeliveryMethod = SmtpDeliveryMethod.Network;
            //            //s.UseDefaultCredentials = false;
            //            //s.Credentials = new System.Net.NetworkCredential(myMail.From.ToString(), password);
            //            //s.EnableSsl = true;
            //            s.Send(myMail);
            //        }
            //        return 1;
            //    }
            //    catch (Exception ex)
            //    {
            //        throw ex;
            //    }

            //}
            //2019-11-28 End lock

            //2019-11-28 Start add: cách gửi mail mới
            string[] mailAddresses = toAddress.Replace(',', ';').Split(';');
            List<string> _mailsCc = new List<string>();
            if (!string.IsNullOrEmpty(mailsCc))
            {
                _mailsCc.AddRange(mailsCc.Replace(',', ';').Split(';'));
            }

            int result = 0;
            
                try
                {
                    using (System.Net.Mail.MailMessage myMail = new System.Net.Mail.MailMessage())
                    {
                        myMail.From = new MailAddress(fromEmail);

                        foreach (var address in mailAddresses)
                        {
                            myMail.To.Add(address.Trim());
                        }



                    //add mails cc
                    foreach (var mCc in _mailsCc)
                    {
                        myMail.CC.Add(new MailAddress(mCc.Trim()));
                    }


                        myMail.Subject = subject;
                        myMail.IsBodyHtml = true;
                        myMail.Body = body;
                      
                        //2019-12-24 Start edit: thay đổi mailnew.spclt.com.vn thành smtp.office365.com, port 587, tạm thời dùng mail chị Hạnh vì PC chưa có mail group
                        using (System.Net.Mail.SmtpClient s = new System.Net.Mail.SmtpClient("smtp.office365.com", 587))
                        {
                            
                            s.DeliveryMethod = SmtpDeliveryMethod.Network;
                            s.UseDefaultCredentials = false;
                            s.Credentials = new System.Net.NetworkCredential("honghanh@spclt.com.vn", "Hanhnguyen1312");
                            s.EnableSsl = true;                            

                            s.Send(myMail);
                        }
                    }
                    result = 1;
                }
                catch (Exception ex)
                {
                    result = -1;
                }

            
            return result;
            //2019-11-28 End add

            ////Send email

            //MailMessage mail = new MailMessage();
            //// set the addresses
            //mail.From = new MailAddress("PC_SPCF4@spclt.com.vn", "HuaVanBe");
            //// mail.[to].add("be.hua@spclt.com.vn")

            ////Get email address 
            //string EmailAdress = "";
            //EmailAdress = ms.getScalar(String.Format("Select link from {0} where sapno ='{1}'", database, CameraCode));

            //if (EmailAdress != "")
            //{
            //    mail.To.Add(EmailAdress);
            //    mail.cc.add("be.hua@spclt.com.vn,hanthanh");


            //    // set the content

            //    mail.Subject = "[AutoEmail] "
            //                + DateTime.Now.ToString("[yyyyMMddHHmmss]") + "_Chương trình Camera bị tắt khi đang khóa: " + CameraCode;
            //    string lockedOrder = "";
            //    try
            //    { //lockedOrder = dgv1.Rows[0].Cells[0].Value.ToString(); 
            //    }
            //    catch { }
            //    mail.Body = "OrderNo bị khóa: " + lockedOrder
            //                + "<br>EmpID: " + EmpID.Text
            //                + "<br>EmpName: " + EmpName.Text
            //                + "<br>User Name: " + Environment.UserName
            //                + "<br>Computer Name: " + Environment.MachineName
            //                ;
            //    mail.IsBodyHtml = true;
            //    SmtpClient smtp = new SmtpClient("mailnew.spclt.com.vn");
            //    // send the message
            //    try
            //    {
            //        smtp.Send();
            //        // MsgBox("Your Email has been sent sucessfully - Thank You")
            //    }
            //    catch (Exception exc)
            //    {
            //        // MsgBox("Send failure: " & exc.ToString())
            //        //Connection.SaveError(10, "DeletePO", "DeletePO", "~~~");
            //    }
            //}
        }
    }
}
