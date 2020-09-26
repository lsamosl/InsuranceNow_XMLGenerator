using Services.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;

namespace Services
{
    public class Tokenizer
    {
        private string naeUser;
        private string naePassword;
        private string dbUser;
        private string dbPswd;
        private string tableName;
        private TokenizerService.SafeNetTokenizerPortTypeClient service;

        public Tokenizer()
        {
            naeUser = ConfigurationManager.AppSettings.Get("NaeUser");
            naePassword = ConfigurationManager.AppSettings.Get("NaePassword");
            dbUser = ConfigurationManager.AppSettings.Get("DbUser");
            dbPswd = ConfigurationManager.AppSettings.Get("DbPswd");
            tableName = ConfigurationManager.AppSettings.Get("TableName");

            service = new TokenizerService.SafeNetTokenizerPortTypeClient();
        }

        public string PolicyNumber { get; set; }
        public string DriverName { get; set; }

        public ResponseDetokenize Detokenize(string Token)
        {
            bool isIntranet = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("isIntranet"));

            ResponseDetokenize response = new ResponseDetokenize() { Result = false };
            string detokenizeString = string.Empty;
            try
            {
                if (isIntranet)
                    detokenizeString = service.GetValue(naeUser, naePassword, dbUser, dbPswd, tableName, Token, 0);
                else
                    detokenizeString = Token;

                response = new ResponseDetokenize()
                {
                    DetokenizedString = detokenizeString,
                    Result = true,
                };
            }

            catch (Exception e)
            {
                var path = ConfigurationManager.AppSettings.Get("LogFileName");

                if (!File.Exists(path))
                {
                    File.WriteAllText(path, string.Format("PolicyNumber    DriverName    License    ErrorMessage    DateTime") + Environment.NewLine);                    
                }
                    

                File.AppendAllText(path, string.Format("{0}    {1}    {2}    {3}    {4}", PolicyNumber, DriverName, Token, e.Message, DateTime.Now) + Environment.NewLine);

                response.Message = e.Message;
            }

            return response;
        }
    }
}
