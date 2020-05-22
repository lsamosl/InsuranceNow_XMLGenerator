using Services.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

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
                response.Message = e.Message;
            }

            return response;
        }
    }
}
