using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Unizap.Addon.QuickCopy
{
    internal class LicenseManager
    {
        private IOrganizationService _crmService = null;
        private IPluginExecutionContext _context = null;
        private const int TRIAL_NO_OF_DAYS = 2;

        public LicenseManager(ref IOrganizationService crmService, ref IPluginExecutionContext context)
        {
            _crmService = crmService;
            _context = context;
        }

        public int ValidateLicense(string licenseFileName, string productName)
        {
            QueryExpression qe = new QueryExpression();
            qe.EntityName = EntityName.Webresource;
            qe.ColumnSet = new ColumnSet() { AllColumns = true };

            ConditionExpression ce = new ConditionExpression();

            ce.AttributeName = EntityAttributes.Name;
            ce.Operator = ConditionOperator.Equal;
            ce.Values.Add(licenseFileName);
            FilterExpression fe = new FilterExpression();
            fe.Conditions.Add(ce);
            qe.Criteria = fe;

            int result = LicenseRetrieve(_crmService.RetrieveMultiple(qe), productName);
            return result;
        }

        private int LicenseRetrieve(EntityCollection result, string ProductName)
        {
            EntityCollection retrievedRecords = result;

            if (retrievedRecords.Entities.Count <= 0)
            {
                //return Result.LicenseNotAvailable;

                if (IsTrialExpired(ProductName, TRIAL_NO_OF_DAYS))
                {
                    return Result.TrialExpried;
                }
                else
                {
                    //trial version
                    return Result.TrialLicense;
                }
            }

            else
            {
                Entity wb = (Entity)retrievedRecords.Entities[0];

                byte[] b = Convert.FromBase64String(wb.GetAttributeValue<string>(EntityAttributes.Content));

                string licenseText = System.Text.Encoding.UTF8.GetString(b, 0, b.Length);

                return ProcessLicense(ProductName, licenseText);
            }
        }

        private int ProcessLicense(string ProductName, string licenseText)
        {
            string token = licenseText.Replace("<token>", "");
            token = token.Replace("</token>", "");
            token = token.Replace("<license>", "");
            token = token.Replace("</license>", "");
            token.Trim();

            string decryptedText = Decrypt(token, true);

            string[] licenseValues = decryptedText.Split('|');

            Dictionary<string, string> dict = new Dictionary<string, string>();
            dict.Add("Expiry", licenseValues[0]);
            dict.Add("Org", licenseValues[1]);
            dict.Add("Product", licenseValues[2]);
            dict.Add("UserCount", licenseValues[3]);

            DateTime licenceExpiredDate = Convert.ToDateTime(dict["Expiry"]);
            DateTime CurrenDateTime = DateTime.Now.Date;
            int IsValidDate = licenceExpiredDate.Subtract(CurrenDateTime).Days;
            if (!ProductName.Equals(dict["Product"], StringComparison.OrdinalIgnoreCase))
                return Result.LicenseInvalid;
            if (IsValidDate < 0)
                return Result.LicenseExpired;
            else
            {
                if (ValidateOrg(dict["Org"]) == Result.LicenseValid)
                {
                    int licenseUser = Convert.ToInt32(dict["UserCount"]);
                    if (licenseUser == 0 || CheckValidUserCount(licenseUser, _crmService, _context))
                    {
                        return Result.LicenseValid;
                    }
                    else
                    {
                        return Result.UserCountError;
                    }
                }
                else
                {
                    return Result.LicenseInvalid;
                }
            }
        }

        private bool IsTrialExpired(string Addon, int TrialDays)
        {
            QueryExpression qe = new QueryExpression();
            qe.EntityName = "solution";
            qe.ColumnSet = new ColumnSet() { AllColumns = true };
            ConditionExpression ce = new ConditionExpression();

            ce.AttributeName = "uniquename";
            ce.Operator = ConditionOperator.Equal;
            ce.Values.Add(Addon);
            FilterExpression fe = new FilterExpression();
            fe.Conditions.Add(ce);
            qe.Criteria = fe;
            EntityCollection retrievedRecords = _crmService.RetrieveMultiple(qe);

            if (retrievedRecords.Entities.Count <= 0)
            {
                throw new SystemException("Could not find add-on solution or solution names mis-match.");
            }
            else
            {
                Entity solution = (Entity)retrievedRecords.Entities[0];
                DateTime installDate = solution.GetAttributeValue<DateTime>("installedon");
                if (installDate.ToUniversalTime().AddDays(TrialDays) > DateTime.Today.ToUniversalTime())
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        private bool CheckValidUserCount(int licenseUser, IOrganizationService _crmService, IPluginExecutionContext executionContext)
        {
            QueryExpression qe = new QueryExpression();
            qe.EntityName = "systemuser";
            qe.ColumnSet = new ColumnSet("firstname");
            ConditionExpression ce = new ConditionExpression();

            ce.AttributeName = "isdisabled";
            ce.Operator = ConditionOperator.Equal;
            ce.Values.Add("0");
            ConditionExpression ce2 = new ConditionExpression();
            ce2.AttributeName = "accessmode";
            ce2.Operator = ConditionOperator.NotEqual;
            ce2.Values.Add("3");
            FilterExpression fe = new FilterExpression();
            fe.Conditions.Add(ce);
            fe.Conditions.Add(ce2);
            qe.Criteria = fe;

            RetrieveMultipleRequest retrieveMultipleRequest = new RetrieveMultipleRequest();
            RetrieveMultipleResponse retrieveMultipleResponse = new RetrieveMultipleResponse();
            retrieveMultipleRequest.Query = qe;

            retrieveMultipleResponse = (RetrieveMultipleResponse)_crmService.Execute(retrieveMultipleRequest);

            EntityCollection entityCollection = retrieveMultipleResponse.EntityCollection;

            if (entityCollection.Entities.Count <= licenseUser)//License found
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private int ValidateOrg(string orgname)
        {
            string org = _context.OrganizationName;
            if (!string.IsNullOrEmpty(orgname))
            {
                if (orgname.Equals(org, StringComparison.OrdinalIgnoreCase))
                    return Result.LicenseValid;
                else
                    return Result.LicenseInvalid;
            }
            else
            {
                return Result.LicenseValid;
            }
        }

        private static string Decrypt(string cipherString, bool useHashing)
        {
            byte[] keyArray;

            //get the byte code of the string

            byte[] toEncryptArray = Convert.FromBase64String(cipherString);

            string key = "zapuni";

            if (useHashing)
            {
                //if hashing was used get the hash code with regards to your key
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));

                //release any resource held by the MD5CryptoServiceProvider

                hashmd5.Clear();
            }
            else
            {
                //if hashing was not implemented get the byte code of the key
                keyArray = UTF8Encoding.UTF8.GetBytes(key);
            }

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();

            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;

            //mode of operation. there are other 4 modes.
            //We choose ECB(Electronic code Book)

            tdes.Mode = CipherMode.ECB;

            //padding mode(if any extra byte added)
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);

            //Release resources held by TripleDes Encryptor
            tdes.Clear();

            //return the Clear decrypted TEXT
            return UTF8Encoding.UTF8.GetString(resultArray);
        }
    }

    public class Result
    {
        public const int TrialExpried = 1;
        public const int TrialLicense = 2;
        public const int LicenseValid = 0;
        public const int LicenseNotAvailable = -1;
        public const int LicenseExpired = -2;
        public const int LicenseInvalid = -3;
        public const int UserCountError = -4;
    }

    public class EntityName
    {
        public static readonly string Webresource = "webresource";
        public static readonly string Systemuser = "systemuser";
        public static readonly string Unit = "uomschedule";
        public static readonly string Solution = "solution";
    }

    public class EntityAttributes
    {
        public static readonly string Content = "content";
        public static readonly string Isdisabled = "isdisabled";
        public static readonly string Statecode = "statecode";

        public static readonly string Fullname = "fullname";
        public static readonly string Name = "name";
        public static readonly string SolutionName = "uniquename";
    }
}