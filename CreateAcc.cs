using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.Configuration;

namespace ADProject
{
    class Program
    {

        static DirectoryEntry de = null;
        static string domain = "LDAP://bill.com/";
        static string queryString = "LDAP://bill.com/OU=dxt,DC=bill,DC=com";
        static string account = "kevin@bill.com";
        static string pwd = "P@ssw0rd123";


        static void Main(string[] args)
        {

            using (de = new DirectoryEntry(queryString, account, pwd))// AD物件
            {
                string value = "";

                try
                {
                    //列出所有帳號及屬性
                    int count = 0;
                    foreach (string key in de.Properties.PropertyNames)//取得所有屬性名稱
                    {
                        count++;
                        foreach (var propVal in de.Properties[key])//取得所有屬性的value
                        {
                            value = key + "=" + propVal.ToString();
                            Console.WriteLine(value);
                        }
                    }
                    Console.WriteLine(count);

                }
                catch (Exception e)
                {
                    Console.WriteLine($"AD : 連線錯誤 : {e}");
                }
                string result = Create_User("dxt", "Ruby47");
                Console.WriteLine(result);
                Console.ReadLine();

            }

        }

        public static string Create_User(string OUName, string userName)
        {
            string status = "fail";
            string DN = Find_OU(OUName);
            if (DN == "null")
            {
                return "OU is not exist";
            }


            string defaultPwd = "P@ssw0rd";
            using (de = new DirectoryEntry(domain + DN, account, pwd)) // AD物件
            {
                using (DirectorySearcher ds = new DirectorySearcher(de))
                {
                    // 確認帳號是否存在, 但最好連同grounp、OU、聯絡人...都一起檢查, 因為一個區域(folder)一個名稱只能出現一次

                    ds.Filter = $"(&(objectClass=user)(cn={userName}))";
                    var result = ds.FindOne();
                    if (result != null)
                    {
                        return "account is exist";
                    }
                }


                //建立帳號物件
                using (DirectoryEntry user = de.Children.Add($"CN={userName}", "user"))
                {
                    // 建立帳號設定屬性
                    try
                    {
                        user.Properties["displayName"].Add(userName); //顯示名稱
                        user.Properties["UserPrincipalName"].Add(userName + "@bill.com"); //用來登入的帳號
                        user.CommitChanges();
                        user.Invoke("SetPassword", defaultPwd);
                        user.CommitChanges();
                        string dn = user.Properties["distinguishedName"][0].ToString();
                        EnableUser(dn);//啟用帳號
                        status = "success";
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    //       Console.WriteLine(user.Properties["distinguishedName"][0].ToString());
                }
            }
            return status;
        }

        //啟用帳號
        public static void EnableUser(string dn)
        {
            using (DirectoryEntry accountStatus = new DirectoryEntry(domain + dn, account, pwd))
            {
                var val = (int)accountStatus.Properties["userAccountControl"].Value;
                accountStatus.Properties["userAccountControl"].Value = val ^ 0x2; //去做XOR，也就是啟動會變停用，停用會變啟用
                accountStatus.CommitChanges();
            }
        }


        //取得使用者 DN
        public static string GetUserDn(string userName)
        {

            using (DirectorySearcher ds = new DirectorySearcher(de))
            {
                ds.Filter = string.Format("(&(objectClass=user)(cn={0}))", userName);

                SearchResult result = ds.FindOne();
                if (result != null)
                {
                    return result.Properties["distinguishedName"][0].ToString();
                }
                return "AD GetUserDN ex: no this account";
            }
        }


        //尋找AD中目標組織單位並回傳DN值
        public static string Find_OU(string OU)
        {
            using (DirectorySearcher sr = new DirectorySearcher(de))
            {
                sr.Filter = string.Format("(&(objectClass=organizationalUnit)(ou={0}))", OU);
                SearchResult result = sr.FindOne();
                if (result != null)
                {
                    Console.WriteLine(result.Properties["distinguishedName"][0].ToString());
                    return result.Properties["distinguishedName"][0].ToString();
                }
                else
                {
                    return "null";
                }
            }
        }
    }
}
