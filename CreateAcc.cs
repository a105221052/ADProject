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
        static string queryString = "LDAP://bill.com/DC=bill,DC=com";
        static string account = "Rabbit@bill.com";
        static string pwd = "P@ssw0rd";


        static void Main(string[] args)
        {

            using (de = new DirectoryEntry(queryString, account, pwd)) // AD物件
            {
                string value = "";
                try
                {
                    //列出所有屬性
                    int count = 0;
                    foreach (string key in de.Properties.PropertyNames)//取得所有屬性名稱, 如管理者帳戶或密碼有誤，在取 Properties 時就會出錯
                    {
                        count++;
                        foreach (var propVal in de.Properties[key])//取得所有屬性的value
                        {
                            value = key + "=" + propVal.ToString();
                            Console.WriteLine(value);
                        }
                    }
                    Console.WriteLine($"屬性數量 : {count}");
                }
                catch (Exception e)
                {
                    Console.WriteLine($"AD : 連線錯誤 : {e}");
                }
                Console.WriteLine("請選擇功能 輸入1:新增使用者，輸入2搬移使用者 : ");
                int n =  int.Parse(Console.ReadLine());
                if (n == 1)
                {
                    Console.Write("請輸入要新增的使用者名稱 : ");
                    string newUser = Console.ReadLine().Trim();
                    Console.Write("請輸入此使用者的 OU : ");
                    string userOU = Console.ReadLine();
                    string result = Create_User(userOU, newUser);
                    Console.WriteLine($"新增結果 : {result}");
                }
                else
                {
                    Console.Write("請輸入搬移對象的 OU : ");
                    string userOU = Console.ReadLine();
                    Console.Write("請輸入搬移對象的名稱 : ");
                    string userName = Console.ReadLine();
                    Console.Write("請輸入要搬移至哪個 OU : ");
                    string destinationOU = Console.ReadLine();
                    string moveResult = Move_User(userName, userOU, destinationOU);
                    Console.WriteLine(moveResult);
                }
                Console.ReadLine();
            }
        }

        public static string Create_User(string OUName, string userName)
        {
            string status = "fail";
            string DN = Find_OU(OUName); //先檢查OU存不存在，存在會回傳 OU 的 DN
            if (DN == "null")
            {
                return "OU is not exist";
            }

            string defaultPwd = "P@ssw0rd";

            using (DirectoryEntry de = new DirectoryEntry(domain + DN, account, pwd)) // AD物件
            {
                using (DirectorySearcher ds = new DirectorySearcher(de))
                {
                    // 確認帳號是否存在, 但最好連同grounp、OU、聯絡人...都一起檢查, 因為一個區域(folder)一個名稱只能出現一次
                    ds.Filter = $"(&(objectClass=user)(cn={userName}))";
                    var result = ds.FindOne();
                    if (result != null) // 這個使用者已經存在
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
                        user.Properties["sAMAccountName"].Add(userName); //用來登入的帳戶
                        user.Properties["UserPrincipalName"].Add(userName + "@bill.com"); //用來登入的帳號
                        user.CommitChanges();
                        user.Invoke("SetPassword", defaultPwd);
                        user.CommitChanges();
                        string dn = user.Properties["distinguishedName"][0].ToString(); //取得剛創建完的 user 的 DN
                        EnableUser(dn);//啟用帳號
                        status = "success";
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                }
            }
            return status;
        }
        
        //尋找AD中目標組織單位並回傳 OU 的 DN 值
        public static string Find_OU(string OU)
        {
            using (DirectorySearcher sr = new DirectorySearcher(de))
            {
                sr.Filter = string.Format("(&(objectClass=organizationalUnit)(ou={0}))", OU);
                try // 使用者帳號或密碼如有錯誤 sr.FindOne會出現錯誤
                {
                    SearchResult result = sr.FindOne();
                    if (result != null)
                    {
                        // Console.WriteLine(result.Properties["distinguishedName"][0].ToString());
                        return result.Properties["distinguishedName"][0].ToString();
                    }
                }
                catch
                {
                    Console.WriteLine("管理者帳戶或密碼錯誤");
                }
                return "null";
            }
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



        #region 搬移帳號

        public static string Move_User(string name, string oldOU, string newOU)
        {
            string oldOU_DN = Find_OU(oldOU);//尋找 OU 有的話回傳 DN
            if (oldOU_DN == "null")
            {
                return "您所選擇的起始OU不存在";
            }
            string msg = "";
            using (DirectoryEntry de = new DirectoryEntry(domain + oldOU_DN, account, pwd))
            {
                try //FindOne 如果管理者帳號有問題會到 catch
                {
                    SearchResult result = Search_User(de, name);
                    if (result != null) //找到目標使用者，接著嘗試搬移
                    {
                        try
                        {
                            newOU = Find_OU(newOU); //查詢這個 OU 存不存在，存在的話回傳 OU 的 DN
                            if (newOU == "null")
                            {
                                return "您所指定的目的地 OU 不存在";
                            }
                            de.Path = domain + newOU;

                            //須設定驗證型別
                            AuthenticationTypes AuthTypes = AuthenticationTypes.Signing | AuthenticationTypes.Sealing | AuthenticationTypes.Secure | AuthenticationTypes.ServerBind;
                            //result.Path 是使用者原先的路徑(DN)
                            DirectoryEntry usr = new DirectoryEntry(result.Path, account, pwd, AuthTypes);
                            DirectoryEntry destination = new DirectoryEntry(de.Path, account, pwd);
                            usr.MoveTo(destination);
                            usr.CommitChanges();
                            msg = "success";
                        }
                        catch (Exception ex)
                        {
                            msg += ex.Message;
                            Console.WriteLine("moveUser error:" + ex.Message);
                        }
                    }
                    else
                    {
                        return $"您所選擇的起始 OU 並不存在使用者 : {name}";
                    }
                }
                catch
                {
                    Console.WriteLine("管理者帳戶或密碼錯誤");
                }

                return msg;
            }

        }

        #endregion

        public static SearchResult Search_User(DirectoryEntry de, string name)
        {
            DirectorySearcher ds = new DirectorySearcher(de);
            ds.Filter = string.Format("(&(objectClass=user)(cn={0}))", name);//在 OU 找尋使用者
            SearchResult result = ds.FindOne();
            return result;
        }
    }
}
