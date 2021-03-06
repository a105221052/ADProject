﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.Configuration;
using System.Collections;

namespace ADProject
{
    class Program
    {

        static DirectoryEntry de = null;
        static string domain = "LDAP://bill.com/";
        static string queryString = "LDAP://bill.com/DC=bill,DC=com";
        static string account = "Rabbit@bill.com";
        static string pwd = "";


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
                string restart = "y";
                while (restart.ToLower() == "y")
                {
                    string[] func = new string[]{
                    "輸入0:退出","輸入1:新增使用者","輸入2:搬移使用者","輸入3:設定帳戶到期日",
                    "輸入4:刪除帳戶","輸入5:加入群組", "輸入6:重設使用者密碼"
                    };
                    bool flag = false;
                    int n = 0;
                    while (!flag)
                    {
                        for (int i = 0; i < func.Length; i += 2)
                        {
                            int length = func[i].Length;
                            while (length < 12)
                            {
                                func[i] = func[i].Insert(func[i].Length, "  ");
                                length++;
                            }
                            try
                            {
                                Console.WriteLine("{0} {1}", func[i], func[i + 1]);
                            }
                            catch
                            {
                                //超出陣列
                                Console.WriteLine(func[i]);
                            }
                        }
                        try
                        {
                            n = int.Parse(Console.ReadLine());
                            flag = true;
                        }
                        catch
                        {
                            Console.WriteLine("請輸入正確代號\n");
                        }
                    }
                    if (n == 0)
                    {
                        Environment.Exit(0);
                       // break;
                    }
                    if (n == 1)
                    {
                        Console.Write("請輸入要新增的使用者名稱 : ");
                        string newUser = Console.ReadLine().Trim();
                        Console.Write("請輸入此使用者的 OU : ");
                        string userOU = Console.ReadLine();
                        string result = Create_User(userOU, newUser);
                        Console.WriteLine($"新增結果 : {result}");
                    }
                    else if (n == 2)
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
                    else if (n == 3)
                    {
                        // 設定帳戶到期日是今日
                        Console.Write("請輸入設定對象的 OU : ");
                        string userOU = Console.ReadLine();
                        Console.Write("請輸入設定對象的名稱 : ");
                        string userName = Console.ReadLine();
                        string setExpiredDate = Set_Expired_Date(userOU, userName);
                        Console.WriteLine(setExpiredDate);
                    }
                    else if (n == 4)
                    {
                        //刪除帳號
                        Console.Write("請輸入刪除對象的 OU : ");
                        string userOU = Console.ReadLine();
                        Console.Write("請輸入刪除對象的名稱 : ");
                        string userName = Console.ReadLine();
                        string delUser = Del_User(userOU, userName);
                        Console.WriteLine(delUser);
                    }
                    else if (n == 5)
                    {
                        //加入群組
                        Console.Write("請輸入使用者的 OU : ");
                        string userOU = Console.ReadLine();
                        Console.Write("請輸入使用者的名稱 : ");
                        string userName = Console.ReadLine();
                        Console.WriteLine("請輸入目標 GROUP");
                        string group = Console.ReadLine();
                        string joinGroup = Join_Group(userOU, userName, group);
                        Console.WriteLine(joinGroup);
                    }
                    else
                    {
                        //重設使用者密碼
                        Console.WriteLine("請輸入使用者的OU : ");
                        string userOU = Console.ReadLine();
                        Console.WriteLine("請輸入使用者的名稱 : ");
                        string userName = Console.ReadLine();
                        Console.WriteLine("請輸入新密碼");
                        string newPwd = Console.ReadLine();
                        string result = Change_Pwd(userOU, userName, newPwd);
                        Console.WriteLine(result);
                    }
                    restart = "";//清空，等待這次回應
                    while (restart.ToLower() != "n" && restart.ToLower() != "y")
                    {
                        Console.WriteLine("continue?");
                        restart = Console.ReadLine();
                    }
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
                var result = Search_User(de, userName);
                if (result != null) // 這個使用者已經存在
                {
                    return "account is exist";
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
                //直接改0x200也可以
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
            using (DirectorySearcher ds = new DirectorySearcher(de))
            {
                // 確認帳號是否存在, 但最好連同grounp、OU、聯絡人...都一起檢查, 因為一個區域(folder)一個名稱只能出現一次
                ds.Filter = string.Format("(&(objectClass=user)(cn={0}))", name);//在 OU 找尋使用者
                SearchResult result = ds.FindOne();
                return result;
            }
        }

        public static string Set_Expired_Date(string OU, string name)
        {
            string message = "fail";
            string DN = Find_OU(OU); //先檢查OU存不存在，存在會回傳 OU 的 DN
            if (DN == "null")
            {
                return "OU is not exist";
            }
            using (DirectoryEntry de = new DirectoryEntry(domain + DN, account, pwd))
            {
                var result = Search_User(de, name);
                if (result == null) // 這個使用者不存在
                {
                    return "OU 裡不存在此 User.";
                }
                //取得使用者物件
                using (DirectoryEntry userObject = new DirectoryEntry(result.Path, account, pwd))
                {
                    DateTime date = new DateTime();
                    date = DateTime.Now;
                    //停用帳號
                    userObject.Properties["accountExpires"].Value = date.ToFileTime().ToString();
                    // 如要啟用帳戶
                    // userObject.Properties["accountExpires"].Value = 0;
                    userObject.CommitChanges();
                    message = "success";
                }
            }
            return message;
        }

        public static string Del_User(string OU, string name)
        {
            string DN = Find_OU(OU); //先檢查OU存不存在，存在會回傳 OU 的 DN
            if (DN == "null")
            {
                return "OU is not exist";
            }
            using (DirectoryEntry de = new DirectoryEntry(domain + DN, account, pwd))
            {
                var result = Search_User(de, name);
                if (result == null) // 這個使用者不存在
                {
                    return "OU 裡不存在此 User.";
                }
                // 如果 OU 不存在 USER 沒 return 會出錯, 因為result.Path就沒有值了
                using (DirectoryEntry userObject = new DirectoryEntry(result.Path, account, pwd))
                {
                    de.Children.Remove(userObject);
                }
            }
            return "success";
        }

        public static string Join_Group(string OU, string name, string group)
        {
            string result = "fail";
            string userPath = "";
            string userDN = "";
            string OUDN = Find_OU(OU);
            if (OUDN == "null")
            {
                return "OU is not exist";
            }

            using (DirectoryEntry de = new DirectoryEntry(domain + OUDN, account, pwd))
            {
                var user = Search_User(de, name);
                if (user == null)
                {
                    return "OU 不存在此 USER";
                }
                userPath = user.Path;
                userDN = user.Properties["distinguishedName"][0].ToString(); ;
                string groupDN = Find_Group(group);//目前的寫法是整個 domian 的 group 都會找到
                // 如要限定 OU 多帶一個參數 Find_Group(group, de) 即可
                if (groupDN == "null")
                {
                    return "此 Group 不存在";
                }
                // 確定 OU、USER、GROUP 都存在
                else
                {
                    using (DirectoryEntry groupDE = new DirectoryEntry(domain + groupDN, account, pwd))
                    {
                        groupDE.Properties["member"].Add(userDN);
                        groupDE.CommitChanges();
                        result = $"使用者 {name} 成功加入群組 : {group}";
                    }
                }
            }
            return result;
        }

        public static string Find_Group(string Group)
        {
            using (DirectorySearcher sr = new DirectorySearcher(de))
            {
                sr.Filter = string.Format("(&(objectClass=group)(cn={0}))", Group);
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

        public static string Change_Pwd(string OU, string name, string newPwd)
        {
            string result = "fail";
            string OUDN = Find_OU(OU);
            if (OUDN == "null")
            {
                return $"OU : {OU} 不存在";
            }
            using (DirectoryEntry de = new DirectoryEntry(domain + OUDN, account, pwd))
            {
                var user = Search_User(de, name);
                if (user == null)
                {
                    return $"User : {name} 不存在";
                }
                using (DirectoryEntry changePwd = new DirectoryEntry(user.Path, account, pwd))
                {
                    try
                    {
                        changePwd.Invoke("SetPassword", new object[] { newPwd });
                        changePwd.CommitChanges();
                    }
                    catch
                    {
                        return "密碼不符合複雜性原則!";
                    }
                    result = $"使用者 {name} 密碼已變更!";
                }
            }
            return result;
        }
    }
}


