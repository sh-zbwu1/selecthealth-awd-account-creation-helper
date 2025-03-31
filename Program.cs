using NLog;
using SelectHealth.Ops.AWD.Core;
using SelectHealth.Ops.AWD.Entities;
using SelectHealth.Ops.AWD.Logger;
using System.Net;
using NPOI;
using NPOI.SS.UserModel;
using System.Text;

namespace AWD_Create_Helper
{
    internal class Program
    {
        private static string awd_url;
        private static string auth_username;
        private static string template_username;
        private static Logger logger;
        private static SH_NLogger wrapped_logger;
        private static RequestService request;
        private static AuthenticationService auth;
        private static AccountService account;
        private static HttpClient client;
        private static Credentials credentials = null;

        static async Task Main(string[] args)
        {
            try
            {
                // set up the environment
                await Setup_Env();

                var users = Read_And_Get_Users_From_Spreadsheet("TestUsers.xlsx");
                //var example_user = Return_Example_Template_User();
                foreach (var user in users)
                {
                    //var success_in_creating_user = await Create_User(user);
                    await Delete_User(user);
                    //if (success_in_creating_user)
                    //{
                    //    await Setup_Security_Group(user.Username);
                    //await Set_Workspace_to_Processor(user.Username);
                    //}

                }
            }
            finally
            {
                if (credentials != null)
                    await auth.SignOut(credentials);
                Console.WriteLine("Signed out");
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadLine();
        }

        private static async Task<bool> Delete_User(AWD_User template_user)
        {
            Console.WriteLine($"Deleting user {template_user.Username}");
            var (success, result) = await account.RemoveUserAccount(template_user.Username, 
                credentials);

            if (!success)
            {
                Console.WriteLine($"Failed to delete user {template_user.Username}, {result}");
            }
            else
            {
                Console.WriteLine($"Deleted user {template_user.Username}");
            }

            return success;
        }

        /// <summary>
        /// Adds the STANDARD security group to the user. 
        /// </summary>
        /// <param name="username"></param>
        /// <returns></returns>
        private static async Task Setup_Security_Group(string username)
        {
            var request = new HttpRequestMessage(HttpMethod.Post, $"{awd_url}awdweb");
            var payload = $"<SecurityGroupViewRequest><add><securityGroup><userId>{username}</userId><securityGroup>STANDARD</securityGroup></securityGroup></add></SecurityGroupViewRequest>";
            var content = new StringContent(payload, Encoding.UTF8, "text/xml");
            request.Content = content;
            request.Headers.Add("csrf_token", credentials.csrf_token);
            request.Headers.Add("Cookie", "JSESSIONID=" + credentials.jsession_cookie);

            var response = await client.SendAsync(request);
            var response_content = await response.Content.ReadAsStringAsync();
            if (response.StatusCode == HttpStatusCode.OK)
                Console.WriteLine($"Successfully added STANDARD security group to {username}");
            else
                Console.WriteLine($"Failed to add STANDARD security group to {username}");
        }

        /// <summary>
        /// Set the default workspace to Administrator. The user must be on cyberark for this method to work. 
        /// </summary>
        /// <param name="username"></param>
        private static async Task Set_Workspace_to_Processor(string username)
        {
            var password = await auth.GetCyberarkPassword(username, "DEV01");
            Credentials _credentials = null;
            try
            {
                _credentials = await auth.SignIn(username, password);
                var request = new HttpRequestMessage(HttpMethod.Put, $"{awd_url}services/v1/user/workspace?name=WSPROCSR&_=1733331099811");
                request.Headers.Add("csrf_token", _credentials.csrf_token);
                request.Headers.Add("Cookie", "JSESSIONID=" + _credentials.jsession_cookie);

                var response = await client.SendAsync(request);
                var response_content = await response.Content.ReadAsStringAsync();
                if (response.StatusCode == HttpStatusCode.NoContent)
                    Console.WriteLine($"Successfully set workspace for {username} to Processor");
                else
                    Console.WriteLine("Failed to set workspace to Processor");
            }
            finally
            {
                if (_credentials != null)
                    await auth.SignOut(_credentials);
            }
        }


        private static AWD_User Return_Example_Template_User()
        {
            return new AWD_User()
            {
                Username = "ZHITEST1",
                Alias = "ZHITEST1",
                Country = "1",
                Work_Select = "1",
                Work_Group = "AWD FACIL",
                First_Name = "BATCH",
                Last_Name = "TEST",
                Status = AccountService.Status.Available,
            };
        }

        private static AWD_User Quick_Create_Template_User(string username, string firstname = "", string lastname = "")
        {

            return new AWD_User()
            {
                Username = username,
                Alias = username,
                Country = "1",
                Work_Group = "AWD FACIL",
                First_Name = string.IsNullOrWhiteSpace(firstname) ? username : firstname,
                Last_Name = string.IsNullOrWhiteSpace(lastname) ? username : lastname,
                Status = AccountService.Status.Available,
            };
        }

        private static List<AWD_User> Read_And_Get_Users_From_Spreadsheet(string xlsx_file_path)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream file = new FileStream(xlsx_file_path, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(file);
                }

                var sheet = workbook.GetSheetAt(0);
                if (sheet == null)
                {
                    Console.WriteLine($"Could not find spreadsheet at {xlsx_file_path}");
                    return null;
                }

                int username_index = 0;
                int firstname_index = 1;
                int lastname_index = 2;

                var row = sheet.GetRow(0);

                var row_i = 1;
                var awd_users_to_create = new List<AWD_User>();
                var consecutively_skipped_rows = 0;

                while ((row = sheet.GetRow(row_i)) != null)
                {
                    if (!RowHasValues(row)) // skip this row if it doesn't have any values
                    {
                        // if we've skipped more than 20 rows in a row, then we should probably stop reading from this spreadsheet
                        if (consecutively_skipped_rows > 20)
                            break;

                        // keep count of the number of consecutive rows we've skipped
                        consecutively_skipped_rows++;
                        continue;
                    }
                    else
                    {
                        consecutively_skipped_rows = 0;
                    }

                    string username = row.GetCell(username_index).ToString();
                    string firstname = row.GetCell(firstname_index)?.ToString();
                    string lastname = row.GetCell(lastname_index)?.ToString();

                    awd_users_to_create.Add(Quick_Create_Template_User(username, firstname, lastname));

                    row_i++;
                }
                return awd_users_to_create;
            }
            finally
            {
                workbook?.Close();
            }
        }
        private static bool RowHasValues(IRow row)
        {
            foreach (var i in row.Cells) // iterate over cells
            {
                if (!string.IsNullOrWhiteSpace(i.StringCellValue))
                    return true; // if any cell is not empty, then return true
            }

            return false;
        }

        private static async Task Setup_Env()
        {
            awd_url = "https://awdwebdev.co.ihc.com/awdServer/awd/";
            auth_username = "AWDWBSRV";
            template_username = "AWDWBSRV";

            logger = NLog.LogManager.GetCurrentClassLogger();
            wrapped_logger = new SH_NLogger(logger);

            request = new RequestService(awd_url, wrapped_logger);
            auth = new AuthenticationService(request, wrapped_logger);
            account = new AccountService(request, wrapped_logger);

            var auth_user_password = await auth.GetCyberarkPassword(auth_username, "DEV01");
            credentials = await auth.SignIn(auth_username, auth_user_password);

            client = new HttpClient();

            Console.WriteLine("Signed in");
        }


        private static async Task<bool> Create_User(AWD_User template_user)
        {
            Console.WriteLine($"Creating user {template_user.Username}");
            var (success, result) = await account.CreateUserAccount(template_user.Username, template_user.First_Name, template_user.Middle_Name, template_user.Last_Name, template_user.Alias,
                template_user.Business_Area, template_user.Work_Select, template_user.Work_Group, template_user.Security_Level, template_user.Redirect,
                template_user.Country, template_user.In_By, template_user.Out_By, template_user.Phone, template_user.Status,
                template_user.Forward_Queue, template_user.Work_Action, template_user.Personal_Queue,
                credentials);

            if (!success)
            {
                Console.WriteLine($"Failed to create user {template_user.Username}, {result}");
            }
            else
            {
                Console.WriteLine($"Successively created user {template_user.Username}");
            }

            return success;
        }
    }
}
