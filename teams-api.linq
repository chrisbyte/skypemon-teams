<Query Kind="Program">
  <Reference Relative="components\Newtonsoft.Json.dll">D:\apps\skypemon\skypemon-teams\components\Newtonsoft.Json.dll</Reference>
  <Reference Relative="components\System.Net.Http.dll">D:\apps\skypemon\skypemon-teams\components\System.Net.Http.dll</Reference>
  <Reference Relative="components\System.Net.Http.Formatting.dll">D:\apps\skypemon\skypemon-teams\components\System.Net.Http.Formatting.dll</Reference>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.IO.Ports</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Net.Http.Headers</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

public class Settings 
{
    public string client_id { get; set; }
    public string tenant { get; set; }
}
public static class Config 
{
    public static Settings Settings { get; set; }
 
    public static string basePath = Path.GetDirectoryName(Util.CurrentQueryPath);
    public static string comport = "COM15";
    public static string scopes = "offline_access chat.read mail.read presence.read presence.read.all";
}

void Main()
{
    if (File.Exists(Path.Combine(Config.basePath, "settings.json"))) 
    {
        Config.Settings = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(Path.Combine(Config.basePath, "settings.json")));

        // Device.ScanPorts();
    
        // create graph service
        using (GraphService service = new GraphService(Config.Settings.tenant, Config.Settings.client_id, Config.scopes, Config.basePath)) 
        {   
            //service.PrintUser();    // user
            service.PrintPresence(set_leds: false);    // user presence
            service.PrintMailbox(); // user mailbox
            //service.PrintManager(); // manager and presence
        }
    }
    else 
    {
        Console.WriteLine("no settings file found");
    }
}

// Define other methods and classes here
public class GraphService : IDisposable
{
    private readonly string authUrl;
    private readonly string clientId;
    private readonly string graphUrl;
    private readonly string scopes;
    private readonly string tenant;
    private readonly string tokenFile;
    
    private AuthAccessToken AccessToken;
    
    public GraphService(string _tenant, string _clientId, string _scopes, string _basePath) 
    {
        clientId = _clientId;
        tenant = _tenant;
        authUrl = string.Format("https://login.microsoftonline.com/{0}/oauth2/v2.0/", _tenant);
        graphUrl = "https://graph.microsoft.com/";
        scopes = _scopes;
        tokenFile = Path.Combine(_basePath, "token.txt");
        
        // load access token
        AccessToken = LoadAccessToken();
    }
    
    public void Dispose() { }
    
    public GraphMailbox GetGraphMailbox(string access_token = "", string id = "") 
    {
        return graphRequest<GraphMailbox>(
            graphHost(access_token), 
            graphEndpoint(id, "mailFolders/inbox")
        );
    }
    
    public GraphUser GetGraphManager(string access_token = "", string id = "") 
    {
        return graphRequest<GraphUser>(
            graphHost(access_token), 
            graphEndpoint(id, "manager")
        );
    }
    
    public GraphPresence GetGraphPresence(string access_token = "", string id = "") 
    {
        return graphRequest<GraphPresence>(
            graphHost(access_token, "beta"), 
            graphEndpoint(id, "presence")
        );
    }
    
    public GraphUser GetGraphUser(string access_token = "", string id = "") 
    {
        return graphRequest<GraphUser>(
            graphHost(access_token), 
            graphEndpoint(id)
        );
    }
    
    public AuthAccessToken LoadAccessToken(bool refresh = false)
    {
        // check for a stored access token
        AuthAccessToken accessToken = readAccessToken();
        double expires = 0;

          // if no token
        if (accessToken == null) 
        {
            // grant permissions and create token
            accessToken = createAccessToken();
        }
        else
        {
            expires = (DateTime.Now - File.GetLastWriteTime(tokenFile)).TotalSeconds;

            // if access token expired
            if (refresh || expires >= accessToken.expires_in)
            {
                // refresh access token
                accessToken = refreshAccessToken(accessToken.refresh_token);
                expires = 0;
            }
        }
        
        // display time until access token expires
        string date = DateTime.Now.AddSeconds(accessToken.expires_in - expires).ToString();
        int min = (int)((accessToken.expires_in - expires) / 60);
        int sec = (int)((accessToken.expires_in - expires) % 60);
        Console.WriteLine($"token refresh: {min} min {sec} sec ({date})");
        
        return accessToken;
    }
    
    private AuthAccessToken createAccessToken()
    {
        Console.WriteLine("request permissions");
        
        // request user permissions
        AuthDeviceCode deviceCode = requestPermissions();
        Console.WriteLine($"user code: {deviceCode.user_code}");
        
        HttpClient client = new HttpClient() 
        {
            Timeout = TimeSpan.FromHours(1)
        };
        string url = authUrl + "token";
        var data = new Dictionary<string, string>() 
        {
            { "tenant", tenant },
            { "client_id", clientId },
            { "grant_type", "device_code" },
            { "device_code", deviceCode.device_code }
        };
        
        AuthAccessToken accessToken = null;
        
        while (accessToken == null) 
        {
            Console.WriteLine("checking permissions");
            
            // wait for access token
            using (var response = client.PostAsync(url, new FormUrlEncodedContent(data)).Result)
            {
                string responseBody = Task.Run(() => response.Content.ReadAsStringAsync()).Result;
                if (responseBody.Contains("access_token"))
                {
                    Console.WriteLine("permissions granted");
                    accessToken = JsonConvert.DeserializeObject<AuthAccessToken>(responseBody);   
                }
            }
            
            Thread.Sleep(deviceCode.interval * 1000);
        }
        
        // save
        saveAccessToken(accessToken);
        
        return accessToken;
    }
    
    private string graphEndpoint(string id = "", string endpoint = "") 
    {
        if (string.IsNullOrEmpty(id))
            return "me/" + endpoint;
        else
            return "users/" + id + "/" + endpoint;
    }
    
    private HttpClient graphHost(string access_token = "", string version = "v1.0")
    {
        if (string.IsNullOrEmpty(access_token))
            access_token = AccessToken.access_token;
        
        var client = new HttpClient()
        {
            BaseAddress = new Uri(graphUrl + version + "/"),
            Timeout = TimeSpan.FromHours(1)
        };
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", access_token);
        
        return client;
    }
    
    private T graphRequest<T>(HttpClient host, string endpoint) 
    {
        HttpResponseMessage response = Task.Run(() => host.GetAsync(endpoint)).Result;
        response.EnsureSuccessStatusCode();
        string responseBody = Task.Run(() => response.Content.ReadAsStringAsync()).Result;
        return JsonConvert.DeserializeObject<T>(responseBody);
    }
    
    private AuthAccessToken readAccessToken() 
    {
        if (File.Exists(tokenFile)) {
            using (var stream = new StreamReader(tokenFile)) 
            {
                return new AuthAccessToken 
                {
                    expires_in = int.Parse(stream.ReadLine()),
                    access_token = stream.ReadLine(),
                    refresh_token = stream.ReadLine(),
                };
            }
        }
        
        return null;
    }
    
    private AuthAccessToken refreshAccessToken(string refresh_token)
    {
        Console.WriteLine("refresh token");
        
        string url = authUrl + "token";
        var data = new Dictionary<string, string>() 
        {
            { "tenant", tenant },
            { "client_id", clientId },
            { "refresh_token", refresh_token },
            { "grant_type", "refresh_token" },
            { "scope", scopes }
        };
        
        // request
        HttpClient client = new HttpClient() 
        {
            Timeout = TimeSpan.FromHours(1)
        };
        
        AuthAccessToken accessToken = new AuthAccessToken();
        
        using (var response = client.PostAsync(url, new FormUrlEncodedContent(data)).Result)
        {
            string responseBody = Task.Run(() => response.Content.ReadAsStringAsync()).Result;
            accessToken = JsonConvert.DeserializeObject<AuthAccessToken>(responseBody);
            
            // save
            saveAccessToken(accessToken);
        }
        
        return accessToken;
    }
    
    private AuthDeviceCode requestPermissions() 
    {
        string url = authUrl + "devicecode";
        var data = new Dictionary<string, string>() 
        {
            { "tenant", tenant },
            { "client_id", clientId },
            { "scope", scopes }
        };
        
        HttpClient client = new HttpClient() 
        {
            Timeout = TimeSpan.FromHours(1)
        };
        
        using (var response = client.PostAsync(url, new FormUrlEncodedContent(data)).Result)
        {
            string responseBody = Task.Run(() => response.Content.ReadAsStringAsync()).Result;
            var mdl = JsonConvert.DeserializeObject<AuthDeviceCode>(responseBody);
            
            // open browser for verification
            System.Diagnostics.Process.Start(mdl.verification_uri);
            
            return mdl;
        }
    }
    
    private void saveAccessToken(AuthAccessToken access_token) 
    {
        using (StreamWriter stream = new StreamWriter(tokenFile)) 
        {
            stream.WriteLine(access_token.expires_in);
            stream.WriteLine(access_token.access_token);
            stream.WriteLine(access_token.refresh_token);
        }
    }
}

//
// Extensions
//
public static class GraphServiceExtensions 
{
    public static void PrintMailbox(this GraphService auth) 
    {
        Console.WriteLine($"unread emails: {auth.GetGraphMailbox().unreadItemCount}");
    }

    public static void PrintManager(this GraphService auth) 
    {
        var manager = auth.GetGraphManager();
        Console.WriteLine($"manager: {manager.mail}; {manager.id}");
        auth.PrintPresence(manager.id, "manager");
    }

    public static void PrintPresence(this GraphService auth, string id = "", string label = "user", bool set_leds = false) 
    {
        var presence = auth.GetGraphPresence();
        string activity = presence.activity != presence.availability ? "; " + presence.activity : "";
        Console.WriteLine($"{label} presence: {presence.availability} {activity}");
        if (set_leds)
            SetLEDStatus(presence.availability);
    }
    
    public static void PrintUser(this GraphService auth) 
    {
        var user = auth.GetGraphUser();
        Console.WriteLine($"user: {user.mail}; {user.id}");
    }
    
    private static void SetLEDStatus(string availability) 
    {
        var py = new Device();
        
        if (py.OpenPort(Config.comport)) 
        {
            switch (availability)
            {    
                case "Available":
                    py.SetLEDs(new byte[] { 255, 0, 0 });
                    break;
                case "Busy":
                    py.SetLEDs(new byte[] { 0, 0, 255 });
                    break;
                case "Away":
                case "BeRightBack":
                    py.SetLEDs(new byte[] { 0, 255, 0 });
                    break;
                    
                default:
                    py.SetLEDs(new byte[] { 0, 0, 0 });
                    break;
            }
            
            py.Dispose();
        }
    }
}

//
// Models
//
public class AuthAccessToken 
{
    public int expires_in { get;set; }
    public string access_token { get;set; }
    public string refresh_token { get;set; }
}
public class AuthDeviceCode
{
    public string device_code { get;set; }
    public int interval { get;set; }
    public string user_code { get;set; }
    public string verification_uri { get; set; }
}
public class GraphMailbox
{
    public string id { get; set; }
    public string displayName { get; set; }
    public int totalItemCount { get; set; }
    public int unreadItemCount { get; set; }
}
public class GraphPresence 
{
    public string id { get; set; }
    public string availability { get; set; }
    public string activity { get; set; }
}
public class GraphUser 
{
    public string id { get; set; }
    public string mail { get; set; }
}

public class Device 
{
    private SerialPort serialPort;
    
    public Device()
    {
        serialPort = new SerialPort();
    }
    
    public void Dispose()
    {
        serialPort.Close();
        serialPort.Dispose();
    }

    public bool OpenPort(string port, string newLine = "\n")
    {
        try
        {
            serialPort.PortName = port;
            serialPort.BaudRate = 9600;
            serialPort.ReadTimeout = 1000;
            serialPort.WriteTimeout = 1000;
            serialPort.NewLine = newLine;
            serialPort.Open();

            return serialPort.IsOpen;
        }

        catch (IOException)
        {
            return false;
        }
    }

    public SerialPort Port
    {
        get { return serialPort; }
    }

    public bool Send(string command)
    {
        if (!Port.IsOpen) return false;

        try
        {
            serialPort.WriteLine(command);
            return true;
        }

        catch (IOException)
        {
            return false;
        }
    }

    public bool SetLEDs(byte[] colors)
    {
        return Send(string.Join(",", colors));
    }
    
    public static void ScanPorts() 
    {
        // Get a list of serial port names.
        string[] ports = SerialPort.GetPortNames();
    
        Console.WriteLine("The following serial ports were found:");
    
        // Display each port name to the console.
        foreach(string port in ports)
        {
            Console.WriteLine(port);
        }
    }
}