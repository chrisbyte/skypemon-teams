<Query Kind="Program">
  <Reference>&lt;ProgramFilesX64&gt;\Microsoft SDKs\Azure\.NET SDK\v2.9\bin\plugins\Diagnostics\Newtonsoft.Json.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Net.Http.dll</Reference>
  <Reference>D:\git\Development\DesktopApplications\PersonImportApp\PersonImportApp\bin\Debug\System.Net.Http.Formatting.dll</Reference>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System</Namespace>
  <Namespace>System.IO.Ports</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Net.Http.Headers</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

public class Config 
{
    public string basePath { get; set; }
    public string client_id { get; set; }
    public string tenant { get; set; }
    public string comport { get; set; }
}

public static Config AppConfig = new Config() 
{
    basePath = Path.GetDirectoryName(Util.CurrentQueryPath),
    comport = "COM15" 
};

void Main()
{
    var settingsPath = Path.Combine(AppConfig.basePath, "settings.json");
    if (File.Exists(settingsPath)) {
        var config = JsonConvert.DeserializeObject<Config>(File.ReadAllText(settingsPath));
        AppConfig.client_id = config.client_id;
        AppConfig.tenant = config.tenant;

        // Device.ScanPorts();
    
        // graph object
        MicroAuth auth = new MicroAuth(AppConfig.tenant, AppConfig.client_id);
        string accessToken = auth.LoadAccessToken().access_token;
        
        // user and presence
        //Console.WriteLine(auth.GetGraphUser(accessToken));
        var presence = auth.GetGraphPresence(accessToken);
        string activity = presence.activity != presence.availability ? presence.activity : "";
        Console.WriteLine("user presence: {0} {1}", presence.availability, activity);
        SetLEDStatus(presence);
        
        // user mailbox
        Console.WriteLine("unread emails: {0}", auth.GetGraphMailbox(accessToken).unreadItemCount);
        
    //    // manager and presence
    //    var manager = auth.GetGraphManager(accessToken);
    //    Console.WriteLine(manager);
    //    Console.WriteLine(auth.GetGraphPresence(accessToken, manager.id));
    }
    else 
    {
        Console.WriteLine("no settings file found");
    }
}

// Define other methods and classes here
public void SetLEDStatus(GraphPresence presence) 
{
    var py = new Device();
    py.OpenPort(AppConfig.comport);
    
    switch (presence.availability)
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

public class MicroAuth 
{
    private readonly string authUrl;
    private readonly string clientId;
    private readonly string graphUrl;
    private readonly string scopes;
    private readonly string tenant;
    private readonly string tokenFile;
    
    public MicroAuth(string _tenant, string _clientId) 
    {
        clientId = _clientId;
        tenant = _tenant;
        authUrl = string.Format("https://login.microsoftonline.com/{0}/oauth2/v2.0/", tenant);
        graphUrl = "https://graph.microsoft.com/";
        scopes = "offline_access presence.read presence.read.all mail.read";
        tokenFile = Path.Combine(AppConfig.basePath, "token.txt");
    }
    
    public GraphMailbox GetGraphMailbox(string access_token, string id = "") 
    {
        return graphRequest<GraphMailbox>(
            graphHost(access_token), 
            graphEndpoint(id, "mailFolders/inbox")
        );
    }
    
    public GraphUser GetGraphManager(string access_token, string id = "") 
    {
        return graphRequest<GraphUser>(
            graphHost(access_token), 
            graphEndpoint(id, "manager")
        );
    }
    
    public GraphPresence GetGraphPresence(string access_token, string id = "") 
    {
        return graphRequest<GraphPresence>(
            graphHost(access_token, "beta"), 
            graphEndpoint(id, "presence")
        );
    }
    
    public GraphUser GetGraphUser(string access_token, string id = "") 
    {
        return graphRequest<GraphUser>(
            graphHost(access_token), 
            graphEndpoint(id)
        );
    }
    
    public AuthAccessToken LoadAccessToken()
    {
        // check for a stored access token
        var accessToken = readAccessToken();

          // if no token
        if (accessToken == null) 
        {
            // grant permissions and create token
            accessToken = createAccessToken();
        }
        else
        {
            var expires = (DateTime.Now - File.GetLastWriteTime(tokenFile)).TotalSeconds;

            // if access token expired
            if (expires >= accessToken.expires_in)
            {
                // refresh access token
                accessToken = refreshAccessToken(accessToken.refresh_token);
            }   
            else
            {
                // display time until access token expires
                int min = (int)((accessToken.expires_in - expires) / 60);
                int sec = (int)((accessToken.expires_in - expires) % 60);
                Console.WriteLine("token refresh: {0} min {1} sec", min, sec);
            }
        }
        
        return accessToken;
    }
    
    private AuthAccessToken createAccessToken()
    {
        Console.WriteLine("request permissions");
        
        // request user permissions
        AuthDeviceCode deviceCode = requestPermissions();
        Console.WriteLine("user code: {0}", deviceCode.user_code);
        
        var client = new HttpClient() 
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
            Thread.Sleep(deviceCode.interval * 1000);
            
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
            access_token = LoadAccessToken().access_token;
        
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
        var client = new HttpClient() 
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
        
        var client = new HttpClient() 
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
        using (var stream = new StreamWriter(tokenFile)) 
        {
            stream.WriteLine(access_token.expires_in);
            stream.WriteLine(access_token.access_token);
            stream.WriteLine(access_token.refresh_token);
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