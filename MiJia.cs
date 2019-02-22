using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace MiJia
{
    public class DeviceInfo
    {
        public string Cmd { get; set; }
        [JsonProperty("mac")]
        public string MAC { get; set; }
        [JsonProperty("password")]
        public string Password { get; set; }
        [JsonProperty("model")]
        public string Model { get; set; }
        public string Sid { get; set; }
        [JsonProperty("short_id")]
        public string ShortId { get; set; }
        [JsonProperty("did")]
        public string DeviceId { get; set; }
        public string Token { get; set; }
        public string Data { get; set; }
        //[JsonProperty("devices")]
        //public AqaraDeviceConfig[] Devices { get; set; }
        //[JsonProperty("gateways")]
        //public AqaraGatewayConfig[] Gateways { get; set; }
    }

    class Gateway
    {
        public string MulticastIP { get; set; } = "224.0.0.50";
        public int MulticastPort { get; set; } = 4321;
        public string RemoteIP { get; set; } = "192.168.1.244";
        public int RemotePort { get; set; } = 9898;
        public string Token { get; set; }

        internal int SendCmd(string cmd)
        {
            IPAddress mip = IPAddress.Parse(MulticastIP);
            IPEndPoint mep = new IPEndPoint(mip, MulticastPort);

            IPAddress rip = IPAddress.Parse(RemoteIP);
            IPEndPoint rep = new IPEndPoint(rip, RemotePort);

            UdpClient udp = new UdpClient();
            udp.JoinMulticastGroup(mip);
            udp.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            udp.ExclusiveAddressUse = false;
            udp.Connect("127.0.0.1", RemotePort);

            byte[] bs = Encoding.Default.GetBytes(cmd);
            return udp.Send(bs, bs.Length);
        }

        internal async Task<string> SendCmd(string cmd, string server, int port)
        {
            string result = string.Empty;

            UdpClient udp = new UdpClient();
            udp.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            udp.ExclusiveAddressUse = false;
            udp.Connect(server, port);

            byte[] bs = Encoding.Default.GetBytes(cmd);
            var ret = udp.Send(bs, bs.Length);
            var received = await udp.ReceiveAsync();
            result = Encoding.Default.GetString(received.Buffer);

            return (result);
        }

        public async Task<List<string>> Listen()
        {
            List<string> result = new List<string>();

            IPAddress mip = IPAddress.Parse(MulticastIP);
            IPEndPoint mep = new IPEndPoint(mip, MulticastPort);

            IPAddress rip = IPAddress.Parse(RemoteIP);
            IPEndPoint rep = new IPEndPoint(rip, RemotePort);

            UdpClient udp = new UdpClient();
            udp.JoinMulticastGroup(mip);
            udp.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            udp.ExclusiveAddressUse = false;

            udp.Client.Bind(new IPEndPoint(IPAddress.Any, RemotePort));

            while (true)
            {
                var received = await udp.ReceiveAsync();
                var line = Encoding.Default.GetString(received.Buffer);
                if (!string.IsNullOrEmpty(line))
                {
                    result.Add(line);
                    break;
                }
            }

            return (result);
        }   

        public async Task<IList<string>> GetDevice()
        {
            List<string> result = new List<string>();

            //var ret = SendCmd("{\"cmd\":\"whois\"}");
            var ret = await SendCmd("{\"cmd\" : \"get_id_list\"}", RemoteIP, RemotePort);

            JsonSerializer serializer = new JsonSerializer();
            DeviceInfo token = JsonConvert.DeserializeObject<DeviceInfo>(ret);
            if (!string.IsNullOrEmpty(token.Data))
            {
                var data = JsonConvert.DeserializeObject<string[]>(token.Data);
                foreach (var device in data)
                    result.Add(device);
            }

            return (result);
        }

        public async Task<IList<string>> GetDevice(string deviceId)
        {
            List<string> result = new List<string>();

            if (!string.IsNullOrEmpty(deviceId))
            {
                var ret = await SendCmd($"{{\"cmd\":\"read\", \"sid\":\"{deviceId}\"}}", RemoteIP, RemotePort);

                JsonSerializer serializer = new JsonSerializer();
                DeviceInfo token = JsonConvert.DeserializeObject<DeviceInfo>(ret);
                if (!string.IsNullOrEmpty(token.Data))
                {
                    JToken data = JsonConvert.DeserializeObject<JToken>(token.Data);
                    if (data["status"] != null)
                    {
                        var status = data["status"];
                    }
                }
            }

            return (result);
        }
    }
}
